"""Folder processing inside the logged-in user's desktop session."""
import os
import queue
import tempfile
import threading

from PySide6.QtCore import QThread, Signal
from watchdog.events import FileSystemEventHandler
from watchdog.observers.polling import PollingObserver


class QueueHandler(FileSystemEventHandler):
    def __init__(self, enqueue):
        self.enqueue = enqueue

    def on_created(self, event):
        if not event.is_directory:
            self.enqueue(event.src_path)

    def on_moved(self, event):
        if not event.is_directory:
            self.enqueue(event.dest_path)


class Processor(QThread):
    ready = Signal()
    failed = Signal(str)
    file_error = Signal(str)

    def __init__(self, config, process_zip):
        super().__init__()
        self.config = dict(config)
        self.process_zip = process_zip
        self.stopping = threading.Event()
        self.pending = queue.Queue()
        self.seen = set()
        self.lock = threading.Lock()

    def stop(self):
        self.stopping.set()

    def enqueue(self, path):
        name = os.path.basename(path)
        if not name.startswith(self.config.get("zip_prefix", "work-")) or not name.lower().endswith(".zip"):
            return
        try:
            stat = os.stat(path)
        except OSError:
            return
        identity = (os.path.normcase(path), stat.st_size, stat.st_mtime_ns)
        with self.lock:
            if identity in self.seen:
                return
            self.seen.add(identity)
        self.pending.put(path)

    def run(self):
        observer = None
        try:
            required_paths = ("watch_folder", "target_list", "target_program", "archive")
            optional_paths = ("target_pdf",) if (self.config.get("target_pdf") or "").strip() else ()
            for key in (*required_paths, *optional_paths):
                path = self.config.get(key, "")
                if not path or not os.path.isabs(path):
                    raise ValueError("Minden mappához teljes elérési utat adj meg.")
                if key == "watch_folder":
                    if not os.path.isdir(path):
                        raise ValueError(f"A forrásmappa nem érhető el: {path}")
                else:
                    os.makedirs(path, exist_ok=True)
                # Test actual access as this user, rather than trusting ACL metadata.
                with tempfile.TemporaryFile(dir=path) as probe:
                    probe.write(b"access check")
            if self.stopping.is_set():
                return
            observer = PollingObserver(timeout=1)
            observer.schedule(QueueHandler(self.enqueue), self.config["watch_folder"], recursive=False)
            observer.start()
            for entry in os.scandir(self.config["watch_folder"]):
                if entry.is_file():
                    self.enqueue(entry.path)
            self.ready.emit()
            while not self.stopping.is_set():
                try:
                    path = self.pending.get(timeout=0.25)
                except queue.Empty:
                    if not os.path.isdir(self.config["watch_folder"]):
                        raise OSError("A forrásmappa kapcsolata megszakadt. Ellenőrizd az elérést, majd indítsd újra a figyelést.")
                    continue
                if os.path.isfile(path):
                    try:
                        self.process_zip(path, self.config)
                    except Exception as error:
                        self.file_error.emit(f"{os.path.basename(path)}: {error}")
        except Exception as error:
            self.failed.emit(str(error))
        finally:
            if observer is not None and observer.is_alive():
                observer.stop()
                observer.join()
