import json
import os
import time
import threading
import win32timezone
import win32serviceutil
import win32service
import win32event
import servicemanager
import sys

from watchdog.observers.polling import PollingObserver as Observer
from watchdog.events import FileSystemEventHandler

from processor import process_zip
from logger import log


CONFIG_FILE = "config.json"

def get_base_dir():
    if getattr(sys, "frozen", False):
        # PyInstaller exe
        return os.path.dirname(sys.executable)
    else:
        # sima python futtatás
        return os.path.dirname(os.path.abspath(__file__))

def load_config():

    base = os.path.dirname(sys.executable)
    path = os.path.join(base, CONFIG_FILE)

    if not os.path.exists(path):
        raise Exception(f"Config file not found: {path}")

    with open(path, encoding="utf8") as f:
        return json.load(f)
        return json.load(f)


class ZipHandler(FileSystemEventHandler):

    def __init__(self, config):
        self.config = config
        self.prefix = config.get("zip_prefix", "work-")

    def on_created(self, event):

        if event.is_directory:
            return

        path = event.src_path
        name = os.path.basename(path)

        if name.startswith(self.prefix) and name.endswith(".zip"):

            log(f"New ZIP detected: {path}")

            try:
                process_zip(path, self.config)

            except Exception as e:

                log(f"Processing error: {e}")


def run_watcher(stop_event):
    log(f"cwd: {os.getcwd()}")
    config = load_config()

    watch_folder = config["watch_folder"]

    log("Service started")
    log(f"Watching folder: {watch_folder}")

    event_handler = ZipHandler(config)

    observer = Observer()
    observer.schedule(event_handler, watch_folder, recursive=False)

    observer.start()

    try:

        while not stop_event.is_set():

            time.sleep(1)

    finally:

        observer.stop()
        observer.join()

        log("Service stopped")


class AsztalosoftService(win32serviceutil.ServiceFramework):

    _svc_name_ = "asztalosoft_processer"
    _svc_display_name_ = "Asztalosoft File Processor"
    _svc_description_ = "ZIP fájlokat figyel és feldolgozza a lista és program tartalmát."

    def __init__(self, args):

        win32serviceutil.ServiceFramework.__init__(self, args)

        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.thread = None
        self.running = True

    def SvcStop(self):

        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)

        log("Service stop requested")

        self.running = False
        win32event.SetEvent(self.stop_event)

    def SvcDoRun(self):

        servicemanager.LogInfoMsg("Asztalosoft service starting")

        # azonnal jelezzük hogy fut
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)

        base = os.path.dirname(sys.executable)
        os.chdir(base)

        stop_flag = threading.Event()

        def run():
            try:
                run_watcher(stop_flag)
            except Exception as e:
                log(f"SERVICE CRASH: {e}")

        self.thread = threading.Thread(target=run, daemon=True)
        self.thread.start()

        win32event.WaitForSingleObject(self.stop_event, win32event.INFINITE)

        stop_flag.set()

        if self.thread:
            self.thread.join()

        servicemanager.LogInfoMsg("Asztalosoft service stopped")


if __name__ == "__main__":

    win32serviceutil.HandleCommandLine(AsztalosoftService)