import json
import os
import time
import threading

import win32timezone
import win32serviceutil
import win32service
import win32event
import servicemanager

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

from processor import process_zip
from logger import log


CONFIG_FILE = "config.json"


def load_config():
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
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


class AsztalosoftService(win32serviceutil.ServiceFramework):

    _svc_name_ = "AsztalosoftProcessor"
    _svc_display_name_ = "Asztalosoft ZIP Processor"
    _svc_description_ = "Processes incoming ZIP files automatically"

    def __init__(self, args):

        win32serviceutil.ServiceFramework.__init__(self, args)

        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.running = True

        self.observer = None
        self.config = None
        self.config_mtime = None

        self.worker = None

    def SvcStop(self):

        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)

        log("Service stopping")

        self.running = False

        if self.observer:
            self.observer.stop()
            self.observer.join()

        win32event.SetEvent(self.stop_event)

    def SvcDoRun(self):

        servicemanager.LogInfoMsg("Asztalosoft service starting")

        self.ReportServiceStatus(win32service.SERVICE_RUNNING)

        self.worker = threading.Thread(target=self.main)
        self.worker.daemon = True
        self.worker.start()

        win32event.WaitForSingleObject(self.stop_event, win32event.INFINITE)

    def start_observer(self):

        if self.observer:
            self.observer.stop()
            self.observer.join()

        watch_folder = self.config["watch_folder"]

        log(f"Watching folder: {watch_folder}")

        event_handler = ZipHandler(self.config)

        self.observer = Observer()
        self.observer.schedule(event_handler, watch_folder, recursive=False)
        self.observer.start()

    def main(self):

        try:

            self.config = load_config()
            self.config_mtime = os.path.getmtime(CONFIG_FILE)

            self.start_observer()

            while self.running:

                try:

                    current_mtime = os.path.getmtime(CONFIG_FILE)

                    if current_mtime != self.config_mtime:

                        log("Config changed, reloading")

                        self.config = load_config()
                        self.config_mtime = current_mtime

                        self.start_observer()

                except Exception as e:
                    log(f"Config reload error: {e}")

                time.sleep(5)

        except Exception as e:
            log(f"Fatal service error: {e}")


if __name__ == "__main__":
    win32serviceutil.HandleCommandLine(AsztalosoftService)
