import sys
import os
import json
import time
import shutil
import zipfile
import threading
import subprocess
import ctypes
import traceback

# GUI
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit,
    QPushButton, QFileDialog, QGridLayout,
    QVBoxLayout, QGroupBox, QMessageBox
)
from PySide6.QtGui import QIcon

# Service
import win32serviceutil
import win32service
import win32event
import servicemanager
import win32timezone

# Watchdog
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler


# -------------------------------------------------
# PATHS
# -------------------------------------------------

BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
LOG_FILE = os.path.join(BASE_DIR, "log.txt")
ICON_FILE = os.path.join(BASE_DIR, "logo.ico")

SERVICE_NAME = "asztalosoft_processer"


# -------------------------------------------------
# DEFAULT CONFIG
# -------------------------------------------------

DEFAULT_CONFIG = {
    "watch_folder": "",
    "target_list": "",
    "target_program": "",
    "archive": "",
    "zip_prefix": "work-"
}


# -------------------------------------------------
# FILE INITIALIZATION
# -------------------------------------------------

def ensure_files():

    if not os.path.exists(LOG_FILE):
        open(LOG_FILE, "w", encoding="utf8").close()

    if not os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "w", encoding="utf8") as f:
            json.dump(DEFAULT_CONFIG, f, indent=2)


# -------------------------------------------------
# LOGGER
# -------------------------------------------------

def log(text):

    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {text}"

    print(line)

    try:
        with open(LOG_FILE, "a", encoding="utf8") as f:
            f.write(line + "\n")
    except:
        pass


def log_exception(prefix="ERROR"):
    log(prefix)
    log(traceback.format_exc())


# -------------------------------------------------
# CONFIG
# -------------------------------------------------

def load_config():

    ensure_files()

    log(f"Loading config from: {CONFIG_FILE}")

    try:
        with open(CONFIG_FILE, encoding="utf8") as f:
            data = json.load(f)

        # hiányzó kulcsok pótlása
        changed = False
        for k, v in DEFAULT_CONFIG.items():
            if k not in data:
                data[k] = v
                changed = True

        if changed:
            save_config(data)

        log(f"Config loaded: {data}")
        return data

    except:
        log_exception("Failed to load config")
        return DEFAULT_CONFIG.copy()


def save_config(data):

    log(f"Saving config: {data}")

    with open(CONFIG_FILE, "w", encoding="utf8") as f:
        json.dump(data, f, indent=2)


# -------------------------------------------------
# ZIP PROCESSOR
# -------------------------------------------------

def wait_for_download(path, retries=10):

    last_size = -1

    for _ in range(retries):

        if not os.path.exists(path):
            return False

        size = os.path.getsize(path)

        if size == last_size:
            return True

        last_size = size
        time.sleep(1)

    return False


def process_zip(path, config):

    try:

        log(f"Processing ZIP: {path}")

        if not wait_for_download(path):
            log("File still downloading")
            return

        temp_folder = path + "_tmp"

        if os.path.exists(temp_folder):
            shutil.rmtree(temp_folder)

        os.makedirs(temp_folder)

        with zipfile.ZipFile(path, 'r') as zip_ref:
            zip_ref.extractall(temp_folder)

        folders = [
            f for f in os.listdir(temp_folder)
            if os.path.isdir(os.path.join(temp_folder, f))
        ]

        if len(folders) == 0:
            log("ZIP is empty")
            return

        root_folder = os.path.join(temp_folder, folders[0])

        subfolders = [
            f for f in os.listdir(root_folder)
            if os.path.isdir(os.path.join(root_folder, f))
        ]

        folder_list = None
        folder_program = None

        for f in subfolders:

            name = f.lower()

            if "lista" in name:
                folder_list = os.path.join(root_folder, f)

            if "program" in name:
                folder_program = os.path.join(root_folder, f)

        if not folder_list or not folder_program:
            log("Missing required folders: program / lista")
            return

        target_list = config["target_list"]
        target_program = config["target_program"]

        os.makedirs(target_list, exist_ok=True)
        os.makedirs(target_program, exist_ok=True)

        for file in os.listdir(folder_list):
            shutil.move(os.path.join(folder_list, file),
                        os.path.join(target_list, file))

        for file in os.listdir(folder_program):
            shutil.move(os.path.join(folder_program, file),
                        os.path.join(target_program, file))

        archive = config["archive"]
        os.makedirs(archive, exist_ok=True)

        archive_path = os.path.join(archive, os.path.basename(path))

        shutil.copy2(path, archive_path)

        os.remove(path)

        shutil.rmtree(temp_folder)

        log("ZIP processed successfully")

    except:
        log_exception("ZIP processing failed")


# -------------------------------------------------
# WATCHDOG
# -------------------------------------------------

class ZipHandler(FileSystemEventHandler):

    def __init__(self, config):
        self.config = config
        self.prefix = config.get("zip_prefix", "work-")

    def handle_zip(self, path):

        name = os.path.basename(path)

        if name.startswith(self.prefix) and name.endswith(".zip"):
            process_zip(path, self.config)

    def on_created(self, event):

        if event.is_directory:
            return

        self.handle_zip(event.src_path)

    def on_moved(self, event):

        if event.is_directory:
            return

        self.handle_zip(event.dest_path)


# -------------------------------------------------
# SERVICE
# -------------------------------------------------

class AsztalosoftService(win32serviceutil.ServiceFramework):

    _svc_name_ = SERVICE_NAME
    _svc_display_name_ = "Asztalosoft ZIP Processor"
    _svc_description_ = "Processes incoming ZIP files automatically"

    def __init__(self, args):

        win32serviceutil.ServiceFramework.__init__(self, args)

        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.running = True
        self.observer = None
        self.worker = None

    def SvcStop(self):

        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)

        self.running = False

        if self.observer:
            self.observer.stop()
            self.observer.join()

        win32event.SetEvent(self.stop_event)

    def SvcDoRun(self):

        self.ReportServiceStatus(win32service.SERVICE_START_PENDING)

        self.worker = threading.Thread(target=self.main)
        self.worker.start()

        self.ReportServiceStatus(win32service.SERVICE_RUNNING)

        win32event.WaitForSingleObject(self.stop_event, win32event.INFINITE)

    def main(self):

        config = load_config()

        watch_folder = config.get("watch_folder")

        if not watch_folder:
            while self.running:
                time.sleep(5)
            return

        handler = ZipHandler(config)

        self.observer = Observer()
        self.observer.schedule(handler, watch_folder, recursive=False)
        self.observer.start()

        while self.running:
            time.sleep(5)

        self.observer.stop()
        self.observer.join()


# -------------------------------------------------
# ADMIN CHECK
# -------------------------------------------------

def ensure_admin():

    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
    except:
        is_admin = False

    if not is_admin:

        params = " ".join(sys.argv)

        ctypes.windll.shell32.ShellExecuteW(
            None,
            "runas",
            sys.executable,
            params,
            None,
            1
        )

        sys.exit()


# -------------------------------------------------
# GUI
# -------------------------------------------------

class App(QWidget):

    def __init__(self):

        super().__init__()

        self.setWindowTitle("Asztalosoft - Fájlkezelő")
        self.setMinimumWidth(700)

        if os.path.exists(ICON_FILE):
            self.setWindowIcon(QIcon(ICON_FILE))

        config = load_config()

        layout = QVBoxLayout()

        box = QGroupBox("Mappa beállítások")
        grid = QGridLayout()

        self.watch = QLineEdit(config.get("watch_folder", ""))
        self.target_list = QLineEdit(config.get("target_list", ""))
        self.target_program = QLineEdit(config.get("target_program", ""))
        self.archive = QLineEdit(config.get("archive", ""))

        rows = [
            ("Forrás mappa", self.watch),
            ("Cél mappa - Lista", self.target_list),
            ("Cél mappa - Program", self.target_program),
            ("Archívum mappa", self.archive)
        ]

        for i, (text, field) in enumerate(rows):

            label = QLabel(text)
            button = QPushButton("Tallózás")

            button.clicked.connect(lambda _, f=field: self.browse(f))

            grid.addWidget(label, i, 0)
            grid.addWidget(field, i, 1)
            grid.addWidget(button, i, 2)

        box.setLayout(grid)
        layout.addWidget(box)

        self.status = QLabel()
        layout.addWidget(self.status)

        buttons = QGridLayout()

        save_btn = QPushButton("Mentés")
        start_btn = QPushButton("Indítás")
        stop_btn = QPushButton("Megállítás")
        log_btn = QPushButton("Log")

        save_btn.clicked.connect(self.save)
        start_btn.clicked.connect(self.start_service)
        stop_btn.clicked.connect(self.stop_service)
        log_btn.clicked.connect(self.open_log)

        buttons.addWidget(save_btn, 0, 0)
        buttons.addWidget(start_btn, 0, 1)
        buttons.addWidget(stop_btn, 0, 2)
        buttons.addWidget(log_btn, 0, 3)

        layout.addLayout(buttons)

        self.setLayout(layout)

        self.update_status()

    def browse(self, field):

        folder = QFileDialog.getExistingDirectory(self)

        if folder:
            field.setText(folder)

    def save(self):

        data = {
            "watch_folder": self.watch.text(),
            "target_list": self.target_list.text(),
            "target_program": self.target_program.text(),
            "archive": self.archive.text(),
            "zip_prefix": "work-"
        }

        save_config(data)

        QMessageBox.information(self, "Info", "Beállítások elmentve")

    def service_cmd(self, cmd):

        exe = sys.argv[0]

        subprocess.run(f'"{exe}" service {cmd}', shell=True)

        self.update_status()

    def start_service(self):

        exe = sys.argv[0]

        r = subprocess.run(
            f"sc query {SERVICE_NAME}",
            shell=True,
            capture_output=True,
            text=True
        )

        if r.returncode != 0:

            # install
            subprocess.run(
                f'"{exe}" service install',
                shell=True
            )

            # autostart beállítás
            subprocess.run(
                f'sc config {SERVICE_NAME} start= auto',
                shell=True
            )

        # start
        subprocess.run(
            f'"{exe}" service start',
            shell=True
        )

        self.update_status()

    def stop_service(self):

        self.service_cmd("stop")

    def update_status(self):

        r = subprocess.run(
            f"sc query {SERVICE_NAME}",
            shell=True,
            capture_output=True,
            text=True
        )

        if "RUNNING" in r.stdout:
            self.status.setText("Service: FUT")
        else:
            self.status.setText("Service: NEM FUT")

    def open_log(self):

        if os.path.exists(LOG_FILE):
            os.startfile(LOG_FILE)


# -------------------------------------------------
# MAIN
# -------------------------------------------------

def run_gui():

    app = QApplication(sys.argv)

    if os.path.exists(ICON_FILE):
        app.setWindowIcon(QIcon(ICON_FILE))

    w = App()
    w.show()

    sys.exit(app.exec())


def main():

    ensure_files()

    # ha CLI service parancs
    if len(sys.argv) > 1 and sys.argv[1] == "service":

        sys.argv.pop(1)
        win32serviceutil.HandleCommandLine(AsztalosoftService)
        return

    # ha a Windows service manager indította
    if len(sys.argv) == 1:

        try:
            servicemanager.Initialize()
            servicemanager.PrepareToHostSingle(AsztalosoftService)
            servicemanager.StartServiceCtrlDispatcher()
            return
        except Exception:
            pass

    # különben GUI
    ensure_admin()
    run_gui()


if __name__ == "__main__":
    main()