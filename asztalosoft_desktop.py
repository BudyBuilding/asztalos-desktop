import sys
import os
import json
import time
import shutil
import zipfile
import threading
import subprocess
import ctypes

# GUI
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit,
    QPushButton, QFileDialog, QGridLayout,
    QVBoxLayout, QGroupBox, QMessageBox
)

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

BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
os.chdir(BASE_DIR)

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
LOG_FILE = os.path.join(BASE_DIR, "log.txt")

SERVICE_NAME = "asztalosoft_processer"


# -------------------------------------------------
# LOGGER
# -------------------------------------------------

def log(text):

    ts = time.strftime("%Y-%m-%d %H:%M:%S")

    line = f"[{ts}] {text}"

    print(line)

    with open(LOG_FILE, "a", encoding="utf8") as f:
        f.write(line + "\n")


# -------------------------------------------------
# CONFIG
# -------------------------------------------------

def load_config():

    if not os.path.exists(CONFIG_FILE):
        return {}

    with open(CONFIG_FILE, encoding="utf8") as f:
        return json.load(f)


def save_config(data):

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

        src = os.path.join(folder_list, file)
        dst = os.path.join(target_list, file)

        shutil.move(src, dst)

    for file in os.listdir(folder_program):

        src = os.path.join(folder_program, file)
        dst = os.path.join(target_program, file)

        shutil.move(src, dst)

    archive = config["archive"]
    os.makedirs(archive, exist_ok=True)

    archive_path = os.path.join(archive, os.path.basename(path))

    shutil.copy2(path, archive_path)

    log(f"ZIP archived: {archive_path}")

    os.remove(path)

    log(f"Original ZIP deleted: {path}")

    shutil.rmtree(temp_folder)

    log(f"ZIP processed successfully: {path}")


# -------------------------------------------------
# WATCHDOG
# -------------------------------------------------

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


# -------------------------------------------------
# SERVICE
# -------------------------------------------------

class AsztalosoftService(win32serviceutil.ServiceFramework):

    _svc_name_ = "asztalosoft_processer"
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

        log("Service stopping")

        self.running = False

        if self.observer:
            self.observer.stop()
            self.observer.join()

        win32event.SetEvent(self.stop_event)

    def SvcDoRun(self):

        servicemanager.LogInfoMsg("Service starting")

        # jelzés hogy indul
        self.ReportServiceStatus(win32service.SERVICE_START_PENDING)

        try:
            self.worker = threading.Thread(target=self.main)
            self.worker.daemon = True
            self.worker.start()

            # most már fut
            self.ReportServiceStatus(win32service.SERVICE_RUNNING)

            win32event.WaitForSingleObject(self.stop_event, win32event.INFINITE)

        except Exception as e:
            log(f"Service failed: {e}")

    def main(self):

        try:

            config = load_config()
            watch_folder = config.get("watch_folder")

            if not watch_folder:
                log("watch_folder missing in config")

                while self.running:
                    time.sleep(5)

                return

            log(f"Watching folder: {watch_folder}")

            handler = ZipHandler(config)

            self.observer = Observer()
            self.observer.schedule(handler, watch_folder, recursive=False)
            self.observer.start()

            while self.running:
                time.sleep(5)

        except Exception as e:

            log(f"Service error: {e}")


# -------------------------------------------------
# ADMIN CHECK
# -------------------------------------------------

def is_admin():

    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

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

        # ellenőrizzük létezik-e
        r = subprocess.run(
            f"sc query {SERVICE_NAME}",
            shell=True,
            capture_output=True,
            text=True
        )

        if r.returncode != 0:
            # service nincs telepítve
            subprocess.run(f'"{exe}" service install', shell=True)

        # indítás
        subprocess.run(f'"{exe}" service start', shell=True)

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

        out = r.stdout

        if "RUNNING" in out:
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

    w = App()
    w.show()

    sys.exit(app.exec())


def main():

    ensure_admin()

    if len(sys.argv) > 1 and sys.argv[1] == "service":

        sys.argv[0] = sys.executable
        sys.argv.pop(1)

        win32serviceutil.HandleCommandLine(AsztalosoftService)

    else:

        run_gui()


if __name__ == "__main__":
    main()