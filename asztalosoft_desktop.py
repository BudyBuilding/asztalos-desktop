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
import win32com.client
import win32serviceutil
import win32service
import win32event
import servicemanager
import win32timezone
import win32api
import win32process
from win32com.shell import shell, shellcon
import win32wnet
import pywintypes

# Watchdog
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler


# -------------------------------------------------
# PATHS
# -------------------------------------------------

BASE_DIR = os.path.dirname(
    os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__)
)

APP_DATA_DIR = os.path.join(
    os.environ.get("APPDATA", os.path.expanduser("~")),
    "Asztalosoft",
)
CONFIG_FILE = os.path.join(APP_DATA_DIR, "config.json")
LOG_FILE = os.path.join(APP_DATA_DIR, "log.txt")
ICON_FILE = os.path.join(BASE_DIR, "logo.ico")

SERVICE_NAME = "asztalosoft_processer"


# -------------------------------------------------
# DEFAULT CONFIG
# -------------------------------------------------

DEFAULT_CONFIG = {
    "watch_folder": "",
    "target_list": "",
    "target_program": "",
    "target_pdf": "",
    "archive": "",
    "zip_prefix": "work-",
    "autostart": False,
}


# -------------------------------------------------
# FILE INITIALIZATION
# -------------------------------------------------

def ensure_files():

    os.makedirs(APP_DATA_DIR, exist_ok=True)

    if not os.path.exists(LOG_FILE):
        open(LOG_FILE, "w", encoding="utf8").close()

    if not os.path.exists(CONFIG_FILE):
        data = DEFAULT_CONFIG.copy()
        legacy_config = os.path.join(BASE_DIR, "config.json")
        if os.path.abspath(legacy_config) != os.path.abspath(CONFIG_FILE) and os.path.isfile(legacy_config):
            try:
                with open(legacy_config, encoding="utf8") as f:
                    legacy_data = json.load(f)
                if isinstance(legacy_data, dict):
                    data.update(legacy_data)
                    log("Migrated configuration from the application folder to user AppData.")
                else:
                    log("Legacy config migration skipped: config.json does not contain an object.")
            except (OSError, ValueError) as error:
                log(f"Legacy config migration failed: {error}")
        with open(CONFIG_FILE, "w", encoding="utf8") as f:
            json.dump(data, f, indent=2)


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

    os.makedirs(APP_DATA_DIR, exist_ok=True)
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
        folder_pdf = None

        for f in subfolders:

            name = f.lower()

            if "lista" in name:
                folder_list = os.path.join(root_folder, f)

            if "program" in name:
                folder_program = os.path.join(root_folder, f)

            if "pdf" in name:
                folder_pdf = os.path.join(root_folder, f)

        if not folder_list or not folder_program:
            log("Missing required folders: program / lista")
            return

        target_pdf = (config.get("target_pdf") or "").strip()
        if folder_pdf and not target_pdf:
            raise ValueError("A ZIP PDF mappájához nincs célmappa beállítva.")

        target_list = config["target_list"]
        target_program = config["target_program"]

        os.makedirs(target_list, exist_ok=True)
        os.makedirs(target_program, exist_ok=True)
        if folder_pdf:
            os.makedirs(target_pdf, exist_ok=True)

        for file in os.listdir(folder_list):
            shutil.move(os.path.join(folder_list, file),
                        os.path.join(target_list, file))

        for file in os.listdir(folder_program):
            shutil.move(os.path.join(folder_program, file),
                        os.path.join(target_program, file))

        if folder_pdf:
            for file in os.listdir(folder_pdf):
                shutil.move(os.path.join(folder_pdf, file),
                            os.path.join(target_pdf, file))

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

def normalize_folder_path(path):
    """Resolve mapped drives to UNC paths, preserving local folders."""
    path = path.strip().strip('"')
    if not path:
        return ""
    path = os.path.normpath(path)
    try:
        return win32wnet.WNetGetUniversalName(path)
    except pywintypes.error:
        return path


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
        self.target_pdf = QLineEdit(config.get("target_pdf", ""))
        self.archive = QLineEdit(config.get("archive", ""))

        rows = [
            ("Forrás mappa", self.watch),
            ("Cél mappa - Lista", self.target_list),
            ("Cél mappa - Program", self.target_program),
            ("Cél mappa - PDF", self.target_pdf),
            ("Archívum mappa", self.archive)
        ]

        for i, (text, field) in enumerate(rows):

            field.setPlaceholderText(r"C:\mappa vagy \\szerver\megosztás\mappa")
            field.setToolTip("Helyi vagy hálózati mappa. A hálózati útvonal közvetlenül is beírható.")
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

        account_btn = QPushButton("Hálózati hozzáférés / szolgáltatásfiók")
        account_btn.clicked.connect(lambda: self.service_cmd("account"))
        layout.addWidget(account_btn)

        self.setLayout(layout)

        self.update_status()

    def browse(self, field):
        try:
            # Native Explorer picker supports Network and UNC address entry.
            selected_path = QFileDialog.getExistingDirectory(
                self,
                "Válassz mappát",
                field.text().strip().strip('"'),
                QFileDialog.Option.ShowDirsOnly
            )

            if selected_path:
                field.setText(normalize_folder_path(selected_path))

        except Exception as e:
            log(f"Mappaválasztó hiba: {e}")
            QMessageBox.critical(
                self,
                "Hiba",
                f"Nem sikerült megnyitni a mappaválasztót:\n{e}"
            )

    def save(self):

        data = {
            "watch_folder": normalize_folder_path(self.watch.text()),
            "target_list": normalize_folder_path(self.target_list.text()),
            "target_program": normalize_folder_path(self.target_program.text()),
            "target_pdf": normalize_folder_path(self.target_pdf.text()),
            "archive": normalize_folder_path(self.archive.text()),
            "zip_prefix": "work-"
        }

        save_config(data)

        QMessageBox.information(self, "Info", "Beállítások elmentve")

    def service_cmd(self, cmd):
        # Keep Explorer and the picker in the user's normal logon session.
        args = [] if getattr(sys, "frozen", False) else [os.path.abspath(__file__)]
        args.extend(["--manage-service", cmd])
        try:
            process = shell.ShellExecuteEx(
                fMask=shellcon.SEE_MASK_NOCLOSEPROCESS,
                hwnd=int(self.winId()),
                lpVerb="runas",
                lpFile=sys.executable,
                lpParameters=subprocess.list2cmdline(args),
                nShow=1 if cmd == "account" else 0,
            )["hProcess"]
            try:
                win32event.WaitForSingleObject(process, win32event.INFINITE)
                if win32process.GetExitCodeProcess(process) != 0:
                    details = ""
                    try:
                        with open(LOG_FILE, encoding="utf8") as f:
                            details = "".join(f.readlines()[-15:])
                    except OSError:
                        pass
                    raise RuntimeError("A szolgáltatás művelete sikertelen.\n" + details)
            finally:
                win32api.CloseHandle(process)
        except Exception as e:
            QMessageBox.warning(self, "Szolgáltatás", str(e))
        self.update_status()

    def start_service(self):
        self.service_cmd("start")

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


def manage_service(command):
    """Run only service mutations in the elevated helper process."""
    if command not in ("start", "stop", "register"):
        raise ValueError("Ismeretlen szolgáltatásművelet")
    def run(args):
        result = subprocess.run(args, capture_output=True, text=True)
        if result.returncode:
            raise RuntimeError(f"{subprocess.list2cmdline(args)}\n{result.stdout}{result.stderr}")

    if command in ("start", "register"):
        if not getattr(sys, "frozen", False):
            raise RuntimeError("A szolgáltatás indításához a buildelt EXE-t használd.")
        result = subprocess.run(
            ["sc.exe", "query", SERVICE_NAME], capture_output=True, text=True
        )
        binary_path = subprocess.list2cmdline([sys.executable])
        if result.returncode == 1060:
            run(["sc.exe", "create", SERVICE_NAME, "binPath=", binary_path,
                 "start=", "auto", "DisplayName=", AsztalosoftService._svc_display_name_])
        elif result.returncode:
            raise RuntimeError(result.stdout + result.stderr)
        else:
            # Repair stale registrations left by earlier executable names/builds.
            run(["sc.exe", "config", SERVICE_NAME, "binPath=", binary_path,
                 "start=", "auto"])

    if command == "register":
        return
    desired = win32service.SERVICE_RUNNING if command == "start" else win32service.SERVICE_STOPPED
    if win32serviceutil.QueryServiceStatus(SERVICE_NAME)[1] == desired:
        return
    run(["sc.exe", command, SERVICE_NAME])
    deadline = time.monotonic() + 30
    while time.monotonic() < deadline:
        status = win32serviceutil.QueryServiceStatus(SERVICE_NAME)
        if status[1] == desired:
            log(f"Service {command}: OK")
            return
        if command == "start" and status[1] == win32service.SERVICE_STOPPED:
            raise RuntimeError(f"A szolgáltatás leállt. Windows hibakód: {status[3]}, szolgáltatás hibakód: {status[4]}")
        time.sleep(0.25)
    raise RuntimeError("A szolgáltatás 30 másodpercen belül nem érte el a kért állapotot.")


def retire_legacy_service():
    """Disable and stop the former service so it cannot process files in parallel."""
    manager = win32service.OpenSCManager(None, None, win32service.SC_MANAGER_CONNECT)
    try:
        try:
            service = win32service.OpenService(
                manager,
                SERVICE_NAME,
                win32service.SERVICE_CHANGE_CONFIG
                | win32service.SERVICE_QUERY_STATUS
                | win32service.SERVICE_STOP,
            )
        except win32service.error as error:
            if getattr(error, "winerror", None) == 1060:
                return
            raise
        try:
            win32service.ChangeServiceConfig(
                service,
                win32service.SERVICE_NO_CHANGE,
                win32service.SERVICE_DISABLED,
                win32service.SERVICE_NO_CHANGE,
                None,
                None,
                0,
                None,
                None,
                None,
                None,
            )
            if win32service.QueryServiceStatus(service)[1] != win32service.SERVICE_STOPPED:
                win32service.ControlService(service, win32service.SERVICE_CONTROL_STOP)
                deadline = time.monotonic() + 30
                while time.monotonic() < deadline:
                    if win32service.QueryServiceStatus(service)[1] == win32service.SERVICE_STOPPED:
                        log("Legacy service disabled and stopped")
                        return
                    time.sleep(0.25)
                raise RuntimeError("A régi szolgáltatás 30 másodpercen belül nem állt le.")
            log("Legacy service disabled; it was already stopped")
        finally:
            win32service.CloseServiceHandle(service)
    finally:
        win32service.CloseServiceHandle(manager)


def main():

    ensure_files()

    if len(sys.argv) == 2 and sys.argv[1] == "--retire-service":
        try:
            retire_legacy_service()
        except Exception:
            log(traceback.format_exc())
            sys.exit(1)
        return

    if len(sys.argv) == 3 and sys.argv[1] == "--manage-service":
        try:
            if sys.argv[2] == "account":
                from service_account import configure_account
                configure_account(SERVICE_NAME, manage_service)
                return
            manage_service(sys.argv[2])
        except Exception:
            log(traceback.format_exc())
            sys.exit(1)
        return

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

    from desktop_ui import run_gui as run_desktop_gui

    run_desktop_gui(sys.modules[__name__])


if __name__ == "__main__":
    main()
