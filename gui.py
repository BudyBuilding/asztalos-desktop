import sys
import json
import os
import subprocess
import ctypes

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit,
    QPushButton, QFileDialog, QGridLayout,
    QVBoxLayout, QGroupBox, QMessageBox
)

CONFIG_FILE = "config.json"
SERVICE_NAME = "asztalosoft_processer"


def load_config():

    if not os.path.exists(CONFIG_FILE):
        return {}

    with open(CONFIG_FILE, encoding="utf8") as f:
        return json.load(f)


def save_config(data):

    with open(CONFIG_FILE, "w", encoding="utf8") as f:
        json.dump(data, f, indent=2)


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


class App(QWidget):

    def __init__(self):

        super().__init__()

        self.setWindowTitle("Asztalosoft - Fájlkezelő")
        self.setMinimumWidth(700)

        if os.path.exists("logo.ico"):
            self.setWindowIcon(QIcon("logo.ico"))

        config = load_config()

        layout = QVBoxLayout()

        # -------------------------
        # PATH SETTINGS
        # -------------------------

        paths_box = QGroupBox("Mappa beállítások")
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

            button.clicked.connect(
                lambda _, f=field: self.browse(f)
            )

            grid.addWidget(label, i, 0)
            grid.addWidget(field, i, 1)
            grid.addWidget(button, i, 2)

        paths_box.setLayout(grid)
        layout.addWidget(paths_box)

        # -------------------------
        # STATUS
        # -------------------------

        self.status = QLabel()
        layout.addWidget(self.status)

        # -------------------------
        # BUTTONS
        # -------------------------

        box = QGroupBox("Háttérfolyamat vezérlés")
        grid = QGridLayout()

        save_btn = QPushButton("Mentés")
        start_btn = QPushButton("Indítás")
        stop_btn = QPushButton("Megállítás")
        log_btn = QPushButton("Log")

        save_btn.clicked.connect(self.save)
        start_btn.clicked.connect(self.start_service)
        stop_btn.clicked.connect(self.stop_service)
        log_btn.clicked.connect(self.open_log)

        grid.addWidget(save_btn, 0, 0)
        grid.addWidget(start_btn, 0, 1)
        grid.addWidget(stop_btn, 0, 2)
        grid.addWidget(log_btn, 0, 3)

        box.setLayout(grid)
        layout.addWidget(box)

        self.setLayout(layout)

        if not is_admin():

            QMessageBox.warning(
                self,
                "Figyelmeztetés",
                "A programot admin joggal kell futtatni!"
            )

        self.update_status()

    # -------------------------

    def service_exe(self):

        base = os.path.dirname(os.path.abspath(sys.argv[0]))
        exe = os.path.join(base, "dist", "asztalosoft_processer.exe")

        return exe

    # -------------------------

    def browse(self, field):

        folder = QFileDialog.getExistingDirectory(self)

        if folder:
            field.setText(folder)

    # -------------------------

    def save(self):

        data = {
            "watch_folder": self.watch.text(),
            "target_list": self.target_list.text(),
            "target_program": self.target_program.text(),
            "archive": self.archive.text(),
            "zip_prefix": "work-"
        }

        save_config(data)

        QMessageBox.information(self, "Info", "Beállítások elmentve.")

    # -------------------------

    def service_status(self):

        r = subprocess.run(
            f"sc query {SERVICE_NAME}",
            shell=True,
            capture_output=True,
            text=True
        )

        return r.stdout + r.stderr

    # -------------------------

    def service_exists(self):

        r = subprocess.run(
            f"sc query {SERVICE_NAME}",
            shell=True,
            capture_output=True,
            text=True
        )

        return r.returncode == 0

    # -------------------------

    def is_running(self):

        out = self.service_status()

        return "RUNNING" in out

    # -------------------------

    def start_service(self):

        exe = self.service_exe()

        if not os.path.exists(exe):
            QMessageBox.critical(self, "Hiba", f"Exe nem található:\n{exe}")
            return

        if not self.service_exists():

            subprocess.run(f'"{exe}" install', shell=True)

        subprocess.run(f'"{exe}" start', shell=True)

        self.update_status()

    # -------------------------

    def stop_service(self):

        exe = self.service_exe()

        subprocess.run(f'"{exe}" stop', shell=True)

        self.update_status()

    def update_status(self):

        out = self.service_status()

        if "FAILED 1060" in out:

            self.status.setText("Service: nincs telepítve")
            return

        if "RUNNING" in out:

            self.status.setText("Service: FUT")

        else:

            self.status.setText("Service: TELEPÍTVE - NEM FUT")

    # -------------------------

    def open_log(self):

        base = os.path.dirname(os.path.abspath(sys.argv[0]))
        log = os.path.join(base, "log.txt")

        if os.path.exists(log):

            os.startfile(log)

        else:

            QMessageBox.information(
                self,
                "Info",
                "Még nincs log fájl."
            )


def run_gui():

    app = QApplication(sys.argv)

    w = App()
    w.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    run_gui()