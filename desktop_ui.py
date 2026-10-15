import os
import subprocess
import sys
import winreg

import win32api
import win32event
import win32process
import win32service
import win32serviceutil
from win32com.shell import shell, shellcon
from PySide6.QtCore import QTimer
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QGridLayout, QLabel, QLineEdit,
    QPushButton, QFileDialog, QMessageBox, QCheckBox, QSystemTrayIcon,
    QMenu, QStyle,
)
from desktop_worker import Processor

RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
RUN_NAME = "AsztalosoftDesktop"


def set_autostart(enabled):
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as key:
        if enabled:
            args = [sys.executable]
            if not getattr(sys, "frozen", False):
                args.append(os.path.join(os.path.dirname(__file__), "asztalosoft_desktop.py"))
            args.append("--autostart")
            winreg.SetValueEx(key, RUN_NAME, 0, winreg.REG_SZ, subprocess.list2cmdline(args))
        else:
            try:
                winreg.DeleteValue(key, RUN_NAME)
            except FileNotFoundError:
                pass


def legacy_needs_retirement(name):
    manager = win32service.OpenSCManager(None, None, win32service.SC_MANAGER_CONNECT)
    try:
        try:
            service = win32service.OpenService(manager, name, win32service.SERVICE_QUERY_CONFIG | win32service.SERVICE_QUERY_STATUS)
        except win32service.error as error:
            if error.winerror == 1060:
                return False
            raise
        try:
            return (win32service.QueryServiceConfig(service)[1] != win32service.SERVICE_DISABLED
                    or win32service.QueryServiceStatus(service)[1] != win32service.SERVICE_STOPPED)
        finally:
            win32service.CloseServiceHandle(service)
    finally:
        win32service.CloseServiceHandle(manager)


class App(QWidget):
    def __init__(self, backend):
        super().__init__()
        self.backend = backend
        self.worker = None
        self.exiting = False
        self.has_error = False
        self.setWindowTitle("Asztalosoft – Fájlkezelő")
        self.setMinimumWidth(720)
        config = backend.load_config()
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("A feldolgozó a saját Windows-felhasználóddal fut. Bezáráskor a tálcán marad."))
        grid = QGridLayout()
        self.fields = {}
        self.browse_buttons = []
        for row, (key, title) in enumerate((
            ("watch_folder", "Forrásmappa"), ("target_list", "Célmappa – Lista"),
            ("target_program", "Célmappa – Program"), ("target_pdf", "Célmappa – PDF"),
            ("archive", "Archívum"),
        )):
            field = QLineEdit(config.get(key, ""))
            field.setPlaceholderText(r"C:\mappa vagy \\szerver\megosztás\mappa")
            button = QPushButton("Tallózás")
            button.clicked.connect(lambda checked=False, f=field: self.browse(f))
            self.fields[key] = field
            self.browse_buttons.append(button)
            grid.addWidget(QLabel(title), row, 0)
            grid.addWidget(field, row, 1)
            grid.addWidget(button, row, 2)
        layout.addLayout(grid)
        self.autostart = QCheckBox("Automatikus indítás és figyelés Windows-bejelentkezéskor")
        self.autostart.setChecked(config.get("autostart", False))
        layout.addWidget(self.autostart)
        self.status = QLabel("Figyelés: LEÁLLÍTVA")
        self.status.setWordWrap(True)
        layout.addWidget(self.status)
        buttons = QGridLayout()
        self.save_button = QPushButton("Mentés")
        self.start_button = QPushButton("Indítás")
        self.stop_button = QPushButton("Megállítás")
        self.stop_button.setEnabled(False)
        log_button = QPushButton("Napló")
        self.save_button.clicked.connect(self.save)
        self.start_button.clicked.connect(self.start)
        self.stop_button.clicked.connect(self.stop)
        log_button.clicked.connect(lambda: os.startfile(backend.LOG_FILE))
        for col, button in enumerate((self.save_button, self.start_button, self.stop_button, log_button)):
            buttons.addWidget(button, 0, col)
        layout.addLayout(buttons)
        icon = QIcon(backend.ICON_FILE)
        if icon.isNull():
            icon = self.style().standardIcon(QStyle.StandardPixmap.SP_DriveHDIcon)
        self.setWindowIcon(icon)
        self.tray = QSystemTrayIcon(icon, self)
        self.tray.setToolTip("Asztalosoft – figyelés leállítva")
        menu = QMenu(self)
        menu.addAction("Megnyitás", self.reveal)
        menu.addAction("Indítás", self.start)
        menu.addAction("Megállítás", self.stop)
        menu.addSeparator()
        menu.addAction("Kilépés", self.exit_app)
        self.tray.setContextMenu(menu)
        self.tray.activated.connect(lambda reason: self.reveal() if reason == QSystemTrayIcon.ActivationReason.DoubleClick else None)
        self.tray.show()

    def reveal(self):
        self.showNormal()
        self.raise_()
        self.activateWindow()

    def browse(self, field):
        path = QFileDialog.getExistingDirectory(self, "Válassz mappát", field.text(), QFileDialog.Option.ShowDirsOnly)
        if path:
            field.setText(self.backend.normalize_folder_path(path))

    def config(self):
        return {**{key: self.backend.normalize_folder_path(field.text()) for key, field in self.fields.items()},
                "zip_prefix": "work-", "autostart": self.autostart.isChecked()}

    def persist(self):
        config = self.config()
        set_autostart(config["autostart"])
        self.backend.save_config(config)
        return config

    def save(self):
        try:
            self.persist()
            QMessageBox.information(self, "Mentés", "Beállítások elmentve.")
        except Exception as error:
            self.error(str(error))

    def retire_legacy(self):
        if not legacy_needs_retirement(self.backend.SERVICE_NAME):
            return
        self.status.setText("A korábbi szolgáltatás kikapcsolása – egyszeri rendszergazdai engedély szükséges.")
        args = [] if getattr(sys, "frozen", False) else [os.path.abspath(self.backend.__file__)]
        args.append("--retire-service")
        process = shell.ShellExecuteEx(
            fMask=shellcon.SEE_MASK_NOCLOSEPROCESS, hwnd=int(self.winId()),
            lpVerb="runas", lpFile=sys.executable,
            lpParameters=subprocess.list2cmdline(args), nShow=0,
        )["hProcess"]
        try:
            win32event.WaitForSingleObject(process, win32event.INFINITE)
            if win32process.GetExitCodeProcess(process) != 0:
                raise RuntimeError("Nem sikerült kikapcsolni a régi szolgáltatást. Részletek a naplóban.")
        finally:
            win32api.CloseHandle(process)
        if legacy_needs_retirement(self.backend.SERVICE_NAME):
            raise RuntimeError("A régi szolgáltatás még aktív; a kettős feldolgozás elkerüléséhez ki kell kapcsolni.")

    def start(self):
        if self.worker is not None or self.exiting:
            return
        try:
            config = self.persist()
            self.retire_legacy()
            self.has_error = False
            self.worker = Processor(config, self.backend.process_zip)
            self.worker.ready.connect(self.ready)
            self.worker.failed.connect(self.error)
            self.worker.file_error.connect(self.error)
            self.worker.finished.connect(self.finished)
            self.set_busy(True)
            self.status.setText("Mappák ellenőrzése és figyelés indítása…")
            self.worker.start()
        except Exception as error:
            self.error(str(error))

    def set_busy(self, busy):
        for widget in (*self.fields.values(), *self.browse_buttons, self.autostart, self.save_button, self.start_button):
            widget.setEnabled(not busy)
        self.stop_button.setEnabled(busy)

    def ready(self):
        self.status.setText("Figyelés: FUT – a saját Windows-fiókoddal")
        self.tray.setToolTip("Asztalosoft – figyelés fut")

    def error(self, message):
        self.has_error = True
        self.backend.log(message)
        self.status.setText("Hiba: " + message)
        self.tray.setToolTip("Asztalosoft – feldolgozási hiba")
        self.tray.showMessage("Asztalosoft", message, QSystemTrayIcon.MessageIcon.Warning)
        if self.isVisible():
            QMessageBox.warning(self, "Feldolgozási hiba", message)

    def stop(self):
        if self.worker is not None:
            self.status.setText("Leállítás – a folyamatban lévő fájl feldolgozása még befejeződik…")
            self.stop_button.setEnabled(False)
            self.worker.stop()

    def finished(self):
        self.worker.deleteLater()
        self.worker = None
        self.set_busy(False)
        if not self.has_error:
            self.status.setText("Figyelés: LEÁLLÍTVA")
            self.tray.setToolTip("Asztalosoft – figyelés leállítva")
        if self.exiting:
            self.tray.hide()
            QApplication.instance().quit()

    def closeEvent(self, event):
        if QSystemTrayIcon.isSystemTrayAvailable():
            self.hide()
            event.ignore()
            self.tray.showMessage("Asztalosoft", "Az alkalmazás a tálcán maradt. Kilépés a tálcaikon menüjéből.")
        else:
            event.ignore()
            self.exit_app()

    def exit_app(self):
        self.exiting = True
        if self.worker is not None:
            self.stop()
        else:
            self.tray.hide()
            QApplication.instance().quit()


def run_gui(backend):
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    mutex = win32event.CreateMutex(None, False, "Local\\AsztalosoftDesktopProcessor")
    if win32api.GetLastError() == 183:
        QMessageBox.information(None, "Asztalosoft", "Az alkalmazás már fut. Nyisd meg a tálcaikonjáról.")
        win32api.CloseHandle(mutex)
        return
    try:
        window = App(backend)
        if "--autostart" in sys.argv:
            if not QSystemTrayIcon.isSystemTrayAvailable():
                window.show()
            QTimer.singleShot(0, window.start)
        else:
            window.show()
        app.exec()
    finally:
        win32api.CloseHandle(mutex)
