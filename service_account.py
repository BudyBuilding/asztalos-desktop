"""Service account setup in a separate elevated GUI; secrets never cross argv/files."""
import win32api
import win32con
import win32security
import win32service
import win32serviceutil
from PySide6.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QFormLayout, QLabel, QLineEdit,
    QDialogButtonBox, QMessageBox,
)


def set_service_account(service_name, username, password):
    if "\\" in username:
        domain, user = username.split("\\", 1)
        if domain == ".":
            domain = win32api.GetComputerName()
        username = domain + "\\" + user
    elif "@" in username:
        user, domain = username, None
    else:
        user, domain = username, win32api.GetComputerName()
        username = domain + "\\" + user
    # Validate before changing the service or granting logon rights.
    token = win32security.LogonUser(
        user, domain, password, win32con.LOGON32_LOGON_NETWORK,
        win32con.LOGON32_PROVIDER_DEFAULT,
    )
    token.Close()
    sid = win32security.LookupAccountName(None, username)[0]
    policy = win32security.LsaOpenPolicy(
        None, win32security.POLICY_LOOKUP_NAMES | win32security.POLICY_CREATE_ACCOUNT
    )
    try:
        win32security.LsaAddAccountRights(policy, sid, ("SeServiceLogonRight",))
    finally:
        policy.Close()
    manager = win32service.OpenSCManager(None, None, win32service.SC_MANAGER_CONNECT)
    try:
        service = win32service.OpenService(manager, service_name, win32service.SERVICE_CHANGE_CONFIG)
        try:
            win32service.ChangeServiceConfig(
                service, win32service.SERVICE_NO_CHANGE, win32service.SERVICE_NO_CHANGE,
                win32service.SERVICE_NO_CHANGE, None, None, 0, None,
                username, password, None,
            )
        finally:
            win32service.CloseServiceHandle(service)
    finally:
        win32service.CloseServiceHandle(manager)


def configure_account(service_name, manage_service):
    app = QApplication.instance() or QApplication([])
    dialog = QDialog()
    dialog.setWindowTitle("Hálózati hozzáférés – szolgáltatásfiók")
    dialog.setMinimumWidth(540)
    layout = QVBoxLayout(dialog)
    description = QLabel(
        "A feldolgozó ezzel a Windows-fiókkal fut majd. Olyan fiókot adj meg, "
        "amely hozzáfér a forrásmappához és írhat a hálózati célmappákba.\n\n"
        "A jelszót a Windows szolgáltatáskezelő tárolja; az alkalmazás nem menti "
        "fájlba. Futó szolgáltatás esetén a mentés újraindítja a feldolgozót."
    )
    description.setWordWrap(True)
    layout.addWidget(description)
    form = QFormLayout()
    username = QLineEdit()
    username.setPlaceholderText(r"GÉPNÉV\felhasználó vagy TARTOMÁNY\felhasználó")
    password = QLineEdit()
    password.setEchoMode(QLineEdit.EchoMode.Password)
    form.addRow("Windows-fiók", username)
    form.addRow("Jelszó (nem PIN-kód)", password)
    layout.addLayout(form)
    buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
    buttons.button(QDialogButtonBox.StandardButton.Save).setText("Beállítás alkalmazása")
    buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Mégse")
    layout.addWidget(buttons)

    def apply():
        if not username.text().strip() or not password.text():
            QMessageBox.warning(dialog, "Hiányzó adat", "Add meg a Windows-fiókot és a jelszavát.")
            return
        try:
            manage_service("register")
            was_running = win32serviceutil.QueryServiceStatus(service_name)[1] == win32service.SERVICE_RUNNING
            set_service_account(service_name, username.text().strip(), password.text())
            password.clear()
            if was_running:
                manage_service("stop")
                manage_service("start")
            QMessageBox.information(dialog, "Kész", "A szolgáltatásfiók beállítva. A feldolgozó az Indítás gombbal indítható.")
            dialog.accept()
        except Exception as error:
            # Never log credentials or include arbitrary exception arguments.
            code = getattr(error, "winerror", None)
            message = getattr(error, "strerror", None)
            QMessageBox.warning(dialog, "A beállítás nem fejeződött be",
                                f"Windows hibakód: {code}\n{message or 'Ellenőrizd a fiókot, a jelszót és a szolgáltatás állapotát.'}")
        finally:
            password.clear()

    buttons.accepted.connect(apply)
    buttons.rejected.connect(dialog.reject)
    dialog.exec()
    password.clear()
