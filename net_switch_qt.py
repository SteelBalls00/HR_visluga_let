import sys
import subprocess
import ctypes
import re

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout,
    QPushButton, QLabel, QTextEdit
)
from PyQt5.QtCore import Qt


# --- Проверка / запрос прав администратора ---
def ensure_admin():
    if ctypes.windll.shell32.IsUserAnAdmin():
        return True
    else:
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, " ".join(sys.argv), None, 1
        )
        return False


class NetSwitcher(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Переключение сети")
        self.setFixedSize(420, 420)

        layout = QVBoxLayout()

        self.ip_label = QLabel("Текущий IP: определение...")
        self.ip_label.setAlignment(Qt.AlignCenter)
        self.ip_label.setStyleSheet("font-size: 14px; font-weight: bold;")
        layout.addWidget(self.ip_label)

        self.btn_local = QPushButton("Сеть суда")
        self.btn_local.clicked.connect(self.set_local)
        layout.addWidget(self.btn_local)

        self.btn_internet = QPushButton("Интернет")
        self.btn_internet.clicked.connect(self.set_internet)
        layout.addWidget(self.btn_internet)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #00ff88;
                font-family: Consolas;
                font-size: 11px;
            }
        """)
        layout.addWidget(self.log)

        self.setLayout(layout)

        self.iface = self.get_active_interface()
        self.update_ip()


    def log_message(self, text):
        self.log.append(text)


    def run_cmd(self, cmd):
        self.log_message(f"> {cmd}")
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding="cp866")
        if result.stdout:
            self.log_message(result.stdout.strip())
        if result.stderr:
            self.log_message("ERROR: " + result.stderr.strip())


    def get_active_interface(self):
        result = subprocess.check_output("netsh interface show interface", shell=True, encoding="cp866")
        for line in result.splitlines():
            if "Connected" in line or "Подключен" in line:
                parts = line.split()
                return parts[-1]
        return "Ethernet"


    def get_current_ip(self):
        result = subprocess.check_output("ipconfig", shell=True, encoding="cp866")
        match = re.search(r"IPv4.*?:\s*([\d\.]+)", result)
        if match:
            return match.group(1)
        return "Не найден"


    def update_ip(self):
        ip = self.get_current_ip()
        self.ip_label.setText(f"Текущий IP: {ip}")


    def apply_settings(self, ip, mask, gateway, dns):
        self.log.clear()
        self.log_message(f"Интерфейс: {self.iface}")
        self.log_message("Применяем настройки...\n")

        self.run_cmd(
            f'netsh interface ip set address name="{self.iface}" source=static addr={ip} mask={mask} gateway={gateway}'
        )
        self.run_cmd(
            f'netsh interface ip set dns name="{self.iface}" source=static addr={dns} register=PRIMARY'
        )

        self.log_message("\nГотово.")
        self.update_ip()


    def set_local(self):
        self.apply_settings(
            "192.168.0.192",
            "255.255.0.0",
            "0.0.0.0",
            "192.168.0.215"
        )


    def set_internet(self):
        self.apply_settings(
            "10.10.0.71",
            "255.255.255.0",
            "10.10.0.150",
            "8.8.8.8"
        )


if __name__ == "__main__":
    if not ensure_admin():
        sys.exit()

    app = QApplication(sys.argv)
    window = NetSwitcher()
    window.show()
    sys.exit(app.exec_())