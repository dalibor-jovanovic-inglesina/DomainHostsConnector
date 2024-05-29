import socket
import win32com.client
import win32net
import sys
import os
import subprocess
import ctypes
from concurrent.futures import ThreadPoolExecutor
from PyQt6.QtWidgets import QApplication, QWidget, QGridLayout, QPushButton, QDialog, QMessageBox, QLabel, QScrollArea
from PyQt6.QtGui import QColor, QPalette
from PyQt6.QtCore import Qt, QTimer

# Devices to exclude from the scan
excluded_devices = {"Device1", "Device2"} #fqdn of devices to exclude from the table eq: excluded_devices = {"Device1", "Device2"...}

def get_domain_name():
    try:
        return socket.getfqdn().split('.', 1)[1].upper()
    except Exception as e:
        print(f"Error during the gathering of the domain name: {str(e)}")
        return None

# Function to get a devices list from the domain
def machines_in_domain(domain_name):
    adsi = win32com.client.Dispatch("ADsNameSpaces")
    nt = adsi.GetObject("", "WinNT:")
    result = nt.OpenDSObject("WinNT://%s" % domain_name, "", "", 0)
    result.Filter = ["computer"]

    for machine in result:
        machine_name = machine.Name
        # Exclude specified devices in excluded_devices
        if machine_name not in excluded_devices:
            yield machine_name

class CustomDialog(QDialog):
    def __init__(self, host_name):
        super().__init__()

        self.setWindowTitle("Host Commands")
        self.setFixedSize(400, 200)  # Change the height to reduce the space
        self.host_name = host_name

        self.init_ui()

    def init_ui(self):
        layout = QGridLayout(self)

        host_label = QLabel(f"<b><font size='6'>Nome host:</font></b> <font size='6'>{self.host_name}</font>")
        host_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(host_label, 0, 0, 1, 2)

        # Adding the ip address
        try:
            ip_address = socket.gethostbyname(self.host_name)
            ip_label = QLabel(f"<b><font size='5'>IP Address:</font></b> <font size='5'>{ip_address}</font>")
        except socket.gaierror:
            ip_label = QLabel("<b><font size='5'>IP Address:</font></b> <font size='5'>Invalid Address</font>")
        ip_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(ip_label, 1, 0, 1, 2)  # Positioning the ip address on the same line

        info_button = QPushButton("Infromation", self)
        info_button.setStyleSheet("background-color: white; color: black; font-size: 18px;")
        info_button.clicked.connect(self.show_info_prompt)
        layout.addWidget(info_button, 2, 0)

        connect_button = QPushButton("Connetti", self)
        connect_button.setStyleSheet("background-color: white; color: black; font-size: 18px;")
        connect_button.clicked.connect(self.handle_connect_button)
        layout.addWidget(connect_button, 2, 1)

    def show_info_prompt(self):
        hostname = self.host_name

        # Creating a prompt with the gathered informations
        info_prompt = QMessageBox(self)
        info_prompt.setWindowTitle("Host Informations")
        # Set the background color to rgb 53 53 53
        info_prompt.setStyleSheet("background-color: rgb(53, 53, 53); color: white;")
        info_prompt.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)  # Enabling text selection

        info_text = f"Host Informations for: {hostname}\n\n"

        info_text += "\nIP:\n"
        ip_info = self.ip_c()
        info_text += ip_info

        info_text += "\nMAC ADDRESS:\n"
        mac_info = self.mc_c()
        info_text += mac_info

        info_text += "\nMANUFACTURER:\n"
        manufacturer_info = self.mf_c()
        info_text += manufacturer_info

        info_text += "\nMODEL:\n"
        model_info = self.md_c()
        info_text += model_info

        info_text += "\nSERIAL:\n"
        serial_info = self.sn_c()
        info_text += serial_info

        info_prompt.setText(info_text)

        info_prompt.exec()

    def ip_c(self):
        command = f'powershell.exe -ExecutionPolicy Unrestricted -command "Test-Connection -ComputerName {self.host_name} -Count 1 | Select-Object -ExpandProperty IPv4Address | ForEach-Object {{ Write-Host $_.IPAddressToString }}"'
        process = subprocess.Popen(command, stdout=subprocess.PIPE, shell=True)
        output, _ = process.communicate()
        return output.decode("utf-8")

    def mc_c(self):
        command = f'powershell.exe -ExecutionPolicy Unrestricted -command "(Get-WmiObject win32_networkadapterconfiguration -ComputerName {self.host_name} | Select -Expand macaddress | Select -Index 0)"'
        process = subprocess.Popen(command, stdout=subprocess.PIPE, shell=True)
        output, _ = process.communicate()
        return output.decode("utf-8")

    def mf_c(self):
        command = f'powershell.exe -ExecutionPolicy Unrestricted -command "Get-WmiObject Win32_ComputerSystem -ComputerName {self.host_name} | ForEach-Object {{ Write-Host $_.Manufacturer }}"'
        process = subprocess.Popen(command, stdout=subprocess.PIPE, shell=True)
        output, _ = process.communicate()
        return output.decode("utf-8")

    def md_c(self):
        command = f'powershell.exe -ExecutionPolicy Unrestricted -command "Get-WmiObject Win32_ComputerSystem -ComputerName {self.host_name} | ForEach-Object {{ Write-Host $_.Model }}"'
        process = subprocess.Popen(command, stdout=subprocess.PIPE, shell=True)
        output, _ = process.communicate()
        return output.decode("utf-8")

    def sn_c(self):
        command = f'powershell.exe -ExecutionPolicy Unrestricted -command "Get-WmiObject win32_bios -ComputerName {self.host_name} | ForEach-Object {{ Write-Host $_.serialnumber }}"'
        process = subprocess.Popen(command, stdout=subprocess.PIPE, shell=True)
        output, _ = process.communicate()
        return output.decode("utf-8")
    
    def find_ultravnc_path(self):
        possible_paths = [
            "C:/Program Files (x86)/UltraVNC/vncviewer.exe",
            "C:/Program Files/UltraVNC/vncviewer.exe"
        ]

        for path in possible_paths:
            if os.path.exists(path):
                return path

        return None

    def connect_to_vnc(self, ip_address):
        try:
            ultravnc_path = self.find_ultravnc_path()

            if ultravnc_path and not self.host_name.lower().startswith("srv"):
                command = f'start "" "{ultravnc_path}" {ip_address}'
                subprocess.Popen(command, shell=True)
            else:
                subprocess.Popen(["mstsc", "/v:" + self.host_name])
        except Exception as e:
            QMessageBox.critical(self, "Errore", f"Impossibile avviare Ultravnc: {str(e)}")

    def handle_connect_button(self):
        try:
            if self.host_name.lower().startswith("srv"):
                # Avvia il Remote Desktop Protocol (RDP)
                subprocess.Popen(["mstsc", "/v:" + self.host_name])
            else:
                # Avvia Ultravnc
                ip_address = socket.gethostbyname(self.host_name)
                self.connect_to_vnc(ip_address)
        except socket.gaierror:
            QMessageBox.critical(self, "Errore", "Indirizzo IP non valido")

class HostListApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Domain Hosts Scanner')
        self.setGeometry(100, 100, 400, 300)

        self.layout = QGridLayout(self)  # Initializing the layout
        self.executor = ThreadPoolExecutor()  # ThreadPoolExecutor for executing the ping on a different thread

        self.init_ui()
        self.start_ping_timer()

    def init_ui(self):
        self.populate_host_list()

    def ping_host(self, host_name, button):
        result = subprocess.run(["ping", "-n", "1", "-w", "1000", host_name], stdout=subprocess.DEVNULL)
        color = "#CCFFCC" if result.returncode == 0 else "#FF9999"
        self.set_button_style(button, color)

    def set_button_style(self, button, color):
        button.setStyleSheet(f"background-color: {color}; color: black; text-align: left; padding: 5px 10px;")

    def show_host_dialog(self, host_name):
        dialog = CustomDialog(host_name)
        dialog.exec()

    def populate_host_list(self):
        domain_name = get_domain_name() #or " domain_name = DOMAIN_NAME"

        try:
            host_names = list(machines_in_domain(domain_name))
            num_columns = 10  # Set columns number

            num_rows = (len(host_names) + num_columns - 1) // num_columns

            for i, host_name in enumerate(host_names):
                row = i % num_rows
                col = i // num_rows

                button = QPushButton(host_name)
                if button.text().lower() in excluded_devices:
                    button.setStyleSheet("background-color: gray; color: white;")
                else:
                    button.setStyleSheet("text-align: left; padding: 5px 10px; color: black;")
                self.layout.addWidget(button, row, col)
                button.clicked.connect(lambda checked, name=host_name: self.show_host_dialog(name))

        except Exception as e:
            error_label = QPushButton(f"Error gathering the domain name '{domain_name}': {str(e)}")
            error_label.setStyleSheet("padding: 10px; background-color: red;")
            self.layout.addWidget(error_label, 0, 0, 1, num_columns)
    
    def start_ping_timer(self):
        self.check_host_status()

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_host_status)
        self.timer.start(5 * 60 * 1000)  # do a ping every 5 minutes
    
    def check_host_status(self):
        for row in range(self.layout.rowCount()):
            for col in range(self.layout.columnCount()):
                item = self.layout.itemAtPosition(row, col)
                if item is not None:
                    button = item.widget()
                    host_name = button.text()
                    self.executor.submit(self.ping_host, host_name, button)

if __name__ == '__main__':
    app = QApplication([])

    # Set the Dark Fusion theme
    dark_palette = app.palette()
    dark_palette.setColor(QPalette.ColorGroup.All, QPalette.ColorRole.Window, QColor(53, 53, 53))
    dark_palette.setColor(QPalette.ColorGroup.All, QPalette.ColorRole.WindowText, QColor(255, 255, 255))
    dark_palette.setColor(QPalette.ColorGroup.All, QPalette.ColorRole.Base, QColor(25, 25, 25))
    dark_palette.setColor(QPalette.ColorGroup.All, QPalette.ColorRole.AlternateBase, QColor(53, 53, 53))
    dark_palette.setColor(QPalette.ColorGroup.All, QPalette.ColorRole.ToolTipBase, QColor(255, 255, 220))
    dark_palette.setColor(QPalette.ColorGroup.All, QPalette.ColorRole.ToolTipText, QColor(0, 0, 0))
    dark_palette.setColor(QPalette.ColorGroup.All, QPalette.ColorRole.Text, QColor(255, 255, 255))
    dark_palette.setColor(QPalette.ColorGroup.All, QPalette.ColorRole.Button, QColor(53, 53, 53))
    dark_palette.setColor(QPalette.ColorGroup.All, QPalette.ColorRole.ButtonText, QColor(255, 255, 255))
    dark_palette.setColor(QPalette.ColorGroup.All, QPalette.ColorRole.BrightText, QColor(255, 0, 0))
    dark_palette.setColor(QPalette.ColorGroup.All, QPalette.ColorRole.Link, QColor(42, 130, 218))
    dark_palette.setColor(QPalette.ColorGroup.All, QPalette.ColorRole.Highlight, QColor(42, 130, 218))
    dark_palette.setColor(QPalette.ColorGroup.All, QPalette.ColorRole.HighlightedText, QColor(0, 0, 0))
    app.setPalette(dark_palette)

    host_list_app = HostListApp()
    host_list_app.show()
    app.exec()

