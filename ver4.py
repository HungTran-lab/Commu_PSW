import sys
import os
import serial
import serial.tools.list_ports
from PyQt5 import QtWidgets, uic
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QTimer


# ============================================================
#  Hỗ trợ load file .ui / icon khi đóng gói PyInstaller
# ============================================================
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# ============================================================
#  CLASS PSW CONTROLLER (SCPI)
# ============================================================
class PSWController:
    def __init__(self):
        self.ser = None
        self.target_v = 0.0
        self.target_i = 0.0

    def open(self, port):
        self.ser = serial.Serial(
            port=port,
            baudrate=9600,
            timeout=1,
            bytesize=8,
            parity=serial.PARITY_NONE,
            stopbits=1
        )

    def close(self):
        if self.ser and self.ser.is_open:
            self.ser.close()
        self.ser = None

    def is_open(self):
        return self.ser is not None and self.ser.is_open

    def _write(self, cmd):
        if not self.is_open():
            raise RuntimeError("Serial not open")
        self.ser.write((cmd + "\r\n").encode())

    def query(self, cmd):
        self._write(cmd)
        return self.ser.readline().decode(errors="ignore").strip()

    # API
    def get_idn(self):
        return self.query("*IDN?")

    def set_voltage(self, v):
        self.target_v = v
        self._write(f"APPLy {v:.4f},{self.target_i:.4f}")

    def set_current(self, i):
        self.target_i = i
        self._write(f"APPLy {self.target_v:.4f},{i:.4f}")

    def output_on(self):
        self._write("OUTPut:STATe ON")

    def output_off(self):
        self._write("OUTPut:STATe OFF")

    def measure_all(self):
        raw = self.query("MEASure:SCALar:ALL?")
        try:
            v, i = raw.split(",")
            return float(v), float(i), raw
        except:
            return None, None, raw


# ============================================================
#  MAIN WINDOW
# ============================================================
class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()

        # Load UI from packaged exe
        uic.loadUi(resource_path("subline.ui"), self)
        self.setWindowTitle("PSW Controller Ver3")
        self.setWindowIcon(QIcon(resource_path("psw.png")))

        self.psw = PSWController()

        # Widgets mapping
        self.combo_ports = self.findChild(QtWidgets.QComboBox, "comboBox")
        self.log_widget = self.findChild(QtWidgets.QTextEdit, "logg")

        self.btn_connect = self.findChild(QtWidgets.QPushButton, "connect")
        self.btn_clean = self.findChild(QtWidgets.QPushButton, "clean")
        self.btn_test = self.findChild(QtWidgets.QPushButton, "test")
        self.btn_read = self.findChild(QtWidgets.QPushButton, "read")

        self.txtVoltage = self.findChild(QtWidgets.QLineEdit, "txtVoltage")
        self.txtCurrent = self.findChild(QtWidgets.QLineEdit, "txtCurrent")
        self.btn_set = self.findChild(QtWidgets.QPushButton, "set")

        self.btn_on = self.findChild(QtWidgets.QPushButton, "on")
        self.btn_off = self.findChild(QtWidgets.QPushButton, "off")

        self.btn_vs70 = self.findChild(QtWidgets.QPushButton, "vs70")
        self.btn_vs25 = self.findChild(QtWidgets.QPushButton, "vs25")
        self.lbl_70 = self.findChild(QtWidgets.QLabel, "status_18")
        self.lbl_25 = self.findChild(QtWidgets.QLabel, "status_25")

        # Timer for blinking RUN
        self.blink_timer = QTimer()
        self.blink_timer.timeout.connect(self.blink_run)
        self.blink_state = False
        self.active_label = None

        # SIGNALS
        self.btn_connect.clicked.connect(self.handle_connect)
        self.btn_clean.clicked.connect(lambda: self.log_widget.clear())
        self.btn_test.clicked.connect(self.handle_test)
        self.btn_read.clicked.connect(self.handle_read)
        self.btn_set.clicked.connect(self.handle_set_v_i)

        self.btn_on.clicked.connect(self.handle_on)
        self.btn_off.clicked.connect(self.handle_off)

        self.btn_vs70.clicked.connect(lambda: self.set_stick(70))
        self.btn_vs25.clicked.connect(lambda: self.set_stick(25))

        self.populate_ports()

    # ----- Helpers -----
    def log(self, msg):
        self.log_widget.append(msg)

    def populate_ports(self):
        self.combo_ports.clear()
        for p in serial.tools.list_ports.comports():
            self.combo_ports.addItem(p.device, p.device)

    # ----- Handlers -----
    def handle_connect(self):
        if self.psw.is_open():
            self.psw.close()
            self.log("Disconnected.")
            self.btn_connect.setText("Connect")
            return

        port = self.combo_ports.currentData()
        if not port:
            self.log("No COM selected")
            return

        try:
            self.psw.open(port)
            self.log(f"Connected {port}")
            self.btn_connect.setText("Disconnect")
        except Exception as e:
            self.log(f"Error: {e}")

    def handle_test(self):
        if not self.psw.is_open():
            self.log("Not connected")
            return
        self.log(self.psw.get_idn())

    def handle_read(self):
        if not self.psw.is_open():
            self.log("Not connected")
            return
        v, i, raw = self.psw.measure_all()
        if v is None:
            self.log(raw)
        else:
            self.log(f"V={v:.3f}  I={i:.3f}")

    def handle_set_v_i(self):
        if not self.psw.is_open():
            self.log("Not connected")
            return
        try:
            v = float(self.txtVoltage.text())
            a = float(self.txtCurrent.text())
            self.psw.set_voltage(v)
            self.psw.set_current(a)
            self.log(f"Set V={v}, I={a}")
        except:
            self.log("Invalid V/I")

    # ----- Output control -----
    def handle_on(self):
        if not self.psw.is_open():
            self.log("Not connected")
            return
        self.psw.output_on()
        self.log("Output ON")

        self.btn_on.setStyleSheet("background-color: red; color: white; font-weight: bold;")

    def handle_off(self):
        if not self.psw.is_open():
            self.log("Not connected")
            return

        self.psw.output_off()
        self.log("Output OFF")

        # Disable Auto ON visual effects
        self.blink_timer.stop()
        self.active_label = None

        self.lbl_70.setText("...")
        self.lbl_25.setText("...")
        self.lbl_70.setStyleSheet("")
        self.lbl_25.setStyleSheet("")

        self.btn_on.setStyleSheet("")

        self.log("AUTO ON disabled")

    # ----- Stick mode VS70 / VS25 -----
    def set_stick(self, value):
        if not self.psw.is_open():
            self.log("Not connected")
            return

        # Stop previous blinking
        self.blink_timer.stop()
        self.lbl_70.setStyleSheet("")
        self.lbl_25.setStyleSheet("")
        self.active_label = None

        if value == 70:
            self.psw.set_voltage(18.0)
            self.psw.set_current(10.0)
            self.log("VS70: Set 18V / 10A")

            self.psw.output_on()
            self.log("Output AUTO ON")

            self.active_label = self.lbl_70
            self.lbl_70.setText("RUN")
            self.lbl_25.setText("...")

        elif value == 25:
            self.psw.set_voltage(25.0)
            self.psw.set_current(10.0)
            self.log("VS25: Set 25V / 10A")

            self.psw.output_on()
            self.log("Output AUTO ON")

            self.active_label = self.lbl_25
            self.lbl_25.setText("RUN")
            self.lbl_70.setText("...")

        # Start blinking
        self.blink_state = False
        self.blink_timer.start(500)

        self.log(f"Stick mode VS{value}")

    # ----- Blink effect -----
    def blink_run(self):
        if self.active_label is None:
            return

        if self.blink_state:
            self.active_label.setStyleSheet(
                "background-color: red; color: white; font-weight: bold; padding: 4px; border-radius: 4px;"
            )
        else:
            self.active_label.setStyleSheet("")

        self.blink_state = not self.blink_state


# ============================================================
# MAIN
# ============================================================
def main():
    app = QtWidgets.QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
