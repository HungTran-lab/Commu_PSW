import sys
import os
import serial
import serial.tools.list_ports
from PyQt5 import QtWidgets, uic
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QTimer, QSettings

from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QSlider, QMessageBox


# ============================================================
#  Hỗ trợ load file .ui / icon khi đóng gói PyInstaller
# ============================================================
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # type: ignore
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# ============================================================
#  Lớp điều khiển PSW qua Serial / SCPI
# ============================================================
class PSWController:
    def __init__(self):
        self.ser = None
        self.target_v = 0.0
        self.target_i = 0.0

    def is_open(self):
        return self.ser is not None and self.ser.is_open

    def open(self, port, baudrate=9600, timeout=1):
        self.close()
        self.ser = serial.Serial(port, baudrate=baudrate, timeout=timeout)

    def close(self):
        if self.ser is not None:
            try:
                if self.ser.is_open:
                    self.ser.close()
            except Exception:
                pass
        self.ser = None

    def _write(self, cmd):
        if not self.is_open():
            raise RuntimeError("Serial not open")
        if not cmd.endswith("\n"):
            cmd = cmd + "\n"
        try:
            self.ser.write(cmd.encode("ascii"))
        except Exception as e:
            self.close()
            raise RuntimeError(f"Serial write error: {e}")

    def _readline(self):
        if not self.is_open():
            raise RuntimeError("Serial not open")
        try:
            line = self.ser.readline().decode(errors="ignore").strip()
            return line
        except Exception as e:
            self.close()
            raise RuntimeError(f"Serial read error: {e}")

    def query(self, cmd):
        self._write(cmd)
        return self._readline()

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
        except Exception:
            return None, None, raw


# ============================================================
#  MAIN WINDOW
# ============================================================
class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()

        # Load UI từ file .ui
        uic.loadUi(resource_path("subline.ui"), self)
        self.setWindowTitle("PSW Controller")
        self.setWindowIcon(QIcon(resource_path("psw.png")))

        self.psw = PSWController()

        self.actionInfor.triggered.connect(self.show_about_message)
        # Cố định kích thước cửa sổ
        self.setFixedSize(771, 609)

        # QSettings: nhớ COM lần cuối
        self.settings = QSettings("songhung.tr", "PSWController")
        last_port = self.settings.value("last_port", "", type=str)

        # Widgets mapping
        self.combo_ports = self.findChild(QtWidgets.QComboBox, "comboBox")
        self.log_widget = self.findChild(QtWidgets.QTextEdit, "logg")

        self.btn_connect = self.findChild(QtWidgets.QPushButton, "connect")
        self.btn_clean = self.findChild(QtWidgets.QPushButton, "clean")
        self.btn_test = self.findChild(QtWidgets.QPushButton, "test")
        self.btn_read = self.findChild(QtWidgets.QPushButton, "read")
        self.btn_refresh = self.findChild(QtWidgets.QPushButton, "refresh")

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
        self.btn_refresh.clicked.connect(self.handle_refresh)
        self.btn_set.clicked.connect(self.handle_set_v_i)

        self.btn_on.clicked.connect(self.handle_on)
        self.btn_off.clicked.connect(self.handle_off)

        self.btn_vs70.clicked.connect(lambda: self.set_stick(70))
        self.btn_vs25.clicked.connect(lambda: self.set_stick(25))

        # Scan COM lần đầu, ưu tiên chọn lại COM đã lưu
        self.populate_ports(preferred_port=last_port)
        self.set_controls_enabled(False)

    # ----- Helpers -----
    def log(self, msg):
        self.log_widget.append(msg)

    def set_controls_enabled(self, enabled):
        for w in [
            self.btn_test,
            self.btn_read,
            self.btn_set,
            self.btn_on,
            self.btn_off,
            self.btn_vs70,
            self.btn_vs25,
        ]:
            w.setEnabled(enabled)

    def populate_ports(self, preferred_port=None):
        """Scan danh sách COM, nếu preferred_port còn tồn tại thì tự chọn nó."""
        current_data = self.combo_ports.currentData()
        self.combo_ports.clear()

        index_to_select = -1
        ports = list(serial.tools.list_ports.comports())
        for idx, p in enumerate(ports):
            self.combo_ports.addItem(p.device, p.device)
            if preferred_port and p.device == preferred_port:
                index_to_select = idx
            elif not preferred_port and current_data and p.device == current_data:
                index_to_select = idx

        if index_to_select >= 0:
            self.combo_ports.setCurrentIndex(index_to_select)

    def handle_refresh(self):
        """Nút Refresh: scan lại COM, cố gắng giữ COM đang chọn."""
        preferred = self.combo_ports.currentData()
        self.populate_ports(preferred_port=preferred)

        count = self.combo_ports.count()
        self.log(f"Refreshed Done:{count} port found")

    def handle_connection_lost(self, err):
        """Gọi hàm này khi có lỗi giao tiếp -> tự chuyển về trạng thái Disconnected."""
        self.log(f"Connection lost: {err}")
        self.psw.close()
        self.btn_connect.setText("Connect")
        self.set_controls_enabled(False)

        # Dừng blink, reset label/status
        self.blink_timer.stop()
        self.blink_state = False
        self.active_label = None

        self.lbl_70.setText("STOP")
        self.lbl_25.setText("STOP")
        self.lbl_70.setStyleSheet("")
        self.lbl_25.setStyleSheet("")

        self.btn_on.setStyleSheet("")
        self.btn_off.setStyleSheet("background-color: lightgray;")

    # ----- Handlers -----
    def handle_connect(self):
        # Nếu đang mở -> bấm lần nữa là Disconnect
        if self.psw.is_open():
            self.psw.close()
            self.log("Disconnected.")
            self.btn_connect.setText("Connect")
            self.set_controls_enabled(False)
            return

        port = self.combo_ports.currentData()
        if not port:
            self.log("No COM selected")
            return

        try:
            self.psw.open(port)
            self.log(f"Connected {port}")
            self.btn_connect.setText("Disconnect")
            self.set_controls_enabled(True)

            # Lưu COM vừa connect thành công
            self.settings.setValue("last_port", port)
        except Exception as e:
            self.log(f"Error: {e}")
            self.set_controls_enabled(False)

    def handle_test(self):
        if not self.psw.is_open():
            self.log("Not connected")
            return
        try:
            self.log(self.psw.get_idn())
        except Exception as e:
            self.handle_connection_lost(e)

    def handle_read(self):
        if not self.psw.is_open():
            self.log("Not connected")
            return
        try:
            v, i, raw = self.psw.measure_all()
        except Exception as e:
            self.handle_connection_lost(e)
            return

        if v is None:
            self.log(f"Read error: {raw}")
        else:
            self.log(f"Read: V={v:.3f} V, I={i:.3f} A  (raw: {raw})")

    def handle_set_v_i(self):
        if not self.psw.is_open():
            self.log("Not connected")
            return

        try:
            v = float(self.txtVoltage.text())
            a = float(self.txtCurrent.text())
        except Exception:
            self.log("Invalid V/I")
            return

        try:
            self.psw.set_voltage(v)
            self.psw.set_current(a)
        except Exception as e:
            self.handle_connection_lost(e)
            return

        self.log(f"Set V={v}, I={a}")

    def handle_on(self):
        if not self.psw.is_open():
            self.log("Not connected")
            return

        try:
            self.psw.output_on()
        except Exception as e:
            self.handle_connection_lost(e)
            return

        self.log("Output ON")
        self.btn_on.setStyleSheet("background-color: red; color: white;")
        self.btn_off.setStyleSheet("")

    def handle_off(self):
        if not self.psw.is_open():
            self.log("Not connected")
            return

        try:
            self.psw.output_off()
        except Exception as e:
            self.handle_connection_lost(e)
            return

        self.log("Output OFF")
        self.btn_on.setStyleSheet("")
        self.btn_off.setStyleSheet("background-color: lightgray;")

        # Stop blinking and reset labels
        self.blink_timer.stop()
        self.blink_state = False
        self.active_label = None
        self.lbl_70.setText("STOP")
        self.lbl_25.setText("STOP")
        self.lbl_70.setStyleSheet("")
        self.lbl_25.setStyleSheet("")

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

        try:
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
        except Exception as e:
            self.handle_connection_lost(e)
            return

        # Start blinking
        if self.active_label:
            self.blink_state = False
            self.blink_timer.start(500)  # 500 ms

    # ----- Blink handler -----
    def blink_run(self):
        if not self.active_label:
            return

        self.blink_state = not self.blink_state
        if self.blink_state:
            self.active_label.setStyleSheet("background-color: yellow; color: black;")
        else:
            self.active_label.setStyleSheet("")

    def show_about_message(self):
        QMessageBox.information(self, "Infor", "Ver 5.\nDec-25\nPIC. songhung.tr")


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
