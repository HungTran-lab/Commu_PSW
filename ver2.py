import sys
import serial
import serial.tools.list_ports

from PyQt5 import QtWidgets, uic
from PyQt5.QtGui import QIcon



class PSWController:
    """
    Điều khiển GW Instek PSW qua cổng COM (USB-CDC).
    """
    def __init__(self):
        self.ser = None
        self.target_v = 0.0
        self.target_i = 0.0

    # ====== Serial cơ bản ======
    def open(self, port, baudrate=9600, timeout=1.0):
        """
        Mở cổng COM.
        """
        self.ser = serial.Serial(
            port=port,
            baudrate=baudrate,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=timeout
        )

    def close(self):
        """
        Đóng cổng COM.
        """
        if self.ser and self.ser.is_open:
            self.ser.close()
        self.ser = None

    def is_open(self):
        return self.ser is not None and self.ser.is_open

    def _write(self, cmd: str):
        """
        Gửi 1 lệnh SCPI.
        Dùng CR+LF: '\r\n'
        """
        if not self.is_open():
            raise RuntimeError("Serial port not open")
        line = (cmd + "\r\n").encode("ascii")
        self.ser.write(line)

    def _readline(self) -> str:
        """
        Đọc 1 dòng trả lời từ PSW.
        """
        if not self.is_open():
            raise RuntimeError("Serial port not open")
        data = self.ser.readline().decode("ascii", errors="ignore").strip()
        return data

    def query(self, cmd: str) -> str:
        """
        Gửi lệnh query (có dấu ?) và đọc 1 dòng trả về.
        """
        self._write(cmd)
        return self._readline()

    # ====== API cơ bản ======
    def get_idn(self) -> str:
        return self.query("*IDN?")

    def apply_vi(self):
        """
        Gửi lệnh APPLy với điện áp & dòng hiện tại.
        """
        self._write(f"APPLy {self.target_v:.4f},{self.target_i:.4f}")

    def set_voltage(self, v: float):
        self.target_v = v
        self.apply_vi()

    def set_current(self, i: float):
        self.target_i = i
        self.apply_vi()

    def output_on(self):
        self._write("OUTPut:STATe ON")

    def output_off(self):
        self._write("OUTPut:STATe OFF")

    def measure_all(self):
        """
        Gửi MEASure:SCALar:ALL? → trả về (V, I, raw_string).
        """
        reply = self.query("MEASure:SCALar:ALL?")
        try:
            v_str, i_str = reply.split(",")
            v = float(v_str)
            i = float(i_str)
            return v, i, reply
        except Exception:
            # nếu parse lỗi, vẫn trả raw để debug
            return None, None, reply

    # ====== REMOTE / LOCAL ======
    def set_remote(self, remote: bool):
        """
        remote = True  -> SYSTem:REMote
        remote = False -> SYSTem:LOCal
        """
        if remote:
            self._write("SYSTem:REMote")
        else:
            self._write("SYSTem:LOCal")

    # ====== OVP / OCP ======
    def set_ovp(self, v: float):
        """
        Đặt mức OVP (Over Voltage Protection) theo Volt.
        """
        self._write(f"SOURce:VOLTage:PROTection:LEVel {v:.4f}")

    def get_ovp(self) -> float:
        reply = self.query("SOURce:VOLTage:PROTection:LEVel?")
        return float(reply)

    def disable_ovp(self):
        """
        “Tắt” OVP bằng cách set level = MAX.
        """
        self._write("SOURce:VOLTage:PROTection:LEVel MAX")

    def set_ocp(self, i: float):
        """
        Đặt mức OCP (Over Current Protection) theo Ampere.
        """
        self._write(f"SOURce:CURRent:PROTection:LEVel {i:.4f}")

    def get_ocp(self) -> float:
        reply = self.query("SOURce:CURRent:PROTection:LEVel?")
        return float(reply)

    def disable_ocp(self):
        """
        “Tắt” OCP bằng cách set level = MAX.
        """
        self._write("SOURce:CURRent:PROTection:LEVel MAX")

    def get_protection_status(self):
        """
        Đọc STATus:QUEStionable:CONDition?
        Ví dụ: bit 0 = OVP, bit 1 = OCP (tùy theo manual model)
        """
        s = int(self.query("STATus:QUEStionable:CONDition?"))
        ov = bool(s & 1)   # bit 0
        oc = bool(s & 2)   # bit 1
        return ov, oc, s


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi("gui.ui", self)   # đảm bảo gui.ui nằm cùng thư mục
        self.setWindowTitle("PSW Communication")
        self.setWindowIcon(QIcon("psw.png"))  # hoặc "psw.png"


        # ==== Widgets theo objectName trong gui.ui ====
        self.combo_ports = self.findChild(QtWidgets.QComboBox, "comboBox")
        self.btn_connect = self.findChild(QtWidgets.QPushButton, "connect")
        self.btn_test = self.findChild(QtWidgets.QPushButton, "test")
        self.btn_on = self.findChild(QtWidgets.QPushButton, "on")
        self.btn_off = self.findChild(QtWidgets.QPushButton, "off")
        self.edit_volt = self.findChild(QtWidgets.QLineEdit, "volt")
        self.edit_ampe = self.findChild(QtWidgets.QLineEdit, "ampe")
        self.btn_set_v = self.findChild(QtWidgets.QPushButton, "set_v")
        self.btn_set_a = self.findChild(QtWidgets.QPushButton, "set_a")
        self.log_widget = self.findChild(QtWidgets.QTextEdit, "logg")
        self.btn_clean = self.findChild(QtWidgets.QPushButton, "clean")
        self.btn_read = self.findChild(QtWidgets.QPushButton, "read")
        self.btn_remote = self.findChild(QtWidgets.QPushButton, "remote")

        # OVP / OCP
        self.edit_ovp = self.findChild(QtWidgets.QLineEdit, "ovp")
        self.edit_ocp = self.findChild(QtWidgets.QLineEdit, "ocp")
        self.btn_set_ovp = self.findChild(QtWidgets.QPushButton, "set_ovp")
        self.btn_set_ocp = self.findChild(QtWidgets.QPushButton, "set_ocp")
        self.chk_ovp_en = self.findChild(QtWidgets.QCheckBox, "ovp_en")
        self.chk_ocp_en = self.findChild(QtWidgets.QCheckBox, "ocp_en")
        self.btn_prot_status = self.findChild(QtWidgets.QPushButton, "prot_status")

        # Controller nguồn
        self.psw = PSWController()

        # trạng thái remote hiện tại (mặc định LOCAL)
        self.is_remote = False

        # ===== Kết nối signal (chỉ connect tay, KHÔNG dùng on_xxx_clicked) =====
        if self.btn_connect is not None:
            self.btn_connect.clicked.connect(self.handle_connect_clicked)
        if self.btn_test is not None:
            self.btn_test.clicked.connect(self.handle_test_clicked)
        if self.btn_on is not None:
            self.btn_on.clicked.connect(self.handle_on_clicked)
        if self.btn_off is not None:
            self.btn_off.clicked.connect(self.handle_off_clicked)
        if self.btn_set_v is not None:
            self.btn_set_v.clicked.connect(self.handle_set_v_clicked)
        if self.btn_set_a is not None:
            self.btn_set_a.clicked.connect(self.handle_set_a_clicked)
        if self.btn_clean is not None:
            self.btn_clean.clicked.connect(self.handle_clean_clicked)
        if self.btn_read is not None:
            self.btn_read.clicked.connect(self.handle_read_clicked)
        if self.btn_remote is not None:
            self.btn_remote.setCheckable(True)
            self.btn_remote.clicked.connect(self.handle_remote_clicked)

        # OVP/OCP signals
        if self.btn_set_ovp is not None:
            self.btn_set_ovp.clicked.connect(self.handle_set_ovp_clicked)
        if self.btn_set_ocp is not None:
            self.btn_set_ocp.clicked.connect(self.handle_set_ocp_clicked)
        if self.chk_ovp_en is not None:
            self.chk_ovp_en.toggled.connect(self.handle_ovp_enable_toggled)
        if self.chk_ocp_en is not None:
            self.chk_ocp_en.toggled.connect(self.handle_ocp_enable_toggled)
        if self.btn_prot_status is not None:
            self.btn_prot_status.clicked.connect(self.handle_prot_status_clicked)

        # ===== Khởi tạo =====
        self.populate_ports()

        if self.edit_volt is not None:
            self.edit_volt.setText("5.0")
            self.psw.target_v = 5.0
        if self.edit_ampe is not None:
            self.edit_ampe.setText("1.0")
            self.psw.target_i = 1.0

        if self.edit_ovp is not None:
            self.edit_ovp.setText("15.0")   # ví dụ
        if self.edit_ocp is not None:
            self.edit_ocp.setText("2.0")    # ví dụ

        self.set_on_button_color(False)
        self.update_remote_button()
        self.log("App started. Chọn cổng COM rồi nhấn 'connect'.")

    # ===== TIỆN ÍCH UI =====
    def log(self, text: str):
        """
        Thêm 1 dòng log vào QTextEdit logg.
        """
        if self.log_widget is not None:
            self.log_widget.append(text)
        else:
            print(text)

    def populate_ports(self):
        """
        Quét và đổ danh sách COM vào comboBox.
        """
        if self.combo_ports is None:
            return

        self.combo_ports.clear()
        ports = serial.tools.list_ports.comports()
        for p in ports:
            display = f"{p.device} - {p.description}"
            self.combo_ports.addItem(display, p.device)
        if not ports:
            self.log("Không tìm thấy cổng COM nào.")

    def get_selected_port(self):
        """
        Lấy tên cổng COM (COM3, COM5, ...) được chọn trong comboBox.
        """
        if self.combo_ports is None:
            return None
        idx = self.combo_ports.currentIndex()
        if idx < 0:
            return None
        return self.combo_ports.itemData(idx)

    def set_on_button_color(self, is_on: bool):
        """
        Đổi màu nút ON:
        - is_on = True  -> đỏ
        - is_on = False -> màu mặc định
        """
        if self.btn_on is None:
            return
        if is_on:
            self.btn_on.setStyleSheet("background-color: red; color: white;")
        else:
            self.btn_on.setStyleSheet("")

    def update_remote_button(self):
        """
        Cập nhật giao diện nút REMOTE/LOCAL theo self.is_remote.
        """
        if self.btn_remote is None:
            return

        if self.is_remote:
            self.btn_remote.setText("REMOTE")
            self.btn_remote.setStyleSheet("background-color: orange; color: black;")
            self.btn_remote.setChecked(True)
        else:
            self.btn_remote.setText("LOCAL")
            self.btn_remote.setStyleSheet("")
            self.btn_remote.setChecked(False)

    # ===== HANDLERS (slot) =====
    def handle_connect_clicked(self):
        """
        Nhấn connect:
        - Nếu chưa mở port → mở
        - Nếu đang mở → đóng
        """
        if self.psw.is_open():
            # đang mở -> đóng
            self.psw.close()
            if self.btn_connect is not None:
                self.btn_connect.setText("connect")
            self.set_on_button_color(False)
            self.is_remote = False
            self.update_remote_button()
            self.log("Đã ngắt kết nối.")
            return

        port = self.get_selected_port()
        if not port:
            self.log("Chưa chọn cổng COM.")
            return

        try:
            self.psw.open(port)
            self.log(f"Đã kết nối {port}.")
            if self.btn_connect is not None:
                self.btn_connect.setText("disconnect")
        except Exception as e:
            self.log(f"Lỗi mở {port}: {e}")

    def handle_test_clicked(self):
        """
        Nhấn test → gửi *IDN? xuống COM, hiển thị trả lời lên logg.
        """
        if not self.psw.is_open():
            self.log("Chưa kết nối. Nhấn 'connect' trước.")
            return

        try:
            idn = self.psw.get_idn()
            self.log(f"*IDN? → {idn}")
        except Exception as e:
            self.log(f"Lỗi gửi *IDN?: {e}")

    def handle_on_clicked(self):
        """
        Nhấn on → OUTPut:STATe ON
        """
        if not self.psw.is_open():
            self.log("Chưa kết nối. Nhấn 'connect' trước.")
            return
        try:
            self.psw.output_on()
            self.log("OUTPut:STATe ON")
            self.set_on_button_color(True)
        except Exception as e:
            self.log(f"Lỗi ON: {e}")

    def handle_off_clicked(self):
        """
        Nhấn off → OUTPut:STATe OFF
        """
        if not self.psw.is_open():
            self.log("Chưa kết nối. Nhấn 'connect' trước.")
            return
        try:
            self.psw.output_off()
            self.log("OUTPut:STATe OFF")
            self.set_on_button_color(False)
        except Exception as e:
            self.log(f"Lỗi OFF: {e}")

    def handle_set_v_clicked(self):
        """
        Nhấn set_v → đọc volt, gửi APPLy V, I
        """
        if not self.psw.is_open():
            self.log("Chưa kết nối. Nhấn 'connect' trước.")
            return
        if self.edit_volt is None:
            self.log("Không tìm thấy ô volt.")
            return

        try:
            v = float(self.edit_volt.text())
        except ValueError:
            self.log("Giá trị volt không hợp lệ.")
            return

        try:
            self.psw.set_voltage(v)
            self.log(f"Đã set V = {v:.4f} V (APPLy {v:.4f},{self.psw.target_i:.4f})")
        except Exception as e:
            self.log(f"Lỗi set_v: {e}")

    def handle_set_a_clicked(self):
        """
        Nhấn set_a → đọc ampe, gửi APPLy V, I
        """
        if not self.psw.is_open():
            self.log("Chưa kết nối. Nhấn 'connect' trước.")
            return
        if self.edit_ampe is None:
            self.log("Không tìm thấy ô ampe.")
            return

        try:
            i = float(self.edit_ampe.text())
        except ValueError:
            self.log("Giá trị ampe không hợp lệ.")
            return

        try:
            self.psw.set_current(i)
            self.log(f"Đã set I = {i:.4f} A (APPLy {self.psw.target_v:.4f},{i:.4f})")
        except Exception as e:
            self.log(f"Lỗi set_a: {e}")

    def handle_clean_clicked(self):
        """
        Nhấn clean → xóa sạch logg.
        """
        if self.log_widget is not None:
            self.log_widget.clear()

    def handle_read_clicked(self):
        """
        Nhấn read → gửi MEASure:SCALar:ALL? và hiển thị U/I đo được.
        """
        if not self.psw.is_open():
            self.log("Chưa kết nối. Nhấn 'connect' trước.")
            return

        try:
            v, i, raw = self.psw.measure_all()
            if v is None:
                self.log(f"MEASure:SCALar:ALL? → {raw}")
            else:
                self.log(
                    f"MEASure:SCALar:ALL? → V={v:.4f} V, I={i:.4f} A (raw: {raw})"
                )
        except Exception as e:
            self.log(f"Lỗi MEAS:SCALar:ALL?: {e}")

    def handle_remote_clicked(self):
        """
        Nhấn nút remote (toggle):
        - nếu đang LOCAL  -> chuyển sang REMOTE
        - nếu đang REMOTE -> chuyển sang LOCAL
        """
        if not self.psw.is_open():
            self.log("Chưa kết nối. Nhấn 'connect' trước.")
            # trả nút về trạng thái cũ
            self.update_remote_button()
            return

        target = not self.is_remote
        mode_str = "REMOTE" if target else "LOCAL"

        try:
            self.psw.set_remote(target)
            self.is_remote = target
            self.update_remote_button()
            self.log(f"SYSTem:{mode_str} mode")
        except Exception as e:
            self.log(f"Lỗi set {mode_str}: {e}")
            self.update_remote_button()

    # ===== OVP / OCP handlers =====
    def handle_set_ovp_clicked(self):
        """
        Nhấn set OVP -> đọc ovp (V) và gửi SOUR:VOLT:PROT:LEV.
        """
        if not self.psw.is_open():
            self.log("Chưa kết nối. Nhấn 'connect' trước.")
            return
        if self.edit_ovp is None:
            self.log("Không tìm thấy ô OVP.")
            return

        try:
            v_ovp = float(self.edit_ovp.text())
        except ValueError:
            self.log("Giá trị OVP không hợp lệ.")
            return

        try:
            self.psw.set_ovp(v_ovp)
            self.log(f"OVP level = {v_ovp:.4f} V")
        except Exception as e:
            self.log(f"Lỗi set OVP: {e}")

    def handle_set_ocp_clicked(self):
        """
        Nhấn set OCP -> đọc ocp (A) và gửi SOUR:CURR:PROT:LEV.
        """
        if not self.psw.is_open():
            self.log("Chưa kết nối. Nhấn 'connect' trước.")
            return
        if self.edit_ocp is None:
            self.log("Không tìm thấy ô OCP.")
            return

        try:
            i_ocp = float(self.edit_ocp.text())
        except ValueError:
            self.log("Giá trị OCP không hợp lệ.")
            return

        try:
            self.psw.set_ocp(i_ocp)
            self.log(f"OCP level = {i_ocp:.4f} A")
        except Exception as e:
            self.log(f"Lỗi set OCP: {e}")

    def handle_ovp_enable_toggled(self, checked: bool):
        """
        Bật/tắt OVP:
        - checked = True  -> dùng ngưỡng trong edit_ovp
        - checked = False -> set Level = MAX (coi như tắt bảo vệ áp)
        """
        if not self.psw.is_open():
            self.log("Chưa kết nối. Nhấn 'connect' trước.")
            # revert lại checkbox
            if self.chk_ovp_en is not None:
                self.chk_ovp_en.blockSignals(True)
                self.chk_ovp_en.setChecked(not checked)
                self.chk_ovp_en.blockSignals(False)
            return

        try:
            if checked:
                if self.edit_ovp is None:
                    self.log("Không tìm thấy ô OVP.")
                    return
                try:
                    v_ovp = float(self.edit_ovp.text())
                except ValueError:
                    self.log("Giá trị OVP không hợp lệ.")
                    return
                self.psw.set_ovp(v_ovp)
                self.log(f"OVP ENABLE, level = {v_ovp:.4f} V")
            else:
                self.psw.disable_ovp()
                self.log("OVP DISABLE (set level = MAX)")
        except Exception as e:
            self.log(f"Lỗi OVP enable/disable: {e}")

    def handle_ocp_enable_toggled(self, checked: bool):
        """
        Bật/tắt OCP:
        - checked = True  -> dùng ngưỡng trong edit_ocp
        - checked = False -> set Level = MAX
        """
        if not self.psw.is_open():
            self.log("Chưa kết nối. Nhấn 'connect' trước.")
            # revert lại checkbox
            if self.chk_ocp_en is not None:
                self.chk_ocp_en.blockSignals(True)
                self.chk_ocp_en.setChecked(not checked)
                self.chk_ocp_en.blockSignals(False)
            return

        try:
            if checked:
                if self.edit_ocp is None:
                    self.log("Không tìm thấy ô OCP.")
                    return
                try:
                    i_ocp = float(self.edit_ocp.text())
                except ValueError:
                    self.log("Giá trị OCP không hợp lệ.")
                    return
                self.psw.set_ocp(i_ocp)
                self.log(f"OCP ENABLE, level = {i_ocp:.4f} A")
            else:
                self.psw.disable_ocp()
                self.log("OCP DISABLE (set level = MAX)")
        except Exception as e:
            self.log(f"Lỗi OCP enable/disable: {e}")

    def handle_prot_status_clicked(self):
        """
        Nhấn PROT status -> đọc STAT:QUES:COND? và giải mã OVP/OCP.
        """
        if not self.psw.is_open():
            self.log("Chưa kết nối. Nhấn 'connect' trước.")
            return

        try:
            ov, oc, raw = self.psw.get_protection_status()
            msg_bits = []
            if ov:
                msg_bits.append("OVP TRIPPED")
            if oc:
                msg_bits.append("OCP TRIPPED")
            if not msg_bits:
                msg = "Không có OVP/OCP trip."
            else:
                msg = ", ".join(msg_bits)

            self.log(f"STAT:QUES:COND? → {raw}  → {msg}")
        except Exception as e:
            self.log(f"Lỗi đọc trạng thái PROT: {e}")


def main():
    app = QtWidgets.QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
