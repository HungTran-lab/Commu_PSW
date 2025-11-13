import sys
import serial
import serial.tools.list_ports

from PyQt5 import QtWidgets, uic


class PSWController:
    """
    Điều khiển GW Instek PSW qua cổng COM (USB-CDC).
    """
    def __init__(self):
        self.ser = None
        self.target_v = 0.0
        self.target_i = 0.0

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

    # ===== API tiện dùng =====
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


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi("gui.ui", self)   # đảm bảo gui.ui nằm cùng thư mục

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

        # Controller nguồn
        self.psw = PSWController()

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

        # ===== Khởi tạo =====
        self.populate_ports()
        if self.edit_volt is not None:
            self.edit_volt.setText("5.0")
            self.psw.target_v = 5.0
        if self.edit_ampe is not None:
            self.edit_ampe.setText("1.0")
            self.psw.target_i = 1.0

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

    # ===== HANDLERS (slot) =====
    def handle_connect_clicked(self):
        """
        Nhấn connect:
        - Nếu chưa mở port → mở
        - Nếu đang mở → đóng
        """
        if self.psw.is_open():
            self.psw.close()
            if self.btn_connect is not None:
                self.btn_connect.setText("connect")
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
                # parse lỗi, log raw để xem PSW trả về gì
                self.log(f"MEASure:SCALar:ALL? → {raw}")
            else:
                self.log(
                    f"MEASure:SCALar:ALL? → V={v:.4f} V, I={i:.4f} A (raw: {raw})"
                )
        except Exception as e:
            self.log(f"Lỗi MEAS:SCALar:ALL?: {e}")


def main():
    app = QtWidgets.QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
