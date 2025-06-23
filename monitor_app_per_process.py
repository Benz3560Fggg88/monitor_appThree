import sys, psutil, time, threading, csv, os
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel,
    QFileDialog, QHBoxLayout, QDoubleSpinBox, QCheckBox,
    QTableWidget, QTableWidgetItem
)
from PyQt5.QtGui import QPainter, QPen, QFont, QColor
from PyQt5.QtCore import Qt, QRectF
from openpyxl import Workbook

# หน้าครึ่งวงกลมแสดงค่า %/MB หรือค่าที่แปลงแล้วของ CPU/RAM
class HalfCircleGauge(QWidget):
    def __init__(self, label="CPU", show_label=True):
        super().__init__()
        self.value = 0
        self.label = label #CPU / RAM
        self.setMinimumSize(200, 150) #ขนาด widget
        self.show_label = show_label

    def setValue(self, val):
        self.value = val
        self.update()  #อัพเดตค่า

    def paintEvent(self, event):   # ครึ่งวงกลม
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        size = min(self.width(), self.height()) - 30 #40
        left = (self.width() - size) / 2
        top = 40 #70
        rect = QRectF(left, top, size, size)

        angle_span = int(180 * 16 * (self.value / 100))

        #พื้นหลัง
        back_pen = QPen(QColor(220, 220, 220), 30)
        back_pen.setCapStyle(Qt.RoundCap)
        painter.setPen(back_pen)
        painter.drawArc(rect, 180 * 16, -180 * 16)
        
        #เเสดง
        color = QColor(200, 0, 0) if self.label == "CPU" else QColor(0, 150, 0)
        arc_pen = QPen(color, 30)
        arc_pen.setCapStyle(Qt.RoundCap)
        painter.setPen(arc_pen)
        painter.drawArc(rect, 180 * 16, -angle_span)


class MonitorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CPU/RAM Monitor by psutil")
        self.resize(800, 600)

        #ตัวแปรควบคุมสถานะการ monitor
        self.monitoring = False
        self.training_source = "Manual"
        self.training_pid = None
        self.data = []
        self.sampling_rate = 1.0
        
        #ตัวเเสดงค่าที่ได้ 
        self.cpu_gauge = HalfCircleGauge("CPU")
        self.ram_gauge = HalfCircleGauge("RAM")
        
        self.cpu_label = QLabel("CPU (%): 0.0 %")
        self.ram_label = QLabel("RAM (MB): 0 MB")

        self.cpu_label.setAlignment(Qt.AlignCenter)
        self.cpu_label.setFont(QFont("Arial", 11, QFont.Bold))
        self.ram_label.setAlignment(Qt.AlignCenter)
        self.ram_label.setFont(QFont("Arial", 11, QFont.Bold))
        # ตารางแสดงข้อมูล
        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["Time", "CPU (%)", "RAM (MB)"])
        # widget UI อื่น ๆ
        self.status_label = QLabel("Status: Idle")
        self.sampling_spinbox = QDoubleSpinBox()
        self.sampling_spinbox.setRange(0.1, 10.0)
        self.sampling_spinbox.setValue(1.0)

        self.auto_start_checkbox = QCheckBox("Auto Start When Training Detected")

        self.btn_reset = QPushButton("Reset Table")
        self.btn_export_excel = QPushButton("Export to Excel")
        self.btn_export_csv = QPushButton("Export to CSV")
        self.btn_exit = QPushButton("Exit")
        # เชื่อมปุ่มเข้ากับฟังก์ชัน
        self.btn_reset.clicked.connect(self.reset_table)
        self.btn_export_excel.clicked.connect(self.export_excel)
        self.btn_export_csv.clicked.connect(self.export_csv)
        self.btn_exit.clicked.connect(self.close)
        # ตั้งค่า layout
        self.setup_ui()
        threading.Thread(target=self.monitor_loop, daemon=True).start()  # สร้าง thread สำหรับ monitor loop

    def setup_ui(self):  #layout ของหน้าจอ
        layout = QVBoxLayout()
        gauge_layout = QHBoxLayout()
        # ฝั่งซ้าย: CPU
        left = QVBoxLayout()
        left.addWidget(self.cpu_label)
        left.addWidget(self.cpu_gauge)
        # ฝั่งขวา: RAM
        right = QVBoxLayout()
        right.addWidget(self.ram_label)
        right.addWidget(self.ram_gauge)

        gauge_layout.addLayout(left)
        gauge_layout.addLayout(right)
        # Sampling Rate
        control_layout = QHBoxLayout()
        control_layout.addWidget(QLabel("Sampling Rate (s):"))
        control_layout.addWidget(self.sampling_spinbox)
        control_layout.addWidget(self.btn_reset)
        control_layout.addWidget(self.btn_export_excel)
        control_layout.addWidget(self.btn_export_csv)
        control_layout.addWidget(self.btn_exit)
        #layout หลัก
        layout.addWidget(self.status_label)
        layout.addWidget(self.auto_start_checkbox)
        layout.addLayout(gauge_layout)
        layout.addWidget(self.table)
        layout.addLayout(control_layout)
        self.setLayout(layout)
    
    def reset_table(self):  # reset ตาราง
        self.data.clear() 
        self.table.setRowCount(0)
        self.status_label.setText("Table reset.")

    def detect_flag_file(self):  # ตรวจสอบว่า MATLAB สร้างไฟล์ flag หรือเปล่า
        return os.path.exists("C:\\temp\\monitoring_flag.txt")

    def detect_training_process(self):   # ตรวจจับ process ที่ train 
        if self.detect_flag_file():
            for proc in psutil.process_iter(['pid', 'name', 'memory_info']):
                try:
                    if "matlab" in proc.info['name'].lower():
                        if proc.memory_info().rss > 200 * 1024 * 1024:
                            self.training_source = "MATLAB (via flag)"
                            self.training_pid = proc.pid
                            return True
                except:
                    continue

        my_pid = psutil.Process().pid     # ตรวจหา process ที่เป็น Python script
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                if proc.pid == my_pid:
                    continue
                name = proc.info['name'].lower()
                cmdline = ' '.join(proc.info.get('cmdline', [])).lower()
                if "python" in name and ".py" in cmdline:
                    self.training_source = "Python"
                    self.training_pid = proc.pid
                    return True
            except:
                continue
        return False

    def get_training_process_resource(self): # อ่านค่า CPU/RAM ของ process ที่จับได้
        try:
            if self.training_pid:
                proc = psutil.Process(self.training_pid)
                proc.cpu_percent(interval=None)
                time.sleep(self.sampling_rate)
                cpu = proc.cpu_percent(interval=None) / psutil.cpu_count()
                ram = proc.memory_info().rss / (1024 * 1024)
                ram_percent = (proc.memory_info().rss / psutil.virtual_memory().total) * 100
                return cpu, ram, ram_percent
        except:
            pass
        return None, None, None

    def monitor_loop(self):  # วนลูปตรวจจับข้อมูลทุก sampling rate
        while True:
            if not self.monitoring and self.auto_start_checkbox.isChecked():
                if self.detect_training_process():
                    self.start_monitoring()

            if self.monitoring:
                if self.auto_start_checkbox.isChecked() and not self.detect_training_process():
                    self.monitoring = False
                    self.status_label.setText("Training stopped. Waiting...")
                    continue

                self.sampling_rate = self.sampling_spinbox.value()
                cpu, ram, ram_percent = self.get_training_process_resource()
                if cpu is None:
                    continue
                
                # อัพเดตขอมูล
                timestamp = datetime.now().strftime("%H:%M:%S")
                self.cpu_gauge.setValue(cpu)
                self.ram_gauge.setValue(ram_percent)
                self.cpu_label.setText(f"CPU (%): {cpu:.1f} %")
                self.ram_label.setText(f"RAM (MB): {int(ram)} MB")
                # เก็บข้อมูล
                self.data.append((timestamp, cpu, ram, self.training_source))
                row = self.table.rowCount()
                self.table.insertRow(row)
                self.table.setItem(row, 0, QTableWidgetItem(timestamp))
                self.table.setItem(row, 1, QTableWidgetItem(f"{cpu:.1f}"))
                self.table.setItem(row, 2, QTableWidgetItem(f"{ram:.1f}"))

                self.status_label.setText(
                    f"{self.training_source}: {timestamp} CPU: {cpu:.1f}% RAM: {ram:.1f} MB"
                )
            else:
                time.sleep(0.3)

    def start_monitoring(self): # เริ่มการ monitor
        self.sampling_rate = self.sampling_spinbox.value()
        self.monitoring = True
        self.data.clear()
        self.cpu_gauge.setValue(0)
        self.ram_gauge.setValue(0)
        self.table.setRowCount(0)
        self.status_label.setText("Monitoring started (Auto).")

    def export_excel(self): # บันทึกข้อมูลเป็น Excel
        if not self.data:
            self.status_label.setText("Status: No data to export")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx)")
        if path:
            wb = Workbook()
            ws = wb.active
            ws.append(["Time", "CPU (%)", "RAM (MB)", "Source"])
            for row in self.data:
                ws.append(row)
            wb.save(path)
            self.status_label.setText(f"Excel saved to {path}")

    def export_csv(self): # บันทึกข้อมูลเป็น CSV
        if not self.data:
            self.status_label.setText("Status: No data to export")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Save CSV File", "", "CSV Files (*.csv)")
        if path:
            with open(path, mode='w', newline='') as file:
                writer = csv.writer(file)
                writer.writerow(["Time", "CPU (%)", "RAM (MB)", "Source"])
                writer.writerows(self.data)
            self.status_label.setText(f"CSV saved to {path}")

# ตัวรัน
if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MonitorApp()
    win.show()
    sys.exit(app.exec_())
