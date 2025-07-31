import sys, psutil, time, threading, csv, os
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel,
    QFileDialog, QHBoxLayout, QDoubleSpinBox, QCheckBox,
    QTableWidget, QTableWidgetItem, QSplitter, QHeaderView
)
from PyQt5.QtCore import Qt
from openpyxl import Workbook
from matplotlib.backends.backend_qt5agg import (
    FigureCanvasQTAgg as FigureCanvas,
    NavigationToolbar2QT as NavigationToolbar
)
from matplotlib.figure import Figure

class PlotCanvas(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.figure = Figure()
        self.ax = self.figure.add_subplot(111)
        self.canvas = FigureCanvas(self.figure)
        self.toolbar = NavigationToolbar(self.canvas, self)

        layout = QVBoxLayout()
        layout.addWidget(self.toolbar)
        layout.addWidget(self.canvas)
        self.setLayout(layout)
        self.ax.set_title("CPU and RAM Usage Over Time")
        self.ax.set_xlabel("Time")
        self.ax.set_ylabel("Usage")
        self.figure.tight_layout()

    def plot(self, timestamps, cpu_vals, ram_vals):
        self.ax.clear()
        self.ax.plot(timestamps, cpu_vals, '-o', label='CPU (%)')
        self.ax.plot(timestamps, ram_vals, '-o', label='RAM (MB)')
        self.ax.legend()
        self.ax.set_xlabel("Time")
        self.ax.set_ylabel("Usage")
        self.ax.set_title("CPU and RAM Usage Over Time")
        self.ax.grid(True)
        self.figure.autofmt_xdate()
        self.canvas.draw()

class MonitorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CPU/RAM Monitor by psutil")
        self.resize(1100, 600)

        self.monitoring = False
        self.training_source = "Manual"
        self.training_pid = None
        self.data = []
        self.buffered_data = []
        self.sampling_rate = 1.0
        self.training_start_time = None
        self.last_update_time = time.time()
        self.update_interval = 2
        self.initial_buffer_flushed = False # เพิ่มตัวแปรสถานะ

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Time", "CPU (%)", "RAM (MB)", "Source"])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)

        self.status_label = QLabel("Status: Idle")
        self.source_label = QLabel("")

        self.sampling_spinbox = QDoubleSpinBox()
        self.sampling_spinbox.setRange(0.1, 10.0)
        self.sampling_spinbox.setValue(1.0)

        self.auto_start_checkbox = QCheckBox("Auto Start When Training Detected")
        self.plot_mode_checkbox = QCheckBox("Plot only after training finished")
        self.buffer_mode_checkbox = QCheckBox("Use sampling-based update (tick = sampling rate, untick = buffered)")
        self.buffer_mode_checkbox.setChecked(False)

        self.btn_reset = QPushButton("Reset Table")
        self.btn_export_excel = QPushButton("Export to Excel")
        self.btn_export_csv = QPushButton("Export to CSV")
        self.btn_save_graph = QPushButton("Save Graph")
        self.btn_exit = QPushButton("Exit")

        self.btn_reset.clicked.connect(self.reset_table)
        self.btn_export_excel.clicked.connect(self.export_excel)
        self.btn_export_csv.clicked.connect(self.export_csv)
        self.btn_save_graph.clicked.connect(self.save_graph)
        self.btn_exit.clicked.connect(self.close)

        self.graph = PlotCanvas(self)
        self.setup_ui()
        threading.Thread(target=self.monitor_loop, daemon=True).start()

    def setup_ui(self):
        layout = QVBoxLayout()
        control_layout = QHBoxLayout()

        control_layout.addWidget(QLabel("Sampling Rate (s):"))
        control_layout.addWidget(self.sampling_spinbox)
        control_layout.addWidget(self.btn_reset)
        control_layout.addWidget(self.btn_export_excel)
        control_layout.addWidget(self.btn_export_csv)
        control_layout.addWidget(self.btn_save_graph)
        control_layout.addWidget(self.btn_exit)

        layout.addWidget(self.status_label)
        layout.addWidget(self.source_label)
        layout.addWidget(self.auto_start_checkbox)
        layout.addWidget(self.plot_mode_checkbox)
        layout.addWidget(self.buffer_mode_checkbox)

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self.table)
        splitter.addWidget(self.graph)
        layout.addWidget(splitter)
        layout.addLayout(control_layout)
        self.setLayout(layout)

    def reset_table(self):
        self.data.clear()
        self.buffered_data.clear()
        self.table.setRowCount(0)
        self.graph.ax.clear()
        self.graph.canvas.draw()
        self.status_label.setText("Table reset.")
        self.source_label.setText("")

    def detect_training_process(self):
        try:
            with open("C:\\temp\\training_pid.txt", "r") as f:
                pid = int(f.read().strip())
                proc = psutil.Process(pid)
                if proc.is_running():
                    cmd = ' '.join(proc.cmdline())
                    self.training_source = f"MATLAB (PID: {pid}) CMD: {cmd}"
                    self.training_pid = pid
                    return True
        except:
            pass

        my_pid = psutil.Process().pid
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                if proc.pid == my_pid:
                    continue
                name = proc.info['name'].lower()
                cmdline = ' '.join(proc.info.get('cmdline', [])).lower()
                if "python" in name and ".py" in cmdline:
                    self.training_source = f"Python: {cmdline}"
                    self.training_pid = proc.pid
                    return True
            except:
                continue
        return False

    def get_training_process_resource(self):
        try:
            if self.training_pid:
                proc = psutil.Process(self.training_pid)
                proc.cpu_percent(interval=None)
                cpu = proc.cpu_percent(interval=self.sampling_rate) / psutil.cpu_count()
                ram = proc.memory_info().rss / (1024 * 1024)
                return cpu, ram
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            self.finish_monitoring()
        except Exception:
            pass
        return None, None

    def flush_buffer_to_table_and_graph(self):
        if not self.buffered_data:
            return
        for rowdata in self.buffered_data:
            row = self.table.rowCount()
            self.table.insertRow(row)
            for i, val in enumerate(rowdata):
                if i == 1 or i == 2:
                    self.table.setItem(row, i, QTableWidgetItem(f"{val:.2f}"))
                else:
                    self.table.setItem(row, i, QTableWidgetItem(str(val)))

        self.data.extend(self.buffered_data)

        if not self.plot_mode_checkbox.isChecked():
            timestamps = [d[0] for d in self.data]
            cpu_vals = [d[1] for d in self.data]
            ram_vals = [d[2] for d in self.data]
            self.graph.plot(timestamps, cpu_vals, ram_vals)

        self.buffered_data.clear()

    def get_dynamic_update_interval(self, elapsed_seconds):
        """คำนวณช่วงเวลาการแสดงผลแบบไดนามิกตามเงื่อนไขใหม่"""
        if elapsed_seconds <= 10: return 10
        if elapsed_seconds <= 20: return 2
        if elapsed_seconds <= 60: return 5
        if elapsed_seconds <= 300: return 10
        if elapsed_seconds <= 900: return 20
        if elapsed_seconds <= 3600: return 30
        if elapsed_seconds <= 10800: return 60
        if elapsed_seconds <= 21600: return 120
        if elapsed_seconds <= 43200: return 300
        if elapsed_seconds <= 86400: return 600
        return 1800

    def monitor_loop(self):
        while True:
            if not self.monitoring and self.auto_start_checkbox.isChecked():
                if self.detect_training_process():
                    self.start_monitoring()

            if self.monitoring:
                if 'matlab' in self.training_source.lower():
                    if not os.path.exists("C:\\temp\\training_pid.txt"):
                        self.finish_monitoring()
                        continue

                if not psutil.pid_exists(self.training_pid):
                    self.finish_monitoring()
                    continue

                self.sampling_rate = self.sampling_spinbox.value()
                cpu, ram = self.get_training_process_resource()

                if cpu is not None and ram is not None:
                    timestamp = datetime.now().strftime("%H:%M:%S")
                    self.buffered_data.append((timestamp, cpu, ram, self.training_source))
                    
                    is_sampling_mode = self.buffer_mode_checkbox.isChecked()

                    if is_sampling_mode:
                        self.flush_buffer_to_table_and_graph()
                        self.last_update_time = time.time()
                    else:
                        elapsed = time.time() - self.training_start_time
                        
                        if not self.initial_buffer_flushed and elapsed >= 10:
                            self.flush_buffer_to_table_and_graph()
                            self.last_update_time = time.time()
                            self.initial_buffer_flushed = True
                        
                        elif self.initial_buffer_flushed:
                            self.update_interval = self.get_dynamic_update_interval(elapsed)
                            if time.time() - self.last_update_time >= self.update_interval:
                                self.flush_buffer_to_table_and_graph()
                                self.last_update_time = time.time()
            else:
                time.sleep(0.3)

    def finish_monitoring(self):
        self.monitoring = False
        self.status_label.setText("Training stopped. Showing result...")
        self.flush_buffer_to_table_and_graph()
        if self.plot_mode_checkbox.isChecked():
            timestamps = [d[0] for d in self.data]
            cpu_vals = [d[1] for d in self.data]
            ram_vals = [d[2] for d in self.data]
            self.graph.plot(timestamps, cpu_vals, ram_vals)
        self.source_label.setText(f"Detected from: {self.training_source}")

    def start_monitoring(self):
        self.sampling_rate = self.sampling_spinbox.value()
        self.monitoring = True
        self.buffered_data.clear()
        self.data.clear()
        self.table.setRowCount(0)
        self.graph.ax.clear()
        self.graph.canvas.draw()
        self.training_start_time = time.time()
        self.last_update_time = time.time()
        self.initial_buffer_flushed = False # รีเซ็ตตัวแปรสถานะ
        self.status_label.setText("Monitoring started (Auto).")
        self.source_label.setText(f"Detected from: {self.training_source}")

    def export_excel(self):
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
            ws.append(["", "", "", f"Command/Source: {self.training_source}"])
            wb.save(path)
            self.status_label.setText(f"Excel saved to {path}")

    def export_csv(self):
        if not self.data:
            self.status_label.setText("Status: No data to export")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Save CSV File", "", "CSV Files (*.csv)")
        if path:
            with open(path, mode='w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(["Time", "CPU (%)", "RAM (MB)", "Source"])
                writer.writerows(self.data)
                writer.writerow(["", "", "", f"Command/Source: {self.training_source}"])
            self.status_label.setText(f"CSV saved to {path}")

    def save_graph(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save Graph as Image", "", "PNG Files (*.png)")
        if path:
            self.graph.figure.savefig(path)
            self.status_label.setText(f"Graph saved to {path}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MonitorApp()
    win.show()
    sys.exit(app.exec_())