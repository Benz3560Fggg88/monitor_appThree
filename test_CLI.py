import time
import psutil
import csv
import os
from openpyxl import Workbook
from datetime import datetime

def get_pid():
    """
    ตรวจหา PID ของโปรเซสที่กำลังเทรน
    - ตรวจหาจากไฟล์ C:\\temp\\training_pid.txt สำหรับ MATLAB ก่อน
    - หากไม่เจอ จะค้นหาโปรเซส Python ที่กำลังรันไฟล์ .py
    """
    # --- ตรวจสอบ MATLAB ก่อน ---
    pid_file_path = "C:\\temp\\training_pid.txt"
    try:
        if os.path.exists(pid_file_path):
            with open(pid_file_path, "r") as f:
                pid = int(f.read().strip())
            proc = psutil.Process(pid)
            if proc.is_running() and "matlab" in proc.name().lower():
                return pid, f"MATLAB (PID: {pid}) CMD: {' '.join(proc.cmdline())}"
    except (FileNotFoundError, psutil.NoSuchProcess, ValueError, psutil.AccessDenied):
        pass # หากมีปัญหา ให้ข้ามไปหา Python

    # --- หากไม่เจอ MATLAB ให้ตรวจสอบ Python ---
    current_pid = psutil.Process().pid
    for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            if proc.pid == current_pid:
                continue

            name = proc.info['name'].lower()
            cmdline_list = proc.info.get('cmdline')
            cmdline = ' '.join(cmdline_list or []).lower()
            
            if ("python" in name or "python.exe" in name) and ".py" in cmdline:
                return proc.pid, f"Python: {cmdline}"
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
            
    return None, None

def get_update_interval(elapsed):
    """คำนวณช่วงเวลาการแสดงผลแบบ Buffered ตามเวลาที่ผ่านไป"""
    if elapsed < 10: return 10
    elif elapsed <= 60: return 5
    elif elapsed <= 300: return 10
    elif elapsed <= 900: return 20
    elif elapsed <= 3600: return 30
    else: return 60

def monitor(samrate, display_mode):
    """
    ฟังก์ชันหลักสำหรับติดตามและบันทึกข้อมูล CPU/RAM
    """
    print("🔍 Waiting for training process...")
    pid_file_path = "C:\\temp\\training_pid.txt"
    
    while True:
        pid, source = get_pid()
        if pid:
            break
        time.sleep(1)

    print(f"\n✅ Detected training from: {source}")
    print(f"{'Time':<10} {'CPU (%)':<10} {'RAM (MB)':<12} Source")

    training_start = time.time()
    last_display_time = training_start
    data, buffer, samples = [], [], []
    is_matlab = "matlab" in source.lower()

    while True:
        # --- เงื่อนไขการหยุด Monitor ---
        # 1. (สำหรับ MATLAB) ตรวจสอบว่าไฟล์ PID ถูกลบไปหรือยัง (สัญญาณที่ชัดเจนที่สุด)
        if is_matlab and not os.path.exists(pid_file_path):
           # print("\nℹ️ MATLAB PID file deleted. Task is complete.")                  ลบได้  --------------------------#
            break

        # 2. ตรวจสอบว่าโปรเซสหายไปจากระบบหรือไม่ (สำหรับ Python หรือกรณี MATLAB ปิดตัวเอง)
        if not psutil.pid_exists(pid):
            print("\nℹ️ Process PID not found. Stopping.")
            break
        
        # --- เก็บข้อมูล CPU/RAM ---
        try:
            proc = psutil.Process(pid)
            proc.cpu_percent(interval=None) # เรียกครั้งแรกเพื่อเริ่มต้น
            time.sleep(0.1)
            cpu = proc.cpu_percent(interval=None) / psutil.cpu_count()
            ram = proc.memory_info().rss / (1024 * 1024)
        except psutil.NoSuchProcess:
            break # ออกจากลูปหากโปรเซสหายไประหว่างทำงาน

        # --- ประมวลผลและแสดงข้อมูล ---
        samples.append((cpu, ram))
        if (len(samples) * 0.1) >= samrate:
            avg_cpu = sum(x[0] for x in samples) / len(samples) if samples else 0
            avg_ram = sum(x[1] for x in samples) / len(samples) if samples else 0
            samples.clear()
            timestamp = datetime.now().strftime("%H:%M:%S")
            row = (timestamp, avg_cpu, avg_ram, source)

            if display_mode == 1:
                print(f"{timestamp:<10} {avg_cpu:<10.2f} {avg_ram:<12.2f} {source}")
                data.append(row)
            else:
                buffer.append(row)
                if time.time() - last_display_time >= get_update_interval(time.time() - training_start):
                    for b in buffer:
                        print(f"{b[0]:<10} {b[1]:<10.2f} {b[2]:<12.2f} {b[3]}")
                    data.extend(buffer)
                    buffer.clear()
                    last_display_time = time.time()

    if display_mode == 2 and buffer:
        for b in buffer:
            print(f"{b[0]:<10} {b[1]:<10.2f} {b[2]:<12.2f} {b[3]}")
        data.extend(buffer)

    print("\n⏹️ Training stopped.")
    return data, source

def export_excel(data, source):
    """ส่งออกข้อมูลเป็นไฟล์ Excel"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Monitoring_Log"
    ws.append(["Time", "CPU (%)", "RAM (MB)", "Source"])
    for row in data:
        ws.append(row)
    ws.append([])
    ws.append(["Command/Source:", source])
    filename = f"monitor_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    try:
        wb.save(filename)
        print(f"📁 Saved Excel to {os.path.abspath(filename)}")
    except Exception as e:
        print(f"❌ Error saving Excel file: {e}")


def export_csv(data, source):
    """ส่งออกข้อมูลเป็นไฟล์ CSV"""
    filename = f"monitor_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    try:
        with open(filename, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(["Time", "CPU (%)", "RAM (MB)", "Source"])
            writer.writerows(data)
            writer.writerow([])
            writer.writerow(["Command/Source:", source])
        print(f"📁 Saved CSV to {os.path.abspath(filename)}")
    except Exception as e:
        print(f"❌ Error saving CSV file: {e}")


def main():
    """ฟังก์ชันหลักในการควบคุมโปรแกรม"""
    while True: # Loop สำหรับ "Restart from beginning"
        # --- 1. รับค่า Sampling Rate ---
        while True:
            try:
                s_input = input("⏱️ Set sampling rate (0.1–10.0) sec (recommended: 1.0): ")
                s = float(s_input)
                if 0.1 <= s <= 10.0: break
                else: print("❌ Invalid range. Try again.")
            except ValueError:
                print("❌ Invalid input. Try again.")

        # --- 2. Loop สำหรับเลือก Display Mode และ Action ---
        display_mode_loop = True
        while display_mode_loop:
            print("\n📺 Select display mode:")
            print("1. Real-time display")
            print("2. Buffered display")
            print("3. Back to sampling rate")
            m = input("Choice: ").strip()

            if m == '3':
                display_mode_loop = False # ออกจากลูปนี้เพื่อกลับไปถาม sampling rate
                continue

            if m in ['1', '2']:
                mode = int(m)
                # --- [หน้าจอใหม่] ให้เลือกว่าจะเริ่มหรือจะย้อนกลับ ---
                while True:
                    print("\n▶️ Select action:")
                    print("1. Wait for training detection")
                    print("2. Back to display mode selection")
                    action = input("Choice: ").strip()

                    if action == '2':
                        break # ย้อนกลับไปหน้า Select display mode

                    if action == '1':
                        # --- เริ่ม Monitor และจัดการผลลัพธ์ ---
                        records, source = monitor(s, mode)
                        
                        # --- เมนูหลังจบการ Monitor ---
                        while True:
                            print("\n✅ Monitoring finished. What next?")
                            print("1. Wait for new training")
                            print("2. Export to Excel")
                            print("3. Export to CSV")
                            print("4. Restart from beginning")
                            print("5. Exit")
                            post = input("Choice: ").strip()

                            if post == '1':
                                print("\n" + "-"*40 + "\n")
                                # กลับไปรัน monitor ใหม่ โดยใช้ค่า s และ mode เดิม
                                records, source = monitor(s, mode)
                                continue
                            elif post == '2':
                                export_excel(records, source)
                            elif post == '3':
                                export_csv(records, source)
                            elif post == '4':
                                # ออกจากทุก Loop เพื่อไปเริ่มใหม่ทั้งหมด
                                display_mode_loop = False
                                break
                            elif post == '5':
                                print("👋 Exiting...")
                                return
                            else:
                                print("❌ Invalid choice.")
                        
                        if not display_mode_loop:
                            break # ออกจาก action loop ถ้าเลือก restart
                    else:
                        print("❌ Invalid choice.")
                
                if not display_mode_loop:
                    break # ออกจาก display mode loop ถ้าเลือก restart

            else:
                print("❌ Invalid choice.")
        
        print("\n" + "="*40 + "\n") # พิมพ์เส้นคั่นเมื่อ Restart
if __name__ == "__main__":
    main()