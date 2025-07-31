# monitor_appThree
# pip install PyQt5 psutil matplotlib openpyxl
# ถ้าจะจับMATLAB อย่าลืมสร้างไฟล์ temp ที่ drive c เเละใน code MATLAB ต้องมี 
% ---------- เริ่มตรวจจับ ---------- 
pid = feature('getpid');  % ดึง PID ของ MATLAB เอง
fid = fopen('C:\temp\training_pid.txt', 'w');
if fid == -1
    error('ไม่สามารถเปิดไฟล์ C:\temp\training_pid.txt เพื่อเขียนได้');
end
fprintf(fid, '%d\n', pid);
fclose(fid);
% ----------------------------------- 
.
.
.
end

% ---------- จบการตรวจจับ ----------
pause(1);  % รอให้ Python ตรวจจับให้ทัน
if exist('C:\temp\training_pid.txt', 'file')
    delete('C:\temp\training_pid.txt');
    fprintf('ลบไฟล์ PID เรียบร้อย\n');
end
% ------------------------------------



!!!! การตรวจจับโดยให้เเสดงผลทุก0.1ยังมีปัญหา
!!!! monitor_app_per_process.py ---> old version
