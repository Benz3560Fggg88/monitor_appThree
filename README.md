# monitor_appThree
# ถ้าจะจับMATLAB อย่าลืมสร้างไฟล์ temp ที่ drive c เเละใน code MATLAB ต้องมี 
% ---------- เริ่มตรวจจับ ----------
fid = fopen('C:\temp\monitoring_flag.txt', 'w'); fclose(fid);
% -----------------------------------
.
.
.
end

% ----------- จบการตรวจจับ ----------
pause(1);  % เผื่อเวลาให้ Python จับได้
delete('C:\temp\monitoring_flag.txt');
% ------------------------------------
