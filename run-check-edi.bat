@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ===== Check EDI - ENTER ชุดรวมหลายฟอร์ม (Various forms) =====
echo วางไฟล์ MANIFEST / MF (.xls, .xlsx หรือ .pdf) และ ENTER (.pdf) ไว้ในโฟลเดอร์ input ก่อน
echo.

rem หา Python ที่รันได้จริง (python.exe ของ Microsoft Store บางเครื่องเป็นแค่ทางลัดที่รันไม่ได้)
set PY=
py -3 -c "import sys" >nul 2>nul && set PY=py -3
if not defined PY python -c "import sys" >nul 2>nul && set PY=python
if not defined PY (
  echo !! ไม่พบ Python - ติดตั้งจาก https://www.python.org/downloads/ แล้วติ๊ก "Add python.exe to PATH"
  pause
  exit /b 1
)

%PY% -c "import openpyxl, xlrd, pymupdf" >nul 2>nul
if errorlevel 1 (
  echo กำลังติดตั้งไลบรารีที่ต้องใช้ ครั้งแรกครั้งเดียว...
  %PY% -m pip install -r requirements.txt
)

rem ชื่อรายงานใส่วันเวลา ไม่เขียนทับไฟล์เดิม
for /f %%i in ('%PY% -c "import datetime;print(datetime.datetime.now().strftime('%%Y%%m%%d-%%H%%M'))"') do set STAMP=%%i
set NAME=EDI-%STAMP%.xlsx
if not exist output mkdir output

%PY% scripts\bundle_edi.py --indir . --outdir output --name %NAME%
echo.
echo รายงานอยู่ที่ output\%NAME%
pause
