# Check EDI — ENTER ชุดรวมหลายฟอร์ม (Various forms)

โปรแกรมตรวจเอกสารขาเข้าเรือก่อนทำใบขนสินค้าขาเข้า สำหรับ **ENTER ที่เป็น PDF ไฟล์เดียวรวมหลายฟอร์ม**
(จดหมายแก้ไขของ forwarder หลายเจ้า + ARRIVAL NOTICE / สำเนา B/L + หน้าสแกน) เทียบกับ **MF (MANIFEST)**
แล้วออกรายงาน **`EDI.xlsx`** แบบใส่สี ทีละ B/L

โปรแกรมแยก ENTER ทีละหน้า → รู้ว่าเป็นฟอร์มไหน → อ่านด้วยตัวอ่านของฟอร์มนั้น → เทียบกับ MF

- CONSIGNEE, CONTAINER NO., STATUS (CY / LCL / LCL-CFS)
- TOTAL PACKAGE + ชนิดหีบห่อ (WOODEN CASE ≠ CASE), GROSS WEIGHT, MEASUREMENT (แจ้ง MF = 0.000)
- MARKS, DESCRIPTION (ก่อน `AS PER ATTACHED SHEET` ตรวจ 30 ตัวอักษรแรก ; กรอบสีแดงของลูกค้า = ตรวจเฉพาะในกรอบ)
- REEFER TEMP, DG (CLASS/UN), TRANSIT
- SHED NO. (ทุก B/L ต้องตรงเลขที่แจ้ง ยกเว้น LAOS 0124 หรือแยก LCL/CFS กับ CY), ประเทศปลายทาง (ลาว = LA), CARGO MOVEMENT
- ลงผล **เฉพาะ B/L ที่มีใน ENTER** — B/L ที่มีใน ENTER แต่ไม่มีใน MF ขึ้นแถววิกฤต
- B/L แตก (เช่น …664, …664A–L) เรียงต่อกัน ; หมายเหตุทุกแถวบอกหน้า ENTER `[ENTER หน้า n]`

**สี:** 🟥 แดง = ไม่ตรง · 🟧 ส้ม = ต้องให้คนตรวจยืนยัน (รวมหน้าที่อ่านด้วย OCR) · 🟨 เหลือง = ไม่มี ENTER ให้เทียบ · 🟩 เขียว = ผ่าน

ฟอร์มที่อ่านได้ทั้งหมด (และวิธีอ่าน) อยู่ใน [`SKILL.md`](SKILL.md) หัวข้อ **Form registry**

> ไฟล์ลูกค้า (MF / ENTER / รายงาน) อยู่ในเครื่องผู้ใช้เท่านั้น — `.gitignore` กันไม่ให้ขึ้น GitHub

---

## ใช้ออนไลน์ (ส่งลิงก์ให้เพื่อนได้ ไม่ต้องติดตั้ง)

เปิด **<https://rattanao.github.io/Various-forms/>** → ลากไฟล์ MF + ENTER ใส่ → ใส่ SHED NO. → กด **ตรวจสอบ** → กด **ดาวน์โหลด EDI.xlsx**

- ไฟล์ถูกประมวลผลในเบราว์เซอร์ของผู้ใช้เท่านั้น ไม่อัปโหลดไปที่ใด (ใช้ Pyodide รันโปรแกรม Python ชุดเดียวกับในเครื่อง)
- ครั้งแรกเบราว์เซอร์โหลดตัวประมวลผล ~40 MB · ใช้ Chrome / Edge รุ่นใหม่
- หน้าที่เป็นภาพสแกนใช้ OCR ของ Tesseract (หน้าละหลายวินาที) — อ่านได้น้อยกว่า OCR ของ Windows ในเวอร์ชันติดตั้งเล็กน้อย
  หน้าที่อ่านไม่ออกจะแจ้งเลขหน้าให้ตรวจด้วยตา ; งานที่มีหน้าสแกนมากใช้เวอร์ชันติดตั้ง (ด้านล่าง) จะแม่นและเร็วกว่า

---

## ติดตั้งครั้งแรก (ทำครั้งเดียว)

1. ติดตั้ง **Python 3.10+** จาก <https://www.python.org/downloads/> — หน้าแรกของตัวติดตั้ง **ติ๊ก ☑ "Add python.exe to PATH"**
2. ดาวน์โหลดโปรแกรม: หน้า repo กด **Code → Download ZIP** แล้วแตกไฟล์ไว้ที่ไหนก็ได้
   (หรือ `git clone https://github.com/Rattanao/Various-forms.git`)
3. ครั้งแรกที่รัน `run-check-edi.bat` จะติดตั้งไลบรารีให้เอง (`pymupdf`, `openpyxl`, `xlrd`)

ต้องใช้ **Windows 10/11** (OCR ของ Windows มีมาในเครื่อง ไม่ต้องติดตั้งเพิ่ม)

## ใช้งานทุกครั้ง

1. วางไฟล์ **MF** (`.xls` / `.xlsx` / MF `.pdf`) และ **ENTER** (`.pdf`) ในโฟลเดอร์ `input`
2. ดับเบิลคลิก **`run-check-edi.bat`**
3. ใส่ SHED NO. — เลขเดียวทุก B/L (เช่น `0141`) หรือเว้นว่างแล้วใส่แยก LCL/CFS กับ CY (เช่น `0141` / `0121`)
4. รายงานอยู่ที่ `output\EDI-<วันเวลา>.xlsx` (ไม่เขียนทับไฟล์เดิม)

ใช้ผ่าน Command Prompt:

```
python scripts\bundle_edi.py --shed 0141 --name EDI.xlsx
python scripts\bundle_edi.py --shed-lcl 0141 --shed-other 0121 --name EDI.xlsx
```

ตัวเลือก: `--manifest` / `--enter` ระบุไฟล์เอง · `--all-mf` แสดงทุก B/L ของ MF · `--outdir` โฟลเดอร์ผลลัพธ์

## ใช้กับ Claude Code (skill)

คัดลอก `SKILL.md` และโฟลเดอร์ `scripts` ไปไว้ที่ `C:\Users\<ชื่อผู้ใช้>\.claude\skills\bundle-enter-check\`
แล้วพิมพ์ `recheck input` / `check-edi input` ในโฟลเดอร์งาน

## ไฟล์

| ไฟล์ | หน้าที่ |
|---|---|
| `scripts/bundle_edi.py` | ตัวอ่าน ENTER ชุดรวม (แยกชนิดหน้า, โปรไฟล์ฟอร์ม, OCR) + ปรับรายงาน |
| `scripts/build_edi.py` | ตัวเทียบ MF/ENTER และตัวอ่านฟอร์มของ forwarder ที่มีอยู่แล้ว (สำเนาของ skill ตรวจ EDI แต่ละเจ้า + ส่วนที่แก้เพิ่ม `[bundle-enter-check]`) |
| `scripts/win_ocr.ps1` | OCR ของ Windows สำหรับหน้าสแกน |
| `scripts/verify_trace.py` | พิสูจน์ว่าค่าทุกช่องมาจากไฟล์จริง |
| `scripts/run_tests.py`, `scripts/test_forwarders.py` | ทดสอบย้อนหลังกับงานตัวอย่างในเครื่อง (ไฟล์งานไม่ได้อยู่ใน repo) |
| `SKILL.md` | กฎการตรวจ + ทะเบียนฟอร์มทุกแบบ |
