# Check EDI — ENTER ชุดรวมหลายฟอร์ม (Various forms)

โปรแกรมตรวจเอกสารขาเข้าเรือก่อนทำใบขนสินค้าขาเข้า สำหรับ **ENTER ที่เป็น PDF ไฟล์เดียวรวมหลายฟอร์ม**
(จดหมายแก้ไขของ forwarder หลายเจ้า + ARRIVAL NOTICE / สำเนา B/L + หน้าสแกน) เทียบกับ **MF (MANIFEST)**
แล้วออกรายงาน **`EDI.xlsx`** แบบใส่สี ทีละ B/L

โปรแกรมแยก ENTER ทีละหน้า → รู้ว่าเป็นฟอร์มไหน → อ่านด้วยตัวอ่านของฟอร์มนั้น → เทียบกับ MF

- CONSIGNEE, CONTAINER NO., STATUS (CY / LCL / LCL-CFS)
- TOTAL PACKAGE + ชนิดหีบห่อ (WOODEN CASE ≠ CASE, FLEXIBAG ≠ FLEXITANK), GROSS WEIGHT, MEASUREMENT (แจ้ง MF = 0.000)
- MARKS, DESCRIPTION (ก่อน `AS PER ATTACHED SHEET` ตรวจ 30 ตัวอักษรแรก ; กรอบแดง / ป้ายสี / ขีดเส้นใต้ของลูกค้า = สิ่งที่ลูกค้าแจ้ง ต้องเท่ากับ MF (MARKS / MEAS. ที่ไม่ได้ทำเครื่องหมาย = ไม่ได้แจ้ง ; ไม่มีเครื่องหมาย + ลูกค้าเขียน "ENTER ตามหน้า BL" = ตรวจทุกช่อง))
- DG: CLASS / UN ตรงกันครบ = ไม่เตือน (`CL9` = CLASS 9) — ENTER ไม่ระบุ CLASS หรือ UN ที่ MF มี = เตือน · REEFER: TEMP ตรงกัน = ไม่เตือน แต่ MF ไม่มี VENT = เตือน · TRANSIT/TRANSHIPMENT ตรงกันและ MF ถูกต้องตามกฎ = `-`
- ประเทศปลายทาง (ลาว = LA), CARGO MOVEMENT, TAX ID เทียบ NOTIFY — **ถูกต้องตามกฎ = แสดง `-`** (แสดงค่าจริงเฉพาะตอนไม่ตรง/ต้องตรวจ)
- **SHED (DG_M/F)** ตรวจเฉพาะ DG ที่ DISCHARGE LAEM CHABANG ต้องเป็น **2826** (ไม่ตรง = เตือน) นอกนั้นไม่ตรวจ
- แสดง **ทุก B/L ของ MF เรียงตามลำดับใน MF** — B/L ที่ไม่มี ENTER = **ช่อง B/L NO. สีเหลืองเข้ม** ; B/L ที่มีใน ENTER แต่ไม่มีใน MF ขึ้นแถววิกฤต
- B/L แตก (เช่น …664, …664A–L) เรียงต่อกัน ; หมายเหตุทุกแถวบอกหน้า ENTER `[ENTER หน้า n]`

**ช่องที่ ENTER ตรงกับ MF = `-`** (ไม่ตรง = แสดงค่าจริง + สี)

**สี:** 🟥 แดง = ไม่ตรง · 🟧 ส้ม = ต้องให้คนตรวจยืนยัน (รวมหน้าที่อ่านด้วย OCR) · 🟨 เหลืองเข้ม (ช่อง B/L NO.) = ไม่มี ENTER ให้เทียบ · 🟩 เขียว = ผ่าน

ฟอร์มที่อ่านได้ทั้งหมด (และวิธีอ่าน) อยู่ใน [`SKILL.md`](SKILL.md) หัวข้อ **Form registry**

> ไฟล์ลูกค้า (MF / ENTER / รายงาน) อยู่ในเครื่องผู้ใช้เท่านั้น — `.gitignore` กันไม่ให้ขึ้น GitHub

---

## ใช้ออนไลน์ (ส่งลิงก์ให้เพื่อนได้ ไม่ต้องติดตั้ง)

เปิด **<https://rattanao.github.io/Various-forms/>** → ลากไฟล์ MF + ENTER ใส่ → กด **ตรวจสอบ** → กด **ดาวน์โหลด EDI.xlsx**

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
3. รายงานอยู่ที่ `output\EDI-<วันเวลา>.xlsx` (ไม่เขียนทับไฟล์เดิม)

ใช้ผ่าน Command Prompt:

```
python scripts\bundle_edi.py --name EDI.xlsx
```

ตัวเลือก: `--manifest` / `--enter` ระบุไฟล์เอง · `--enter-only` แสดงเฉพาะ B/L ที่มีใน ENTER (แบบเดิม) · `--outdir` โฟลเดอร์ผลลัพธ์

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
