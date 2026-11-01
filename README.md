# excel-qa-tools

สคริปต์ของชุด skill ตรวจไฟล์ Excel ก่อนส่งลูกค้า (excel-qa, excel-fix, excel-recheck, excel-lock, excel-final, excel-feedback)
ทุกตัว **อ่านอย่างเดียว** ไม่แก้ไฟล์ต้นฉบับ — การแก้ทำผ่าน Excel จริงด้วยไฟล์ใน `fix/` เท่านั้น
ไม่มีข้อมูลลูกค้า ไม่มีรหัสผ่านใด ๆ ใน repo นี้

## ติดตั้งใน workspace ของ Claude
```bash
git clone -q --depth 1 https://github.com/keypakinn/excel-qa-tools /home/claude/xqa
pip install -q oletools msoffcrypto-tool --break-system-packages
```

## ไฟล์
| ไฟล์ | ใช้ทำอะไร |
|---|---|
| `excel_qa_check.py` | ตรวจไฟล์ → JSON กลุ่ม A–G (`--excel 2019`, `--calendar ce\|be`, `--out qa.json`) |
| `qa_extra.py` | check เสริม: เกณฑ์ SUMIFS ซ้ำ/ไม่มีในข้อมูล, ป้ายสลับคีย์, dropdown จากข้อมูลดิบ (รวม x14), Power Query path, ข้อความเดือน/ปีตายตัว, ปฏิทินปน |
| `excel_qa_tools.py` | `dump` `diff` `lookup` `find-sheet` `buttons` `vba-diff` `lock` `final` |
| `xparts.py` | อ่าน XML ที่ openpyxl ไม่รองรับ + เปิดไฟล์มีรหัส |
| `fix/template.vbs` `fix/make_vbs.py` `fix/RUN_FIX.bat` | รอบแก้ด้วย Excel จริง (VBScript late binding) |
| `tests/` | `python3 tests/run_tests.py` ต้องได้ ALL PASS ก่อน push ทุกครั้ง |

## ไฟล์มีรหัส
ส่งรหัสทาง env เท่านั้น: `QA_PASSWORD` (ไฟล์แรก) / `QA_PASSWORD2` (ไฟล์ที่สองของ diff/final) — ห้าม echo ห้ามเขียนลงไฟล์

## กลุ่มผลตรวจ
A สูตรพัง/ยอดผิด · B สูตรเสี่ยง · C VBA · D เวอร์ชัน Excel · E วันที่/ปฏิทิน · F ความเรียบร้อย · G ถามยืนยัน (อาจถูกหรือผิด ให้ user ตัดสิน)
ทุก finding มี `sev` (P0–P3, ASK, OK)
