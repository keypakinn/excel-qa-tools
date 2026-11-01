"""สร้างไฟล์ทดสอบจำลอง (ไม่มีข้อมูลลูกค้าจริง)
bad.xlsx  = มีปัญหาครบทุก check ใหม่
good.xlsx = โครงเดียวกันแต่ถูกต้อง ต้องไม่ถูกเตือนเรื่องใหม่
ใช้: python3 tests/make_samples.py <out_dir>
"""
import base64, io, os, re, sys, zipfile
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Protection

out = sys.argv[1] if len(sys.argv) > 1 else "."
os.makedirs(out, exist_ok=True)

ROWS = [("Jan", "AAA", "T1", 100), ("Jan", "BBB", "T2", 200), ("Feb", "AAA", "T2", 50),
        ("Feb", "CCC", "T1", 70), ("Mar", "BBB", "T1", 30), ("Mar", "AAA", "T1", 10)]


def build(path, bad):
    wb = openpyxl.Workbook()
    rp = wb.active; rp.title = "Report"
    d = wb.create_sheet("Data")
    lst = wb.create_sheet("Lists")
    d.append(["Month", "Company", "Type", "Amount"])
    for r in ROWS:
        d.append(list(r))
    for i, m in enumerate(["Jan", "Feb", "Mar", "Apr"], 2):
        lst.cell(i, 1, m)
    lst["A1"] = "Months"

    rp["B1"] = "รายงานยอดขาย"
    rp["B2"] = "ยอดขายเดือนมีนาคม 2569" if bad else "ยอดขายตามเดือนที่เลือก"
    rp["C3"] = "Jan"
    dv = DataValidation(type="list", formula1="=Data!$A$2:$A$7" if bad else "=Lists!$A$2:$A$5")
    rp.add_data_validation(dv); dv.add("C3")
    heads = ["AAA", "BBB", "AAA"] if bad else ["AAA", "BBB", "CCC"]
    for j, h in enumerate(heads):
        rp.cell(5, 4 + j, h)
    rp["H5"] = "key"
    labels = [("T1", "T1"), ("T2", "T2"), ("T9", "T9")] if bad else [("T1", "T1"), ("T2", "T2")]
    if bad:
        labels = [("T1", "T1"), ("T2", "T1"), ("T9", "T9")]   # แถว 7: ป้าย T2 แต่คีย์ T1 / แถว 8: T9 ไม่มีในข้อมูล
    r = 6
    for lab, key in labels:
        rp.cell(r, 2, lab); rp.cell(r, 8, key)
        for j in range(3):
            c = 4 + j
            L = openpyxl.utils.get_column_letter(c)
            rp.cell(r, c, f"=SUMIFS(Data!$D:$D,Data!$B:$B,{L}$5,Data!$C:$C,$H{r},Data!$A:$A,$C$3)")
        r += 1
    rp.cell(r, 2, "Manual adjust")
    for j in range(3):
        rp.cell(r, 4 + j, 0)
    r += 1
    rp.cell(r, 2, "รวม")
    for j in range(3):
        L = openpyxl.utils.get_column_letter(4 + j)
        rp.cell(r, 4 + j, f"=SUM({L}6:{L}{r-1})")
    rp["J1"] = 2026
    rp["B20"] = '="ปี "&J1'
    # การป้องกัน: bad = สูตรบางเซลล์ไม่ซ่อน / dropdown ล็อก ; good = ถูกต้อง
    for row in rp.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                c.protection = Protection(locked=True, hidden=True)
    rp["C3"].protection = Protection(locked=bad)
    if bad:
        rp["D6"].protection = Protection(locked=False, hidden=False)
    rp.protection.sheet = True
    rp.protection.formatColumns = False if bad else True
    lst.protection.sheet = True
    wb.save(path)


def patch(path, add_x14=False, add_pq=False, cached=None):
    """ฉีด XML ที่ openpyxl เขียนไม่ได้: x14 DV, DataMashup, ค่าผลลัพธ์ของสูตร"""
    zin = zipfile.ZipFile(path)
    items = {n: zin.read(n) for n in zin.namelist()}
    zin.close()
    sheet1 = "xl/worksheets/sheet1.xml"
    x = items[sheet1].decode("utf-8")
    if cached:
        for addr, val in cached.items():
            x = re.sub(r'(<c r="%s"[^>]*?)(?: t="\w+")?>(<f>.*?</f>)<v\s*/?>(?:</v>)?' % addr,
                       lambda m: f'{m.group(1)} t="str">{m.group(2)}<v>{val}</v>', x)
    if add_x14:
        ext = ('<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" '
               'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
               '<x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
               '<x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>Data!$B$2:$B$7</xm:f></x14:formula1>'
               '<xm:sqref>K3</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>')
        x = x.replace("</worksheet>", ext + "</worksheet>")
    items[sheet1] = x.encode("utf-8")
    if add_pq:
        m = ('section Section1;\r\nshared Sales = let\r\n Source = Excel.Workbook(File.Contents("C:\\Users\\dev\\Desktop\\db.xlsx"), null, true)\r\n'
             'in Source;\r\nshared Local = let Source = Excel.CurrentWorkbook(){[Name="T"]}[Content] in Source;')
        bio = io.BytesIO()
        with zipfile.ZipFile(bio, "w") as pz:
            pz.writestr("Formulas/Section1.m", m)
        pkg = bio.getvalue()
        blob = (0).to_bytes(4, "little") + len(pkg).to_bytes(4, "little") + pkg + (0).to_bytes(4, "little")
        xml = ('<?xml version="1.0" encoding="utf-16"?><DataMashup xmlns="http://schemas.microsoft.com/DataMashup">'
               + base64.b64encode(blob).decode() + "</DataMashup>")
        items["customXml/item1.xml"] = xml.encode("utf-16")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in items.items():
            z.writestr(n, b)


bad = os.path.join(out, "bad.xlsx"); good = os.path.join(out, "good.xlsx")
build(bad, True); patch(bad, add_x14=True, add_pq=True, cached={"B20": "ปี 2026"})
build(good, False); patch(good, cached={"B20": "ปี 2569"})
print("created", bad, good)
