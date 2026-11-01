"""xparts.py — ตัวช่วยอ่านส่วนต่าง ๆ ของไฟล์ .xlsx/.xlsm จาก XML โดยตรง (อ่านอย่างเดียว)
ใช้ร่วมกันโดย excel_qa_check.py และ excel_qa_tools.py

- open_book(path, pw_env)   : คืน path ที่อ่านได้ (ถอดรหัสเป็นไฟล์ชั่วคราวถ้ามีรหัส) + ฟังก์ชันลบทิ้ง
- sheet_xml_map(z)          : ชื่อชีต -> path XML ในไฟล์ zip
- x14_validations(z)        : Data Validation แบบ x14 (list อ้างชีตอื่น) ที่ openpyxl อ่านไม่ได้
- power_queries(z)          : ชื่อ query + source (File.Contents / Folder.Files / Web.Contents ...) จาก DataMashup
- buttons(z)                : ปุ่ม/รูปที่ผูกมาโคร + ตำแหน่ง
- split_args / iter_calls   : แยกอาร์กิวเมนต์ของฟังก์ชันในสูตร
"""
import base64, io, os, re, tempfile, zipfile, posixpath
from xml.etree import ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# ---------- เปิดไฟล์ (รองรับรหัสผ่าน) ----------
def file_status(path):
    """ok / encrypted / legacy_xls / unreadable"""
    if zipfile.is_zipfile(path):
        return "ok"
    try:
        import olefile
        if olefile.isOleFile(path):
            o = olefile.OleFileIO(path)
            if o.exists("EncryptionInfo") and o.exists("EncryptedPackage"):
                return "encrypted"
            return "legacy_xls"
    except Exception:
        pass
    return "unreadable"


def open_book(path, pw_env="QA_PASSWORD"):
    """คืน (status, readable_path, cleanup) — ถ้ามีรหัสและตั้ง env ไว้ จะถอดเป็นไฟล์ชั่วคราว
    ห้าม print รหัส ห้ามเขียนรหัสลงไฟล์"""
    st = file_status(path)
    noop = lambda: None
    if st != "encrypted":
        return st, (path if st == "ok" else None), noop
    pw = os.environ.get(pw_env)
    if not pw:
        return "encrypted", None, noop
    try:
        import msoffcrypto
    except ImportError:
        return "encrypted", None, noop
    fd, tmp = tempfile.mkstemp(suffix=os.path.splitext(path)[1])
    os.close(fd)
    try:
        with open(path, "rb") as fi, open(tmp, "wb") as fo:
            of = msoffcrypto.OfficeFile(fi); of.load_key(password=pw); of.decrypt(fo)
    except Exception:
        os.remove(tmp)
        return "wrong_password", None, noop
    return "ok", tmp, (lambda: os.path.exists(tmp) and os.remove(tmp))


# ---------- โครงสร้างไฟล์ ----------
def _rels(z, part):
    d, b = posixpath.split(part)
    rp = posixpath.join(d, "_rels", b + ".rels")
    if rp not in z.namelist():
        return {}
    out = {}
    for m in re.finditer(r'<Relationship\b[^>]*>', z.read(rp).decode("utf-8", "ignore")):
        t = m.group(0)
        rid = re.search(r'Id="([^"]+)"', t); tg = re.search(r'Target="([^"]+)"', t)
        if rid and tg:
            target = tg.group(1)
            if 'TargetMode="External"' not in t:
                target = posixpath.normpath(posixpath.join(d, target)) if not target.startswith("/") else target.lstrip("/")
            out[rid.group(1)] = target
    return out


def sheet_xml_map(z):
    """{ชื่อชีต: 'xl/worksheets/sheetN.xml'} ตามลำดับในไฟล์"""
    wbx = z.read("xl/workbook.xml").decode("utf-8", "ignore")
    rels = _rels(z, "xl/workbook.xml")
    out = {}
    for m in re.finditer(r"<(?:\w+:)?sheet\b([^>]*)/?>", wbx):
        a = m.group(1)
        nm = re.search(r'name="([^"]*)"', a); rid = re.search(r'r:id="([^"]+)"', a) or re.search(r'\bid="([^"]+)"', a)
        if nm and rid and rid.group(1) in rels:
            out[_xml_unescape(nm.group(1))] = rels[rid.group(1)]
    return out


def sheet_names_fast(path):
    """อ่านชื่อชีตจาก workbook.xml (เร็ว ไม่โหลดทั้งไฟล์) — None ถ้าอ่านไม่ได้"""
    try:
        with zipfile.ZipFile(path) as z:
            return list(sheet_xml_map(z))
    except Exception:
        return None


def _xml_unescape(s):
    return (s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
             .replace("&apos;", "'").replace("&amp;", "&"))


# ---------- Data Validation แบบ x14 ----------
X14_RE = re.compile(r"<x14:dataValidation\b([^>]*)>(.*?)</x14:dataValidation>", re.S)


def x14_validations(z):
    """[{sheet, sqref, type, formula1}] — DV ที่เก็บใน extLst (openpyxl ข้ามไป)"""
    out = []
    for sn, part in sheet_xml_map(z).items():
        if part not in z.namelist():
            continue
        x = z.read(part).decode("utf-8", "ignore")
        if "x14:dataValidation" not in x:
            continue
        for m in X14_RE.finditer(x):
            attrs, body = m.groups()
            t = re.search(r'type="([^"]+)"', attrs)
            f1 = re.search(r"<x14:formula1>\s*<xm:f>(.*?)</xm:f>", body, re.S)
            sq = re.search(r"<xm:sqref>(.*?)</xm:sqref>", body, re.S)
            out.append({"sheet": sn, "sqref": sq.group(1).strip() if sq else "",
                        "type": t.group(1) if t else "", "formula1": _xml_unescape(f1.group(1)) if f1 else ""})
    return out


# ---------- Power Query ----------
PQ_SRC_RE = re.compile(r'\b(File\.Contents|Folder\.Files|Folder\.Contents|Web\.Contents|Sql\.Database|'
                       r'Odbc\.DataSource|OleDb\.DataSource|SharePoint\.Files|SharePoint\.Contents|Csv\.Document|'
                       r'Excel\.Workbook|Excel\.CurrentWorkbook)\s*\(\s*("(?:[^"]|"")*")?', re.S)
LOCAL_PATH_RE = re.compile(r'^(?:[A-Za-z]:\\|\\\\)')


def power_queries(z):
    """[{name, sources:[(func, arg)], local_paths:[...]}] — ถอด DataMashup ใน customXml"""
    out = []
    for n in z.namelist():
        if not (n.startswith("customXml/item") and n.endswith(".xml")):
            continue
        raw = z.read(n)
        txt = None
        for enc in ("utf-16", "utf-8-sig", "utf-8"):
            try:
                t = raw.decode(enc)
                if "DataMashup" in t:
                    txt = t; break
            except Exception:
                continue
        if not txt:
            continue
        m = re.search(r"<DataMashup[^>]*>(.*?)</DataMashup>", txt, re.S)
        if not m:
            continue
        try:
            b = base64.b64decode(re.sub(r"\s", "", m.group(1)))
            ln = int.from_bytes(b[4:8], "little")
            pkg = zipfile.ZipFile(io.BytesIO(b[8:8 + ln]))
            code = pkg.read("Formulas/Section1.m").decode("utf-8-sig", "ignore")
        except Exception as e:
            out.append({"name": "(ถอด DataMashup ไม่ได้)", "error": str(e)[:120], "sources": [], "local_paths": []})
            continue
        # แยกเป็นแต่ละ query: shared #"Name" = ... ;
        for qm in re.finditer(r'shared\s+(#"(?:[^"]|"")*"|[A-Za-z_][\w.]*)\s*=(.*?)(?=\bshared\s|\Z)', code, re.S):
            name = qm.group(1).strip('#"')
            body = qm.group(2)
            srcs = []
            for sm in PQ_SRC_RE.finditer(body):
                arg = sm.group(2)[1:-1].replace('""', '"') if sm.group(2) else ""
                srcs.append((sm.group(1), arg))
            local = [a for f, a in srcs if LOCAL_PATH_RE.match(a)]
            out.append({"name": name, "sources": srcs, "local_paths": local})
    return out


# ---------- ปุ่ม / shape ที่ผูกมาโคร ----------
def buttons(z):
    """[{sheet, kind, name, macro, anchor}]"""
    out = []
    names = z.namelist()
    for sn, part in sheet_xml_map(z).items():
        rels = _rels(z, part)
        for rid, tgt in rels.items():
            if tgt not in names:
                continue
            x = z.read(tgt).decode("utf-8", "ignore")
            if tgt.startswith("xl/drawings/drawing"):
                for am in re.finditer(r"<xdr:(?:twoCellAnchor|oneCellAnchor|absoluteAnchor)\b.*?</xdr:(?:twoCellAnchor|oneCellAnchor|absoluteAnchor)>", x, re.S):
                    a = am.group(0)
                    for mm in re.finditer(r'<xdr:(sp|pic|grpSp|graphicFrame)\b[^>]*\bmacro="([^"]+)"', a):
                        nm = re.search(r'<xdr:cNvPr\b[^>]*name="([^"]*)"', a[mm.start():])
                        fc = re.search(r"<xdr:from>\s*<xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>", a, re.S)
                        anchor = _a1(int(fc.group(1)), int(fc.group(2))) if fc else ""
                        out.append({"sheet": sn, "kind": "shape", "name": _xml_unescape(nm.group(1)) if nm else "",
                                    "macro": _xml_unescape(mm.group(2)), "anchor": anchor})
            elif "vmlDrawing" in tgt:
                for sm in re.finditer(r"<v:shape\b.*?</v:shape>", x, re.S):
                    s = sm.group(0)
                    mac = re.search(r"<x:FmlaMacro>(.*?)</x:FmlaMacro>", s, re.S)
                    if not mac:
                        continue
                    an = re.search(r"<x:Anchor>\s*([\d\s,]+)</x:Anchor>", s)
                    anchor = ""
                    if an:
                        v = [int(t) for t in an.group(1).replace(" ", "").split(",") if t]
                        if len(v) >= 4:
                            anchor = _a1(v[0], v[2])
                    ot = re.search(r'ObjectType="([^"]+)"', s)
                    out.append({"sheet": sn, "kind": "form:" + (ot.group(1) if ot else "?"), "name": "",
                                "macro": _xml_unescape(mac.group(1).strip()), "anchor": anchor})
    return out


def _a1(col0, row0):
    from openpyxl.utils import get_column_letter
    return f"{get_column_letter(col0 + 1)}{row0 + 1}"


# ---------- แยกอาร์กิวเมนต์ฟังก์ชันในสูตร ----------
def split_args(s):
    """แยกสตริงอาร์กิวเมนต์ด้วย , ระดับบนสุด (เคารพ วงเล็บ/สตริง/'ชื่อชีต')"""
    out, depth, cur, q, sq = [], 0, [], False, False
    i = 0
    while i < len(s):
        ch = s[i]
        if q:
            cur.append(ch)
            if ch == '"':
                if i + 1 < len(s) and s[i + 1] == '"':
                    cur.append('"'); i += 1
                else:
                    q = False
        elif sq:
            cur.append(ch)
            if ch == "'":
                sq = False
        elif ch == '"':
            q = True; cur.append(ch)
        elif ch == "'":
            sq = True; cur.append(ch)
        elif ch in "({":
            depth += 1; cur.append(ch)
        elif ch in ")}":
            depth -= 1; cur.append(ch)
        elif ch == "," and depth == 0:
            out.append("".join(cur).strip()); cur = []
        else:
            cur.append(ch)
        i += 1
    out.append("".join(cur).strip())
    return out


def iter_calls(formula, funcs):
    """ให้ (ชื่อฟังก์ชัน, [args]) ของทุกการเรียกฟังก์ชันใน funcs (รวมที่ซ้อนกัน)"""
    up = formula
    pat = re.compile(r"(?<![A-Za-z0-9_.])(?:_xlfn\.)?(" + "|".join(funcs) + r")\s*\(", re.I)
    for m in pat.finditer(up):
        start = m.end(); depth = 1; i = start; q = False; sq = False
        while i < len(up) and depth:
            ch = up[i]
            if q:
                if ch == '"':
                    if i + 1 < len(up) and up[i + 1] == '"':
                        i += 1          # "" = เครื่องหมายคำพูดในสตริง
                    else:
                        q = False
            elif sq:
                if ch == "'": sq = False
            elif ch == '"': q = True
            elif ch == "'": sq = True
            elif ch == "(": depth += 1
            elif ch == ")": depth -= 1
            i += 1
        if depth == 0:
            yield m.group(1).upper(), split_args(up[start:i - 1])
