#!/usr/bin/env python3
"""excel_qa_check.py — ตรวจความสมบูรณ์ไฟล์ Excel (อ่านอย่างเดียว ไม่แก้ไฟล์)
ใช้: python3 excel_qa_check.py <file.xlsx|.xlsm> [--excel 2016|2019|2021|2024|365] [--calendar ce|be] [--out result.json]
--calendar be = ลูกค้าใช้ปี พ.ศ. (ไม่เตือนรูปแบบ พ.ศ. แต่ตรวจปีปนกันแทน)
ไฟล์มีรหัส: ตั้ง env QA_PASSWORD แล้วสคริปต์จะถอดเป็นสำเนาชั่วคราว ตรวจ และลบทิ้ง
ผลลัพธ์: JSON แบ่งกลุ่ม A-F ให้ Claude นำไปเขียนรายงาน
"""
import argparse, json, os, re, sys, tempfile, zipfile, datetime
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import warnings; warnings.filterwarnings("ignore", message="Data Validation extension")
import openpyxl
import qa_extra
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.formula.translate import Translator

MAX_LOC = 30  # เก็บตัวอย่างตำแหน่งสูงสุดต่อรายการ

# ---------- ตัวช่วย ----------
ERRORS = ("#REF!", "#DIV/0!", "#N/A", "#VALUE!", "#NAME?", "#NUM!", "#NULL!", "#SPILL!", "#CALC!", "#GETTING_DATA")
CELL_RE = re.compile(r"(?<![A-Za-z0-9_.\]!:$])(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(])")
STR_RE = re.compile(r'"(?:[^"]|"")*"')
FUNC_RE = re.compile(r"(?<![A-Za-z0-9_.])((?:_xlfn\.)?(?:_xlws\.)?[A-Z][A-Z0-9.]*)\s*\(")

def formula_text(v):
    if isinstance(v, ArrayFormula):
        return v.text or ""
    if isinstance(v, str) and v.startswith("="):
        return v
    t = getattr(v, "text", None)  # DataTableFormula ฯลฯ
    return t if isinstance(t, str) else None

def strip_strings(f):
    return STR_RE.sub('""', f)

def rel_pattern(f, row, col):
    """แปลงสูตรเป็นรูปแบบสัมพัทธ์ (แบบ R1C1) เพื่อเทียบความสม่ำเสมอ"""
    f = strip_strings(f)
    def rep(m):
        cabs, cl, rabs, rw = m.groups()
        c = column_index_from_string(cl); r = int(rw)
        cs = f"C{c}" if cabs else f"C[{c-col}]"
        rs = f"R{r}" if rabs else f"R[{r-row}]"
        return rs + cs
    return CELL_RE.sub(rep, f)

class Loc:
    def __init__(self):
        self.items = []; self.count = 0
    def add(self, x):
        self.count += 1
        if len(self.items) < MAX_LOC:
            self.items.append(x)
    def out(self):
        return {"count": self.count, "examples": self.items}

SEV_DEFAULT = {"grouped_sheets": "P1", "external_link": "P1", "vba_no_option_explicit": "P3"}
TOTAL_RE = re.compile(r"รวม|สรุป|total|subtotal|grand|sum\b", re.I)
MANUAL_RE = re.compile(r"manual|กรอก|ปรับ|ใส่เอง|key\s*in|input|adjust", re.I)


def finding(group, key, title, loc, **extra):
    d = {"group": group, "key": key, "title": title}
    d.update(loc.out() if isinstance(loc, Loc) else {"count": len(loc), "examples": loc[:MAX_LOC]})
    d.update(extra)
    return d

# ---------- ฟังก์ชันที่แต่ละเวอร์ชันไม่รองรับ (เวอร์ชันแรกที่รองรับ) ----------
FUNC_MIN_VER = {}
for f in "CONCAT TEXTJOIN IFS SWITCH MAXIFS MINIFS".split():
    FUNC_MIN_VER[f] = 2019
for f in "XLOOKUP XMATCH FILTER SORT SORTBY UNIQUE SEQUENCE RANDARRAY LET".split():
    FUNC_MIN_VER[f] = 2021
for f in ("LAMBDA MAP REDUCE SCAN BYROW BYCOL MAKEARRAY ISOMITTED TEXTSPLIT TEXTBEFORE TEXTAFTER "
          "VSTACK HSTACK TOCOL TOROW WRAPROWS WRAPCOLS TAKE DROP CHOOSEROWS CHOOSECOLS EXPAND "
          "ARRAYTOTEXT VALUETOTEXT IMAGE").split():
    FUNC_MIN_VER[f] = 2024
for f in ("GROUPBY PIVOTBY PERCENTOF TRIMRANGE REGEXTEST REGEXEXTRACT REGEXREPLACE STOCKHISTORY "
          "COPILOT TRANSLATE DETECTLANGUAGE").split():
    FUNC_MIN_VER[f] = 365
VER_ORDER = {"2016": 2016, "2019": 2019, "2021": 2021, "2024": 2024, "365": 9999}

# ---------- วันที่ ----------
THAI_MONTHS = ("มกราคม กุมภาพันธ์ มีนาคม เมษายน พฤษภาคม มิถุนายน กรกฎาคม สิงหาคม กันยายน ตุลาคม พฤศจิกายน ธันวาคม "
               "ม.ค. ก.พ. มี.ค. เม.ย. พ.ค. มิ.ย. ก.ค. ส.ค. ก.ย. ต.ค. พ.ย. ธ.ค.").split()
TEXT_DATE_RE = re.compile(r"^\s*(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2}|\d{4})\s*$")
ISO_DATE_RE = re.compile(r"^\s*(\d{4})-(\d{1,2})-(\d{1,2})\s*$")
THAI_TEXT_DATE_RE = re.compile(r"^\s*(\d{1,2})\s*(" + "|".join(re.escape(m) for m in THAI_MONTHS) + r")\s*(\d{2,4})\s*$")
LCID_RE = re.compile(r"\[\$-([0-9A-Fa-f]{5,8})\]")
LOCALE_RE = re.compile(r"\[\$-[A-Za-z\-]+,([0-9A-Fa-f]+)\]")

def is_buddhist_format(fmt):
    if not fmt:
        return False
    for m in LCID_RE.finditer(fmt):
        if (int(m.group(1), 16) >> 16) & 0xFF == 7:
            return True
    for m in LOCALE_RE.finditer(fmt):
        if int(m.group(1), 16) & 0xFF == 7:
            return True
    bare = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", fmt)
    return bool(re.search(r"(?<![A-Za-z])bb(bb)?(?![A-Za-z])", bare, re.I))

def to_ce_format(fmt):
    """แปลงรูปแบบ พ.ศ. เป็น ค.ศ.: ตัดแท็กปฏิทินไทย, bbbb->yyyy, bb->yy"""
    thai = bool(re.search(r"\[\$-(?:th|[0-9A-Fa-f]*41[Ee])", fmt))
    out = LCID_RE.sub("", fmt)
    out = LOCALE_RE.sub("", out)
    # มีชื่อเดือน (mmm/mmmm) และเดิมเป็นภาษาไทย -> คงภาษาไทยไว้ แต่ใช้ปฏิทิน ค.ศ.
    if thai and re.search(r"(?i)mmm", re.sub(r'"[^"]*"', "", out)):
        out = "[$-th-TH]" + out
    parts = re.split(r'("[^"]*")', out)
    parts = [p if p.startswith('"') else re.sub(r"(?i)bbbb", "yyyy", re.sub(r"(?i)(?<![b])bb(?![b])", "yy", p)) for p in parts]
    out = "".join(parts).strip()
    return out or "dd/mm/yyyy"

# ---------- VBA ----------
def check_vba(path, sheetnames):
    res = {"has_vba": False, "status": "no_vba", "locked": False, "modules": [], "findings": []}
    with zipfile.ZipFile(path) as z:
        if "xl/vbaProject.bin" not in z.namelist():
            return res
        res["has_vba"] = True
        binary = z.read("xl/vbaProject.bin")
    try:
        import olefile, io
        ole = olefile.OleFileIO(io.BytesIO(binary))
        if ole.exists("PROJECT"):
            proj = ole.openstream("PROJECT").read().decode("latin-1", "ignore")
            m = re.search(r'DPB="([0-9A-Fa-f]*)"', proj)
            if m and len(m.group(1)) > 8:
                res["locked"] = True
    except Exception:
        pass
    try:
        from oletools.olevba import VBA_Parser
    except ImportError:
        res["status"] = "cannot_read"; res["reason"] = "ไม่มีไลบรารี oletools"; return res
    code_by_mod = {}
    try:
        vp = VBA_Parser(path)
        for (_f, _s, vname, code) in vp.extract_macros():
            if isinstance(code, bytes):
                code = code.decode("cp874", "ignore")
            code_by_mod[vname] = code
        vp.close()
    except Exception as e:
        res["status"] = "cannot_read"; res["reason"] = str(e)[:200]; return res
    real = {k: v for k, v in code_by_mod.items()
            if [l for l in v.splitlines() if l.strip() and not l.strip().startswith("Attribute ")]}
    if not real:
        # มี vbaProject แต่อ่านโค้ดไม่ได้เลย -> ถือว่าตรวจไม่ได้ (มักเป็นเพราะล็อก/เข้ารหัส)
        res["status"] = "cannot_read" if res["locked"] or code_by_mod == {} else "empty"
        res["reason"] = "ตรวจ VBA ไม่ได้เพราะล็อก" if res["locked"] else "ไม่มีโค้ดในโมดูล"
        return res
    res["status"] = "ok"
    res["modules"] = sorted(real)
    path_re = re.compile(r'"[^"]*(?:[A-Za-z]:\\|\\\\[A-Za-z0-9_]|/Users/|\\Users\\)[^"]*"')
    sheet_re = re.compile(r'(?:Sheets|Worksheets)\s*(?:\.Item)?\s*\(\s*"([^"]+)"\s*\)', re.I)
    be_re = re.compile(r'[+\-]\s*543\b|"[^"]*bbbb[^"]*"', re.I)
    no_explicit = []; paths = Loc(); badsheets = Loc(); be = Loc()
    for mod, code in real.items():
        lines = code.splitlines()
        body = [l for l in lines if l.strip() and not l.strip().startswith("Attribute ")]
        if not any(re.match(r"\s*Option\s+Explicit", l, re.I) for l in body):
            no_explicit.append(mod)
        for i, l in enumerate(lines, 1):
            s = l.split("'")[0] if '"' not in l else l
            if path_re.search(s):
                paths.add(f"{mod} บรรทัด {i}: {l.strip()[:120]}")
            for m in sheet_re.finditer(s):
                if m.group(1) not in sheetnames:
                    badsheets.add(f'{mod} บรรทัด {i}: อ้างชีต "{m.group(1)}" ซึ่งไม่มีในไฟล์')
            if be_re.search(s):
                be.add(f"{mod} บรรทัด {i}: {l.strip()[:120]}")
    if paths.count:
        res["findings"].append(finding("C", "vba_hardcoded_path", "VBA อ้าง path ตายตัว (อาจเป็นเครื่องคุณ)", paths))
    if badsheets.count:
        res["findings"].append(finding("C", "vba_missing_sheet", "VBA อ้างชื่อชีตที่ไม่มีแล้ว", badsheets))
    if no_explicit:
        res["findings"].append(finding("C", "vba_no_option_explicit", "โมดูลที่ไม่มี Option Explicit", no_explicit))
    res["be_lines"] = be.out()
    return res

# ---------- ตัวตรวจหลัก ----------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--excel", default="2019", choices=list(VER_ORDER))
    ap.add_argument("--calendar", default="ce", choices=["ce", "be"])
    ap.add_argument("--out")
    a = ap.parse_args()
    path = a.file

    # ไฟล์ที่ตั้งรหัสผ่านเปิด = OLE container ไม่ใช่ zip -> ตรวจไม่ได้จนกว่าจะได้รหัส/สำเนาที่ไม่มีรหัส
    if not zipfile.is_zipfile(path):
        status, reason = "unreadable", "ไม่ใช่ไฟล์ .xlsx/.xlsm ที่อ่านได้ (อาจเสียหาย หรือเป็น .xls เก่า)"
        try:
            import olefile
            if olefile.isOleFile(path):
                o = olefile.OleFileIO(path)
                if o.exists("EncryptionInfo") and o.exists("EncryptedPackage"):
                    status, reason = "encrypted", "ไฟล์ตั้งรหัสผ่านเปิดไว้ (Password to open)"
                else:
                    status, reason = "legacy_xls", "เป็นไฟล์ .xls รุ่นเก่า (binary) ไม่ใช่ .xlsx/.xlsm"
        except Exception:
            pass
        pw = os.environ.get("QA_PASSWORD")
        if status == "encrypted" and pw:
            # ถอดรหัสเป็นสำเนาชั่วคราว ตรวจ แล้วลบทิ้งทันที (ต้นฉบับไม่ถูกแตะ)
            try:
                import msoffcrypto
            except ImportError:
                status, reason = "encrypted", "ต้องติดตั้ง msoffcrypto-tool ก่อน"
            else:
                fd, tmp = tempfile.mkstemp(suffix=os.path.splitext(path)[1])
                os.close(fd)
                try:
                    with open(path, "rb") as fi, open(tmp, "wb") as fo:
                        of = msoffcrypto.OfficeFile(fi); of.load_key(password=pw); of.decrypt(fo)
                except Exception:
                    os.remove(tmp)
                    status, reason = "wrong_password", "รหัสผ่านไม่ถูกต้อง (หรือถอดรหัสไม่ได้)"
                else:
                    try:
                        return run(tmp, a, display_path=path, encrypted=True)
                    finally:
                        os.remove(tmp)
        res = {"file": path, "status": status, "reason": reason}
        print(json.dumps(res, ensure_ascii=False))
        if a.out:
            open(a.out, "w", encoding="utf-8").write(json.dumps(res, ensure_ascii=False))
        sys.exit(2)
    return run(path, a)


def run(path, a, display_path=None, encrypted=False):
    target = VER_ORDER[a.excel]

    wb = openpyxl.load_workbook(path, keep_vba=True)           # สูตร
    wbv = openpyxl.load_workbook(path, data_only=True)          # ค่าที่ Excel คำนวณเก็บไว้
    findings = []
    stats = {"sheets": wb.sheetnames, "formula_cells": 0, "formula_no_cached": 0}

    err_cells = defaultdict(Loc)          # A
    err_in_formula = Loc()                # A: #REF! อยู่ในตัวสูตร
    self_ref = Loc()                      # A: วนอ้างตัวเอง (แบบง่าย)
    inconsistent = Loc(); hardcode = Loc(); intended = Loc()  # B
    ext_formula = Loc()                   # B
    ver_funcs = defaultdict(Loc); spill = Loc()  # D
    be_fmt = Loc(); be_year_val = Loc(); text_date_be = Loc(); text_date = Loc(); f543 = Loc()  # E
    text_date_fix = []                    # สำหรับรอบแก้
    be_fmt_fix = []; be_year_fix = []; formula_fix = []
    narrow_cols = {}; not_a1 = []; comments = Loc()  # F

    for ws in wb.worksheets:
        wsv = wbv[ws.title]
        sn = ws.title
        col_formulas = defaultdict(dict)  # col -> {row: pattern}
        col_values = defaultdict(dict)    # col -> {row: value(non-formula)}
        col_ftext = defaultdict(dict)     # col -> {row: formula text}
        merged = set()
        for rng in ws.merged_cells.ranges:
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    merged.add((r, c))
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                addr = f"{sn}!{cell.coordinate}"
                fmt = cell.number_format
                if cell.comment is not None:
                    comments.add(addr)
                if is_buddhist_format(fmt):
                    be_fmt.add(f"{addr} [{fmt}]" + (" (เซลล์ว่าง)" if v is None else ""))
                    be_fmt_fix.append({"sheet": sn, "cell": cell.coordinate, "format": fmt, "new_format": to_ce_format(fmt)})
                if v is None:
                    continue
                f = formula_text(v)
                cv = wsv[cell.coordinate].value
                if f:
                    stats["formula_cells"] += 1
                    if cv is None:
                        stats["formula_no_cached"] += 1
                    if isinstance(cv, str) and cv in ERRORS:
                        err_cells[cv].add(addr)
                    fs = strip_strings(f)
                    if "#REF!" in fs:
                        err_in_formula.add(f"{addr}: {f[:100]}")
                    if re.search(r"\[[^\]]+\.(xls[xmb]?|csv)\]|'[A-Za-z]:\\|\[\d+\]", f, re.I):
                        ext_formula.add(f"{addr}: {f[:100]}")
                    # อ้างตัวเอง (เฉพาะอ้างในชีตเดียวกันแบบไม่ระบุชื่อชีต)
                    for m in re.finditer(r"(?<![!A-Za-z0-9_.$])\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?(?![A-Za-z0-9_(!])", fs):
                        c1 = column_index_from_string(m.group(1)); r1 = int(m.group(2))
                        c2 = column_index_from_string(m.group(3)) if m.group(3) else c1
                        r2 = int(m.group(4)) if m.group(4) else r1
                        if min(c1, c2) <= cell.column <= max(c1, c2) and min(r1, r2) <= cell.row <= max(r1, r2):
                            self_ref.add(f"{addr}: {f[:100]}"); break
                    # เวอร์ชัน
                    for m in FUNC_RE.finditer(fs):
                        name = m.group(1).replace("_xlfn.", "").replace("_xlws.", "").upper()
                        mv = FUNC_MIN_VER.get(name)
                        if mv and mv > target:
                            ver_funcs[name].add(addr)
                    if re.search(r"(?<![A-Za-z])\$?[A-Z]{1,3}\$?\d+#", fs) or "ANCHORARRAY" in fs.upper():
                        spill.add(addr)
                    # วันที่ พ.ศ. ในสูตร
                    if re.search(r"[+\-]\s*543\b", fs) or re.search(r'"[^"]*bbbb[^"]*"', f, re.I) or \
                       any(is_buddhist_format(s.strip('"')) for s in STR_RE.findall(f)):
                        f543.add(f"{addr}: {f[:100]}")
                    col_formulas[cell.column][cell.row] = rel_pattern(f, cell.row, cell.column)
                    col_ftext[cell.column][cell.row] = f
                else:
                    col_values[cell.column][cell.row] = v
                    if isinstance(v, str) and v.strip() in ERRORS:
                        err_cells[v.strip() + " (พิมพ์เป็นข้อความ)"].add(addr)
                # วันที่
                dv = cv if f else v
                if isinstance(dv, (datetime.datetime, datetime.date)) and dv.year >= 2400:
                    be_year_val.add(f"{addr} = {dv:%d/%m/%Y}")
                    if not f:
                        try:
                            be_year_fix.append({"sheet": sn, "cell": cell.coordinate, "was": f"{dv:%Y-%m-%d}",
                                                "iso": dv.replace(year=dv.year - 543).strftime("%Y-%m-%d")})
                        except ValueError:
                            pass
                if not f and isinstance(v, str):
                    s = v.strip(); y = None; d = mth = None
                    m = TEXT_DATE_RE.match(s)
                    if m:
                        d, mth, y = int(m.group(1)), int(m.group(2)), m.group(3)
                    else:
                        m = ISO_DATE_RE.match(s)
                        if m:
                            y, mth, d = m.group(1), int(m.group(2)), int(m.group(3))
                        else:
                            m = THAI_TEXT_DATE_RE.match(s)
                            if m:
                                d = int(m.group(1)); mth = THAI_MONTHS.index(m.group(2)) % 12 + 1; y = m.group(3)
                    if y is not None and 1 <= mth <= 12 and 1 <= d <= 31:
                        yi = int(y)
                        if len(y) == 2:
                            yi = 2500 + yi if yi >= 40 else 2000 + yi  # 2 หลัก: >=40 เดาเป็น พ.ศ. 25xx
                        is_be = yi >= 2400
                        ce = yi - 543 if is_be else yi
                        try:
                            iso = datetime.date(ce, mth, d).isoformat()
                        except ValueError:
                            iso = None
                        tag = f'{addr} = "{s}"'
                        (text_date_be if is_be else text_date).add(tag)
                        if iso:
                            text_date_fix.append({"sheet": sn, "cell": cell.coordinate, "text": s,
                                                  "iso": iso, "was_be": is_be,
                                                  "ambiguous_dm": bool(TEXT_DATE_RE.match(s)) and d <= 12 and mth <= 12 and d != mth})
                # #### (ประมาณการ) เฉพาะตัวเลข/วันที่
                if (cell.row, cell.column) not in merged and dv is not None and isinstance(dv, (int, float, datetime.date)) and not isinstance(dv, bool):
                    w = ws.column_dimensions[cell.column_letter].width if cell.column_letter in ws.column_dimensions and ws.column_dimensions[cell.column_letter].width else 8.43
                    if ws.column_dimensions.get(cell.column_letter) is not None and ws.column_dimensions[cell.column_letter].hidden:
                        continue
                    if isinstance(dv, datetime.date):
                        txt = "00/00/0000"
                    else:
                        dec = 0
                        m = re.search(r"0\.(0+)", fmt or "")
                        if m: dec = len(m.group(1))
                        txt = f"{dv:,.{dec}f}" if "," in (fmt or "") else (f"{dv:.{dec}f}" if dec or fmt not in (None, "General") else str(dv))
                        if "%" in (fmt or ""): txt = f"{dv*100:.{dec}f}%"
                    need = sum(0.6 if ch in ",./:-" else 1 for ch in txt)
                    if need > w:
                        k = (sn, cell.column_letter)
                        if k not in narrow_cols or need > narrow_cols[k][1]:
                            narrow_cols[k] = (w, need, cell.coordinate)

        # B: สูตรไม่สม่ำเสมอ + hardcode ในคอลัมน์สูตร
        for col, rows in col_formulas.items():
            rs = sorted(rows)
            # แบ่งเป็นช่วงต่อเนื่อง (อนุญาตให้มีช่องว่าง/ค่าคั่นได้ไม่เกิน 1 แถว)
            runs, cur = [], [rs[0]]
            for r in rs[1:]:
                if r - cur[-1] <= 2: cur.append(r)
                else: runs.append(cur); cur = [r]
            runs.append(cur)
            for run in runs:
                if len(run) < 3: continue
                pats = Counter(rows[r] for r in run)
                top, n = pats.most_common(1)[0]
                ref_row = next(r for r in run if rows[r] == top)
                cl = get_column_letter(col)
                def suggest(r):
                    try:
                        return Translator(col_ftext[col][ref_row], origin=f"{cl}{ref_row}").translate_formula(f"{cl}{r}")
                    except Exception:
                        return None
                def maybe_intended(r, ftxt=None):
                    lab = " ".join(str(vals[r]) for c2, vals in col_values.items()
                                   if c2 < col and isinstance(vals.get(r), str)).lower()
                    if TOTAL_RE.search(lab) or MANUAL_RE.search(lab):
                        return True
                    m = re.match(r"=\s*SUM\(\s*\$?([A-Z]{1,3})\$?(\d+)\s*:\s*\$?([A-Z]{1,3})\$?(\d+)\s*\)\s*$", ftxt or "", re.I)
                    return bool(m and m.group(1) == m.group(3) == cl and int(m.group(4)) < r)
                if n / len(run) >= 0.6 and n < len(run):
                    for r in run:
                        if rows[r] != top:
                            sg = suggest(r)
                            if maybe_intended(r, col_ftext[col][r]):
                                intended.add(f"{sn}!{cl}{r}: {col_ftext[col][r][:60]}")
                                continue
                            inconsistent.add(f"{sn}!{cl}{r}: {col_ftext[col][r][:80]} (เซลล์อื่นเป็นแบบ {sg})")
                            formula_fix.append({"sheet": sn, "cell": f"{cl}{r}", "kind": "inconsistent",
                                                "current": col_ftext[col][r], "suggested": sg})
                if n / len(run) >= 0.6:
                    for r in range(run[0], run[-1] + 1):
                        v = col_values[col].get(r)
                        if isinstance(v, (int, float)) and not isinstance(v, bool):
                            sg = suggest(r)
                            if maybe_intended(r):
                                intended.add(f"{sn}!{cl}{r} = {v}")
                                continue
                            hardcode.add(f"{sn}!{cl}{r} = {v} (ควรเป็น {sg})")
                            formula_fix.append({"sheet": sn, "cell": f"{cl}{r}", "kind": "hardcode",
                                                "current": v, "suggested": sg})

        # F: ตำแหน่งเคอร์เซอร์/มุมมอง
        sv = ws.sheet_view
        ac = sv.selection[0].activeCell if sv.selection else "A1"
        tl = sv.topLeftCell or "A1"
        has_freeze = ws.freeze_panes is not None
        if (ac or "A1") != "A1" or (tl != "A1" and not has_freeze):
            not_a1.append(f"{sn} (เคอร์เซอร์ {ac or 'A1'}, มุมบนซ้าย {tl})")

    # ---------- สรุปเป็น findings ----------
    for e, loc in err_cells.items():
        findings.append(finding("A", "error_value", f"เซลล์แสดง {e}", loc))
    if err_in_formula.count:
        findings.append(finding("A", "ref_in_formula", "สูตรมี #REF! อยู่ในตัวสูตร (อ้างอิงถูกลบไปแล้ว)", err_in_formula))
    if self_ref.count:
        findings.append(finding("A", "self_reference", "สูตรอ้างอิงเซลล์ตัวเอง (circular)", self_ref))
    if inconsistent.count:
        findings.append(finding("B", "inconsistent_formula", "สูตรต่างจากเซลล์ข้างเคียงในคอลัมน์เดียวกัน", inconsistent))
    if hardcode.count:
        findings.append(finding("B", "hardcode_in_formula_column", "ตัวเลขพิมพ์ทับในคอลัมน์ที่เป็นสูตร", hardcode))
    if intended.count:
        findings.append(finding("B", "likely_intended", "สูตรต่าง/ตัวเลขคงที่ในแถวยอดรวมหรือแถวกรอกเอง (อาจตั้งใจ — รายงานบรรทัดเดียว)", intended, sev="P3"))

    # external links / defined names / validation / CF
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        ext_targets = []
        for n in names:
            if re.match(r"xl/externalLinks/_rels/externalLink\d+\.xml\.rels", n):
                ext_targets += re.findall(r'Target="([^"]+)"', z.read(n).decode("utf-8", "ignore"))
        threaded = [n for n in names if n.startswith("xl/threadedComments/")]
        app_xml = z.read("docProps/app.xml").decode("utf-8", "ignore") if "docProps/app.xml" in names else ""
        parts = {"charts": len([n for n in names if n.startswith("xl/charts/chart")]),
                 "drawings": len([n for n in names if n.startswith("xl/drawings/drawing")]),
                 "images": len([n for n in names if n.startswith("xl/media/")]),
                 "pivot_tables": len([n for n in names if n.startswith("xl/pivotTables/")]),
                 "slicers": len([n for n in names if n.startswith("xl/slicers/")]),
                 "tables": len([n for n in names if n.startswith("xl/tables/")]),
                 "ctrl_props": len([n for n in names if n.startswith("xl/ctrlProps/")]),
                 "power_query": int(any("connections.xml" in n for n in names) or any(n.startswith("customXml/") and "DataMashup" in z.read(n).decode("utf-16", "ignore") for n in names if n.endswith(".xml") and n.startswith("customXml/item"))),
                 "vba": int("xl/vbaProject.bin" in names)}
    if ext_targets or ext_formula.count:
        loc = Loc()
        for t in ext_targets: loc.add(f"ลิงก์ไปไฟล์: {t}")
        for x in ext_formula.items: loc.add(x)
        loc.count = len(ext_targets) + ext_formula.count
        findings.append(finding("B", "external_link", "อ้างอิงไฟล์ภายนอก (ลูกค้าเปิดแล้วจะเตือน/ค่าเพี้ยน)", loc))

    bad_names = Loc()
    try:
        items = list(wb.defined_names.items())
    except Exception:
        items = []
    for ws in wb.worksheets:
        try: items += [(f"{ws.title}!{k}", v) for k, v in ws.defined_names.items()]
        except Exception: pass
    for k, dn in items:
        val = getattr(dn, "attr_text", None) or getattr(dn, "value", "") or ""
        if k.startswith("_xlnm.") and "#REF!" not in val:
            continue
        if "#REF!" in val or re.search(r"\[\d+\]|\.xls", val, re.I):
            bad_names.add(f"{k} = {val[:80]}")
    if bad_names.count:
        findings.append(finding("B", "bad_defined_name", "Named Range เสีย (#REF! หรือชี้ไปไฟล์อื่น)", bad_names))

    bad_dv = Loc(); bad_cf = Loc()
    for ws in wb.worksheets:
        for dv in ws.data_validations.dataValidation:
            if "#REF!" in f"{dv.formula1 or ''}{dv.formula2 or ''}":
                bad_dv.add(f"{ws.title}!{dv.sqref}")
        for rng, rules in ws.conditional_formatting._cf_rules.items():
            for rule in rules:
                if any("#REF!" in (x or "") for x in (rule.formula or [])):
                    bad_cf.add(f"{ws.title}!{rng.sqref}")
    if bad_dv.count:
        findings.append(finding("B", "bad_validation", "Data Validation อ้างอิงเสีย (#REF!)", bad_dv))
    if bad_cf.count:
        findings.append(finding("B", "bad_conditional_format", "Conditional Formatting อ้างอิงเสีย (#REF!)", bad_cf))

    # C: VBA
    vba = check_vba(path, wb.sheetnames)
    findings += vba.pop("findings", [])

    # D
    for name, loc in sorted(ver_funcs.items()):
        findings.append(finding("D", "unsupported_function", f"ฟังก์ชัน {name} (ต้องใช้ Excel {FUNC_MIN_VER[name] if FUNC_MIN_VER[name] != 365 else '365'} ขึ้นไป)", loc, function=name))
    if spill.count and target < 2021:
        findings.append(finding("D", "spill_reference", "อ้างอิงแบบ spill (A1#) ใช้ไม่ได้ใน Excel ก่อน 2021", spill))

    # E
    if be_fmt.count:
        findings.append(finding("E", "be_number_format", "รูปแบบวันที่ตั้งเป็นปี พ.ศ.", be_fmt))
    if be_year_val.count:
        findings.append(finding("E", "be_year_in_date_value", "ค่าวันที่ปีเกิน 2400 (น่าจะพิมพ์ปี พ.ศ. ลงในช่องวันที่)", be_year_val))
    if text_date_be.count:
        findings.append(finding("E", "text_date_be", "วันที่พิมพ์เป็นข้อความ ปี พ.ศ.", text_date_be))
    if text_date.count:
        findings.append(finding("E", "text_date", "วันที่พิมพ์เป็นข้อความ (ปี ค.ศ.) สูตรคำนวณวันจะใช้ไม่ได้", text_date))
    if f543.count:
        findings.append(finding("E", "be_in_formula", "สูตรแปลงเป็น พ.ศ. (+543 / รูปแบบ bbbb)", f543))
    if vba.get("be_lines", {}).get("count"):
        loc = vba["be_lines"]
        findings.append({"group": "E", "key": "be_in_vba", "title": "VBA แปลงเป็น พ.ศ.", **loc})

    # F
    hidden = [f"{ws.title} ({ws.sheet_state})" for ws in wb.worksheets if ws.sheet_state != "visible"]
    if hidden:
        findings.append(finding("F", "hidden_sheet", "ชีตที่ซ่อนอยู่", hidden))
    first_visible = next((i for i, ws in enumerate(wb.worksheets) if ws.sheet_state == "visible"), 0)
    active = wb.active.title if wb.active else None
    if wb.active is not None and wb.worksheets.index(wb.active) != first_visible:
        findings.append(finding("F", "active_sheet", "เปิดไฟล์มาไม่ได้อยู่ชีตแรก", [f"ชีตที่เปิดอยู่: {active}"]))
    grouped = [ws.title for ws in wb.worksheets if ws.sheet_view.tabSelected]
    if len(grouped) > 1:
        findings.append(finding("F", "grouped_sheets", "มีหลายชีตถูกเลือกพร้อมกัน (Group) ลูกค้าพิมพ์แล้วจะเขียนทับหลายชีต", grouped))
    if not_a1:
        findings.append(finding("F", "cursor_not_a1", "ชีตที่เคอร์เซอร์/มุมมองไม่อยู่ที่ A1", not_a1))
    if comments.count or threaded:
        loc = comments
        if threaded: loc.add(f"มี threaded comment {len(threaded)} ไฟล์")
        findings.append(finding("F", "comments", "Comment/Note ค้างในไฟล์", loc))
    p = wb.properties
    company = (re.search(r"<Company>([^<]*)</Company>", app_xml) or [None, ""])[1]
    meta = [f"{k}: {v}" for k, v in (("ผู้สร้าง", p.creator), ("แก้ไขล่าสุดโดย", p.lastModifiedBy), ("บริษัท", company), ("Title", p.title)) if v]
    if meta:
        findings.append(finding("F", "metadata", "ข้อมูลผู้สร้างใน metadata (ดูว่าต้องการให้ลูกค้าเห็นไหม)", meta))
    narrow = Loc()
    for (sn_, cl), (w, need, ex) in sorted(narrow_cols.items()):
        narrow.add(f"{sn_} คอลัมน์ {cl} (กว้าง {w:.1f} ต้องการราว {need:.0f} เช่น {ex})")
    if narrow.count:
        findings.append(finding("F", "narrow_column", "คอลัมน์แคบจนอาจแสดง #### (ประมาณการ ควรเปิดดูยืนยัน)", narrow))

    # ---------- check เพิ่มเติม (qa_extra.py) ----------
    extra_info = {}
    with zipfile.ZipFile(path) as z:
        for name_, fn_ in (("criteria", lambda: qa_extra.criteria_checks(wb, wbv, formula_text)),
                           ("dropdown", lambda: qa_extra.dropdown_checks(wb, wbv, z)),
                           ("power_query", lambda: qa_extra.power_query_checks(z))):
            try:
                r_ = fn_()
            except Exception as e:  # check เสริมพังต้องไม่ทำให้ทั้งรอบพัง
                findings.append({"group": "I", "key": f"check_failed_{name_}", "title": f"check {name_} รันไม่ผ่าน: {str(e)[:120]}", "sev": "P3", "count": 1, "examples": []})
                continue
            if isinstance(r_, tuple):
                findings += r_[0]; extra_info[name_] = r_[1]
            else:
                findings += r_
    try:
        findings += qa_extra.period_text_checks(wb, extra_info.get("dropdown", []), formula_text)
        findings += qa_extra.calendar_checks(wb, wbv, a.calendar, is_buddhist_format, formula_text)
    except Exception as e:
        findings.append({"group": "I", "key": "check_failed_period", "title": f"check ช่วงเวลา/ปฏิทิน รันไม่ผ่าน: {str(e)[:120]}", "sev": "P3", "count": 1, "examples": []})
    if extra_info.get("power_query"):
        parts["power_query_sources"] = extra_info["power_query"]
    if any(f["key"] == "bad_validation_x14" for f in findings):
        findings = [f for f in findings if f["key"] != "bad_validation"]  # x14 รวมแบบปกติแล้ว

    # ปฏิทิน พ.ศ. ตามลูกค้า: ยุบเรื่องรูปแบบ พ.ศ. เหลือบรรทัดเดียว
    if a.calendar == "be":
        quiet = [f for f in findings if f["key"] in ("be_number_format", "be_in_formula", "be_in_vba")]
        findings = [f for f in findings if f not in quiet]
        if quiet:
            findings.append({"group": "E", "key": "be_by_policy", "sev": "OK",
                             "title": "ใช้ปี พ.ศ. ตามที่ลูกค้าต้องการ (" + ", ".join(f"{f['key']} {f['count']}" for f in quiet) + ")",
                             "count": sum(f["count"] for f in quiet), "examples": []})

    for f_ in findings:
        f_.setdefault("sev", SEV_DEFAULT.get(f_["key"], {"A": "P0", "B": "P1", "C": "P1", "D": "P2", "E": "P2", "F": "P3", "G": "P2"}.get(f_["group"], "P3")))

    notes = []
    if stats["formula_cells"] and stats["formula_no_cached"] == stats["formula_cells"]:
        notes.append("ไฟล์ยังไม่เคยถูกคำนวณใน Excel (ไม่มีค่าผลลัพธ์ของสูตรเลย) กลุ่ม A ตรวจค่า error ไม่ได้ ต้องคำนวณใหม่ก่อน")
    elif stats["formula_no_cached"]:
        notes.append(f"สูตร {stats['formula_no_cached']} เซลล์ไม่มีค่าผลลัพธ์ที่บันทึกไว้")
    notes.append("circular ตรวจเฉพาะกรณีสูตรอ้างตัวเองโดยตรง ไม่ตรวจวงอ้างข้ามหลายเซลล์")

    out = {"file": display_path or path, "status": "ok", "encrypted": encrypted, "excel_target": a.excel, "calendar": a.calendar, "stats": stats, "parts": parts,
           "vba": vba, "findings": findings, "notes": notes,
           "fix_data": {"text_dates": text_date_fix, "be_formats": be_fmt_fix,
                        "be_year_values": be_year_fix, "formulas": formula_fix}}
    s = json.dumps(out, ensure_ascii=False, indent=1, default=str)
    if a.out:
        open(a.out, "w", encoding="utf-8").write(s)
        # สรุปสั้นบนหน้าจอ
        c = Counter(f["group"] for f in findings)
        print(json.dumps({"groups": dict(c), "vba": vba.get("status"), "notes": notes, "parts": parts}, ensure_ascii=False))
    else:
        print(s)

if __name__ == "__main__":
    main()
