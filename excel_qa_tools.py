#!/usr/bin/env python3
"""excel_qa_tools.py — เครื่องมือเสริมของชุด excel-qa (อ่านอย่างเดียว ไม่แก้ไฟล์)

  dump   <file> [--out f.tsv]            ทุกเซลล์: ค่า/สูตร + Locked/Hidden + สถานะชีต
  diff   <old> <new> [--all] [--json]    เทียบ 2 เวอร์ชัน (ค่า สูตร การล็อก ชีต VBA ชิ้นส่วน)
  lookup <file> <Sheet!A1[:C3]> ...      อ่านค่าจริง (ค่าที่คำนวณไว้ + สูตร) — ใช้ก่อนพูดตัวเลข/ชื่อ
  find-sheet <folder> <ชื่อชีต> ...       หาว่าชื่อชีตอยู่ไฟล์ไหน
  buttons <file>                         ปุ่ม/รูปที่ผูกมาโคร + ตำแหน่ง
  vba-diff <old> <new>                   เทียบโค้ด VBA (ตัดบรรทัดที่ไม่ใช่โค้ดออกก่อน)
  lock   <file> [--plan S1,S2|all]       ตรวจก่อนล็อกชีต/ซ่อนสูตร
  final  <file> --base <file_ที่ตรวจผ่าน>   ตรวจไฟล์ส่งฉบับสุดท้าย

ไฟล์มีรหัส: QA_PASSWORD (ไฟล์แรก) / QA_PASSWORD2 (ไฟล์ที่สอง) — ห้าม echo รหัส
"""
import argparse, datetime, difflib, json, os, re, sys, warnings, zipfile
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
warnings.filterwarnings("ignore")
import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries
import xparts
import qa_extra

MANUAL_RE = re.compile(r"manual|กรอก|ใส่เอง|ปรับ|key\s*in|input|adjust", re.I)
PROT_FLAGS = ["selectLockedCells", "selectUnlockedCells", "formatCells", "formatColumns", "formatRows",
              "insertColumns", "insertRows", "insertHyperlinks", "deleteColumns", "deleteRows",
              "sort", "autoFilter", "pivotTables", "objects", "scenarios"]


# ---------- เปิดไฟล์ ----------
class Opened:
    def __init__(self, path, env="QA_PASSWORD"):
        self.path = path
        self.status, self.real, self._clean = xparts.open_book(path, env)
        if self.status != "ok":
            msg = {"encrypted": "ไฟล์มีรหัส ตั้ง %s ก่อน" % env, "wrong_password": "รหัสไม่ถูก",
                   "legacy_xls": "ไฟล์ .xls รุ่นเก่า", "unreadable": "อ่านไฟล์ไม่ได้"}[self.status]
            print(f"STATUS {self.status}: {path} — {msg}")
            sys.exit(2)
        self._wb = self._wbv = None

    @property
    def wb(self):
        if self._wb is None:
            self._wb = openpyxl.load_workbook(self.real)
        return self._wb

    @property
    def wbv(self):
        if self._wbv is None:
            self._wbv = openpyxl.load_workbook(self.real, data_only=True)
        return self._wbv

    def zip(self):
        return zipfile.ZipFile(self.real)

    def close(self):
        self._clean()


def ftext(v):
    if isinstance(v, str) and v.startswith("="):
        return v
    t = getattr(v, "text", None)
    return ("{=" + t.lstrip("=") + "}") if isinstance(t, str) else None


def vrepr(v):
    if v is None:
        return ""
    f = ftext(v)
    if f:
        return f
    if isinstance(v, datetime.datetime):
        return v.isoformat(sep=" ")
    if isinstance(v, (datetime.date, datetime.time)):
        return v.isoformat()
    return repr(v) if isinstance(v, float) else str(v)


def prot_info(ws):
    p = ws.protection
    has_pw = bool(getattr(p, "password", None) or getattr(p, "hashValue", None))
    # openpyxl/XML: ค่า True = "ห้าม" ทำสิ่งนั้น (ระวังอ่านกลับด้าน) -> แปลงเป็นรายการที่ "อนุญาต"
    allowed = [f for f in PROT_FLAGS if getattr(p, f, None) is False]
    return {"protected": bool(p.sheet), "password": has_pw, "allowed": allowed}


# ---------- dump ----------
def dump_book(op, fmt=False):
    wb = op.wb
    lines = []
    for ws in wb.worksheets:
        pi = prot_info(ws)
        lines.append(("S", ws.title, ws.sheet_state, int(pi["protected"]), int(pi["password"]), ",".join(pi["allowed"])))
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                L = "L" if c.protection.locked else "U"
                H = "H" if c.protection.hidden else "-"
                if v is None and L == "L" and H == "-":
                    continue
                rec = ("C", ws.title, c.coordinate, vrepr(v), L + H)
                if fmt and c.number_format != "General":
                    rec += (c.number_format,)
                lines.append(rec)
    return lines


def cmd_dump(a):
    op = Opened(a.file)
    try:
        lines = dump_book(op, a.fmt)
        txt = "\n".join("\t".join(str(x) for x in l) for l in lines)
        if a.out:
            open(a.out, "w", encoding="utf-8").write(txt + "\n")
            print(f"เขียน {len(lines)} บรรทัด -> {a.out}")
        else:
            print(txt)
    finally:
        op.close()


# ---------- VBA ----------
def vba_modules(path):
    """{module: code} ตัดบรรทัด Attribute/ว่าง (คืน None ถ้าไม่มี VBA, {} ถ้าอ่านไม่ได้)"""
    with zipfile.ZipFile(path) as z:
        if "xl/vbaProject.bin" not in z.namelist():
            return None
    try:
        from oletools.olevba import VBA_Parser
        vp = VBA_Parser(path)
        out = {}
        for (_f, _s, name, code) in vp.extract_macros():
            if isinstance(code, bytes):
                code = code.decode("cp874", "ignore")
            body = [l.rstrip() for l in code.splitlines() if l.strip() and not l.strip().startswith("Attribute ")]
            if body:
                out[name] = "\n".join(body)
        vp.close()
        return out
    except Exception:
        return {}


def vba_locked(path):
    try:
        import olefile, io
        with zipfile.ZipFile(path) as z:
            if "xl/vbaProject.bin" not in z.namelist():
                return None
            o = olefile.OleFileIO(io.BytesIO(z.read("xl/vbaProject.bin")))
        proj = o.openstream("PROJECT").read().decode("latin-1", "ignore")
        m = re.search(r'DPB="([0-9A-Fa-f]*)"', proj)
        return bool(m and len(m.group(1)) > 8)
    except Exception:
        return None


def vba_diff_text(p1, p2):
    a, b = vba_modules(p1), vba_modules(p2)
    if a is None and b is None:
        return "ไม่มี VBA ทั้งสองไฟล์", True
    if a is None or b is None:
        return f"VBA: ไฟล์เก่า{'ไม่มี' if a is None else 'มี'} / ไฟล์ใหม่{'ไม่มี' if b is None else 'มี'}", False
    if a == {} or b == {}:
        return "อ่านโค้ด VBA ไม่ได้ (อาจล็อก) เทียบไม่ได้", False
    out = []
    for m in sorted(set(a) | set(b)):
        if a.get(m) == b.get(m):
            continue
        if m not in a:
            out.append(f"+ โมดูลใหม่: {m} ({len(b[m].splitlines())} บรรทัด)"); continue
        if m not in b:
            out.append(f"- โมดูลหาย: {m}"); continue
        d = list(difflib.unified_diff(a[m].splitlines(), b[m].splitlines(), m + " (เก่า)", m + " (ใหม่)", lineterm="", n=1))
        out.append("\n".join(d))
    return ("VBA เหมือนเดิมทุกโมดูล" if not out else "\n".join(out)), not out


def cmd_vba_diff(a):
    o1, o2 = Opened(a.old), Opened(a.new, "QA_PASSWORD2")
    try:
        print(vba_diff_text(o1.real, o2.real)[0])
    finally:
        o1.close(); o2.close()


# ---------- parts ----------
def parts_of(path):
    with zipfile.ZipFile(path) as z:
        n = z.namelist()
        return {"charts": sum(x.startswith("xl/charts/chart") for x in n),
                "images": sum(x.startswith("xl/media/") for x in n),
                "pivot_tables": sum(x.startswith("xl/pivotTables/") for x in n),
                "slicers": sum(x.startswith("xl/slicers/") for x in n),
                "tables": sum(x.startswith("xl/tables/") for x in n),
                "ctrl_props": sum(x.startswith("xl/ctrlProps/") for x in n),
                "buttons": len(xparts.buttons(z)),
                "vba": int("xl/vbaProject.bin" in n)}


# ---------- diff ----------
def diff_books(o1, o2, limit=20):
    A = dump_book(o1, fmt=True); B = dump_book(o2, fmt=True)
    def split(lines):
        sh, cells = {}, {}
        for l in lines:
            if l[0] == "S":
                sh[l[1]] = l[2:]
            else:
                cells[(l[1], l[2])] = l[3:]
        return sh, cells
    sa, ca = split(A); sb, cb = split(B)
    res = {"sheets_added": [s for s in sb if s not in sa], "sheets_removed": [s for s in sa if s not in sb],
           "sheet_changes": [], "cells": {}}
    for s in sa:
        if s in sb and sa[s] != sb[s]:
            old, new = sa[s], sb[s]
            ch = []
            if old[0] != new[0]: ch.append(f"สถานะ {old[0]} → {new[0]}")
            if old[1] != new[1]: ch.append("ล็อกชีต " + ("เปิด" if new[1] else "ปลด"))
            if old[2] != new[2]: ch.append("รหัสชีต " + ("ตั้ง" if new[2] else "ลบ"))
            if old[3] != new[3]: ch.append(f"ตัวเลือกที่อนุญาต [{old[3]}] → [{new[3]}]")
            res["sheet_changes"].append(f"{s}: " + ", ".join(ch))
    keys = set(ca) | set(cb)
    per = defaultdict(lambda: defaultdict(list))
    for k in sorted(keys, key=lambda k: (k[0], k[1])):
        s, coord = k
        if s in res["sheets_added"] or s in res["sheets_removed"]:
            continue
        va = ca.get(k, ("", "L-", "General")); vb = cb.get(k, ("", "L-", "General"))
        va = tuple(va) + ("General",) * (3 - len(va)); vb = tuple(vb) + ("General",) * (3 - len(vb))
        if va[0] != vb[0]:
            kind = "สูตร" if (va[0].startswith("=") or vb[0].startswith("=")) else "ค่า"
            per[s][kind].append(f"{coord}: {va[0][:50] or '(ว่าง)'} → {vb[0][:50] or '(ว่าง)'}")
        if va[1] != vb[1]:
            per[s]["ล็อก/ซ่อน"].append(f"{coord}: {va[1]} → {vb[1]}")
        if va[2] != vb[2]:
            per[s]["รูปแบบตัวเลข"].append(f"{coord}: {va[2]} → {vb[2]}")
    for s, kinds in per.items():
        res["cells"][s] = {k: {"count": len(v), "examples": v[:limit] if limit else v} for k, v in kinds.items()}
    vtxt, vsame = vba_diff_text(o1.real, o2.real)
    pa, pb = parts_of(o1.real), parts_of(o2.real)
    res["vba"] = {"same": vsame, "detail": vtxt}
    res["parts_changed"] = {k: f"{pa[k]} → {pb[k]}" for k in pa if pa[k] != pb[k]}
    return res


def print_diff(res):
    if res["sheets_added"]: print("ชีตใหม่:", ", ".join(res["sheets_added"]))
    if res["sheets_removed"]: print("ชีตหาย:", ", ".join(res["sheets_removed"]))
    for x in res["sheet_changes"]: print("ชีตเปลี่ยน:", x)
    if not res["cells"]:
        print("เซลล์: ไม่มีอะไรเปลี่ยน")
    for s, kinds in res["cells"].items():
        print(f"\n[{s}]")
        for k, d in kinds.items():
            print(f"  {k} {d['count']} เซลล์")
            for e in d["examples"]: print("    " + e)
            if d["count"] > len(d["examples"]): print(f"    … และอีก {d['count'] - len(d['examples'])} (ใช้ --all)")
    print("\nVBA:", res["vba"]["detail"] if len(res["vba"]["detail"]) < 3000 else res["vba"]["detail"][:3000] + "\n…")
    print("ชิ้นส่วน (chart/รูป/ปุ่ม…):", res["parts_changed"] or "เท่าเดิม")


def cmd_diff(a):
    o1, o2 = Opened(a.old), Opened(a.new, "QA_PASSWORD2")
    try:
        res = diff_books(o1, o2, 0 if a.all else 20)
        print(json.dumps(res, ensure_ascii=False, indent=1)) if a.json else print_diff(res)
    finally:
        o1.close(); o2.close()


# ---------- lookup ----------
def cmd_lookup(a):
    op = Opened(a.file)
    try:
        for ref in a.refs:
            sn, rng = ref.rsplit("!", 1)
            sn = sn.strip("'")
            if sn not in op.wb.sheetnames:
                print(f"{ref}: ไม่มีชีต {sn} (มี: {', '.join(op.wb.sheetnames)})"); continue
            ws, wsv = op.wb[sn], op.wbv[sn]
            c1, r1, c2, r2 = range_boundaries(rng.replace("$", ""))
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    cell = ws.cell(r, c); v = cell.value; f = ftext(v)
                    shown = wsv.cell(r, c).value
                    print(f"{sn}!{cell.coordinate}\t{vrepr(shown) if f else vrepr(v)}" + (f"\t{f}" if f else ""))
    finally:
        op.close()


# ---------- find-sheet ----------
def cmd_find_sheet(a):
    want = [w.lower() for w in a.names]
    for root, _d, files in os.walk(a.folder):
        for fn in sorted(files):
            if not re.search(r"\.xls[xmb]?$", fn, re.I) or fn.startswith("~$"):
                continue
            p = os.path.join(root, fn)
            st = xparts.file_status(p)
            if st != "ok":
                print(f"?  {p} ({'มีรหัส ตรวจไม่ได้' if st == 'encrypted' else st})"); continue
            names = xparts.sheet_names_fast(p) or []
            hit = [n for n in names if any(w in n.lower() for w in want)]
            if hit:
                print(f"✓  {p}: {', '.join(hit)}")


# ---------- buttons ----------
def cmd_buttons(a):
    op = Opened(a.file)
    try:
        with op.zip() as z:
            bs = xparts.buttons(z)
        if not bs:
            print("ไม่มีปุ่ม/รูปที่ผูกมาโคร")
        for b in bs:
            print(f"{b['sheet']}!{b['anchor']}\t{b['kind']}\t{b['name']}\t→ {b['macro']}")
    finally:
        op.close()


# ---------- lock audit ----------
WRITE_RE = re.compile(r"\.(ClearContents|ClearFormats|Clear|Delete|Insert|Sort|AutoFilter|PasteSpecial|Cut|"
                      r"Value2?\s*=|Formula\w*\s*=|NumberFormat\s*=|Interior\.|Font\.|RemoveDuplicates|"
                      r"ListObjects|Columns\(.*\)\.AutoFit|EntireRow\.(?:Delete|Insert))", re.I)
SHEET_REF_RE = re.compile(r'(?:Sheets|Worksheets)\s*\(\s*"([^"]+)"\s*\)', re.I)


def scan_vba(modules, codenames, sheetnames):
    """แยกเป็น procedure: อ้างชีตไหน, เขียนไหม, มี Unprotect ไหม, ใช้ ActiveSheet, Sheet.Copy, Sheets.Add"""
    procs = []
    for mod, code in (modules or {}).items():
        cur = None
        for line in code.splitlines():
            s = line.split("'")[0] if '"' not in line else line
            m = re.match(r"\s*(?:Public\s+|Private\s+)?(?:Sub|Function)\s+(\w+)", s, re.I)
            if m:
                cur = {"module": mod, "name": m.group(1), "sheets": set(), "writes": False, "unprotect": False,
                       "active": False, "sheet_copy": False, "sheet_add": False, "lines": []}
                procs.append(cur); continue
            if cur is None:
                continue
            for sm in SHEET_REF_RE.finditer(s):
                cur["sheets"].add(sm.group(1))
            for cn, sn in codenames.items():
                if re.search(r"(?<![\w.])" + re.escape(cn) + r"\.", s):
                    cur["sheets"].add(sn)
            if WRITE_RE.search(s):
                cur["writes"] = True; cur["lines"].append(s.strip()[:100])
            if re.search(r"\.Unprotect\b", s, re.I): cur["unprotect"] = True
            if re.search(r"\bActive(?:Sheet|Cell)\b|\bSelection\.", s, re.I): cur["active"] = True
            if re.search(r"\.Copy\s*(?:$|Before|After|:)", s, re.I) and re.search(r"Sheets|Worksheets|ActiveSheet|\w+\.Copy\s*$", s, re.I):
                cur["sheet_copy"] = True
            if re.search(r"(?:Sheets|Worksheets)\.Add\b|\.Move\b", s, re.I): cur["sheet_add"] = True
    return procs


def lock_audit(op, plan):
    wb, z = op.wb, op.zip()
    dvs = qa_extra.all_validations(wb, z)
    btns = xparts.buttons(z)
    z.close()
    vis = [ws.title for ws in wb.worksheets]
    if plan == "all":
        targets = [ws.title for ws in wb.worksheets if ws.sheet_state == "visible"]
    elif plan:
        targets = [s.strip() for s in plan.split(",") if s.strip()]
    else:
        targets = [ws.title for ws in wb.worksheets if ws.protection.sheet]
    out = {"targets": targets, "sheets": {}, "issues": []}
    def issue(sev, key, title, items):
        if items:
            out["issues"].append({"sev": sev, "key": key, "title": title, "count": len(items), "examples": items[:30]})

    for ws in wb.worksheets:
        pi = prot_info(ws)
        out["sheets"][ws.title] = {"state": ws.sheet_state, **pi, "planned": ws.title in targets}

    # 1 สูตรต้อง Locked + Hidden
    f_bad = []
    referenced = defaultdict(set)   # sheet -> {(col,row)} เซลล์คงที่ที่สูตรอ้าง
    bk = qa_extra.Book(wb, op.wbv)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                f = ftext(c.value)
                if not f:
                    continue
                if ws.title in targets and (not c.protection.locked or not c.protection.hidden):
                    f_bad.append(f"{ws.title}!{c.coordinate} ({'ไม่ล็อก' if not c.protection.locked else ''}{' ' if not c.protection.locked and not c.protection.hidden else ''}{'ไม่ซ่อน' if not c.protection.hidden else ''})")
                for m in qa_extra.REF_RE.finditer(qa_extra.STR_RE.sub('""', f)):
                    res = bk.resolve(m.group(1), m.group(2), m.group(3), ws.title)
                    if res and (res[3] - res[1] + 1) * (res[4] - res[2] + 1) <= 60:
                        for cc in range(res[1], res[3] + 1):
                            for rr in range(res[2], res[4] + 1):
                                referenced[res[0]].add((cc, rr))
    issue("P1", "formula_not_hidden", "สูตรที่ยังไม่ Locked+Hidden (ลูกค้าจะเห็น/แก้สูตรได้หลังล็อก)", f_bad)

    # 2 dropdown ต้องปลดล็อก
    dv_bad = []
    for d in dvs:
        if d["sheet"] not in targets or not d["sqref"]:
            continue
        ws = wb[d["sheet"]]
        for part in d["sqref"].split():
            try:
                c1, r1, c2, r2 = range_boundaries(part)
            except Exception:
                continue
            for r in range(r1, min(r2, r1 + 200) + 1):
                for c in range(c1, c2 + 1):
                    if ws.cell(r, c).protection.locked:
                        dv_bad.append(f"{d['sheet']}!{get_column_letter(c)}{r}"); break
                else:
                    continue
                break
    issue("P1", "dropdown_locked", "Dropdown/ช่องเลือกที่ยังล็อก (ลูกค้าจะเลือกไม่ได้)", dv_bad)

    # 3 ค่าคงที่ที่สูตรอ้าง (ช่องกรอก/คีย์/หัวคอลัมน์) และแถวกรอกเอง ที่ยังล็อก -> ถาม user
    inp, manual = [], []
    skip = {(d["sheet"], c) for d in dvs for part in (d["sqref"] or "").split() for c in _cells(part)}
    for sn in targets:
        if sn not in wb.sheetnames:
            continue
        ws = wb[sn]
        for (c, r) in sorted(referenced.get(sn, ()), key=lambda x: (x[1], x[0])):
            cell = ws.cell(r, c)
            if (sn, cell.coordinate) in skip or (sn, cell.coordinate) in manual_cells(ws):
                continue
            if cell.value is not None and not ftext(cell.value) and cell.protection.locked:
                inp.append(f"{sn}!{cell.coordinate} = {str(cell.value)[:30]}")
        for row in ws.iter_rows():
            lab = " ".join(str(c.value) for c in row if isinstance(c.value, str) and not c.value.startswith("="))
            if lab and MANUAL_RE.search(lab):
                locked = [c.coordinate for c in row if isinstance(c.value, (int, float)) and not isinstance(c.value, bool) and c.protection.locked]
                if locked:
                    manual.append(f"{sn} แถว {row[0].row} ({lab[:30]}): {', '.join(locked[:6])}")
    issue("ASK", "referenced_constants_locked", "ค่าคงที่ที่สูตรอ้าง (คีย์/หัวคอลัมน์/พารามิเตอร์) ถูกล็อก — ลูกค้าต้องแก้ไหม?", inp)
    issue("P1", "manual_rows_locked", "แถวที่ป้ายบอกว่ากรอกเอง แต่ตัวเลขถูกล็อก", manual)

    # 4 ชีตรายการที่ลูกค้าอาจเพิ่มแถว (เป็นแหล่งของ dropdown) แต่จะถูกล็อก
    lists = []
    for d in dvs:
        f1 = (d.get("formula1") or "").lstrip("=")
        rng = bk.first_range(f1, d["sheet"]) if d.get("type") == "list" and not f1.startswith('"') else None
        if rng and rng[0] in targets:
            ws = wb[rng[0]]
            last = max((r for r in range(rng[2], min(rng[4], ws.max_row) + 1) if ws.cell(r, rng[1]).value is not None), default=rng[2])
            nxt = ws.cell(last + 1, rng[1])
            msg = f"{fmt(rng)} (ใช้โดย dropdown {d['sheet']}!{d['sqref'][:20]})"
            if rng[4] <= last:
                msg += " — ช่วงจบที่ข้อมูลแถวสุดท้ายพอดี เพิ่มรายการแล้ว dropdown ไม่เห็น"
            elif nxt.protection.locked:
                msg += f" — แถวว่างถัดไป ({nxt.coordinate}) ล็อกอยู่ ลูกค้าเพิ่มรายการไม่ได้"
            else:
                continue
            if msg not in lists:
                lists.append(msg)
    issue("P1", "list_sheet_locked", "ชีตรายการที่ลูกค้าอาจต้องเพิ่มแถว", lists)

    # 5 ตัวเลือกตอนล็อกให้เหมือนกัน
    combos = defaultdict(list)
    for ws in wb.worksheets:
        if ws.protection.sheet:
            combos[",".join(prot_info(ws)["allowed"]) or "(ค่าเริ่มต้น)"].append(ws.title)
    if len(combos) > 1:
        issue("P3", "inconsistent_options", "ตัวเลือกตอนล็อกไม่เหมือนกันทุกชีต", [f"[{k}]: {', '.join(v)}" for k, v in combos.items()])
    nopw = [ws.title for ws in wb.worksheets if ws.protection.sheet and not prot_info(ws)["password"]]
    issue("P2", "protected_without_password", "ชีตที่ล็อกแต่ไม่มีรหัส (ใครก็ปลดได้)", nopw)

    # 6 มาโครกับชีตที่ล็อก
    mods = vba_modules(op.real)
    codenames = {ws.sheet_properties.codeName: ws.title for ws in wb.worksheets if ws.sheet_properties.codeName}
    procs = scan_vba(mods, codenames, vis)
    btn_by_macro = defaultdict(list)
    for b in btns:
        btn_by_macro[b["macro"].split("!")[-1].split(".")[-1]].append(f"{b['sheet']}!{b['anchor']}")
    mac = []
    struct = bool(getattr(wb.security, "lockStructure", False)) if wb.security else False
    for p in procs:
        hit = sorted(s for s in p["sheets"] if s in targets)
        bt = btn_by_macro.get(p["name"])
        where = f" (ปุ่มที่ {', '.join(bt)})" if bt else ""
        if p["writes"] and hit and not p["unprotect"]:
            mac.append(f"{p['module']}.{p['name']}{where} เขียนลงชีตที่จะล็อก {', '.join(hit)} โดยไม่ Unprotect — จะ error (เช่น {p['lines'][0]})")
        elif p["writes"] and p["active"] and not p["unprotect"] and targets:
            mac.append(f"{p['module']}.{p['name']}{where} เขียนลง ActiveSheet/Selection — ถ้ากดจากชีตที่ล็อกจะ error")
        if p["sheet_copy"]:
            mac.append(f"{p['module']}.{p['name']}{where} คัดลอกชีต (Sheet.Copy) — การล็อกติดไปกับชีตที่คัดลอก ต้อง Unprotect ชีตใหม่ก่อนวาง/แก้")
        if struct and (p["sheet_copy"] or p["sheet_add"]):
            mac.append(f"{p['module']}.{p['name']} เพิ่ม/คัดลอก/ย้ายชีต แต่ไฟล์ล็อกโครงสร้าง (Protect Workbook) ไว้ — จะ error")
    issue("P0" if mac else "OK", "macro_vs_protection", "มาโครที่จะชนกับการล็อก", mac)
    if mods == {}:
        out["issues"].append({"sev": "ASK", "key": "vba_unreadable", "title": "อ่าน VBA ไม่ได้ ตรวจเรื่องมาโครกับชีตที่ล็อกไม่ได้", "count": 1, "examples": []})

    # 7 สถานะชีต
    out["hidden"] = [f"{ws.title} ({ws.sheet_state})" for ws in wb.worksheets if ws.sheet_state != "visible"]
    out["workbook_structure_locked"] = struct
    return out


def _cells(part, cap=400):
    try:
        c1, r1, c2, r2 = range_boundaries(part)
    except Exception:
        return []
    out = []
    for r in range(r1, min(r2, r1 + cap) + 1):
        for c in range(c1, c2 + 1):
            out.append(f"{get_column_letter(c)}{r}")
    return out


_MANUAL_CACHE = {}


def manual_cells(ws):
    """ตัวเลขในแถวที่ป้ายบอกว่ากรอกเอง (รายงานแยกแล้ว ไม่ต้องซ้ำ)"""
    k = id(ws)
    if k not in _MANUAL_CACHE:
        s = set()
        for row in ws.iter_rows():
            lab = " ".join(str(c.value) for c in row if isinstance(c.value, str) and not c.value.startswith("="))
            if lab and MANUAL_RE.search(lab):
                s |= {(ws.title, c.coordinate) for c in row if isinstance(c.value, (int, float)) and not isinstance(c.value, bool)}
        _MANUAL_CACHE[k] = s
    return _MANUAL_CACHE[k]


def fmt(rng):
    return qa_extra.fmt_rng(rng)


def print_lock(res):
    print("ชีตที่จะล็อก:", ", ".join(res["targets"]) or "(ไม่มี — ระบุ --plan)")
    print("สถานะชีต:")
    for s, d in res["sheets"].items():
        print(f"  {s}: {d['state']}, {'ล็อก' if d['protected'] else 'ไม่ล็อก'}{' มีรหัส' if d['password'] else ''}"
              f"{' (อนุญาต: ' + ','.join(d['allowed']) + ')' if d['protected'] and d['allowed'] else ''}"
              f"{' ← แผนจะล็อก' if d['planned'] and not d['protected'] else ''}")
    if res["workbook_structure_locked"]:
        print("  ไฟล์ล็อกโครงสร้าง (Protect Workbook)")
    for i in res["issues"]:
        print(f"\n[{i['sev']}] {i['title']} — {i['count']}")
        for e in i["examples"]:
            print("   " + e)


def cmd_lock(a):
    op = Opened(a.file)
    try:
        res = lock_audit(op, a.plan)
        print(json.dumps(res, ensure_ascii=False, indent=1, default=list)) if a.json else print_lock(res)
    finally:
        op.close()


# ---------- final ----------
def cmd_final(a):
    o1, o2 = Opened(a.base, "QA_PASSWORD2"), Opened(a.file)
    try:
        print("=== 1) เนื้อหาเทียบกับไฟล์ที่ตรวจผ่านล่าสุด ===")
        print_diff(diff_books(o1, o2))
        print("\n=== 2) การล็อก/ซ่อน ===")
        print_lock(lock_audit(o2, None))
        print("\n=== 3) ความเรียบร้อย ===")
        wb = o2.wb
        print("VBA project ล็อก:", {None: "ไม่มี VBA/อ่านไม่ได้", True: "ล็อก", False: "ไม่ล็อก ⚠"}[vba_locked(o2.real)])
        print("เปิดไฟล์ที่ชีต:", wb.active.title if wb.active else "?")
        bad = []
        for ws in wb.worksheets:
            if ws.sheet_state != "visible":
                continue
            sv = ws.sheet_view
            ac = sv.selection[0].activeCell if sv.selection else "A1"
            if (ac or "A1") != "A1" or ((sv.topLeftCell or "A1") != "A1" and ws.freeze_panes is None):
                bad.append(f"{ws.title} ({ac})")
        print("เคอร์เซอร์ไม่อยู่ A1:", ", ".join(bad) or "-")
        grouped = [ws.title for ws in wb.worksheets if ws.sheet_view.tabSelected]
        if len(grouped) > 1:
            print("⚠ ชีตถูก Group:", ", ".join(grouped))
        p = wb.properties
        print("metadata:", "; ".join(f"{k}={v}" for k, v in (("creator", p.creator), ("lastModifiedBy", p.lastModifiedBy), ("title", p.title)) if v) or "-")
        print("\n=== 4) ให้เช็กเพิ่มด้วย device_list_dir: โฟลเดอร์มีไฟล์อื่นที่อาจส่งผิด (ต้นฉบับ, _checked_vN, - Copy, ~$) ===")
    finally:
        o1.close(); o2.close()


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = ap.add_subparsers(dest="cmd", required=True)
    p = sp.add_parser("dump"); p.add_argument("file"); p.add_argument("--out"); p.add_argument("--fmt", action="store_true"); p.set_defaults(fn=cmd_dump)
    p = sp.add_parser("diff"); p.add_argument("old"); p.add_argument("new"); p.add_argument("--all", action="store_true"); p.add_argument("--json", action="store_true"); p.set_defaults(fn=cmd_diff)
    p = sp.add_parser("lookup"); p.add_argument("file"); p.add_argument("refs", nargs="+"); p.set_defaults(fn=cmd_lookup)
    p = sp.add_parser("find-sheet"); p.add_argument("folder"); p.add_argument("names", nargs="+"); p.set_defaults(fn=cmd_find_sheet)
    p = sp.add_parser("buttons"); p.add_argument("file"); p.set_defaults(fn=cmd_buttons)
    p = sp.add_parser("vba-diff"); p.add_argument("old"); p.add_argument("new"); p.set_defaults(fn=cmd_vba_diff)
    p = sp.add_parser("lock"); p.add_argument("file"); p.add_argument("--plan"); p.add_argument("--json", action="store_true"); p.set_defaults(fn=cmd_lock)
    p = sp.add_parser("final"); p.add_argument("file"); p.add_argument("--base", required=True); p.set_defaults(fn=cmd_final)
    a = ap.parse_args()
    a.fn(a)


if __name__ == "__main__":
    main()
