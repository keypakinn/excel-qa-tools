"""qa_extra.py — check เพิ่มเติมของ excel_qa_check.py (อ่านอย่างเดียว)

ทุกฟังก์ชันคืน list ของ finding dict (group, key, title, sev, count, examples)
- criteria_checks   : เกณฑ์ SUMIFS/COUNTIFS ซ้ำ (P0), เกณฑ์ไม่พบในข้อมูล (P1), ป้ายแถวสลับกับคีย์ (G)
- dropdown_checks   : DV ทุกแบบรวม x14, list อ้างช่วงที่มีค่าซ้ำ/ช่วงข้อมูลดิบ, #REF!
- power_query_checks: path ตายตัวในแหล่งข้อมูล Power Query
- period_text_checks: ข้อความเดือน/ปีพิมพ์ตายตัวในชีตที่มีตัวเลือก
- calendar_checks   : ปี ค.ศ./พ.ศ. ปนกันตามนโยบายปฏิทินของงาน
"""
import datetime, re
from collections import defaultdict

from openpyxl.utils import column_index_from_string, get_column_letter

import xparts

MAX_LOC = 30
CRIT_FUNCS = ["SUMIFS", "COUNTIFS", "AVERAGEIFS", "MAXIFS", "MINIFS", "SUMIF", "COUNTIF", "AVERAGEIF"]

SHEET_PREFIX = r"(?:(?:'((?:[^']|'')+)'|([A-Za-z_฀-๿][\w.฀-๿]*))!)?"
REF_RE = re.compile(r"(?<![A-Za-z0-9_.฀-๿'\]])" + SHEET_PREFIX +
                    r"(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d+:\$?\d+)"
                    r"(?![A-Za-z0-9_(])")
TABLE_RE = re.compile(r"([A-Za-z_฀-๿][\w.฀-๿]*)\[(?:\[#[A-Za-z ]+\],\s*)?\[?([^\[\]]+?)\]?\]")
STR_RE = re.compile(r'"(?:[^"]|"")*"')

THAI_MONTHS = ("มกราคม กุมภาพันธ์ มีนาคม เมษายน พฤษภาคม มิถุนายน กรกฎาคม สิงหาคม กันยายน ตุลาคม พฤศจิกายน ธันวาคม "
               "ม.ค. ก.พ. มี.ค. เม.ย. พ.ค. มิ.ย. ก.ค. ส.ค. ก.ย. ต.ค. พ.ย. ธ.ค.").split()
EN_MONTHS = ("January February March April May June July August September October November December "
             "Jan Feb Mar Apr Jun Jul Aug Sep Sept Oct Nov Dec").split()
MONTH_YEAR_RE = re.compile(r"(" + "|".join(re.escape(m) for m in THAI_MONTHS + EN_MONTHS) + r")\.?\s*(?:พ\.ศ\.|ค\.ศ\.)?\s*((?:19|20|25)\d\d)\b", re.I)


class Loc:
    def __init__(self):
        self.items = []; self.count = 0
    def add(self, x):
        self.count += 1
        if len(self.items) < MAX_LOC:
            self.items.append(x)


def F(group, key, title, loc, sev, **extra):
    d = {"group": group, "key": key, "title": title, "sev": sev}
    if isinstance(loc, Loc):
        d.update({"count": loc.count, "examples": loc.items})
    else:
        d.update({"count": len(loc), "examples": list(loc)[:MAX_LOC]})
    d.update(extra)
    return d


def norm(v):
    """ทำค่าให้เทียบกันได้ (ตัวเลข/ข้อความ/วันที่)"""
    if v is None:
        return None
    if isinstance(v, bool):
        return str(v).upper()
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, datetime.datetime):
        return v.date().isoformat() if v.time() == datetime.time(0) else v.isoformat()
    if isinstance(v, datetime.date):
        return v.isoformat()
    s = str(v).strip()
    if s == "":
        return None
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return s.lower()


class Book:
    """ตัวช่วย resolve ช่วงอ้างอิง + อ่านค่าแบบ cache"""
    def __init__(self, wb, wbv):
        self.wb, self.wbv = wb, wbv
        self.cache = {}
        self.tables = {}
        for ws in wb.worksheets:
            for t in ws.tables.values():
                self.tables[t.name.lower()] = (ws.title, t.ref, t.headerRowCount if t.headerRowCount is not None else 1)

    def resolve(self, sheet_q, sheet_u, ref, cur):
        sn = (sheet_q.replace("''", "'") if sheet_q else sheet_u) or cur
        if sn not in self.wb.sheetnames:
            return None
        ws = self.wb[sn]
        r = ref.replace("$", "")
        if re.fullmatch(r"[A-Z]{1,3}:[A-Z]{1,3}", r):
            a, b = r.split(":")
            return (sn, column_index_from_string(a), 1, column_index_from_string(b), max(ws.max_row, 1))
        if re.fullmatch(r"\d+:\d+", r):
            a, b = r.split(":")
            return (sn, 1, int(a), max(ws.max_column, 1), int(b))
        parts = r.split(":")
        def cr(p):
            m = re.fullmatch(r"([A-Z]{1,3})(\d+)", p)
            return column_index_from_string(m.group(1)), int(m.group(2))
        c1, r1 = cr(parts[0]); c2, r2 = cr(parts[-1])
        return (sn, min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2))

    def resolve_table(self, tname, col):
        t = self.tables.get(tname.lower())
        if not t:
            return None
        sn, ref, hdr = t
        ws = self.wb[sn]
        a, b = ref.split(":")
        m1 = re.fullmatch(r"([A-Z]+)(\d+)", a); m2 = re.fullmatch(r"([A-Z]+)(\d+)", b)
        c1, r1, c2, r2 = column_index_from_string(m1.group(1)), int(m1.group(2)), column_index_from_string(m2.group(1)), int(m2.group(2))
        col = col.strip().lstrip("@").strip("[]")
        for c in range(c1, c2 + 1):
            h = ws.cell(r1, c).value
            if h is not None and str(h).strip().lower() == col.lower():
                return (sn, c, r1 + (hdr or 0), c, r2)
        return None

    def first_range(self, text, cur):
        """ช่วงแรกในข้อความอาร์กิวเมนต์ (ref ปกติ หรือ Table[Col])"""
        t = STR_RE.sub('""', text)
        m = TABLE_RE.search(t)
        if m and m.group(1).lower() in self.tables:
            return self.resolve_table(m.group(1), m.group(2))
        m = REF_RE.search(t)
        if m:
            return self.resolve(m.group(1), m.group(2), m.group(3), cur)
        return None

    def values(self, rng):
        if rng in self.cache:
            return self.cache[rng]
        sn, c1, r1, c2, r2 = rng
        ws = self.wbv[sn]
        r2 = min(r2, ws.max_row or r2)
        vals = []
        if r2 >= r1:
            for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2, values_only=True):
                vals.extend(row)
        self.cache[rng] = vals
        return vals

    def value(self, sn, c, r):
        return self.wbv[sn].cell(r, c).value


def fmt_rng(rng):
    sn, c1, r1, c2, r2 = rng
    a = f"{get_column_letter(c1)}{r1}"
    b = f"{get_column_letter(c2)}{r2}"
    return f"{sn}!{a}" + ("" if a == b else f":{b}")


# ---------- เกณฑ์ของ SUMIFS / COUNTIFS ----------
def _crit_pairs(fn, args):
    """[(crit_range_arg, crit_arg)] ตามชนิดฟังก์ชัน"""
    if fn in ("SUMIF", "COUNTIF", "AVERAGEIF"):
        return [(args[0], args[1])] if len(args) >= 2 else []
    start = 0 if fn == "COUNTIFS" else 1
    return [(args[i], args[i + 1]) for i in range(start, len(args) - 1, 2)]


def criteria_checks(wb, wbv, formula_text):
    bk = Book(wb, wbv)
    sig_groups = defaultdict(list)       # (sheet, resolved formula) -> [(cell, frozenset(criteria cells))]
    missing = {}                         # criteria cell -> (value, range)
    empty_ranges = set()
    key_cells = defaultdict(dict)        # sheet -> {row: (col, value)} คีย์แบบคอลัมน์คงที่
    for ws in wb.worksheets:
        sn = ws.title
        for row in ws.iter_rows():
            for cell in row:
                f = formula_text(cell.value)
                if not f or not re.search(r"IFS?\s*\(", f, re.I):
                    continue
                crit_cells = {}
                for fn, args in xparts.iter_calls(f, CRIT_FUNCS):
                    for rarg, carg in _crit_pairs(fn, args):
                        rng = bk.first_range(rarg, sn)
                        cs = STR_RE.sub('""', carg)
                        refs = [m for m in REF_RE.finditer(cs) if ":" not in m.group(3)]
                        for m in refs:
                            res = bk.resolve(m.group(1), m.group(2), m.group(3), sn)
                            if not res:
                                continue
                            csn, cc, cr = res[0], res[1], res[2]
                            crit_cells[(csn, cc, cr)] = bk.value(csn, cc, cr)
                            # คีย์แบบล็อกคอลัมน์ ($P5) เก็บไว้ตรวจป้ายแถว
                            if m.group(3).startswith("$") and not re.search(r"\$\d", m.group(3)) and csn == sn:
                                key_cells[sn][cr] = (cc, bk.value(csn, cc, cr))
                        # เกณฑ์แบบเท่ากับ: อ้างเซลล์เดียว หรือ "="&ref
                        c0 = re.sub(r'^"="\s*&\s*', "", carg.strip())
                        if rng and len(refs) == 1 and REF_RE.fullmatch(c0):
                            res = bk.resolve(refs[0].group(1), refs[0].group(2), refs[0].group(3), sn)
                            if res:
                                v = bk.value(res[0], res[1], res[2])
                                nv = norm(v)
                                if nv is None or (isinstance(nv, str) and re.search(r"[*?~]", nv)):
                                    continue
                                vals = [norm(x) for x in bk.values(rng)]
                                if not any(x is not None for x in vals):
                                    empty_ranges.add(fmt_rng(rng)); continue
                                if nv not in set(vals):
                                    k = f"{res[0]}!{get_column_letter(res[1])}{res[2]}"
                                    missing.setdefault(k, (v, fmt_rng(rng)))
                if not crit_cells:
                    continue
                # สร้างสูตรแบบ "แทนค่าเกณฑ์ด้วยค่าจริง" เพื่อหาเซลล์ที่คำนวณซ้ำกันทุกประการ
                def rep(m):
                    res = bk.resolve(m.group(1), m.group(2), m.group(3), sn)
                    if not res:
                        return m.group(0)
                    if res[1] == res[3] and res[2] == res[4] and (res[0], res[1], res[2]) in crit_cells:
                        return "«" + repr(norm(crit_cells[(res[0], res[1], res[2])])) + "»"
                    return fmt_rng(res)
                resolved = REF_RE.sub(rep, STR_RE.sub(lambda m: m.group(0), f))
                if any(norm(v) is None for v in crit_cells.values()):
                    continue  # เกณฑ์ว่าง = คอลัมน์ที่ยังไม่ใช้ ข้าม
                sig_groups[(sn, resolved)].append((cell.coordinate, frozenset(crit_cells)))

    out = []
    # P0: เซลล์ต่างกัน อ้างเกณฑ์คนละเซลล์ แต่ค่าเกณฑ์เหมือนกันทุกตัว -> ผลซ้ำกัน
    dup = Loc(); dup_cells = set()
    by_value = defaultdict(lambda: {"crit": set(), "cells": []})
    for (sn, _sig), lst in sig_groups.items():
        sets = {s for _, s in lst}
        if len(sets) < 2:
            continue
        # เกณฑ์ที่ต่างกันระหว่างเซลล์ = ตัวที่ซ้ำ
        common = frozenset.intersection(*sets)
        diff_cells = set().union(*sets) - common
        vals = {}
        for c in diff_cells:
            v = bk.value(*c)
            vals.setdefault(repr(norm(v)), []).append(c)
        for v, cs in vals.items():
            if len(cs) < 2:
                continue
            key = (sn, v)
            by_value[key]["crit"].update(cs)
            by_value[key]["cells"].extend(c for c, _ in lst)
            by_value[key]["raw"] = bk.value(*cs[0])
    for (sn, v), d in by_value.items():
        crit = sorted(d["crit"], key=lambda c: (c[0], c[2], c[1]))
        cl = ", ".join(f"{c[0]}!{get_column_letter(c[1])}{c[2]}" for c in crit[:8]) + (f" และอีก {len(crit)-8}" if len(crit) > 8 else "")
        cells = sorted(set(d["cells"]))
        dup.add(f'ค่าเกณฑ์ "{d["raw"]}" ซ้ำ {len(crit)} ช่อง: {cl} → ผลในชีต {sn} ซ้ำกัน {len(cells)} เซลล์ เช่น {", ".join(cells[:4])}')
    if dup.count:
        out.append(F("A", "duplicate_criteria", "เกณฑ์ของ SUMIFS/COUNTIFS ซ้ำกัน ทำให้ยอดนับซ้ำ (ยอดรวมผิด)", dup, "P0"))

    miss = Loc()
    for k, (v, rng) in missing.items():
        miss.add(f'{k} = "{v}" ไม่พบใน {rng}')
    if miss.count:
        out.append(F("B", "criteria_not_in_data", "เกณฑ์ที่ไม่มีอยู่ในข้อมูลเลย (ยอดจะเป็น 0 — สะกดไม่ตรง/เปลี่ยนรหัสแล้วไม่ได้แก้ทุกที่)", miss, "P1",
                     note=("ช่วงข้อมูลว่าง ตรวจไม่ได้: " + ", ".join(sorted(empty_ranges)[:5])) if empty_ranges else ""))

    # G: ป้ายแถวกับคีย์สลับกัน (คีย์ของแถวนี้ = ป้ายของอีกแถว)
    swap = Loc()
    for sn, rows in key_cells.items():
        ws = wb[sn]
        labels = {}
        for r, (kc, kv) in rows.items():
            lab = [c.value.strip() for c in ws[r][: kc - 1] if isinstance(c.value, str) and not c.value.startswith("=") and c.value.strip()]
            labels[r] = lab
        all_labels = {}
        for r, lab in labels.items():
            for t in lab:
                all_labels.setdefault(t.lower(), r)
        for r, (kc, kv) in rows.items():
            if not isinstance(kv, str) or not kv.strip():
                continue
            k = kv.strip().lower()
            mine = [t.lower() for t in labels.get(r, [])]
            if any(k in t or t in k for t in mine):
                continue
            r2 = all_labels.get(k)
            if r2 and r2 != r:
                swap.add(f'{sn}!{get_column_letter(kc)}{r} = "{kv}" แต่ป้ายแถวนี้คือ "{" / ".join(labels[r])[:40]}" (ค่า "{kv}" เป็นป้ายของแถว {r2})')
    if swap.count:
        out.append(F("G", "label_key_mismatch", "ป้ายแถวไม่ตรงกับคีย์ที่สูตรใช้ (อาจสลับแถว)", swap, "P1"))
    return out


# ---------- Dropdown ----------
def all_validations(wb, z):
    out = []
    for ws in wb.worksheets:
        for dv in ws.data_validations.dataValidation:
            out.append({"sheet": ws.title, "sqref": str(dv.sqref), "type": dv.type or "", "formula1": dv.formula1 or ""})
    out += xparts.x14_validations(z)
    return out


def dropdown_checks(wb, wbv, z):
    bk = Book(wb, wbv)
    dvs = all_validations(wb, z)
    out = []
    bad = Loc(); dups = Loc()
    names = {}
    try:
        for k, dn in wb.defined_names.items():
            names[k.lower()] = dn.attr_text
    except Exception:
        pass
    for d in dvs:
        f1 = (d["formula1"] or "").lstrip("=")
        if "#REF!" in f1:
            bad.add(f'{d["sheet"]}!{d["sqref"]}')
            continue
        if d["type"] != "list" or f1.startswith('"'):
            continue
        target = names.get(f1.lower(), f1)
        if target.upper().startswith(("OFFSET(", "INDIRECT(")):
            continue
        rng = bk.first_range(target, d["sheet"])
        if not rng:
            continue
        vals = [v for v in bk.values(rng) if norm(v) is not None]
        seen, rep = set(), []
        for v in vals:
            n = norm(v)
            if n in seen and v not in rep:
                rep.append(v)
            seen.add(n)
        rows = rng[4] - rng[2] + 1
        if rep:
            dups.add(f'{d["sheet"]}!{d["sqref"][:30]} ← {fmt_rng(rng)}: ค่าซ้ำ {len(vals) - len(seen)} รายการ เช่น {", ".join(map(str, rep[:3]))}'
                     + (f" (ช่วงยาว {rows:,} แถว ดูเหมือนอ้างคอลัมน์ข้อมูลดิบ)" if rows > 500 else ""))
    if bad.count:
        out.append(F("B", "bad_validation_x14", "Data Validation อ้างอิงเสีย (#REF!) รวมแบบ x14", bad, "P1"))
    if dups.count:
        out.append(F("B", "dropdown_from_data", "Dropdown อ้างช่วงที่มีค่าซ้ำ/ข้อมูลดิบ (รายการซ้ำ และขาดค่าที่ยังไม่มีในข้อมูล) ควรใช้รายการคงที่", dups, "P1"))
    return out, dvs


# ---------- Power Query ----------
def power_query_checks(z):
    qs = xparts.power_queries(z)
    loc = Loc()
    for q in qs:
        for p in q.get("local_paths", []):
            loc.add(f'query "{q["name"]}" อ่านจาก {p}')
    out = []
    if loc.count:
        out.append(F("B", "power_query_local_path", "Power Query ชี้ path ตายตัว (เครื่องลูกค้าไม่มี path นี้ / ถ้าไฟล์ต้นทางมีรหัส PQ เปิดไม่ได้)", loc, "P1"))
    return out, [{"name": q["name"], "sources": q["sources"]} for q in qs]


# ---------- ข้อความเดือน/ปีพิมพ์ตายตัว ----------
def period_text_checks(wb, dvs, formula_text):
    sheets_with_selector = {d["sheet"] for d in dvs if d.get("type") == "list"}
    loc = Loc()
    for ws in wb.worksheets:
        if ws.title not in sheets_with_selector:
            continue
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and not formula_text(v) and MONTH_YEAR_RE.search(v):
                    loc.add(f'{ws.title}!{c.coordinate} = "{v[:60]}"')
    if loc.count:
        return [F("G", "fixed_period_text", "ข้อความเดือน/ปีพิมพ์ตายตัวในชีตที่มีตัวเลือก (จะไม่เปลี่ยนตามที่เลือก)", loc, "P2")]
    return []


# ---------- ปฏิทิน ----------
DATE_FMT_RE = re.compile(r"(?i)(yy|bb)")


def calendar_checks(wb, wbv, calendar, is_buddhist_format, formula_text):
    out = []
    wrong_text = Loc(); ce_fmt = Loc()
    want_be = calendar == "be"
    bad_year = re.compile(r"(?<!\d)(19|20)\d\d(?!\d)") if want_be else re.compile(r"(?<!\d)25\d\d(?!\d)")
    for ws in wb.worksheets:
        wsv = wbv[ws.title]
        for row in ws.iter_rows():
            for c in row:
                f = formula_text(c.value)
                cv = wsv[c.coordinate].value
                if f and isinstance(cv, str) and bad_year.search(cv) and re.search(r"(?i)YEAR|TEXT|&", f):
                    wrong_text.add(f'{ws.title}!{c.coordinate} = "{cv[:50]}" ({f[:60]})')
                if want_be:
                    dv = cv if f else c.value
                    fmt = c.number_format or ""
                    if isinstance(dv, (datetime.date, datetime.datetime)) and DATE_FMT_RE.search(re.sub(r'"[^"]*"', "", fmt)) \
                            and not is_buddhist_format(fmt) and "y" in fmt.lower():
                        ce_fmt.add(f"{ws.title}!{c.coordinate} [{fmt}]")
    if wrong_text.count:
        t = "ข้อความจากสูตรแสดงปี ค.ศ. ในไฟล์ที่ลูกค้าใช้ พ.ศ." if want_be else "ข้อความจากสูตรแสดงปี พ.ศ."
        out.append(F("E", "year_text_calendar", t, wrong_text, "P2"))
    if ce_fmt.count:
        out.append(F("E", "ce_format_in_be_book", "วันที่แสดงเป็น ค.ศ. ในไฟล์ที่ลูกค้าใช้ พ.ศ. (ปฏิทินปนกัน)", ce_fmt, "P2"))
    return out
