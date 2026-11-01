"""ทดสอบชุดเครื่องมือด้วยไฟล์จำลอง: python3 tests/run_tests.py
ผ่านทั้งหมด = พิมพ์ ALL PASS (exit 0)"""
import json, os, subprocess, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
tmp = tempfile.mkdtemp()
fails = []


def ok(cond, msg):
    print(("  ✓ " if cond else "  ✗ ") + msg)
    if not cond:
        fails.append(msg)


def run(*args, env=None):
    e = dict(os.environ); e.update(env or {})
    return subprocess.run([sys.executable, *args], capture_output=True, text=True, env=e, cwd=ROOT)


subprocess.run([sys.executable, os.path.join(HERE, "make_samples.py"), tmp], check=True, capture_output=True)
bad, good = os.path.join(tmp, "bad.xlsx"), os.path.join(tmp, "good.xlsx")


def check(path, *extra):
    out = os.path.join(tmp, "qa.json")
    r = run("excel_qa_check.py", path, "--out", out, *extra)
    assert r.returncode == 0, r.stderr
    return json.load(open(out, encoding="utf-8"))


NEW = {"duplicate_criteria", "criteria_not_in_data", "label_key_mismatch", "dropdown_from_data",
       "power_query_local_path", "fixed_period_text", "year_text_calendar"}

print("1) check: ไฟล์มีปัญหา (พ.ศ.)")
d = check(bad, "--calendar", "be")
keys = {f["key"]: f for f in d["findings"]}
for k in NEW:
    ok(k in keys, f"เจอ {k}")
ok(keys.get("duplicate_criteria", {}).get("sev") == "P0", "เกณฑ์ซ้ำเป็น P0")
ok("likely_intended" in keys and "inconsistent_formula" not in keys and "hardcode_in_formula_column" not in keys,
   "แถวรวม/แถวกรอกเอง ถูกยุบเป็น 'อาจตั้งใจ' ไม่เตือนเป็น P1")
ok(any("C:\\Users" in x for q in d["parts"].get("power_query_sources", []) for s in q["sources"] for x in s), "แสดง source ของ Power Query")

print("2) check: ไฟล์ถูกต้อง ต้องไม่มี check ใหม่")
d = check(good, "--calendar", "be")
got = {f["key"] for f in d["findings"]} & NEW
ok(not got, f"ไม่เตือนผิด (เจอ: {got or '-'})")

print("3) check: นโยบาย ค.ศ. กับไฟล์ที่แสดง พ.ศ.")
d = check(good)
ok(any(f["key"] == "year_text_calendar" for f in d["findings"]), "จับข้อความปี พ.ศ. เมื่อลูกค้าใช้ ค.ศ.")

print("4) lock")
r = run("excel_qa_tools.py", "lock", bad, "--json"); L = json.loads(r.stdout)
lk = {i["key"] for i in L["issues"]}
for k in ("formula_not_hidden", "dropdown_locked", "manual_rows_locked", "inconsistent_options", "referenced_constants_locked"):
    ok(k in lk, f"lock เจอ {k}")
r = run("excel_qa_tools.py", "lock", good, "--json"); L = json.loads(r.stdout)
lk = {i["key"] for i in L["issues"]}
ok("formula_not_hidden" not in lk and "dropdown_locked" not in lk, "ไฟล์ถูกต้อง: สูตรซ่อนครบ dropdown ปลดล็อก")
ok("list_sheet_locked" in lk, "เตือนชีตรายการที่ช่วง dropdown จบพอดีข้อมูล")

print("5) scan VBA (โค้ดจำลอง)")
import excel_qa_tools as T
mods = {"Module1": 'Sub UpdateData()\n Sheets("Report").Range("A3:M9").ClearContents\nEnd Sub\n'
                   'Sub SafeWrite()\n Sheets("Report").Unprotect "x"\n Sheets("Report").Range("A1").Value = 1\nEnd Sub\n'
                   'Sub Export()\n Sheets("Report").Copy\n ActiveSheet.Range("A1").PasteSpecial\nEnd Sub'}
P = {p["name"]: p for p in T.scan_vba(mods, {}, ["Report"])}
ok(P["UpdateData"]["writes"] and not P["UpdateData"]["unprotect"], "จับมาโครเขียนชีตโดยไม่ Unprotect")
ok(P["SafeWrite"]["unprotect"], "มาโครที่ Unprotect ก่อน ไม่ถูกนับว่าเสี่ยง")
ok(P["Export"]["sheet_copy"], "จับ Sheet.Copy")

print("6) diff / lookup / final")
r = run("excel_qa_tools.py", "diff", good, bad)
ok("F5: CCC → AAA" in r.stdout and "D6: LH → U-" in r.stdout, "diff เห็นค่าและการล็อกที่เปลี่ยน")
r = run("excel_qa_tools.py", "diff", good, good)
ok("ไม่มีอะไรเปลี่ยน" in r.stdout, "diff ไฟล์เดียวกัน = ไม่เปลี่ยน")
r = run("excel_qa_tools.py", "lookup", bad, "Report!F5")
ok("Report!F5\tAAA" in r.stdout, "lookup อ่านค่าจริง")
r = run("excel_qa_tools.py", "final", good, "--base", good)
ok(r.returncode == 0 and "เปิดไฟล์ที่ชีต" in r.stdout, "final รันได้")

print("7) ไฟล์มีรหัส")
try:
    import msoffcrypto
    from msoffcrypto.format.ooxml import OOXMLFile
    enc = os.path.join(tmp, "enc.xlsx")
    with open(bad, "rb") as fi, open(enc, "wb") as fo:
        OOXMLFile(fi).encrypt("t3st", fo)
    r = run("excel_qa_check.py", enc)
    ok(r.returncode == 2 and "encrypted" in r.stdout, "ไม่มีรหัส -> status encrypted")
    r = run("excel_qa_check.py", enc, env={"QA_PASSWORD": "wrong"})
    ok("wrong_password" in r.stdout, "รหัสผิด -> wrong_password")
    out = os.path.join(tmp, "e.json")
    r = run("excel_qa_check.py", enc, "--out", out, env={"QA_PASSWORD": "t3st"})
    ok(r.returncode == 0 and json.load(open(out, encoding="utf-8"))["encrypted"], "รหัสถูก -> ตรวจได้")
    ok("t3st" not in open(out, encoding="utf-8").read() and "t3st" not in r.stdout, "รหัสไม่รั่วออกมาในผลลัพธ์")
    r = run("excel_qa_tools.py", "lookup", enc, "Report!F5", env={"QA_PASSWORD": "t3st"})
    ok("AAA" in r.stdout, "tools อ่านไฟล์มีรหัสได้")
except (ImportError, AttributeError) as e:
    print("  - ข้าม (msoffcrypto ไม่รองรับการเข้ารหัส):", e)

print("\nALL PASS" if not fails else f"\nFAILED {len(fails)}: {fails}")
sys.exit(1 if fails else 0)
