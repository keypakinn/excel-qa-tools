#!/usr/bin/env python3
"""make_vbs.py - build excel_qa_fix.vbs from template.vbs
usage: python3 make_vbs.py <template.vbs> <job.json> <out.vbs>
job.json: {"source": "file.xlsx", "target": "file_checked_v1.xlsx", "password_required": false, "ops": [[op, arg1, ...], ...]}
op names = the Subs in the template (SetF, SetFormat, SetDate, SheetUnhide, SheetHide, SheetDelete, ClearAllComments,
SetMeta, BreakAllLinks, NewSheetFromLayout, ReplaceInFormulas, SetCellProtection, LockFormulas,
TidyView, OpenOnSheet, RestoreActive)
"""
import json, sys

def lit(v):
    if isinstance(v, bool):
        return "True" if v else "False"
    if isinstance(v, (int, float)):
        return str(v)
    s = "".join(c if ord(c) < 128 else "\\u%04X" % ord(c) for c in str(v))
    return '"' + s.replace('"', '""') + '"'

def main(tpl, job, out):
    t = open(tpl, encoding="ascii").read()
    j = json.load(open(job, encoding="utf-8"))
    cfg = ("src = here & \"\\\" & U(%s)\n" % lit(j["source"]) +
           "dst = here & \"\\\" & U(%s)\n" % lit(j["target"]) +
           "pwRequired = %s" % lit(bool(j.get("password_required"))))
    lines = []
    for op in j["ops"]:
        name, args = op[0], op[1:]
        if name == "NewSheetFromLayout":
            lines.append("Dim newSheet_%d : Set newSheet_%d = NewSheetFromLayout(%s)" % (len(lines), len(lines), ", ".join(lit(a) for a in args)))
        else:
            lines.append(name + ((" " + ", ".join(lit(a) for a in args)) if args else ""))
    t = t.replace("'@@CONFIG@@", cfg).replace("'@@OPS@@", "\n".join(lines))
    open(out, "wb").write(t.replace("\r\n", "\n").replace("\n", "\r\n").encode("ascii"))

if __name__ == "__main__":
    main(*sys.argv[1:4])
