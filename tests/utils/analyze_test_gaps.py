"""Analyze gaps between VBA Test_* assertions and Python crossval coverage.

Usage: python tests/utils/analyze_test_gaps.py [ModuleName]
"""

import os, re, sys

SRC = os.path.join(os.path.dirname(__file__), "..", "..", "src")
CV = os.path.join(os.path.dirname(__file__), "..", "crossval")

VBA_BUILTINS = {
    # Keywords / flow control
    "If","Then","Else","ElseIf","End","Exit","For","Next","Do","Loop","While",
    "Select","Case","With","Each","In","To","Step","Mod","Like","Is","Not",
    "And","Or","Xor","On","GoTo","GoSub","Resume","Call","Dim","Set","ReDim",
    "Preserve","Erase","Static","Const","Private","Public","Optional","ParamArray",
    "ByVal","ByRef","New","Sub","Function","Property","Get","Let","Enum","Type",
    # Literals
    "True","False","Empty","Null","Nothing","Boolean","Long","Double","String",
    "Variant","Integer","Single","Byte","Currency","Date","Object","vbCrLf",
    # VBA runtime
    "Abs","CDbl","CLng","CStr","CInt","CSng","CBool","CDate","CCur","CVar",
    "CVErr","Str","Val","Hex","Oct","Format","Format$","Chr","ChrW","Chr$",
    "Asc","AscW","AscB","Len","LenB","Mid","Mid$","Left","Left$","Right","Right$",
    "Trim","Trim$","LTrim","RTrim","UCase","UCase$","LCase","LCase$","Space",
    "Space$","String","String$","Replace","Replace$","Split","Join","InStr",
    "InStrRev","StrComp","StrReverse","Round","Int","Fix","Sgn","Rnd","Randomize",
    # Type checks
    "IsError","IsArray","IsEmpty","IsNull","IsNumeric","IsDate","IsObject",
    "IsMissing","TypeName","VarType","TypeOf","LBound","UBound","Array",
    # Date/time
    "Date","Time","Now","Year","Month","Day","Hour","Minute","Second","Timer",
    "Weekday","DateSerial","TimeSerial","DateValue","TimeValue","DateAdd",
    "DateDiff","DatePart",
    # Flow
    "IIf","Choose","Switch",
    # Error/object
    "Err","Error","CreateObject","GetObject","Application","Transpose",
    "DoEvents","Debug","MsgBox",
    # File
    "Dir","Environ","CurDir","FreeFile","Open","Close","Input","Print","Write",
    # Excel-specific
    "Range","Cells","Rows","Columns","Sheets","Worksheets","Workbook","ActiveCell",
}

def crossval_funcs(mod):
    """Return set of 'func' values in build_<mod>.py"""
    fs = set()
    for f in os.listdir(CV):
        low = f.lower()
        mlow = mod.lower()
        if low.startswith("build_") and mlow in low:
            path = os.path.join(CV, f)
            with open(path, encoding="utf-8") as fh:
                for m in re.finditer(r'"func":\s*"(\w+)"', fh.read()):
                    fs.add(m.group(1))
    return fs

def analyze(mod):
    path = os.path.join(SRC, f"{mod}.bas")
    with open(path, encoding="utf-8-sig") as f:
        txt = f.read()
    m = re.search(r"Public Sub (Test_\w+)\(\)", txt)
    if not m:
        return {"has_test": False}
    # Test_* is always at end of file — take everything from start
    body = txt[m.start():]
    lines = body.count("\n")
    vba_funcs = set()
    for fm in re.finditer(r'(?:result\s*=\s*)?(\w+)\(', body):
        n = fm.group(1)
        if n not in VBA_BUILTINS and not n.startswith("_"):
            vba_funcs.add(n)
    cv_funcs = crossval_funcs(mod)
    uncovered = sorted(vba_funcs - cv_funcs)
    errors = len(re.findall(r"IsError\(|Err\.Raise\s+5|Err\.Number", body))
    subs = len(re.findall(r'(FilterTable|TransposeTable|SplitColumnToRows|MergeColumns|Unpivot)\s', body))
    return {
        "has_test": True, "lines": lines,
        "vba_funcs": sorted(vba_funcs), "cv_funcs": sorted(cv_funcs),
        "uncovered": uncovered,
        "errors": errors, "subs": subs,
    }

if __name__ == "__main__":
    ci_mode = "--ci" in sys.argv
    target = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith("--") else None
    mods = sorted(f.replace(".bas","") for f in os.listdir(SRC) if f.endswith(".bas"))
    if target:
        mods = [m for m in mods if target.lower() in m.lower()]
    done = []; pending = []; total_uncovered = 0; total_lines = 0
    for mod in mods:
        info = analyze(mod)
        if not info["has_test"]:
            # Module has no VBA Test_* — check if crossval covers it
            cv = crossval_funcs(mod)
            if cv:
                print(f"  {mod:20s}  migrated   (0 VBA Test lines, {len(cv)} crossval funcs)")
            else:
                print(f"  {mod:20s}  NO TESTS   (no VBA Test_* and no crossval coverage)")
                if ci_mode:
                    all_uncovered = [mod]
            done.append(mod); continue
        n = len(info["uncovered"])
        total_uncovered += n; total_lines += info["lines"]
        s = "DONE" if n==0 else f"{n} gaps"
        print(f"  {mod:20s}  {s:10s}  {info['lines']:4d} lines  VBA:{len(info['vba_funcs']):2d} funcs  CV:{len(info['cv_funcs']):2d}")
        if info["uncovered"]:
            print(f"    UNCOVERED: {info['uncovered']}")
        if n == 0: pending.append(mod)
    print(f"\n  Total: {total_lines} VBA test lines, {total_uncovered} uncovered functions")
    if pending:
        print(f"  Ready to delete Test_*: {pending}")
    if ci_mode:
        # Known false-positives from regex limitations:
        #   "result"  — captured from `result = Func(...)` patterns
        #   "Len"     — ambiguous with `Len()` builtin in test assertions
        #   "Range"   — Excel Range object references
        #   Test_SqlUtils / SqlRangeQuery — ADODB-only, cannot be COM-tested
        WHITELIST = {"Len","Range","result","Test_SqlUtils","SqlRangeQuery"}
        all_uncovered = set()
        for m in mods:
            info = analyze(m)
            if info.get("uncovered"):
                all_uncovered.update(f for f in info["uncovered"] if f not in WHITELIST)
        if all_uncovered:
            print(f"  CI FAIL: unexpected uncovered functions: {sorted(all_uncovered)}")
            sys.exit(1)
        print("  CI PASS: no unexpected coverage gaps")
