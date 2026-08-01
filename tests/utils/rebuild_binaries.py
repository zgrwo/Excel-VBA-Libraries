"""Rebuild docs/VBA_Libraries.xlsm and docs/VBA_Libraries.xlam from latest source.

Usage: python tests/utils/rebuild_binaries.py
"""

import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from tests.test_utils import (
    ensure_excel, teardown, SRC_DIR, VBA_CORE_DIR, VBA_CORE_IMPORT_ORDER,
)
import pathlib

DOCS = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "docs")
XLSM = os.path.join(DOCS, "VBA_Libraries.xlsm")
XLAM = os.path.join(DOCS, "VBA_Libraries.xlam")

def main():
    excel = ensure_excel()
    wb = None

    # Collect all module paths in correct import order
    module_paths = [
        os.path.join(VBA_CORE_DIR, name + ".cls")
        for name in VBA_CORE_IMPORT_ORDER
    ]
    bas_files = sorted([p.stem for p in pathlib.Path(SRC_DIR).glob("*.bas")])
    module_paths += [os.path.join(SRC_DIR, f + ".bas") for f in bas_files]

    try:
        # --- Build .xlsm ---
        print(f"Building {XLSM} ...")
        wb = excel.Workbooks.Add()
        vbproj = wb.VBProject

        for src_path in module_paths:
            mod_name = os.path.splitext(os.path.basename(src_path))[0]
            # Remove existing component with same name
            for comp in list(vbproj.VBComponents):
                try:
                    if comp.Name == mod_name:
                        vbproj.VBComponents.Remove(comp)
                except Exception:
                    pass

            ext = os.path.splitext(src_path)[1].lower()
            comp_type = 2 if ext == ".cls" else 1
            with open(src_path, "r", encoding="utf-8-sig") as f:
                content = f.read()
            comp = vbproj.VBComponents.Add(comp_type)
            comp.Name = mod_name
            comp.CodeModule.AddFromString(content)
            print(f"  Imported {mod_name}")

        # Remove default sheets, keep one clean sheet
        for ws in list(wb.Sheets):
            try:
                if wb.Sheets.Count > 1:
                    ws.Delete()
            except Exception:
                pass
        if wb.Sheets.Count > 0:
            wb.Sheets(1).Name = "Sheet1"

        # Save .xlsm (52 = xlOpenXMLWorkbookMacroEnabled)
        wb.SaveAs(XLSM, FileFormat=52)
        print(f"  Saved {XLSM}")

        # Save .xlam (55 = xlOpenXMLAddIn)
        wb.SaveAs(XLAM, FileFormat=55)
        print(f"  Saved {XLAM}")

        wb.Close(SaveChanges=False)
        print("\nDone. Binary files updated successfully.")

    except Exception as e:
        print(f"Error: {e}")
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        raise
    finally:
        teardown(excel, wb)

if __name__ == "__main__":
    main()
