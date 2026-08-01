"""Verify VBA-Core classes via COM — runs Test_* procedures and reports.

Usage: python tests/crossval/build_vba_core.py

VBA class methods cannot be called directly via Application.Run.
This module injects a VBA wrapper that instantiates each class and
runs its Test_* method.

Correctness is verified by the 145 VBA assertions across 3 classes:
  VariantKit (50), ArrayOps (51), DictProxy (44).
"""

import os, sys, tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__)))))

from tests.test_utils import (
    ensure_excel, teardown, create_workbook, inject_testrunner,
    run_macro, read_results, print_report,
    VBA_CORE_DIR, VBA_CORE_IMPORT_ORDER,
)

VBA_CORE_RUNNER = r"""
Option Explicit

Public Sub VBA_Core_RunAll()
    Dim ws As Worksheet, r As Long, errMsg As String
    Set ws = ThisWorkbook.Sheets("TestResults")
    r = ws.Cells(ws.Rows.Count, 1).End(-4162).Row + 1

    On Error Resume Next

    Dim vk As New VariantKit: vk.Test_VariantKit
    errMsg = ""
    If Err.Number <> 0 Then errMsg = "Err " & Err.Number & " - " & Err.Description: Err.Clear
    ws.Cells(r,1)="VariantKit":ws.Cells(r,2)="Test_VariantKit"
    If errMsg = "" Then ws.Cells(r,3)="PASS" Else ws.Cells(r,3)="FAIL": ws.Cells(r,4)=errMsg
    r=r+1

    Dim ao As New ArrayOps: ao.Test_ArrayOps
    errMsg = ""
    If Err.Number <> 0 Then errMsg = "Err " & Err.Number & " - " & Err.Description: Err.Clear
    ws.Cells(r,1)="ArrayOps":ws.Cells(r,2)="Test_ArrayOps"
    If errMsg = "" Then ws.Cells(r,3)="PASS" Else ws.Cells(r,3)="FAIL": ws.Cells(r,4)=errMsg
    r=r+1

    Dim dp As New DictProxy: dp.Test_DictProxy
    errMsg = ""
    If Err.Number <> 0 Then errMsg = "Err " & Err.Number & " - " & Err.Description: Err.Clear
    ws.Cells(r,1)="DictProxy":ws.Cells(r,2)="Test_DictProxy"
    If errMsg = "" Then ws.Cells(r,3)="PASS" Else ws.Cells(r,3)="FAIL": ws.Cells(r,4)=errMsg
    r=r+1

    On Error GoTo 0
End Sub
"""


def main() -> int:
    print("=" * 60)
    print("  VBA-Core Class Verification")
    print("=" * 60)

    paths = [os.path.join(VBA_CORE_DIR, n + ".cls")
             for n in VBA_CORE_IMPORT_ORDER]

    excel = ensure_excel()
    wb = None
    try:
        out = os.path.join(tempfile.gettempdir(), "vba_core_test.xlsm")
        wb = create_workbook(excel, out, paths,
                             import_order=VBA_CORE_IMPORT_ORDER)
        inject_testrunner(wb)
        run_macro(excel, wb, "TestRunner.RunAllTests")

        # Inject VBA wrapper and run
        comp = wb.VBProject.VBComponents.Add(1)
        comp.Name = "VBA_Core_Runner"
        comp.CodeModule.AddFromString(VBA_CORE_RUNNER)
        run_macro(excel, wb, "VBA_Core_Runner.VBA_Core_RunAll")

        passed, failed, details = read_results(wb)
        print_report(passed, failed, details, label="VBA-Core Results")
        return 0 if failed == 0 else 1
    finally:
        teardown(excel, wb)


if __name__ == "__main__":
    sys.exit(main())
