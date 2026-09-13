"""Verify VBA-Core classes via COM — runs Test_* procedures and reports.

Usage: python tests/crossval/build_vba_core.py

VBA class methods cannot be called directly via Application.Run.
This module injects a VBA wrapper that instantiates each class and
runs its Test_* method.

Correctness is verified by the 145 VBA assertions across 3 classes:
  VariantKit (50), ArrayOps (51), DictProxy (44), plus the R5 regression
  probes defined in this builder (tests-only; VBA-Core source untouched).
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

    ' =======================================================================
    ' R5 regression probes (2026-09-13) — tests-only; source modules untouched
    ' =======================================================================
    WriteProbeRow ws, r, "R5-08_multiarea_range_fail_closed", RunR5(8)
    WriteProbeRow ws, r, "R5-09_sort_null_empty", RunR5(9)
    WriteProbeRow ws, r, "R5-10_merge_cross_comparemode", RunR5(10)
    WriteProbeRow ws, r, "R5-58_oern_guard_idiom", RunR5(58)

    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' Probe driver — runs a probe and converts an escaping runtime error into a
' nonzero code so a probe row is always written (never silently skipped).
' ---------------------------------------------------------------------------
Private Function RunR5(ByVal probeId As Long) As Long
    Dim pcode As Long
    pcode = 0
    On Error Resume Next
    Err.Clear
    Select Case probeId
        Case 8: pcode = R5_08_Probe()
        Case 9: pcode = R5_09_Probe()
        Case 10: pcode = R5_10_Probe()
        Case 58: pcode = R5_58_Probe()
        Case Else: pcode = -1
    End Select
    If Err.Number <> 0 Then pcode = Err.Number
    Err.Clear
    On Error GoTo 0
    RunR5 = pcode
End Function

Private Sub WriteProbeRow(ByVal ws As Worksheet, ByRef r As Long, _
                          ByVal tname As String, ByVal pcode As Long)
    ws.Cells(r, 1) = "VBA-Core"
    ws.Cells(r, 2) = tname
    If pcode = 0 Then
        ws.Cells(r, 3) = "PASS"
    Else
        ws.Cells(r, 3) = "FAIL"
        ws.Cells(r, 4) = "probe code " & pcode
    End If
    r = r + 1
End Sub

' ---------------------------------------------------------------------------
' R5-08 — multi-area Range must fail closed (no silent first-area pick);
'         single-area Range must keep working.
' ---------------------------------------------------------------------------
Private Function R5_08_Probe() As Long
    Dim vk As VariantKit
    Dim wsData As Worksheet
    Dim singleArea As Range, multiArea As Range
    Dim v As Variant, res As Variant
    Dim nr As Long, nc As Long
    Dim pErr As Long
    Dim isE As Boolean

    Set vk = New VariantKit
    Set wsData = ThisWorkbook.Sheets("TestData")
    wsData.Range("A1:B2").Value = 1
    Set singleArea = wsData.Range("A1:B2")
    Set multiArea = Union(wsData.Range("A1:A2"), wsData.Range("C1:C2"))

    ' Positive: single-area NormalizeInput still converts in place
    Set v = singleArea
    If Not vk.NormalizeInput(v) Then R5_08_Probe = 1: Exit Function
    If IsObject(v) Then R5_08_Probe = 2: Exit Function
    If Not (v(1, 1) = 1) Then R5_08_Probe = 3: Exit Function

    ' Positive: single-area NormalizeTo2D still returns 2x2
    res = vk.NormalizeTo2D(singleArea, nr, nc)
    If IsError(res) Then R5_08_Probe = 4: Exit Function
    If nr <> 2 Or nc <> 2 Then R5_08_Probe = 5: Exit Function

    ' Negative: multi-area NormalizeInput must return False, v untouched
    Set v = multiArea
    If vk.NormalizeInput(v) Then R5_08_Probe = 6: Exit Function
    If Not IsObject(v) Then R5_08_Probe = 7: Exit Function

    ' Negative: multi-area NormalizeTo2D must return CVErr or raise
    On Error Resume Next
    Err.Clear
    res = vk.NormalizeTo2D(multiArea, nr, nc)
    pErr = Err.Number
    isE = IsError(res)
    Err.Clear
    On Error GoTo 0
    If pErr = 0 And Not isE Then R5_08_Probe = 8: Exit Function

    ' Negative: Normalize1D / Normalize2D keep the CVErr contract
    res = vk.Normalize1D(multiArea)
    If Not IsError(res) Then R5_08_Probe = 9: Exit Function
    res = vk.Normalize2D(multiArea)
    If Not IsError(res) Then R5_08_Probe = 10
End Function

' ---------------------------------------------------------------------------
' R5-09 — Sort/SortIndices with Null/Empty and numeric/text comparers must
'         not raise Error 94 (CStr(Null)); VK value order must apply.
' ---------------------------------------------------------------------------
Private Function R5_09_Probe() As Long
    Dim ao As ArrayOps
    Dim arr As Variant
    Dim idx() As Long

    Set ao = New ArrayOps

    ' numeric comparer x Null-vs-text mix (original R5-09 counterexample)
    arr = Array("b", Null, "a")
    ao.Sort arr, True, "numeric"
    If Not IsNull(arr(0)) Then R5_09_Probe = 1: Exit Function
    If Not (arr(1) = "a") Then R5_09_Probe = 2: Exit Function
    If Not (arr(2) = "b") Then R5_09_Probe = 3: Exit Function

    ' text comparer x Null
    arr = Array(3, Null, 1)
    ao.Sort arr, True, "text"
    If Not IsNull(arr(0)) Then R5_09_Probe = 4: Exit Function
    If Not (arr(1) = 1) Then R5_09_Probe = 5: Exit Function
    If Not (arr(2) = 3) Then R5_09_Probe = 6: Exit Function

    ' text comparer x Empty
    arr = Array("b", Empty, "a")
    ao.Sort arr, True, "text"
    If Not IsEmpty(arr(0)) Then R5_09_Probe = 7: Exit Function
    If Not (arr(1) = "a") Then R5_09_Probe = 8: Exit Function
    If Not (arr(2) = "b") Then R5_09_Probe = 9: Exit Function

    ' SortIndices text comparer x Null
    arr = Array("b", Null, "a")
    ReDim idx(0 To 2)
    idx(0) = 0: idx(1) = 1: idx(2) = 2
    ao.SortIndices arr, idx, True, "text"
    If Not (idx(0) = 1) Then R5_09_Probe = 10: Exit Function
    If Not (idx(1) = 2) Then R5_09_Probe = 11: Exit Function
    If Not (idx(2) = 0) Then R5_09_Probe = 12
End Function

' ---------------------------------------------------------------------------
' R5-10 — Merge across different CompareModes must promote to vbBinaryCompare
'         so text{"a"} + binary{"A"} keeps both keys (count = 2).
' ---------------------------------------------------------------------------
Private Function R5_10_Probe() As Long
    Dim dp As DictProxy
    Dim dText As Object, dBin As Object, m As Object

    Set dp = New DictProxy
    Set dText = dp.Create()
    dText.Add "a", 1
    Set dBin = dp.Create(vbBinaryCompare)
    dBin.Add "A", 2

    Set m = dp.Merge(dText, dBin, False)
    If Not (m.Count = 2) Then R5_10_Probe = 1: Exit Function
    If Not (m.CompareMode = vbBinaryCompare) Then R5_10_Probe = 2: Exit Function

    Set m = dp.Merge(dBin, dText, False)
    If Not (m.Count = 2) Then R5_10_Probe = 3: Exit Function
    If Not (m.CompareMode = vbBinaryCompare) Then R5_10_Probe = 4: Exit Function

    ' same-mode merge stays in the source mode
    Set m = dp.Merge(dText, dp.Create(), False)
    If Not (m.CompareMode = vbTextCompare) Then R5_10_Probe = 5
End Function

' ---------------------------------------------------------------------------
' R5-58 — the fixed `If Err.Number = 0 Then On Error GoTo 0: Err.Raise 5`
'         idiom must propagate under an active On Error Resume Next caller
'         (previously the raise was swallowed and assertions silently passed).
' ---------------------------------------------------------------------------
Private Sub R5_58_Raise()
    On Error Resume Next
    Err.Clear
    If Err.Number = 0 Then On Error GoTo 0: Err.Raise 5
End Sub

Private Function R5_58_Probe() As Long
    Dim caught As Long
    On Error Resume Next
    Err.Clear
    R5_58_Raise
    caught = Err.Number
    Err.Clear
    On Error GoTo 0
    If caught <> 5 Then R5_58_Probe = 1
End Function
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
