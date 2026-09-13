"""Cross-validate SolveUtils (SOLVE.*) against independent Python references.

Python is the judge: forward predictions come from numpy least squares on the
raw fixture (no VBA algorithm copied); inverse anchors are the closed-form
linear solution; equation text is derived symbolically from the fixture.

Usage: python tests/crossval/build_SolveUtils.py
"""

import os
import sys

import numpy as np

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from tests.crossval.build_common import CrossValRunner  # noqa: E402
from tests.test_utils import (  # noqa: E402
    SRC_DIR,
    VBA_CORE_DIR,
    VBA_CORE_IMPORT_ORDER,
    run_macro,
    write_range,
)

MODULE_PATHS = [os.path.join(VBA_CORE_DIR, n + ".cls") for n in VBA_CORE_IMPORT_ORDER]
MODULE_PATHS.append(os.path.join(SRC_DIR, "LinearUtils.bas"))
MODULE_PATHS.append(os.path.join(SRC_DIR, "SolveUtils.bas"))

# VBA vbObjectError — SolveUtils module error codes are vbObjectError + 1601..1699
_VB_OBJECT_ERROR = -2147221504
ERR_SV_INVALID_INPUT = _VB_OBJECT_ERROR + 1601  # SolveUtils.bas:36
ERR_SV_NO_VARIABLE = _VB_OBJECT_ERROR + 1602    # SolveUtils.bas:37
ERR_SV_POLY_LIMIT = _VB_OBJECT_ERROR + 1606     # SolveUtils.bas:41
ERR_SV_LIMIT = _VB_OBJECT_ERROR + 1607          # SolveUtils.bas:42
ERR_SV_NONFINITE = _VB_OBJECT_ERROR + 1610      # SolveUtils.bas:45

# =============================================================================
# Independent fixtures — y = 2 + 0.5·IncomingA + 1.5·VariableU1
# =============================================================================
A = np.arange(1, 11, dtype=float)
U = ((A * 3) % 10) + 2.0
Y = 2.0 + 0.5 * A + 1.5 * U

FWD_EXPECTED = "OutputY1 = 2 + 0.5*IncomingA + 1.5*VariableU1"
INV_EXPECTED = "VariableU1 = (OutputY1 - 2 - 0.5*IncomingA) / 1.5"


def _linear_table(with_request=True, target=13.0):
    rows = [["IncomingA", "VariableU1", "OutputY1"]]
    for i in range(10):
        rows.append([float(A[i]), float(U[i]), float(Y[i])])
    if with_request:
        rows.append([10.0, None, float(target)])
    return rows


def _micro_table():
    xs = [0.0, 0.25, 0.5, 0.75, 1.0]
    ys = [1e-12, 1.5e-12, 2e-12, 2.5e-12, 3e-12]
    rows = [["VariableU1", "OutputY1"]]
    for x, y in zip(xs, ys):
        rows.append([x, y])
    rows.append([None, 4.5e-12])
    return rows


def _py_ols_predict(data, values):
    """Independent OLS reference via numpy (raw units, exact-linear fixture)."""
    arr = np.asarray(data, dtype=object)
    X = np.array([[float(v) for v in row] for row in arr[1:, :2]])
    y = np.array([float(v) for v in arr[1:, 2]])
    V = np.asarray(values, dtype=float)
    A_ = np.column_stack([np.ones(len(X)), X])
    beta, *_ = np.linalg.lstsq(A_, y, rcond=None)
    Av = np.column_stack([np.ones(len(V)), V])
    return (Av @ beta).tolist()


# =============================================================================
# Reconstruction probes (multi-step VBA calls → extracted scalars)
# =============================================================================

def _write_table(ws, rows):
    ws.UsedRange.ClearContents()
    write_range(ws, rows, 1, 1)
    return ws.Range(ws.Cells(1, 1), ws.Cells(len(rows), len(rows[0])))


def _inverse_field(field):
    def probe(excel, wb, ws, runner, tc, args):
        rng = _write_table(ws, _linear_table(True, tc.get("target", 13.0)))
        res = run_macro(excel, wb, "SolveUtils.UDF_SOLVE_INVERSE", rng)
        row = res[1]
        if field == "rec":
            return float(row[1]), 4.0, 1e-4
        if field == "pred":
            return float(row[2]), 13.0, 1e-6
        if field == "status":
            return (0.0 if str(row[4]) == "可达" else 1.0), float(tc["py_val"]), 0.0
        if field == "deterministic":
            res2 = run_macro(excel, wb, "SolveUtils.UDF_SOLVE_INVERSE", rng)
            same = all(float(res[1][c]) == float(res2[1][c]) for c in (1, 2, 3))
            return same, True, 0.0
        raise ValueError(f"unknown field {field}")
    return probe


def _micro_status_probe(excel, wb, ws, runner, tc, args):
    rng = _write_table(ws, _micro_table())
    res = run_macro(excel, wb, "SolveUtils.UDF_SOLVE_INVERSE", rng)
    return (0.0 if str(res[1][4]) == "可达" else 1.0), 1.0, 0.0


def _quality_probe():
    def probe(excel, wb, ws, runner, tc, args):
        rng = _write_table(ws, _linear_table(False))
        res = run_macro(excel, wb, "SolveUtils.UDF_SOLVE_QUALITY", rng, "linear", 42)
        return float(res[1][3]), 1.0, 1e-9
    return probe


def _equation_probe(row_idx):
    def probe(excel, wb, ws, runner, tc, args):
        rng = _write_table(ws, _linear_table(False))
        res = run_macro(excel, wb, "SolveUtils.UDF_SOLVE_EQUATION", rng, "linear")
        return str(res[row_idx][2]), tc["py_val"], 0.0
    return probe


# =============================================================================
# 2026-09-13 第四轮审查回归辅助 (R4-01/07/21): poly 独立参考 / 边界 / 限额
# =============================================================================

def _poly_design(A):
    """Poly 展开 (intercept + 一次项 + 平方项 + 两两交叉项), 与 VBA 展开对应。"""
    A = np.asarray(A, dtype=float)
    n, k = A.shape
    cols = [np.ones(n)]
    cols.extend([A[:, j] for j in range(k)])
    cols.extend([A[:, j] ** 2 for j in range(k)])
    for a in range(k):
        for b in range(a + 1, k):
            cols.append(A[:, a] * A[:, b])
    return np.column_stack(cols)


_POLY3_X = np.random.default_rng(7).uniform(0.5, 1.5, size=(30, 3))
_POLY3_BETA = np.array([1.0, 2.0, -1.0, 0.5, 0.5, 0.75, 1.25, -0.5, 0.3, -0.8])
_POLY3_Y = _poly_design(_POLY3_X) @ _POLY3_BETA
_POLY3_VAL = [[0.7, 1.3, 0.4], [1.1, 0.9, 1.5]]


def _poly3_table():
    rows = [["IncomingA", "VariableU1", "FixedK1", "OutputY1"]]
    for i in range(30):
        rows.append([float(_POLY3_X[i, 0]), float(_POLY3_X[i, 1]),
                     float(_POLY3_X[i, 2]), float(_POLY3_Y[i])])
    return rows


def _py_poly_predict(data, values):
    """Poly 前向预测独立参考: numpy lstsq 在展开设计矩阵上求解 (跳过表头行)。"""
    arr = np.asarray(data[1:], dtype=float)
    X = _poly_design(arr[:, :3])
    y = arr[:, 3]
    beta, *_ = np.linalg.lstsq(X, y, rcond=None)
    return (_poly_design(np.asarray(values, dtype=float)) @ beta).tolist()


def _poly_table_12():
    """12 特征 × 200 行, y 为含全部 90 个 poly 项的二阶函数 (系数非零)。"""
    rng = np.random.default_rng(12345)
    X = rng.uniform(0.5, 1.5, size=(200, 12))
    D = _poly_design(X)
    beta = rng.uniform(-0.5, 0.5, size=D.shape[1])
    y = D @ beta
    names = ["Incoming1", "Incoming2", "Incoming3",
             "Variable4", "Variable5", "Variable6", "Variable7", "Variable8", "Variable9",
             "Fixed10", "Fixed11", "Fixed12"]
    rows = [names + ["OutputY"]]
    for i in range(200):
        rows.append([float(v) for v in X[i]] + [float(y[i])])
    return rows


def _many_output_table():
    """17 个 Output 列 (超过 MAX_OUTPUTS=10): 必须返回限额错误而非 Error 9。"""
    headers = ["VariableU1"] + ["Output" + str(i) for i in range(1, 18)]
    rows = [headers]
    for r in range(1, 6):
        rows.append([float(r)] + [float(r * i) for i in range(1, 18)])
    return rows


def _auto_quality_probe(excel, wb, ws, runner, tc, args):
    """R4-21: auto 模型选择路径 — 每个输出恰好选中 1 个候选。"""
    rng = _write_table(ws, _poly3_table())
    res = run_macro(excel, wb, "SolveUtils.UDF_SOLVE_QUALITY", rng, "auto", 42)
    selected = sum(1 for row in res[1:] if str(row[5]) == "是")
    return float(selected), 1.0, 0.0


def _bounds_probe(excel, wb, ws, runner, tc, args):
    """R4-21: bounds 独立数组路径 — 推荐值必须落在边界内。"""
    rng = _write_table(ws, _linear_table(True, 13.0))
    res = run_macro(excel, wb, "SolveUtils.UDF_SOLVE_INVERSE",
                    rng, [], [[0.0, 100.0]], "linear", 42, 10)
    rec = float(res[1][1])
    return (1.0 if 0.0 <= rec <= 100.0 else 0.0), 1.0, 0.0


def _equation_poly_probe(excel, wb, ws, runner, tc, args):
    """R4-01: 90 项 poly 方程文本构建不得越界 (修复前 UDF 返回 #VALUE!)。"""
    rng = _write_table(ws, _poly_table_12())
    res = run_macro(excel, wb, "SolveUtils.UDF_SOLVE_EQUATION", rng, "poly")
    try:
        txt = str(res[1][2])
        return (1.0 if "Incoming1" in txt else 0.0), 1.0, 0.0
    except Exception:
        return 0.0, 1.0, 0.0


# =============================================================================
# 2026-09-13 第五轮审查回归探针 (R5-29/40/41/42/43)
# =============================================================================

def _independent_request_no_role():
    """R5-40: 独立 request 表无 Variable*/Output* 前缀列 (仅角色映射)。"""
    return [["IncomingA", "OutputY1"], [10.0, 13.0]]


def _independent_request_with_init():
    """R5-29: 独立 request 表显式给出可调变量初值。"""
    return [["IncomingA", "VariableU1", "OutputY1"], [10.0, 4.0, 13.0]]


def _multi_output_table():
    """R5-29: 多输出数据表 — Y1 = 2 + 5U, Y2 = -1 + 3U。"""
    rows = [["VariableU1", "OutputY1", "OutputY2"]]
    for i in range(1, 11):
        rows.append([float(i), 2.0 + 5.0 * i, -1.0 + 3.0 * i])
    return rows


def _inverse_request_probe(field, request_factory):
    """SolveInverse 数组路径 + 独立 request 表 (Python list → COM Variant)。"""
    def probe(excel, wb, ws, runner, tc, args):
        data = _linear_table(False)
        res = run_macro(excel, wb, "SolveUtils.SolveInverse",
                        data, request_factory(), [], "linear", 42, 10)
        row = res[1]
        if field == "rec":
            return float(row[1]), 4.0, 1e-4
        if field == "pred":
            return float(row[2]), 13.0, 1e-6
        if field == "status":
            return (0.0 if str(row[4]) == "可达" else 1.0), 0.0, 0.0
        raise ValueError(f"unknown field {field}")
    return probe


def _inverse_multi_output_probe(excel, wb, ws, runner, tc, args):
    """R5-29: 多输出反解 — U=5 同时命中 Y1=27、Y2=14。"""
    request = [["VariableU1", "OutputY1", "OutputY2"], [5.0, 27.0, 14.0]]
    res = run_macro(excel, wb, "SolveUtils.SolveInverse",
                    _multi_output_table(), request, [], "linear", 42, 10)
    row = res[1]
    dev = (abs(float(row[1]) - 5.0) + abs(float(row[2]) - 27.0)
           + abs(float(row[3]) - 14.0))
    return dev, 0.0, 1e-4


def _quality_array_probe(excel, wb, ws, runner, tc, args):
    """R5-29: SolveQuality Python-list 数组路径 (此前仅 SolvePredict 有)。"""
    res = run_macro(excel, wb, "SolveUtils.SolveQuality",
                    _linear_table(False), "linear", 42)
    return float(res[1][3]), 1.0, 1e-9


def _equation_array_probe(row_idx):
    """R5-29: SolveEquation Python-list 数组路径。"""
    def probe(excel, wb, ws, runner, tc, args):
        res = run_macro(excel, wb, "SolveUtils.SolveEquation",
                        _linear_table(False), "linear")
        return str(res[row_idx][2]), tc["py_val"], 0.0
    return probe


def _g6_edge_probe(excel, wb, ws, runner, tc, args):
    """R5-42: 999999.7 的 G6 文本必须是 1E+06 (经 SolveEquation 格式化)。"""
    rows = [["VariableU1", "OutputY1"]]
    for i in range(1, 7):
        rows.append([float(i), 999999.7 + float(i)])
    res = run_macro(excel, wb, "SolveUtils.SolveEquation", rows, "linear")
    txt = str(res[1][2])
    return (1.0 if "1E+06" in txt else 0.0), 1.0, 0.0


_SOLVE_PROBE_CODE = "\r\n".join([
    "Option Explicit",
    "",
    "Public Function ProbeErr(ByVal kind As String) As Double",
    "    On Error GoTo EH",
    "    Dim res As Variant",
    "    Dim tbl As Variant",
    "    Dim rr As Long, cc As Long",
    "    If kind = \"no_variable\" Then",
    "        ReDim tbl(1 To 3, 1 To 2)",
    "        tbl(1, 1) = \"IncomingA\": tbl(1, 2) = \"OutputY1\"",
    "        tbl(2, 1) = 1#: tbl(2, 2) = 10#",
    "        tbl(3, 1) = 2#: tbl(3, 2) = 20#",
    "        res = SolveInverse(tbl, Empty, Empty, \"linear\", 42, 10)",
    "    ElseIf kind = \"too_many_outputs\" Then",
    "        ReDim tbl(1 To 6, 1 To 18)",
    "        tbl(1, 1) = \"VariableU1\"",
    "        For cc = 2 To 18: tbl(1, cc) = \"Output\" & (cc - 1): Next cc",
    "        For rr = 2 To 6",
    "            tbl(rr, 1) = CDbl(rr)",
    "            For cc = 2 To 18: tbl(rr, cc) = CDbl(rr * (cc - 1)): Next cc",
    "        Next rr",
    "        res = SolveInverse(tbl, Empty, Empty, \"linear\", 42, 10)",
    "    ElseIf kind = \"poly_limit\" Then",
    "        ReDim tbl(1 To 7, 1 To 14)",
    "        For cc = 1 To 7: tbl(1, cc) = \"Incoming\" & cc: Next cc",
    "        For cc = 8 To 10: tbl(1, cc) = \"Variable\" & cc: Next cc",
    "        For cc = 11 To 13: tbl(1, cc) = \"Fixed\" & cc: Next cc",
    "        tbl(1, 14) = \"OutputY1\"",
    "        For rr = 2 To 7",
    "            For cc = 1 To 13: tbl(rr, cc) = CDbl(rr) / 10#: Next cc",
    "            tbl(rr, 14) = CDbl(rr)",
    "        Next rr",
    "        res = SolveEquation(tbl, \"poly\")",
    "    ElseIf kind = \"nonfinite\" Then",
    "        ReDim tbl(1 To 6, 1 To 2)",
    "        tbl(1, 1) = \"VariableU1\": tbl(1, 2) = \"OutputY1\"",
    "        For rr = 2 To 6",
    "            tbl(rr, 1) = 1E+200 * CDbl(rr)",
    "            tbl(rr, 2) = CDbl(rr)",
    "        Next rr",
    "        res = SolveEquation(tbl, \"poly\")",
    "    End If",
    "    ProbeErr = 0#",
    "    Exit Function",
    "EH:",
    "    ProbeErr = CDbl(Err.Number)",
    "End Function",
    "",
    "Public Function ProbeCoreInvalidOpt(ByVal data As Range, ByVal bad As Range) As Double",
    "    On Error GoTo EH",
    "    Dim res As Variant",
    "    res = SolveInverse(data, bad, Empty, \"linear\", 42, 10)",
    "    ProbeCoreInvalidOpt = 0#",
    "    Exit Function",
    "EH:",
    "    ProbeCoreInvalidOpt = CDbl(Err.Number)",
    "End Function",
    "",
    "Public Function ProbeUdfInvalidOpt(ByVal data As Range, ByVal bad As Range) As Double",
    "    On Error GoTo EH",
    "    Dim res As Variant",
    "    res = UDF_SOLVE_INVERSE(data, bad)",
    "    If IsError(res) Then",
    "        ProbeUdfInvalidOpt = -1#",
    "    Else",
    "        ProbeUdfInvalidOpt = 0#",
    "    End If",
    "    Exit Function",
    "EH:",
    "    ProbeUdfInvalidOpt = 99#",
    "End Function",
])


def _inject_solve_probe(wb):
    """Inject SolveProbeR5 (idempotent); save so the injected macro is runnable."""
    vbproj = wb.VBProject
    for comp in list(vbproj.VBComponents):
        if comp.Name == "SolveProbeR5":
            return
    comp = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
    comp.Name = "SolveProbeR5"
    comp.CodeModule.AddFromString(_SOLVE_PROBE_CODE)
    wb.Save()


def _solve_error_probe(excel, wb, ws, runner, tc, args):
    """R5-29: 在 VBA 内部捕获错误码并返回真实 Err.Number (非恒定 1)。"""
    mode = args[0]
    expected = float(args[1])
    _inject_solve_probe(wb)
    err = run_macro(excel, wb, "SolveProbeR5.ProbeErr", mode)
    return float(err), expected, 0.0


def _invalid_optional_probe(which):
    """R5-43: 多区域 Range 作为可选表 — UDF 返回 CVErr, 核心抛模块错误。"""
    def probe(excel, wb, ws, runner, tc, args):
        data_rng = _write_table(ws, _linear_table(False))
        second = ws.Range(ws.Cells(15, 1), ws.Cells(16, 2))
        bad = excel.Union(ws.Range("A1:B2"), second)
        macro = ("SolveProbeR5.ProbeUdfInvalidOpt" if which == "udf"
                 else "SolveProbeR5.ProbeCoreInvalidOpt")
        val = run_macro(excel, wb, macro, data_rng, bad)
        return float(val), float(tc["py_val"]), 0.0
    return probe


# =============================================================================
# Test cases
# =============================================================================

TEST_CASES = [
    # ── Forward prediction: core (Range path) + UDF (Range path) ────────────
    {"name": "Predict_Linear_RangePath", "func": "SolvePredict", "range_path": True,
     "args": lambda: (_linear_table(False), [[10.0, 4.0], [1.0, 5.0]], "linear"),
     "py_ref": lambda a: _py_ols_predict(a[0], a[1]),
     "result_type": "array", "tol": 1e-8},

    {"name": "UDF_Predict_Linear", "func": "UDF_SOLVE_PREDICT", "is_udf": True,
     "args": lambda: (_linear_table(False), [[10.0, 4.0], [1.0, 5.0]], "linear"),
     "py_ref": lambda a: _py_ols_predict(a[0], a[1]),
     "result_type": "array", "tol": 1e-8},

    # ── Inverse anchors (closed-form / status) ──────────────────────────────
    {"name": "Inverse_Recommendation", "func": "UDF_SOLVE_INVERSE",
     "args": lambda: (), "reconstruct": _inverse_field("rec"), "py_ref": lambda a: 4.0},
    {"name": "Inverse_Prediction", "func": "UDF_SOLVE_INVERSE",
     "args": lambda: (), "reconstruct": _inverse_field("pred"), "py_ref": lambda a: 13.0},
    {"name": "Inverse_Reachable", "func": "UDF_SOLVE_INVERSE",
     "args": lambda: (), "reconstruct": _inverse_field("status"), "py_val": 0.0, "py_ref": lambda a: 0.0},
    {"name": "Inverse_Unreachable", "func": "UDF_SOLVE_INVERSE",
     "args": lambda: (), "reconstruct": _inverse_field("status"), "target": 1000.0,
     "py_val": 1.0, "py_ref": lambda a: 1.0},
    {"name": "Inverse_MicroScaleRelativeTolerance", "func": "UDF_SOLVE_INVERSE",
     "args": lambda: (), "reconstruct": _micro_status_probe, "py_ref": lambda a: 1.0},
    {"name": "Inverse_SameSeedDeterministic", "func": "UDF_SOLVE_INVERSE",
     "args": lambda: (), "reconstruct": _inverse_field("deterministic"), "result_type": "bool",
     "py_ref": lambda a: True},

    # ── Model quality (exact-linear fixture ⇒ out-of-fold R² = 1) ───────────
    {"name": "Quality_Linear_ExactFixture", "func": "UDF_SOLVE_QUALITY",
     "args": lambda: (), "reconstruct": _quality_probe(), "py_ref": lambda a: 1.0},

    # ── Equation text (G6, derived symbolically in Python) ──────────────────
    {"name": "Equation_Forward", "func": "UDF_SOLVE_EQUATION",
     "args": lambda: (), "reconstruct": _equation_probe(1), "py_val": FWD_EXPECTED,
     "result_type": "string", "py_ref": lambda a: FWD_EXPECTED},
    {"name": "Equation_Inverse", "func": "UDF_SOLVE_EQUATION",
     "args": lambda: (), "reconstruct": _equation_probe(2), "py_val": INV_EXPECTED,
     "result_type": "string", "py_ref": lambda a: INV_EXPECTED},

    # ── 2026-09-13 第四轮审查回归 (R4-01/07/21) ─────────────────────────────
    # 数组路径 (Python list → COM Variant) — 此前 11 例全部走 Range 路径
    {"name": "Predict_Linear_ArrayPath", "func": "SolvePredict",
     "args": lambda: (_linear_table(False), [[10.0, 4.0], [1.0, 5.0]], "linear"),
     "py_ref": lambda a: _py_ols_predict(a[0], a[1]),
     "result_type": "array", "tol": 1e-8},
    # poly 模型前向预测 (独立 numpy lstsq 参考; VBA 侧 poly 含小岭回归,
    # 与无正则解存在 ~1e-5 量级偏差, 故容差取 1e-4 验证路径正确性)
    {"name": "Predict_Poly_ArrayPath", "func": "SolvePredict",
     "args": lambda: (_poly3_table(), _POLY3_VAL, "poly"),
     "py_ref": lambda a: _py_poly_predict(a[0], a[1]),
     "result_type": "array", "tol": 1e-4},
    # auto 模型选择路径
    {"name": "Quality_Auto_SelectsOne", "func": "UDF_SOLVE_QUALITY",
     "args": lambda: (), "reconstruct": _auto_quality_probe, "py_ref": lambda a: 1.0},
    # bounds 数组路径
    {"name": "Inverse_Bounds_ArrayPath", "func": "UDF_SOLVE_INVERSE",
     "args": lambda: (), "reconstruct": _bounds_probe, "py_ref": lambda a: 1.0},
    # R4-01: 90 项 poly 方程 (缓冲由 POLY_TERM_LIMIT 派生)
    {"name": "Equation_Poly_90Terms", "func": "UDF_SOLVE_EQUATION",
     "args": lambda: (), "reconstruct": _equation_poly_probe, "py_ref": lambda a: 1.0},
    # R4-07: 限额/负例路径 (VBA 内捕获错误, 断言具体错误码)
    {"name": "Inverse_NoVariable_Raises", "func": "SolveInverse",
     "args": lambda: ("no_variable", ERR_SV_NO_VARIABLE), "reconstruct": _solve_error_probe,
     "py_ref": lambda a: float(ERR_SV_NO_VARIABLE)},
    {"name": "Inverse_TooManyOutputs_Raises", "func": "SolveInverse",
     "args": lambda: ("too_many_outputs", ERR_SV_LIMIT), "reconstruct": _solve_error_probe,
     "py_ref": lambda a: float(ERR_SV_LIMIT)},
    {"name": "Equation_PolyLimit_Raises", "func": "SolveEquation",
     "args": lambda: ("poly_limit", ERR_SV_POLY_LIMIT), "reconstruct": _solve_error_probe,
     "py_ref": lambda a: float(ERR_SV_POLY_LIMIT)},

    # =====================================================================
    # 2026-09-13 第五轮审查回归 (R5-29/40/41/42/43)
    # =====================================================================
    # R5-40: 独立 request 表 (仅 IncomingA+OutputY1, 无 Variable*/Output* 角色列)
    {"name": "Inverse_IndependentRequest_NoRole_Rec", "func": "SolveInverse",
     "args": lambda: (),
     "reconstruct": _inverse_request_probe("rec", _independent_request_no_role),
     "py_ref": lambda a: 4.0},
    {"name": "Inverse_IndependentRequest_NoRole_Pred", "func": "SolveInverse",
     "args": lambda: (),
     "reconstruct": _inverse_request_probe("pred", _independent_request_no_role),
     "py_ref": lambda a: 13.0},
    # R5-29: 独立 request 表显式提供可调变量初值
    {"name": "Inverse_RequestInitialValue_ArrayPath", "func": "SolveInverse",
     "args": lambda: (),
     "reconstruct": _inverse_request_probe("rec", _independent_request_with_init),
     "py_ref": lambda a: 4.0},
    # R5-29: 多输出 request (两个 Output 目标同时反解)
    {"name": "Inverse_MultiOutput_ArrayPath", "func": "SolveInverse",
     "args": lambda: (), "reconstruct": _inverse_multi_output_probe,
     "py_ref": lambda a: 0.0},
    # R5-29: SolveQuality/SolveEquation Python-list 数组路径
    {"name": "Quality_Linear_ArrayPath", "func": "SolveQuality",
     "args": lambda: (), "reconstruct": _quality_array_probe,
     "py_ref": lambda a: 1.0},
    {"name": "Equation_Forward_ArrayPath", "func": "SolveEquation",
     "args": lambda: (), "reconstruct": _equation_array_probe(1), "py_val": FWD_EXPECTED,
     "result_type": "string", "py_ref": lambda a: FWD_EXPECTED},
    {"name": "Equation_Inverse_ArrayPath", "func": "SolveEquation",
     "args": lambda: (), "reconstruct": _equation_array_probe(2), "py_val": INV_EXPECTED,
     "result_type": "string", "py_ref": lambda a: INV_EXPECTED},
    # R5-41: poly 巨值 (1e200) → ERR_SV_NONFINITE, 不得裸 Error 6
    {"name": "Equation_Poly_NonFinite_Raises", "func": "SolveEquation",
     "args": lambda: ("nonfinite", ERR_SV_NONFINITE), "reconstruct": _solve_error_probe,
     "py_ref": lambda a: float(ERR_SV_NONFINITE)},
    # R5-42: G6 边界 999999.7 → "1E+06" (经 SolveEquation 格式化路径)
    {"name": "Equation_G6_Edge_1E06", "func": "SolveEquation",
     "args": lambda: (), "reconstruct": _g6_edge_probe, "py_ref": lambda a: 1.0},
    # R5-43: 非法可选表 (多区域 Range) — UDF 返回 CVErr, 核心抛 ERR_SV_INVALID_INPUT
    {"name": "UDF_INVERSE_InvalidOptional_CVErr", "func": "UDF_SOLVE_INVERSE",
     "args": lambda: (), "reconstruct": _invalid_optional_probe("udf"),
     "py_val": -1.0, "py_ref": lambda a: -1.0},
    {"name": "INVERSE_InvalidOptional_Core_Raises", "func": "SolveInverse",
     "args": lambda: (), "reconstruct": _invalid_optional_probe("core"),
     "py_val": float(ERR_SV_INVALID_INPUT), "py_ref": lambda a: float(ERR_SV_INVALID_INPUT)},
]


def main() -> int:
    runner = CrossValRunner("SolveUtils", MODULE_PATHS)
    runner.run_all(TEST_CASES)
    passed, failed = runner.print_summary()
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
