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


def _solve_error_probe(excel, wb, ws, runner, tc, args):
    """R4-07: 限额/负例必须在 VBA 内部捕获错误 (错误不穿透 COM)。"""
    from tests.test_utils import run_macro
    vbproj = wb.VBProject
    for comp in list(vbproj.VBComponents):
        if comp.Name == "SolveR4Probe":
            vbproj.VBComponents.Remove(comp)
    comp = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
    comp.Name = "SolveR4Probe"
    comp.CodeModule.AddFromString(
        "Option Explicit\r\n"
        "Public Function ProbeMode(ByVal which As String) As Double\r\n"
        "    On Error GoTo EH\r\n"
        "    Dim v As Variant\r\n"
        "    Dim r As Long, c As Long\r\n"
        "    If which = \"no_variable\" Then\r\n"
        "        Dim d1(1 To 3, 1 To 2) As Variant\r\n"
        "        d1(1, 1) = \"IncomingA\": d1(1, 2) = \"OutputY1\"\r\n"
        "        d1(2, 1) = 1#: d1(2, 2) = 10#\r\n"
        "        d1(3, 1) = 2#: d1(3, 2) = 20#\r\n"
        "        v = SolveInverse(d1, Empty, Empty, \"linear\", 42, 10)\r\n"
        "    Else\r\n"
        "        Dim d2(1 To 6, 1 To 18) As Variant\r\n"
        "        d2(1, 1) = \"VariableU1\"\r\n"
        "        For c = 2 To 18: d2(1, c) = \"Output\" & (c - 1): Next c\r\n"
        "        For r = 2 To 6\r\n"
        "            d2(r, 1) = CDbl(r)\r\n"
        "            For c = 2 To 18: d2(r, c) = CDbl(r * (c - 1)): Next c\r\n"
        "        Next r\r\n"
        "        v = SolveInverse(d2, Empty, Empty, \"linear\", 42, 10)\r\n"
        "    End If\r\n"
        "    ProbeMode = 0\r\n"
        "    Exit Function\r\n"
        "EH:\r\n"
        "    ProbeMode = 1\r\n"
        "End Function")
    mode = args[0]
    val = run_macro(excel, wb, "SolveR4Probe.ProbeMode", mode)
    return float(val), 1.0, 0.0


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
    # R4-07: 限额/负例路径 (VBA 内捕获错误)
    {"name": "Inverse_NoVariable_Raises", "func": "SolveInverse",
     "args": lambda: ("no_variable",), "reconstruct": _solve_error_probe,
     "py_ref": lambda a: 1.0},
    {"name": "Inverse_TooManyOutputs_Raises", "func": "SolveInverse",
     "args": lambda: ("too_many_outputs",), "reconstruct": _solve_error_probe,
     "py_ref": lambda a: 1.0},
]


def main() -> int:
    runner = CrossValRunner("SolveUtils", MODULE_PATHS)
    runner.run_all(TEST_CASES)
    passed, failed = runner.print_summary()
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
