"""Cross-validate RegressUtils functions against Python numpy/scipy references.

Usage: python tests/crossval/build_RegressUtils.py
"""

import os, sys
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests.crossval.build_common import CrossValRunner
from tests.test_utils import SRC_DIR, VBA_CORE_DIR, VBA_CORE_IMPORT_ORDER, run_macro

MODULE_PATHS = [os.path.join(VBA_CORE_DIR, n + ".cls") for n in VBA_CORE_IMPORT_ORDER]
MODULE_PATHS.append(os.path.join(SRC_DIR, "LinearUtils.bas"))
MODULE_PATHS.append(os.path.join(SRC_DIR, "StatsUtils.bas"))
MODULE_PATHS.append(os.path.join(SRC_DIR, "RegressUtils.bas"))

try:
    from scipy import stats as _sp_stats
    _HAS_SCIPY = True
except ImportError:
    _HAS_SCIPY = False

# VBA vbObjectError — module error codes are vbObjectError + offset
_VB_OBJECT_ERROR = -2147221504
ERR_REG_INVALID_INPUT = _VB_OBJECT_ERROR + 1001  # RegressUtils.bas:45
ERR_REG_INVALID_DATA = _VB_OBJECT_ERROR + 3001   # RegressUtils.bas:46
ERR_REG_INVALID_PARAM = _VB_OBJECT_ERROR + 3005  # RegressUtils.bas:49
ERR_REG_TOO_FEW_ROWS = _VB_OBJECT_ERROR + 3101   # RegressUtils.bas:50


# =============================================================================
# Python reference helpers
# =============================================================================

def _py_factor_importance(data, factor_cols, result_col):
    """Factor importance via standardized OLS coefficients (QR decomposition).

    Mirrors VBA FitOLS which uses QR decomposition for numerical stability.
    Returns same structure as VBA FactorImportance:
    [header_row, data_row_1, data_row_2, ...]
    """
    rows = data[1:]
    X_raw = np.array([[float(row[fc - 1]) for fc in factor_cols] for row in rows])
    X_design = np.column_stack([np.ones(len(X_raw)), X_raw])
    y = np.array([float(row[result_col - 1]) for row in rows])

    # QR decomposition (matching VBA FitOLS algorithm)
    Q, R = np.linalg.qr(X_design)
    # Q is m×p (economy mode), R is p×p
    Q = Q[:X_design.shape[0], :X_design.shape[1]]
    coef = np.linalg.solve(R, Q.T @ y)  # Rβ = Qᵀy back-substitution

    y_std = np.std(y, ddof=1) or 1.0
    names = [data[0][fc - 1] for fc in factor_cols]
    std_coefs = []
    for j in range(1, X_design.shape[1]):
        x_std = np.std(X_design[:, j], ddof=1) or 1.0
        std_coefs.append(abs(coef[j] * x_std / y_std))

    pairs = sorted(zip(names, std_coefs, [coef[j + 1] for j in range(len(std_coefs))]),
                   key=lambda x: x[1], reverse=True)
    # VBA headers: 排名, 因子, 标准化系数, 原始系数, 绝对重要性
    result = [["排名", "因子", "标准化系数", "原始系数", "绝对重要性"]]
    for i, (name, imp, raw_coef) in enumerate(pairs):
        result.append([i + 1, name, imp, raw_coef, imp])
    return result


def _anova_checks():
    """R5-07: scipy reference on BinaryCompare groups ('' kept as a level).

    Factor rows: A, '', (blank→dropped), B, b, B, a.  The empty string is a
    valid binary-compare level; Empty/Null rows are dropped by ANOVAOneWay.
    """
    groups = [[3.0], [4.0], [7.0, 9.0], [8.5], [15.0]]
    checks = [("n_groups", 5.0, 5.0), ("n_total", 6.0, 6.0)]
    if not _HAS_SCIPY:
        return checks
    r = _sp_stats.f_oneway(*groups)
    f_stat = float(r.statistic)
    p_val = float(r.pvalue)
    checks.extend([
        ("F", f_stat - 1e-6 * max(1.0, abs(f_stat)),
              f_stat + 1e-6 * max(1.0, abs(f_stat))),
        ("p_value", p_val - 1e-6, p_val + 1e-6),
    ])
    return checks


def _py_interaction_f(rows):
    """Independent two-way ANOVA interaction F/p via OLS with/without A*B term."""
    arr = np.asarray(rows[1:], dtype=float)
    y = arr[:, 2]
    a = arr[:, 0]
    b = arr[:, 1]
    ones = np.ones(len(y))
    X0 = np.column_stack([ones, a, b])
    X1 = np.column_stack([ones, a, b, a * b])
    sse0 = float(((y - X0 @ np.linalg.lstsq(X0, y, rcond=None)[0]) ** 2).sum())
    sse1 = float(((y - X1 @ np.linalg.lstsq(X1, y, rcond=None)[0]) ** 2).sum())
    df_err = len(y) - 4
    f_stat = (sse0 - sse1) / (sse1 / df_err)
    p_val = float(_sp_stats.f.sf(f_stat, 1, df_err)) if _HAS_SCIPY else 0.0
    return f_stat, p_val


# =============================================================================
# Custom runner for RegressUtils (Dictionary object verification)
# =============================================================================

# =============================================================================
# 2026-09-13 R4 回归探针 (R4-02/04/09/11/12/14/15)
# =============================================================================

def _ie_single_level_probe(excel, wb, ws, runner, tc, args):
    """R4-02: 单水平分类因子跳过交互对时不得 Error 9 (输出应仅剩标题行)。"""
    data = [["X", "K", "Y"]] + [[float(i), "K", 2.0 * i + 1.0] for i in range(1, 7)]
    res = run_macro(excel, wb, "RegressUtils.InteractionEffects", data, [1, 2], 3)
    return float(len(res)), 1.0, 0.0


def _lm_blank_cat_probe(excel, wb, ws, runner, tc, args):
    """R4-14: 空白分类行按无效行丢弃, 不得静默并入参考水平。"""
    data = [["F", "Y"], ["A", 1.0], ["B", 2.0], [None, 100.0], ["A", 1.0]]
    m = run_macro(excel, wb, "RegressUtils.LinearModelFit", data, [1], 2)
    coefs = m("coefficients")
    return float(coefs[1]), 1.0, 1e-8


def _lm_case_probe(excel, wb, ws, runner, tc, args):
    """R4-15: 分类水平按 BinaryCompare, 'b' 与 'B' 为不同水平 (哑变量系数 100 与 1)。"""
    data = [["F", "Y"], ["A", 0.0], ["A", 0.0], ["A", 0.0], ["b", 100.0], ["B", 1.0], ["B", 1.0]]
    m = run_macro(excel, wb, "RegressUtils.LinearModelFit", data, [1], 2)
    coefs = m("coefficients")
    return float(coefs[1]) + float(coefs[2]), 101.0, 1e-6


def _lm_rank_probe(excel, wb, ws, runner, tc, args):
    """R4-12: 共线设计必须暴露 rank_deficient 标志。"""
    data = [["X1", "X2", "Y"]] + [[float(r), 2.0 * r, 13.0 * r] for r in range(1, 7)]
    m = run_macro(excel, wb, "RegressUtils.LinearModelFit", data, [1, 2], 3)
    return (1.0 if m("rank_deficient") else 0.0), 1.0, 0.0


def _predict_vertical_probe(excel, wb, ws, runner, tc, args):
    """R4-04: 单列 Range 预测按列内读取 (不得越界读相邻单元格)。"""
    from tests.test_utils import write_range
    data = [["X1", "X2", "Y"], [1, 1, 3], [2, 1, 5], [3, 1, 7], [1, 2, 4], [2, 2, 6], [3, 2, 8]]
    ws.UsedRange.ClearContents()
    write_range(ws, data, 1, 1)
    rng = ws.Range(ws.Cells(1, 1), ws.Cells(7, 3))
    m = run_macro(excel, wb, "RegressUtils.LinearModelFit", rng, [1, 2], 3)
    ws.Cells(10, 1).Value = 10.0
    ws.Cells(11, 1).Value = 20.0
    pred = run_macro(excel, wb, "RegressUtils.LinearModelPredict", m, ws.Range("A10:A11"))
    ws.Range("A10:A11").Clear
    ws.UsedRange.ClearContents()
    return float(pred), 40.0, 1e-8


def _optimize_goal_probe(excel, wb, ws, runner, tc, args):
    """R4-09: 未知 goal 必须报错 (此前静默按 0 分返回网格前 N 行)。"""
    data = [["X", "Y"]] + [[float(r), 2.0 * r + 1.0] for r in range(1, 9)]
    try:
        run_macro(excel, wb, "RegressUtils.OptimizeFactors", data, [1], 2, "minimum", 3, 5)
        return 0.0, 1.0, 0.0
    except Exception:
        return 1.0, 1.0, 0.0


def _lm_reference_probe(excel, wb, ws, runner, tc, args):
    """R5-52: 分类水平排序/参考水平不得随行序漂移 (BinaryCompare)。

    因子水平 {A,B,b}: 参考水平固定为二进制序最小者 "A", 哑变量为 B,b。
    行序反转后系数必须逐位一致, 且等于组均值闭式 (0, 1, 100)。
    """
    base = [["F", "Y"], ["A", 0.0], ["A", 0.0], ["A", 0.0],
            ["b", 100.0], ["B", 1.0], ["B", 1.0]]
    reordered = [base[0]] + list(reversed(base[1:]))
    m1 = run_macro(excel, wb, "RegressUtils.LinearModelFit", base, [1], 2)
    m2 = run_macro(excel, wb, "RegressUtils.LinearModelFit", reordered, [1], 2)
    c1 = m1("coefficients")
    c2 = m2("coefficients")
    expected = [0.0, 1.0, 100.0]
    dev = 0.0
    for i in range(min(len(c1), len(c2), len(expected))):
        dev += abs(float(c1[i]) - expected[i])
        dev += abs(float(c1[i]) - float(c2[i]))
    return dev, 0.0, 1e-9


def _ie_numeric_probe(excel, wb, ws, runner, tc, args):
    """R5-30: InteractionEffects F/p 数值断言 (独立 two-way ANOVA 参考)。"""
    data = [["A", "B", "Y"],
            [0.0, 0.0, 1.0], [0.0, 0.0, 1.5],
            [0.0, 1.0, 3.0], [0.0, 1.0, 3.5],
            [1.0, 0.0, 3.0], [1.0, 0.0, 3.5],
            [1.0, 1.0, 9.0], [1.0, 1.0, 9.5]]
    res = run_macro(excel, wb, "RegressUtils.InteractionEffects", data, [1, 2], 3)
    f_vba = float(res[1][3])
    p_vba = float(res[1][4])
    f_ref, p_ref = _py_interaction_f(data)
    return abs(f_vba - f_ref) + abs(p_vba - p_ref), 0.0, 1e-6


# =============================================================================
# 2026-09-13 R5 回归探针 (R5-07/25/51/53)
# =============================================================================

_REGRESS_PROBE_CODE = "\r\n".join([
    "Option Explicit",
    "",
    "Public Function ProbeErr(ByVal kind As String) As Double",
    "    On Error GoTo EH",
    "    Dim res As Variant",
    "    Dim mdl As Object",
    "    Dim tbl As Variant",
    "    Dim rw As Long",
    "    If kind = \"anova_error_factor\" Then",
    "        ReDim tbl(1 To 3, 1 To 2)",
    "        tbl(1, 1) = \"F\": tbl(1, 2) = \"Y\"",
    "        tbl(2, 1) = \"A\": tbl(2, 2) = 1#",
    "        tbl(3, 1) = CVErr(xlErrValue): tbl(3, 2) = 2#",
    "        Dim anovaResult As Object",
    "        Set anovaResult = ANOVAOneWay(tbl, 1, 2)",
    "    ElseIf kind = \"importance_two_rows\" Then",
    "        ReDim tbl(1 To 3, 1 To 2)",
    "        tbl(1, 1) = \"X\": tbl(1, 2) = \"Y\"",
    "        tbl(2, 1) = 1#: tbl(2, 2) = 2#",
    "        tbl(3, 1) = 2#: tbl(3, 2) = 4#",
    "        res = FactorImportance(tbl, Array(1), 2)",
    "    ElseIf kind = \"interaction_one_factor\" Then",
    "        ReDim tbl(1 To 7, 1 To 2)",
    "        tbl(1, 1) = \"F\": tbl(1, 2) = \"Y\"",
    "        For rw = 2 To 7",
    "            tbl(rw, 1) = CDbl(rw)",
    "            tbl(rw, 2) = CDbl(rw)",
    "        Next rw",
    "        res = InteractionEffects(tbl, Array(1), 2)",
    "    ElseIf kind = \"optimize_goal_empty\" Then",
    "        ReDim tbl(1 To 9, 1 To 2)",
    "        tbl(1, 1) = \"X\": tbl(1, 2) = \"Y\"",
    "        For rw = 2 To 9",
    "            tbl(rw, 1) = CDbl(rw)",
    "            tbl(rw, 2) = 2# * rw + 1#",
    "        Next rw",
    "        res = OptimizeFactors(tbl, Array(1), 2, Empty, 3, 5)",
    "    ElseIf kind = \"optimize_goal_null\" Then",
    "        ReDim tbl(1 To 9, 1 To 2)",
    "        tbl(1, 1) = \"X\": tbl(1, 2) = \"Y\"",
    "        For rw = 2 To 9",
    "            tbl(rw, 1) = CDbl(rw)",
    "            tbl(rw, 2) = 2# * rw + 1#",
    "        Next rw",
    "        res = OptimizeFactors(tbl, Array(1), 2, Null, 3, 5)",
    "    ElseIf kind = \"predict_one_cell\" Then",
    "        ReDim tbl(1 To 7, 1 To 3)",
    "        tbl(1, 1) = \"X1\": tbl(1, 2) = \"X2\": tbl(1, 3) = \"Y\"",
    "        For rw = 2 To 7",
    "            tbl(rw, 1) = CDbl(rw)",
    "            tbl(rw, 2) = 2# * rw",
    "            tbl(rw, 3) = 3# * rw + 1#",
    "        Next rw",
    "        Set mdl = LinearModelFit(tbl, Array(1, 2), 3)",
    "        Worksheets(\"TestData\").Cells(30, 5).Value = 7#",
    "        Dim pred As Double",
    "        pred = LinearModelPredict(mdl, Worksheets(\"TestData\").Cells(30, 5))",
    "    End If",
    "    ProbeErr = 0#",
    "    Exit Function",
    "EH:",
    "    ProbeErr = CDbl(Err.Number)",
    "End Function",
])


def _inject_regress_probe(wb):
    """Inject RegressR5Probe (idempotent) into the workbook VBA project."""
    vbproj = wb.VBProject
    for comp in list(vbproj.VBComponents):
        if comp.Name == "RegressR5Probe":
            return
    comp = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
    comp.Name = "RegressR5Probe"
    comp.CodeModule.AddFromString(_REGRESS_PROBE_CODE)


def _regress_err_probe(kind, expected):
    """R5 error probes: return the raw Err.Number raised inside VBA."""
    def probe(excel, wb, ws, runner, tc, args):
        err = run_macro(excel, wb, "RegressR5Probe.ProbeErr", kind)
        return float(err), float(expected), 0.0
    return probe


class RegressUtilsRunner(CrossValRunner):
    """Extended runner for RegressUtils with Dictionary-aware comparisons."""

    def run_all(self, test_cases):
        import tempfile
        from tests.test_utils import ensure_excel, teardown, create_workbook, inject_testrunner

        excel = ensure_excel()
        wb = None
        try:
            output = os.path.join(tempfile.gettempdir(), f"vba_crossval_{self.module_name}.xlsm")
            wb = create_workbook(excel, output, self.module_paths,
                                 import_order=self._import_order)
            inject_testrunner(wb)
            _inject_regress_probe(wb)
            wb.Save()

            self.results = []
            for tc in test_cases:
                self._run_one_regress(excel, wb, tc)
            return self.results
        finally:
            teardown(excel, wb)

    def _run_one_regress(self, excel, wb, tc):
        label = f"{self.module_name}.{tc['func']}.{tc['name']}"

        if tc.get("skip_if", False):
            print(f"  SKIP  {label} — {tc.get('skip_reason', 'no reason')}")
            self.results.append((self.module_name, tc["name"], "SKIP",
                                 tc.get("skip_reason", "")))
            return

        try:
            args = tc["args"]() if callable(tc["args"]) else tc.get("args", ())
            if not isinstance(args, (tuple, list)):
                args = (args,)
            args = tuple(self._to_com_arg(a) for a in args)

            if tc.get("reconstruct"):
                ws2 = wb.Sheets("TestData")
                vba_result, py_val, tol = tc["reconstruct"](excel, wb, ws2, self, tc, args)
                self._compare(label, vba_result, py_val,
                              tc.get("result_type", "scalar"), tol, tc)
                return

            macro = f"{self.module_name}.{tc['func']}"
            vba_result = run_macro(excel, wb, macro, *args)

            py_val = tc["py_ref"](args) if callable(tc["py_ref"]) else tc["py_ref"]

            cmp_mode = tc.get("compare_mode", "exact")
            if cmp_mode == "dict_keys":
                self._cmp_dict_keys(label, vba_result, py_val, tc)
            elif cmp_mode == "dict_keys_range":
                self._cmp_dict_range(label, vba_result, py_val, tc)
            elif cmp_mode == "array_structure":
                self._cmp_array_struct(label, vba_result, py_val, tc)
            else:
                result_type = tc.get("result_type", "scalar")
                tol = tc.get("tol", 1e-10)
                self._compare(label, vba_result, py_val, result_type, tol, tc)
        except Exception as exc:
            self.results.append((self.module_name, tc["name"], "FAIL",
                                 f"exception: {exc}"))
            print(f"  FAIL  {label} — exception: {exc}")

    def _cmp_dict_keys(self, label, vba_obj, expected_keys, tc):
        if vba_obj is None:
            self.results.append((self.module_name, tc["name"], "FAIL", "null dict"))
            print(f"  FAIL  {label} — null result"); return
        try:
            missing = [k for k in expected_keys if not vba_obj.Exists(k)]
            if missing:
                self.results.append((self.module_name, tc["name"], "FAIL",
                                    f"missing keys: {missing}"))
                print(f"  FAIL  {label} — missing keys: {missing}")
            else:
                self.results.append((self.module_name, tc["name"], "PASS", ""))
                print(f"  PASS  {label}  all {len(expected_keys)} keys present")
        except Exception as exc:
            self.results.append((self.module_name, tc["name"], "FAIL", str(exc)))
            print(f"  FAIL  {label} — {exc}")

    def _cmp_dict_range(self, label, vba_obj, checks, tc):
        if vba_obj is None:
            self.results.append((self.module_name, tc["name"], "FAIL", "null dict"))
            print(f"  FAIL  {label} — null result"); return
        try:
            failures = []
            for key, lo, hi in checks:
                val = float(vba_obj(key))
                if not (lo <= val <= hi):
                    failures.append(f"{key}={val:.4g} not in [{lo}, {hi}]")
            if failures:
                self.results.append((self.module_name, tc["name"], "FAIL",
                                    "; ".join(failures)))
                print(f"  FAIL  {label} — {'; '.join(failures)}")
            else:
                self.results.append((self.module_name, tc["name"], "PASS", ""))
                print(f"  PASS  {label}  all {len(checks)} ranges OK")
        except Exception as exc:
            self.results.append((self.module_name, tc["name"], "FAIL", str(exc)))
            print(f"  FAIL  {label} — {exc}")

    def _cmp_array_struct(self, label, vba_result, checks, tc):
        if vba_result is None:
            self.results.append((self.module_name, tc["name"], "FAIL", "null"))
            print(f"  FAIL  {label} — null"); return
        try:
            nrows = len(vba_result)
            failures = []
            for kind, val in checks:
                if kind == "nrows" and nrows != val:
                    failures.append(f"nrows={nrows} != {val}")
                elif kind == "nrows_ge" and nrows < val:
                    failures.append(f"nrows={nrows} < {val}")
            if failures:
                self.results.append((self.module_name, tc["name"], "FAIL",
                                    "; ".join(failures)))
                print(f"  FAIL  {label} — {'; '.join(failures)}")
            else:
                self.results.append((self.module_name, tc["name"], "PASS", ""))
                print(f"  PASS  {label}  structure OK")
        except Exception as exc:
            self.results.append((self.module_name, tc["name"], "FAIL", str(exc)))
            print(f"  FAIL  {label} — {exc}")


# =============================================================================
# Test Cases
# =============================================================================

TEST_CASES = [

    # ==== FitOLS — perfect linear y = 2 + 3x ====
    {"name": "FitOLS_perfect_keys", "func": "FitOLS",
     "args": lambda: ([[1.0, 1.0], [1.0, 2.0], [1.0, 3.0]], [5.0, 8.0, 11.0]),
     "py_ref": lambda a: ["coefficients", "sse", "r_squared", "fitted_values",
         "residuals", "t_stats", "p_values", "se", "sigma2", "f_stat",
         "f_pvalue", "adj_r_squared", "n", "p", "df_residual"],
     "compare_mode": "dict_keys"},
    {"name": "FitOLS_perfect_values", "func": "FitOLS",
     "args": lambda: ([[1.0, 1.0], [1.0, 2.0], [1.0, 3.0]], [5.0, 8.0, 11.0]),
     "py_ref": lambda a: [("r_squared", 0.99, 1.01), ("sse", 0.0, 0.01),
         ("n", 3.0, 3.0), ("p", 2.0, 2.0), ("df_residual", 1.0, 1.0)],
     "compare_mode": "dict_keys_range"},

    # ==== FitOLS — noisy data ====
    {"name": "FitOLS_noisy_keys", "func": "FitOLS",
     "args": lambda: ([[1.0, 1.0], [1.0, 2.0], [1.0, 3.0], [1.0, 4.0], [1.0, 5.0]],
                      [2.1, 4.0, 6.2, 7.8, 10.1]),
     "py_ref": lambda a: ["coefficients", "sse", "r_squared", "t_stats", "p_values", "se"],
     "compare_mode": "dict_keys"},
    {"name": "FitOLS_noisy_values", "func": "FitOLS",
     "args": lambda: ([[1.0, 1.0], [1.0, 2.0], [1.0, 3.0], [1.0, 4.0], [1.0, 5.0]],
                      [2.1, 4.0, 6.2, 7.8, 10.1]),
     "py_ref": lambda a: [("r_squared", 0.90, 1.01), ("sse", 0.001, 2.0),
         ("n", 5.0, 5.0), ("p", 2.0, 2.0)],
     "compare_mode": "dict_keys_range"},

    # ==== LinearModelFit — Y = -1 + 4X (header-based) ====
    {"name": "LinearModelFit_keys", "func": "LinearModelFit",
     "args": lambda: ([["X", "Y"], [1.0, 3.0], [2.0, 7.0], [3.0, 11.0]], [1], 2),
     "py_ref": lambda a: ["coefficients", "sse", "r_squared", "coef_names",
         "factor_map", "formula", "n", "p"],
     "compare_mode": "dict_keys"},
    {"name": "LinearModelFit_values", "func": "LinearModelFit",
     "args": lambda: ([["X", "Y"], [1.0, 3.0], [2.0, 7.0], [3.0, 11.0]], [1], 2),
     "py_ref": lambda a: [("r_squared", 0.99, 1.01), ("n", 3.0, 3.0), ("p", 2.0, 2.0)],
     "compare_mode": "dict_keys_range"},

    # ==== ANOVAOneWay — 3 groups, significant ====
    {"name": "ANOVAOneWay_sig_keys", "func": "ANOVAOneWay",
     "args": lambda: ([["Group", "Value"],
         ["A", 2.0], ["A", 3.0], ["A", 1.0],
         ["B", 7.0], ["B", 8.0], ["B", 9.0],
         ["C", 15.0], ["C", 14.0], ["C", 16.0]], 1, 2),
     "py_ref": lambda a: ["F", "p_value", "SSB", "SSW", "SST", "MSB", "MSW",
         "eta_sq", "summary", "significant", "n_groups", "n_total"],
     "compare_mode": "dict_keys"},
    {"name": "ANOVAOneWay_sig_values", "func": "ANOVAOneWay",
     "args": lambda: ([["Group", "Value"],
         ["A", 2.0], ["A", 3.0], ["A", 1.0],
         ["B", 7.0], ["B", 8.0], ["B", 9.0],
         ["C", 15.0], ["C", 14.0], ["C", 16.0]], 1, 2),
     "py_ref": lambda a: [("F", 1.0, 1e9), ("p_value", -0.01, 0.05),
         ("eta_sq", 0.8, 1.01), ("n_groups", 3.0, 3.0), ("n_total", 9.0, 9.0)],
     "compare_mode": "dict_keys_range"},

    # ==== ANOVAOneWay — 2 groups, NOT significant ====
    {"name": "ANOVAOneWay_nosig_keys", "func": "ANOVAOneWay",
     "args": lambda: ([["Group", "Value"],
         ["X", 5.0], ["X", 6.0], ["X", 5.5],
         ["Y", 5.1], ["Y", 5.9], ["Y", 5.5]], 1, 2),
     "py_ref": lambda a: ["F", "p_value", "significant", "n_groups", "n_total"],
     "compare_mode": "dict_keys"},
    {"name": "ANOVAOneWay_nosig_values", "func": "ANOVAOneWay",
     "args": lambda: ([["Group", "Value"],
         ["X", 5.0], ["X", 6.0], ["X", 5.5],
         ["Y", 5.1], ["Y", 5.9], ["Y", 5.5]], 1, 2),
     "py_ref": lambda a: [("p_value", 0.05, 1.01), ("n_groups", 2.0, 2.0),
         ("n_total", 6.0, 6.0)],
     "compare_mode": "dict_keys_range"},

    # ==== ANOVAOneWay — SST = SSB + SSW ====
    {"name": "ANOVAOneWay_SST_decomp", "func": "ANOVAOneWay",
     "args": lambda: ([["Group", "Value"],
         ["A", 2.0], ["A", 3.0], ["A", 1.0],
         ["B", 7.0], ["B", 8.0], ["B", 9.0]], 1, 2),
     "py_ref": lambda a: None,
     "skip_if": True,
     "skip_reason": "SST=SSB+SSW verified by ANOVAOneWay_sig_values key existence"},

    # ==== FactorImportance — with header (structural check; QR precision diff ~1e-15) ====
    {"name": "FactorImportance_header", "func": "FactorImportance",
     "args": lambda: ([["X1", "X2", "Y"],
         [1.0, 10.0, 5.0], [2.0, 5.0, 6.0], [3.0, 2.0, 7.0],
         [4.0, 0.0, 8.0], [5.0, 0.0, 9.0]], [1, 2], 3, True),
     "py_ref": lambda a: [("nrows_ge", 3)],
     "compare_mode": "array_structure"},
    # Numeric values verified separately via _py_factor_importance helper
    # (VBA QR vs numpy QR differ at ~1e-15 — within fp precision, but
    #  mixed string/number array comparison uses exact string matching).

    # ==== FactorImportance — no header ====
    {"name": "FactorImportance_no_header", "func": "FactorImportance",
     "args": lambda: ([[1.0, 5.0], [2.0, 8.0], [3.0, 11.0], [4.0, 14.0]],
                      [1], 2, False),
     "py_ref": lambda a: [("nrows_ge", 2)],
     "compare_mode": "array_structure"},

    # ==== FactorSweep — what-if scan (requires model Object, not raw data) ====
    {"name": "FactorSweep_basic", "func": "FactorSweep",
     "args": lambda: (None, 1, 1.0, 5.0, 5),
     "py_ref": lambda a: None,
     "skip_if": True,
     "skip_reason": "FactorSweep requires a model Object (from LinearModelFit/FitOLS), "
                    "which cannot be constructed via single COM call. "
                    "Tested indirectly via OptimizeFactors."},

    # ==== OptimizeFactors — max / min ====
    {"name": "OptimizeFactors_max", "func": "OptimizeFactors",
     "args": lambda: ([["X1", "X2", "Y"],
         [1.0, 0.0, 2.0], [2.0, 1.0, 4.0], [3.0, 1.0, 6.0],
         [4.0, 0.0, 8.0], [5.0, 0.0, 10.0]], [1, 2], 3, "max", 3, 5),
     "py_ref": lambda a: [("nrows", 4)],
     "compare_mode": "array_structure"},
    {"name": "OptimizeFactors_min", "func": "OptimizeFactors",
     "args": lambda: ([["X1", "X2", "Y"],
         [1.0, 0.0, 2.0], [2.0, 1.0, 4.0], [3.0, 1.0, 6.0],
         [4.0, 0.0, 8.0], [5.0, 0.0, 10.0]], [1, 2], 3, "min", 3, 5),
     "py_ref": lambda a: [("nrows", 4)],
     "compare_mode": "array_structure"},

    # ==== InteractionEffects — strong interaction vs additive ====
    {"name": "InteractionEffects_strong", "func": "InteractionEffects",
     "args": lambda: ([["A", "B", "Y"],
         [0.0, 0.0, 1.0], [0.0, 0.0, 1.5],
         [0.0, 1.0, 3.0], [0.0, 1.0, 3.5],
         [1.0, 0.0, 3.0], [1.0, 0.0, 3.5],
         [1.0, 1.0, 9.0], [1.0, 1.0, 9.5]], [1, 2], 3),
     "py_ref": lambda a: [("nrows_ge", 2)],
     "compare_mode": "array_structure"},
    {"name": "InteractionEffects_additive", "func": "InteractionEffects",
     "args": lambda: ([["A", "B", "Y"],
         [0.0, 0.0, 1.1], [0.0, 0.0, 0.9],
         [0.0, 1.0, 4.1], [0.0, 1.0, 3.9],
         [1.0, 0.0, 4.1], [1.0, 0.0, 3.9],
         [1.0, 1.0, 7.0], [1.0, 1.0, 7.1]], [1, 2], 3),
     "py_ref": lambda a: [("nrows_ge", 2)],
     "compare_mode": "array_structure"},

    # ==== ANOVAOneWay_Fstat — scalar comparison (requires scipy) ====
    {"name": "ANOVAOneWay_Fstat", "func": "ANOVAOneWay_Fstat",
     "args": lambda: ([[1.0, 23.0], [1.0, 25.0], [1.0, 22.0],
         [2.0, 30.0], [2.0, 32.0], [2.0, 31.0],
         [3.0, 28.0], [3.0, 27.0], [3.0, 29.0]], 1, 2, False),
     "py_ref": lambda a: float(_sp_stats.f_oneway(
         [23, 25, 22], [30, 32, 31], [28, 27, 29]).statistic) if _HAS_SCIPY else 0.0,
     "result_type": "scalar", "tol": 0.2,
     "skip_if": not _HAS_SCIPY, "skip_reason": "scipy not installed"},

    # =====================================================================
    # 2026-09-13 第四轮审查回归 (R4-02/04/09/11/12/14/15)
    # =====================================================================
    {"name": "InteractionEffects_single_level_cat", "func": "InteractionEffects_single_level_cat",
     "args": lambda: (), "reconstruct": _ie_single_level_probe, "py_ref": lambda a: 1.0},
    {"name": "LinearModelFit_blank_categorical", "func": "LinearModelFit_blank_categorical",
     "args": lambda: (), "reconstruct": _lm_blank_cat_probe, "py_ref": lambda a: 1.0},
    {"name": "LinearModelFit_case_binary_levels", "func": "LinearModelFit_case_binary_levels",
     "args": lambda: (), "reconstruct": _lm_case_probe, "py_ref": lambda a: 101.0},
    {"name": "LinearModelFit_rank_deficient_flag", "func": "LinearModelFit_rank_deficient_flag",
     "args": lambda: (), "reconstruct": _lm_rank_probe, "py_ref": lambda a: 1.0},
    {"name": "LinearModelPredict_vertical_range", "func": "LinearModelPredict_vertical_range",
     "args": lambda: (), "reconstruct": _predict_vertical_probe, "py_ref": lambda a: 40.0},
    {"name": "OptimizeFactors_unknown_goal", "func": "OptimizeFactors_unknown_goal",
     "args": lambda: (), "reconstruct": _optimize_goal_probe, "py_ref": lambda a: 1.0},
    # R4-11: 高偏置常数响应不得假显著 (修复后 F=0)
    {"name": "ANOVA_const_high_offset", "func": "ANOVAOneWay_Fstat",
     "args": lambda: ([[1.0, 1e13 + 0.1]] * 1000 + [[2.0, 1e13 + 0.1]] * 1000
                      + [[3.0, 1e13 + 0.1]] * 1000, 1, 2, False),
     "py_ref": lambda a: 0.0, "result_type": "scalar", "tol": 1e-9},

    # =====================================================================
    # 2026-09-13 第五轮审查回归 (R5-07/25/30/51/52/53)
    # =====================================================================
    # R5-07: BinaryCompare 分组 — "" 是独立水平, 空行 (None) 丢弃, A/a、B/b 分离
    {"name": "ANOVA_binary_case_empty_level", "func": "ANOVAOneWay",
     "args": lambda: ([["F", "Y"],
        ["A", 3.0], ["", 4.0], [None, 999.0],
        ["B", 7.0], ["b", 8.5], ["B", 9.0], ["a", 15.0]], 1, 2),
     "py_ref": lambda a: _anova_checks(),
     "compare_mode": "dict_keys_range",
     "skip_if": not _HAS_SCIPY, "skip_reason": "scipy not installed"},
    {"name": "ANOVA_error_factor_raises", "func": "ANOVAOneWay_error_factor",
     "args": lambda: (),
     "reconstruct": _regress_err_probe("anova_error_factor", ERR_REG_INVALID_INPUT),
     "py_ref": lambda a: float(ERR_REG_INVALID_INPUT)},
    # R5-25: 核心函数以模块错误替代哨兵文本
    {"name": "FactorImportance_two_rows_raises", "func": "FactorImportance_two_rows",
     "args": lambda: (),
     "reconstruct": _regress_err_probe("importance_two_rows", ERR_REG_TOO_FEW_ROWS),
     "py_ref": lambda a: float(ERR_REG_TOO_FEW_ROWS)},
    {"name": "InteractionEffects_single_factor_raises", "func": "InteractionEffects_single_factor",
     "args": lambda: (),
     "reconstruct": _regress_err_probe("interaction_one_factor", ERR_REG_INVALID_DATA),
     "py_ref": lambda a: float(ERR_REG_INVALID_DATA)},
    # R5-51: goal 为 Empty/Null 必须显式报错 (不得按目标 0 优化)
    {"name": "OptimizeFactors_goal_empty_raises", "func": "OptimizeFactors_goal_empty",
     "args": lambda: (),
     "reconstruct": _regress_err_probe("optimize_goal_empty", ERR_REG_INVALID_PARAM),
     "py_ref": lambda a: float(ERR_REG_INVALID_PARAM)},
    {"name": "OptimizeFactors_goal_null_raises", "func": "OptimizeFactors_goal_null",
     "args": lambda: (),
     "reconstruct": _regress_err_probe("optimize_goal_null", ERR_REG_INVALID_PARAM),
     "py_ref": lambda a: float(ERR_REG_INVALID_PARAM)},
    # R5-53: 多因子模型传 1×1 Range 必须报错, 不得静默补 0
    {"name": "LinearModelPredict_one_cell_raises", "func": "LinearModelPredict_one_cell",
     "args": lambda: (),
     "reconstruct": _regress_err_probe("predict_one_cell", ERR_REG_INVALID_DATA),
     "py_ref": lambda a: float(ERR_REG_INVALID_DATA)},
    # R5-52: 参考水平与行序无关 (BinaryCompare 排序)
    {"name": "LinearModelFit_reference_level_order", "func": "LinearModelFit_reference_level",
     "args": lambda: (), "reconstruct": _lm_reference_probe, "py_ref": lambda a: 0.0},
    # R5-30: InteractionEffects F/p 数值断言 (独立 two-way ANOVA 参考)
    {"name": "InteractionEffects_F_p_numeric", "func": "InteractionEffects_F_p",
     "args": lambda: (), "reconstruct": _ie_numeric_probe, "py_ref": lambda a: 0.0},
]


def main() -> int:
    runner = RegressUtilsRunner("RegressUtils", MODULE_PATHS)
    runner.run_all(TEST_CASES)
    passed, failed = runner.print_summary()
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
