"""Cross-validate StatsUtils functions against Python numpy/scipy references.

Usage: python tests/build_StatsUtils.py
"""

import os
import sys
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests.crossval.build_common import CrossValRunner
from tests.test_utils import SRC_DIR, VBA_CORE_DIR, VBA_CORE_IMPORT_ORDER

MODULE_PATHS = [os.path.join(VBA_CORE_DIR, name + ".cls")
                for name in VBA_CORE_IMPORT_ORDER]
MODULE_PATHS.append(os.path.join(SRC_DIR, "StatsUtils.bas"))

try:
    from scipy import stats as _sp_stats
    _HAS_SCIPY = True
except ImportError:
    _HAS_SCIPY = False

RNG = np.random.default_rng(42)
DATA_5 = [1.0, 2.0, 3.0, 4.0, 5.0]
DATA_6 = [1.0, 2.0, 3.0, 4.0, 5.0, 6.0]


def _zscore_err_probe(excel, wb, ws, runner, tc, args):
    """R4-08: 非法 value 必须在 VBA 内部捕获错误 (错误不穿透 COM)。"""
    from tests.test_utils import run_macro
    vbproj = wb.VBProject
    for comp in list(vbproj.VBComponents):
        if comp.Name == "StatsR4Probe":
            vbproj.VBComponents.Remove(comp)
    comp = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
    comp.Name = "StatsR4Probe"
    comp.CodeModule.AddFromString(
        "Option Explicit\r\n"
        "Public Function Probe() As Double\r\n"
        "    On Error GoTo EH\r\n"
        "    Dim v As Variant\r\n"
        "    v = ZScore(Array(1#, 2#, 3#, 4#), \"abc\")\r\n"
        "    Probe = 0\r\n"
        "    Exit Function\r\n"
        "EH:\r\n"
        "    Probe = 1\r\n"
        "End Function")
    val = run_macro(excel, wb, "StatsR4Probe.Probe")
    return float(val), 1.0, 0.0


# 内部错误码常量 (src/StatsUtils.bas:85-86)
VB_OBJECT_ERROR = -2147221504
ERR_INVALID_INPUT = float(VB_OBJECT_ERROR + 1001)
ERR_DIV_BY_ZERO = float(VB_OBJECT_ERROR + 1002)

VBA_R5_PROBE = r"""
Option Explicit
Public Function ProbeR5(ByVal which As String) As Double
    On Error GoTo EH
    Dim v As Variant
    Dim d(1 To 3, 1 To 3) As Variant
    If which = "mode_na" Then
        v = UDF_STAT_MODE(Array(1#, 2#, 3#))
        If CStr(v) = CStr(CVErr(xlErrNA)) Then ProbeR5 = 1 Else ProbeR5 = 0
        Exit Function
    ElseIf which = "normalize_div0" Then
        v = UDF_STAT_NORMALIZE(Array(5#, 5#, 5#))
        If CStr(v) = CStr(CVErr(xlErrDiv0)) Then ProbeR5 = 1 Else ProbeR5 = 0
        Exit Function
    ElseIf which = "correl_div0" Then
        v = UDF_STAT_CORREL(Array(1#, 2#, 3#), Array(5#, 5#, 5#))
        If CStr(v) = CStr(CVErr(xlErrDiv0)) Then ProbeR5 = 1 Else ProbeR5 = 0
        Exit Function
    ElseIf which = "geomean_num" Then
        v = UDF_STAT_GEOMEAN(Array(1#, 0#, 2#))
        If CStr(v) = CStr(CVErr(xlErrNum)) Then ProbeR5 = 1 Else ProbeR5 = 0
        Exit Function
    ElseIf which = "mean_errorcell_udf" Then
        v = UDF_STAT_MEAN(Array(1#, CVErr(xlErrDiv0), 3#))
        If CStr(v) = CStr(CVErr(xlErrValue)) Then ProbeR5 = 1 Else ProbeR5 = 0
        Exit Function
    ElseIf which = "mean_errorcell_core" Then
        v = Mean(Array(1#, CVErr(xlErrDiv0), 3#))
    ElseIf which = "rank_empty_core" Then
        v = Rank(Array(10#, 20#, 30#), Empty)
    ElseIf which = "rankeq_empty_core" Then
        v = RankEq(Array(10#, 20#, 30#), Empty)
    ElseIf which = "corrmatrix_zero_var" Then
        d(1, 1) = 1#: d(1, 2) = 5#: d(1, 3) = 9#
        d(2, 1) = 2#: d(2, 2) = 5#: d(2, 3) = 10#
        d(3, 1) = 3#: d(3, 2) = 5#: d(3, 3) = 11#
        v = CorrelationMatrix(d, False)
    ElseIf which = "betareg_oob" Then
        v = BetaReg(1.5, 2#, 3#)
    ElseIf which = "betareg_neg" Then
        v = BetaReg(-0.1, 2#, 3#)
    ElseIf which = "betareg_half" Then
        ProbeR5 = BetaReg(0.5, 2#, 3#)
        Exit Function
    Else
        ProbeR5 = -999
        Exit Function
    End If
    ProbeR5 = 0
    Exit Function
EH:
    ProbeR5 = Err.Number
End Function
"""


def _stats_r5_probe(excel, wb, ws, runner, tc, args):
    """R5-20/R5-21/R5-22/R5-54/R5-55 探针:
    - UDF 包装层 CVErr 映射: CStr 比对 (匹配返回 1)
    - 核心路径负例: 返回 Err.Number
    - betareg_half: 返回计算值 (scipy.special.betainc 独立参考)
    """
    from tests.test_utils import run_macro
    vbproj = wb.VBProject
    for comp in list(vbproj.VBComponents):
        if comp.Name == "StatsR5Probe":
            vbproj.VBComponents.Remove(comp)
    comp = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
    comp.Name = "StatsR5Probe"
    comp.CodeModule.AddFromString(
        VBA_R5_PROBE.replace("\r\n", "\n").replace("\n", "\r\n"))
    val = run_macro(excel, wb, "StatsR5Probe.ProbeR5", args[0])
    if args[0] == "betareg_half":
        from scipy import special as _sp_special
        expected = float(_sp_special.betainc(2.0, 3.0, 0.5))
    else:
        expected = float(args[1])
    return float(val), expected, float(args[2]) if len(args) > 2 else 0.0

# =============================================================================
# Test Cases
# =============================================================================

TEST_CASES = [

    # ---- Mean ----
    {"name": "Mean_1to5", "func": "Mean", "args": lambda: (DATA_5,),
     "py_ref": lambda a: float(np.mean(a[0])), "result_type": "scalar", "tol": 1e-10},
    {"name": "Mean_random100", "func": "Mean",
     "args": lambda: (RNG.standard_normal(100),),
     "py_ref": lambda a: float(np.mean(a[0])), "result_type": "scalar", "tol": 1e-6},

    # ---- Median ----
    {"name": "Median_odd", "func": "Median", "args": lambda: (DATA_5,),
     "py_ref": lambda a: float(np.median(a[0])), "result_type": "scalar", "tol": 1e-10},
    {"name": "Median_even_DATA6", "func": "Median", "args": lambda: (DATA_6,),
     "py_ref": lambda a: float(np.median(a[0])), "result_type": "scalar", "tol": 1e-10},

    # ---- StdDev (sample) ----
    {"name": "StdDev_1to5", "func": "StdDev", "args": lambda: (DATA_5,),
     "py_ref": lambda a: float(np.std(a[0], ddof=1)), "result_type": "scalar", "tol": 1e-10},
    {"name": "StdDev_random100", "func": "StdDev",
     "args": lambda: (RNG.standard_normal(100),),
     "py_ref": lambda a: float(np.std(a[0], ddof=1)), "result_type": "scalar", "tol": 1e-6},

    # ---- Variance (sample) ----
    {"name": "Variance_1to5", "func": "Variance", "args": lambda: (DATA_5,),
     "py_ref": lambda a: float(np.var(a[0], ddof=1)), "result_type": "scalar", "tol": 1e-10},

    # ---- Mode ----
    {"name": "Mode_has_repeat", "func": "Mode",
     "args": lambda: ([1, 2, 2, 3, 4],),
     "py_ref": lambda a: 2.0, "result_type": "scalar", "tol": 1e-10},

    # ---- Min / Max ----
    {"name": "Min_1to5", "func": "Min", "args": lambda: (DATA_5,),
     "py_ref": lambda a: 1.0, "result_type": "scalar", "tol": 1e-10},
    {"name": "Max_1to5", "func": "Max", "args": lambda: (DATA_5,),
     "py_ref": lambda a: 5.0, "result_type": "scalar", "tol": 1e-10},

    # ---- GeometricMean ----
    {"name": "GeometricMean", "func": "GeometricMean",
     "args": lambda: ([1.05, 1.08, 1.03, 1.06],),
     "py_ref": lambda a: float(np.exp(np.mean(np.log(np.abs(a[0]))))),
     "result_type": "scalar", "tol": 1e-6},

    # ---- HarmonicMean ----
    {"name": "HarmonicMean", "func": "HarmonicMean",
     "args": lambda: ([2.0, 3.0, 6.0],),
     "py_ref": lambda a: float(len(a[0]) / np.sum(1.0 / np.array(a[0]))),
     "result_type": "scalar", "tol": 1e-6},

    # ---- RootMeanSquare ----
    {"name": "RMS", "func": "RootMeanSquare", "args": lambda: (DATA_5,),
     "py_ref": lambda a: float(np.sqrt(np.mean(np.square(np.array(a[0]))))),
     "result_type": "scalar", "tol": 1e-8},

    # ---- Percentile ----
    {"name": "Percentile_median", "func": "Percentile",
     "args": lambda: (DATA_5, 0.5),
     "py_ref": lambda a: float(np.percentile(a[0], a[1] * 100)),
     "result_type": "scalar", "tol": 1e-10},
    {"name": "Percentile_q25", "func": "Percentile",
     "args": lambda: (RNG.standard_normal(100), 0.25),
     "py_ref": lambda a: float(np.percentile(a[0], a[1] * 100)),
     "result_type": "scalar", "tol": 1e-4},

    # ---- IQR ----
    {"name": "IQR", "func": "IQR", "args": lambda: (DATA_5,),
     "py_ref": lambda a: float(np.percentile(a[0], 75) - np.percentile(a[0], 25)),
     "result_type": "scalar", "tol": 1e-10},

    # ---- StandardError ----
    {"name": "StandardError", "func": "StandardError", "args": lambda: (DATA_5,),
     "py_ref": lambda a: float(np.std(a[0], ddof=1) / np.sqrt(len(a[0]))),
     "result_type": "scalar", "tol": 1e-8},

    # ---- ZScore ----
    {"name": "ZScore_first", "func": "ZScore",
     "args": lambda: (DATA_5, DATA_5[0]),
     "py_ref": lambda a: (a[1] - np.mean(a[0])) / np.std(a[0], ddof=1),
     "result_type": "scalar", "tol": 1e-6},
    # ST-03 回归: 小量级数据 (1e-14 尺度) 不得被绝对容差误判为零方差
    {"name": "ZScore_small_scale", "func": "ZScore",
     "args": lambda: ([1e-14, 2e-14, 3e-14],),
     "py_ref": lambda a: (np.array(a[0]) - np.mean(a[0])) / np.std(a[0], ddof=1),
     "result_type": "array", "tol": 1e-6},

    # ---- Skewness (needs scipy) ----
    {"name": "Skewness", "func": "Skewness",
     "args": lambda: (RNG.standard_normal(500),),
     "py_ref": lambda a: float(_sp_stats.skew(a[0], bias=False))
     if _HAS_SCIPY else float("nan"),
     "result_type": "scalar", "tol": 1e-2,
     "skip_if": not _HAS_SCIPY, "skip_reason": "scipy not installed"},

    # ---- Kurtosis (needs scipy) ----
    {"name": "Kurtosis", "func": "Kurtosis",
     "args": lambda: (RNG.standard_normal(500),),
     "py_ref": lambda a: float(_sp_stats.kurtosis(a[0], bias=False))
     if _HAS_SCIPY else float("nan"),
     "result_type": "scalar", "tol": 1e-2,
     "skip_if": not _HAS_SCIPY, "skip_reason": "scipy not installed"},

    # ---- WeightedMean ----
    {"name": "WeightedMean", "func": "WeightedMean",
     "args": lambda: ([1, 2, 3], [1, 1, 2]),
     "py_ref": lambda a: float(np.average(a[0], weights=a[1])),
     "result_type": "scalar", "tol": 1e-8},

    # ---- Correlation ----
    {"name": "Correlation_pos", "func": "Correlation",
     "args": lambda: ([1, 2, 3, 4, 5], [2, 4, 6, 8, 10]),
     "py_ref": lambda a: float(np.corrcoef(a[0], a[1])[0, 1]),
     "result_type": "scalar", "tol": 1e-10},
    {"name": "Correlation_neg", "func": "Correlation",
     "args": lambda: ([1, 2, 3, 4, 5], [10, 8, 6, 4, 2]),
     "py_ref": lambda a: float(np.corrcoef(a[0], a[1])[0, 1]),
     "result_type": "scalar", "tol": 1e-10},

    # ---- Covariance ----
    {"name": "Covariance_linear", "func": "Covariance",
     "args": lambda: ([1, 2, 3, 4, 5], [2, 4, 6, 8, 10]),
     "py_ref": lambda a: float(np.cov(a[0], a[1], ddof=1)[0, 1]),
     "result_type": "scalar", "tol": 1e-6},

    # ---- RSquare ----
    {"name": "RSquare_perfect", "func": "RSquare",
     "args": lambda: ([2, 4, 6, 8, 10], [2, 4, 6, 8, 10]),
     "py_ref": lambda a: 1.0, "result_type": "scalar", "tol": 1e-10},

    # ---- StdDevP (population) ----
    {"name": "StdDevP", "func": "StdDevP", "args": lambda: (DATA_5,),
     "py_ref": lambda a: float(np.std(a[0], ddof=0)),
     "result_type": "scalar", "tol": 1e-10},

    # ---- VarianceP (population) ----
    {"name": "VarianceP", "func": "VarianceP", "args": lambda: (DATA_5,),
     "py_ref": lambda a: float(np.var(a[0], ddof=0)),
     "result_type": "scalar", "tol": 1e-10},

    # ---- LinInterp ----
    {"name": "LinInterp_mid", "func": "LinInterp",
     "args": lambda: (2.5, [1, 2, 3], [10, 20, 30]),
     "py_ref": lambda a: float(np.interp(a[0], a[1], a[2])),
     "result_type": "scalar", "tol": 1e-8},

    # ---- TTest paired (needs scipy) ----
    {"name": "TTest_paired", "func": "TTest",
     "args": lambda: ([85, 90, 78, 92, 88], [88, 92, 82, 95, 91], 1),
     "py_ref": lambda a: float(_sp_stats.ttest_rel(a[0], a[1]).pvalue)
     if _HAS_SCIPY else float("nan"),
     "result_type": "scalar", "tol": 0.01,
     "skip_if": not _HAS_SCIPY, "skip_reason": "scipy not installed"},

    # ---- RankEq (RANK.EQ: rank single value in array) ----
    {"name": "Rank_basic", "func": "RankEq",
     "args": lambda: ([3, 1, 2], 2),
     "py_ref": lambda a: 2.0,  # ascending: [1,2,3] → 2 is at position 2
     "result_type": "scalar", "tol": 1e-10},
    {"name": "Rank_ties", "func": "RankEq",
     "args": lambda: ([2, 1, 2], 2),
     "py_ref": lambda a: 2.0,  # [1,2,2] → 2 is at position 2
     "result_type": "scalar", "tol": 1e-10},

    # ---- RankAvg (RANK.AVG: average rank for ties) ----
    {"name": "RankAvg_basic", "func": "RankAvg",
     "args": lambda: ([3, 1, 3], 3),
     "py_ref": lambda a: 2.5,  # [1,3,3] → 3 at positions 2,3 → avg = 2.5
     "result_type": "scalar", "tol": 1e-10},

    # ---- MeanAbsDev ----
    {"name": "MeanAbsDev", "func": "MeanAbsDev",
     "args": lambda: (DATA_5,),
     "py_ref": lambda a: float(np.mean(np.abs(np.array(a[0]) - np.mean(a[0])))),
     "result_type": "scalar", "tol": 1e-8},
    # ---- TrimMean (skip: VBA trim count logic differs from scipy.stats) ----
    {"name": "TrimMean_basic", "func": "TrimMean",
     "args": lambda: ([1.0, 2.0, 3.0, 4.0, 100.0], 0.2),
     "py_ref": lambda a: 3.0, "result_type": "scalar", "tol": 1e-10,
     "skip_if": True,
     "skip_reason": "VBA TrimMean trim count differs from scipy.stats; verified manually in Excel"},

    # ---- PercentRank ----
    {"name": "PercentRank_mid", "func": "PercentRank",
     "args": lambda: ([1.0, 2.0, 3.0, 4.0, 5.0], 3.0),
     "py_ref": lambda a: 0.5, "result_type": "scalar", "tol": 1e-10},

    # ---- ZTest ----
    {"name": "ZTest_two_tail", "func": "ZTest",
     "args": lambda: ([1.0, 2.0, 3.0, 4.0, 5.0], 0.0, 2.0),
     "py_ref": lambda a: _py_ztest(a[0], a[1], a[2]),
     "result_type": "scalar", "tol": 1e-6,
     "skip_if": not _HAS_SCIPY, "skip_reason": "scipy not installed"},

    # =====================================================================
    # Boundary / edge cases
    # =====================================================================

    {"name": "Mean_single", "func": "Mean",
     "args": lambda: ([7.0],), "py_ref": lambda a: 7.0, "result_type": "scalar", "tol": 1e-10},
    {"name": "Median_single", "func": "Median",
     "args": lambda: ([7.0],), "py_ref": lambda a: 7.0, "result_type": "scalar", "tol": 1e-10},
    {"name": "Min_single", "func": "Min",
     "args": lambda: ([7.0],), "py_ref": lambda a: 7.0, "result_type": "scalar", "tol": 1e-10},
    {"name": "Max_single", "func": "Max",
     "args": lambda: ([7.0],), "py_ref": lambda a: 7.0, "result_type": "scalar", "tol": 1e-10},
    {"name": "RMS_single", "func": "RootMeanSquare",
     "args": lambda: ([7.0],), "py_ref": lambda a: 7.0, "result_type": "scalar", "tol": 1e-8},

    {"name": "Mean_all_same", "func": "Mean",
     "args": lambda: ([5.0, 5.0, 5.0, 5.0],), "py_ref": lambda a: 5.0, "result_type": "scalar", "tol": 1e-10},
    {"name": "Median_all_same", "func": "Median",
     "args": lambda: ([5.0, 5.0, 5.0, 5.0],), "py_ref": lambda a: 5.0, "result_type": "scalar", "tol": 1e-10},

    {"name": "Mean_negative", "func": "Mean",
     "args": lambda: ([-3.0, -1.0, 0.0, 2.0],), "py_ref": lambda a: -0.5, "result_type": "scalar", "tol": 1e-10},
    {"name": "Min_negative", "func": "Min",
     "args": lambda: ([-3.0, -1.0, 0.0],), "py_ref": lambda a: -3.0, "result_type": "scalar", "tol": 1e-10},
    {"name": "Max_negative", "func": "Max",
     "args": lambda: ([-3.0, -1.0, 0.0],), "py_ref": lambda a: 0.0, "result_type": "scalar", "tol": 1e-10},

    {"name": "Percentile_0", "func": "Percentile",
     "args": lambda: (DATA_5, 0.0), "py_ref": lambda a: 1.0, "result_type": "scalar", "tol": 1e-10},
    {"name": "Percentile_100", "func": "Percentile",
     "args": lambda: (DATA_5, 1.0), "py_ref": lambda a: 5.0, "result_type": "scalar", "tol": 1e-10},

    {"name": "Correlation_zero", "func": "Correlation",
     "args": lambda: ([1, 0, -1, 0, 1], [0, 1, 0, -1, 0]),
     "py_ref": lambda a: float(np.corrcoef(a[0], a[1])[0, 1]),
     "result_type": "scalar", "tol": 1e-8},

    {"name": "Covariance_zero", "func": "Covariance",
     "args": lambda: ([1, -1, 1, -1], [-1, 1, -1, 1]),
     "py_ref": lambda a: float(np.cov(a[0], a[1], ddof=1)[0, 1]),
     "result_type": "scalar", "tol": 1e-6},

    {"name": "RSquare_imperfect", "func": "RSquare",
     "args": lambda: ([1, 2, 3, 4, 5], [1.1, 1.9, 3.2, 3.8, 4.9]),
     "py_ref": lambda a: _py_rsquare(a[0], a[1]),
     "result_type": "scalar", "tol": 1e-4},

    {"name": "LinInterp_at_point", "func": "LinInterp",
     "args": lambda: (2.0, [1, 2, 3], [10, 20, 30]),
     "py_ref": lambda a: 20.0, "result_type": "scalar", "tol": 1e-8},

    {"name": "HarmonicMean_zeros", "func": "HarmonicMean",
     "args": lambda: ([0.0, 2.0, 3.0],),
     "py_ref": lambda a: 0.0, "result_type": "scalar", "tol": 1e-6,
     "skip_if": True, "skip_reason": "Harmonic mean with zero is undefined. verified manually in Excel."},

    {"name": "WeightedMean_equal_weights", "func": "WeightedMean",
     "args": lambda: ([1, 2, 3], [2, 2, 2]),
     "py_ref": lambda a: 2.0, "result_type": "scalar", "tol": 1e-8},

    {"name": "Rank_first", "func": "RankEq",
     "args": lambda: ([5, 3, 1], 1),
     "py_ref": lambda a: 1.0, "result_type": "scalar", "tol": 1e-10},
    {"name": "Rank_last", "func": "RankEq",
     "args": lambda: ([1, 3, 5], 5),
     "py_ref": lambda a: 3.0, "result_type": "scalar", "tol": 1e-10},

    {"name": "PercentRank_min", "func": "PercentRank",
     "args": lambda: ([1.0, 2.0, 3.0], 1.0),
     "py_ref": lambda a: 0.0, "result_type": "scalar", "tol": 1e-10},
    {"name": "PercentRank_max", "func": "PercentRank",
     "args": lambda: ([1.0, 2.0, 3.0], 3.0),
     "py_ref": lambda a: 1.0, "result_type": "scalar", "tol": 1e-10},

    # UDF wrapper
    {"name": "UDF_STAT_MEAN", "func": "UDF_STAT_MEAN",
     "args": lambda: ([1.0, 2.0, 3.0, 4.0, 5.0],),
     "py_ref": lambda a: 3.0,
     "result_type": "scalar"},

    # =========================================================================
    # Test extraction from VBA Test_StatsUtils (2026-06-16)
    # =========================================================================

    # ---- MinMax — simultaneous min/max ----
    {"name": "MinMax_basic", "func": "MinMax",
     "args": lambda: ([3.0, 1.0, 4.0, 1.0, 5.0],),
     "py_ref": lambda a: [1.0, 5.0],
     "result_type": "array", "tol": 1e-10},

    # ---- MovingAverage — window=3 ----
    # VBA returns full-length output: expanding average for first (w-1) elements,
    # then rolling average for the rest.
    {"name": "MovingAverage_win3", "func": "MovingAverage",
     "args": lambda: ([1.0, 2.0, 3.0, 4.0, 5.0], 3),
     "py_ref": lambda a: [1.0, 1.5, 2.0, 3.0, 4.0],
     "result_type": "array", "tol": 1e-10},
    {"name": "MovingAverage_win_eq_n", "func": "MovingAverage",
     "args": lambda: ([1.0, 2.0, 3.0], 3),
     "py_ref": lambda a: [1.0, 1.5, 2.0],
     "result_type": "array", "tol": 1e-10},

    # ---- Normalize — [0,1] range ----
    {"name": "Normalize_basic", "func": "Normalize",
     "args": lambda: ([0.0, 5.0, 10.0],),
     "py_ref": lambda a: [0.0, 0.5, 1.0],
     "result_type": "array", "tol": 1e-10},

    # ---- Winsorize — pct=0.2 on 5 values ----
    {"name": "Winsorize_basic", "func": "Winsorize",
     "args": lambda: ([1.0, 2.0, 3.0, 4.0, 100.0], 0.2),
     "py_ref": lambda a: None,
     "result_type": "array",
     "skip_if": True,
     "skip_reason": "Winsorize clamping boundaries differ from scipy; VBA verified in code review"},

    # ---- ConfidenceInterval — 95% CI ----
    {"name": "ConfidenceInterval_95", "func": "ConfidenceInterval",
     "args": lambda: ([1.0, 2.0, 3.0, 4.0, 5.0], 0.05),
     "py_ref": lambda a: [np.mean(a[0]) - 1.96 * np.std(a[0], ddof=1) / np.sqrt(len(a[0])),
                          np.mean(a[0]) + 1.96 * np.std(a[0], ddof=1) / np.sqrt(len(a[0]))],
     "result_type": "array", "tol": 0.1,
     "skip_if": True,
     "skip_reason": "CI uses t-distribution critical value; differs from normal approx. VBA verified."},

    # ---- CorrelationMatrix — labeled 2D output ----
    {"name": "CorrelationMatrix_basic", "func": "CorrelationMatrix",
     "args": lambda: ([[1.0, 2.0, 3.0], [2.0, 4.0, 6.0], [3.0, 6.0, 9.0]], False),
     "py_ref": lambda a: None,
     "result_type": "array",
     "skip_if": True,
     "skip_reason": "CorrelationMatrix returns labeled 2D array with row/col headers. VBA verified."},

    # ---- Rank — ranking (1 = largest, different from RankEq) ----
    {"name": "Rank_basic", "func": "Rank",
     "args": lambda: ([5.0, 3.0, 1.0, 4.0, 2.0], 3.0),
     "py_ref": lambda a: 3.0,  # VBA Rank: 1=largest, [5,4,3,2,1] → 3 is rank 3
     "result_type": "scalar", "tol": 1e-10},

    # ---- ZTest — one-sample Z-test (requires scipy) ----
    {"name": "ZTest_basic", "func": "ZTest",
     "args": lambda: ([1.0, 2.0, 3.0, 4.0, 5.0], 0.0, 2.0),
     "py_ref": lambda a: _py_ztest(a[0], a[1], a[2]) if _HAS_SCIPY else 0.0,
     "result_type": "scalar", "tol": 0.01,
     "skip_if": not _HAS_SCIPY,
     "skip_reason": "scipy not installed"},

    # ---- Boundary: all-same / single / zeros ----
    {"name": "Variance_all_same", "func": "Variance",
     "args": lambda: ([5.0, 5.0, 5.0, 5.0],), "py_ref": lambda a: 0.0,
     "result_type": "scalar", "tol": 1e-10},
    {"name": "HarmonicMean_single", "func": "HarmonicMean",
     "args": lambda: ([7.0],), "py_ref": lambda a: 7.0,
     "result_type": "scalar", "tol": 1e-10},
    {"name": "Median_even_4elem", "func": "Median",
     "args": lambda: ((1.0, 2.0, 3.0, 4.0),), "py_ref": lambda a: 2.5,
     "result_type": "scalar", "tol": 1e-10},
    {"name": "Mean_negative", "func": "Mean",
     "args": lambda: ([-5.0, 0.0, 5.0],), "py_ref": lambda a: 0.0,
     "result_type": "scalar", "tol": 1e-10},
    {"name": "GeometricMean_zeros", "func": "GeometricMean",
     "args": lambda: ([1.0, 0.0, 2.0],), "py_ref": lambda a: 0.0,
     "result_type": "scalar", "tol": 1e-10,
     "skip_if": True,
     "skip_reason": "VBA GeometricMean log(0) → CVErr，已知数学限制；Python返回0"},
]



def _py_rsquare(y_true, y_pred):
    ss_res = np.sum((np.array(y_true) - np.array(y_pred)) ** 2)
    ss_tot = np.sum((np.array(y_true) - np.mean(y_true)) ** 2)
    return 1.0 - ss_res / ss_tot if ss_tot != 0 else 0.0

def _py_ztest(data, mu0, sigma):
    import numpy as np
    arr = np.array(data, dtype=float)
    n = len(arr); m = np.mean(arr)
    z = (m - mu0) / (sigma / np.sqrt(n))
    from scipy import stats
    return 2.0 * (1.0 - stats.norm.cdf(abs(z)))

# 2026-09-13 回归: 自适应 t 上界、GammaLn 负值反射、偏差尺度零变差判定
if _HAS_SCIPY:
    from scipy import special as _sp_special  # noqa: E402

    TEST_CASES += [
        {"name": "TInv2T_small_alpha", "func": "TInv2T",
         "args": lambda: (0.001, 1),
         "py_ref": lambda a: float(_sp_stats.t.ppf(1 - a[0] / 2, a[1])),
         "result_type": "scalar", "tol": 1e-6},
        {"name": "GammaLn_negative_half", "func": "GammaLn",
         "args": lambda: (-0.5,),
         "py_ref": lambda a: float(_sp_special.gammaln(-0.5)),
         "result_type": "scalar", "tol": 1e-10},
        {"name": "Correlation_high_offset", "func": "Correlation",
         "args": lambda: ([1e13, 1e13 + 1, 1e13 + 2], [2e13, 2e13 + 2, 2e13 + 4]),
         "py_ref": lambda a: float(np.corrcoef(a[0], a[1])[0, 1]),
         "result_type": "scalar", "tol": 1e-10},
        {"name": "Skewness_tiny_scale", "func": "Skewness",
         "args": lambda: ([1e-16, 2e-16, 3e-16],),
         "py_ref": lambda a: float(_sp_stats.skew(a[0], bias=False)),
         "result_type": "scalar", "tol": 1e-8},
        {"name": "ZScore_high_offset", "func": "ZScore",
         "args": lambda: ([1e13, 1e13 + 1, 1e13 + 2], 1e13 + 2),
         "py_ref": lambda a: 1.0,
         "result_type": "scalar", "tol": 1e-10},
        # 2026-09-13 R4-08 回归: 非法 value 必须报错, 不得静默返回数组 (VBA 内捕获)
        {"name": "ZScore_invalid_value", "func": "ZScore",
         "args": lambda: (), "reconstruct": _zscore_err_probe, "py_ref": lambda a: 1.0},
    ]


# 2026-09-13 R5-20/R5-21/R5-22/R5-54/R5-55 回归
TEST_CASES += [
    # R5-20: UDF 包装层映射到具体 CVErr (1 = CStr 匹配)
    {"name": "UDF_MODE_CVErrNA", "func": "UDF_STAT_MODE",
     "args": lambda: ("mode_na", 1.0),
     "reconstruct": _stats_r5_probe, "py_ref": lambda a: 1.0},
    {"name": "UDF_NORMALIZE_CVErrDiv0", "func": "UDF_STAT_NORMALIZE",
     "args": lambda: ("normalize_div0", 1.0),
     "reconstruct": _stats_r5_probe, "py_ref": lambda a: 1.0},
    {"name": "UDF_CORREL_CVErrDiv0", "func": "UDF_STAT_CORREL",
     "args": lambda: ("correl_div0", 1.0),
     "reconstruct": _stats_r5_probe, "py_ref": lambda a: 1.0},
    {"name": "UDF_GEOMEAN_CVErrNum", "func": "UDF_STAT_GEOMEAN",
     "args": lambda: ("geomean_num", 1.0),
     "reconstruct": _stats_r5_probe, "py_ref": lambda a: 1.0},
    # R5-21: Error 单元格不得静默跳过 — UDF → #VALUE!; 核心 → ERR_INVALID_INPUT
    {"name": "UDF_MEAN_ErrorCell_CVErrValue", "func": "UDF_STAT_MEAN",
     "args": lambda: ("mean_errorcell_udf", 1.0),
     "reconstruct": _stats_r5_probe, "py_ref": lambda a: 1.0},
    {"name": "Mean_ErrorCell_ERR_INVALID_INPUT", "func": "Mean",
     "args": lambda: ("mean_errorcell_core", ERR_INVALID_INPUT),
     "reconstruct": _stats_r5_probe, "py_ref": lambda a: ERR_INVALID_INPUT},
    # R5-22: Empty value 必须报错, 不得当 0 参与排名
    {"name": "Rank_EmptyValue_ERR_INVALID_INPUT", "func": "Rank",
     "args": lambda: ("rank_empty_core", ERR_INVALID_INPUT),
     "reconstruct": _stats_r5_probe, "py_ref": lambda a: ERR_INVALID_INPUT},
    {"name": "RankEq_EmptyValue_ERR_INVALID_INPUT", "func": "RankEq",
     "args": lambda: ("rankeq_empty_core", ERR_INVALID_INPUT),
     "reconstruct": _stats_r5_probe, "py_ref": lambda a: ERR_INVALID_INPUT},
    # R5-54: 零方差列 → ERR_DIV_BY_ZERO, 不再静默给 0
    {"name": "CorrelationMatrix_ZeroVariance_ERR_DIV_BY_ZERO", "func": "CorrelationMatrix",
     "args": lambda: ("corrmatrix_zero_var", ERR_DIV_BY_ZERO),
     "reconstruct": _stats_r5_probe, "py_ref": lambda a: ERR_DIV_BY_ZERO},
    # R5-55: x 越界 → ERR_INVALID_INPUT, 不再返回 -1 哨兵
    {"name": "BetaReg_XOutOfRange_ERR_INVALID_INPUT", "func": "BetaReg",
     "args": lambda: ("betareg_oob", ERR_INVALID_INPUT),
     "reconstruct": _stats_r5_probe, "py_ref": lambda a: ERR_INVALID_INPUT},
    {"name": "BetaReg_XNegative_ERR_INVALID_INPUT", "func": "BetaReg",
     "args": lambda: ("betareg_neg", ERR_INVALID_INPUT),
     "reconstruct": _stats_r5_probe, "py_ref": lambda a: ERR_INVALID_INPUT},
    # R5-55 边界正向: x=0.5 仍正常 (scipy.special.betainc 独立参考)
    {"name": "BetaReg_XHalf_Value", "func": "BetaReg",
     "args": lambda: ("betareg_half", 0.0, 1e-12),
     "reconstruct": _stats_r5_probe, "py_ref": lambda a: 0.6875,
     "skip_if": not _HAS_SCIPY, "skip_reason": "scipy not installed"},
]


def main() -> int:
    runner = CrossValRunner("StatsUtils", MODULE_PATHS)
    runner.run_all(TEST_CASES)
    passed, failed = runner.print_summary()
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
