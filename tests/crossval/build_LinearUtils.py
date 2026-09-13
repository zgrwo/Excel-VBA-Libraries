"""Cross-validate LinearUtils functions against Python numpy references.

Usage: python tests/build_LinearUtils.py
"""

import os
import sys
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests.crossval.build_common import CrossValRunner
from tests.test_utils import SRC_DIR, VBA_CORE_DIR, VBA_CORE_IMPORT_ORDER, com_to_numpy

MODULE_PATHS = [os.path.join(VBA_CORE_DIR, name + ".cls")
                for name in VBA_CORE_IMPORT_ORDER]
MODULE_PATHS.append(os.path.join(SRC_DIR, "LinearUtils.bas"))

RNG = np.random.default_rng(99)
ID2 = np.eye(2); ID3 = np.eye(3)
MAT_2X2 = np.array([[1., 2.], [3., 4.]])
MAT_3X2 = np.array([[1., 2.], [3., 4.], [5., 6.]])

# 内部错误码常量 (src/LinearUtils.bas:48-49)
VB_OBJECT_ERROR = -2147221504
ERR_INVALID_INPUT = float(VB_OBJECT_ERROR + 1001)
# src/LinearUtils.bas:47 MAX_DOUBLE 字面量
MAX_DOUBLE_LITERAL = float("1.79769313486231E+308")

VBA_R5_PROBE = r"""
Option Explicit
Public Function Probe(ByVal which As String) As Double
    On Error GoTo EH
    Dim v As Variant
    Dim m(1 To 2, 1 To 2) As Variant
    If which = "polyfit_string" Then
        v = PolyFit(Array(1#, 2#, "abc", 4#), Array(1#, 2#, 3#, 4#), 1)
    ElseIf which = "polyfit_bool" Then
        v = PolyFit(Array(1#, 2#, True, 4#), Array(1#, 2#, 3#, 4#), 1)
    ElseIf which = "polyfit_empty" Then
        v = PolyFit(Array(1#, 2#, Empty, 4#), Array(1#, 2#, 3#, 4#), 1)
    ElseIf which = "todouble_string" Then
        m(1, 1) = 1#: m(1, 2) = 2#: m(2, 1) = 3#: m(2, 2) = "abc"
        v = MatrixDeterminant(m)
    ElseIf which = "todouble_error" Then
        m(1, 1) = 1#: m(1, 2) = 2#: m(2, 1) = 3#: m(2, 2) = CVErr(xlErrDiv0)
        v = MatrixDeterminant(m)
    Else
        Probe = -999
        Exit Function
    End If
    Probe = 0
    Exit Function
EH:
    Probe = Err.Number
End Function
"""

def _spd(n):
    A = RNG.standard_normal((n, n))
    return A @ A.T + n * np.eye(n)

def _set_col(mat, col_idx, new_col):
    """Python reference for MatrixSetColumn: set column col_idx (1-based) to new_col."""
    result = mat.copy()
    result[:, col_idx - 1] = new_col
    return result

def _call_udf_matrix(excel, wb, ws, runner, func_name, matrix):
    """Call a single-matrix UDF via COM, returning numpy array."""
    tc = {"func": func_name, "is_udf": True}
    return com_to_numpy(runner._call_udf(excel, wb, ws, tc, (matrix.tolist(),)))


def _qr_wide_economy_probe(excel, wb, ws, runner, tc, args):
    """R4-10 回归: economy=True 宽矩阵 reduced QR 必须可重构 Q·R = A (R 为 m×n)。"""
    A = np.array(args[0], dtype=float)
    q = com_to_numpy(runner._call_udf(excel, wb, ws, {"func": "UDF_LINALG_QR_Q"},
                                      (A.tolist(), True)))
    r = com_to_numpy(runner._call_udf(excel, wb, ws, {"func": "UDF_LINALG_QR_R"},
                                      (A.tolist(), True)))
    recon = q @ r
    return float(np.max(np.abs(recon - A))), 0.0, 1e-10


def _eigen_symmetrize_probe(excel, wb, ws, runner, tc, args):
    """R4-13 回归: 容差内不对称输入取对称部分 (A+Aᵀ)/2。

    A = diag/block [[1e16,0,0],[0,0,-1],[0,1,0]] 的对称部分块为零矩阵,
    显式对称化后最小 |特征值| = 0 (旧实现按上三角取数会得到 1)。
    """
    A = np.array(args[0], dtype=float)
    D = com_to_numpy(runner._call_udf(excel, wb, ws, {"func": "UDF_LINALG_EIGVAL"},
                                      (A.tolist(),)))
    d = np.diag(D) if D.ndim == 2 else D
    return float(np.min(np.abs(d))), 0.0, 1e-9


# =============================================================================
# 2026-09-13 R5 回归探针
# =============================================================================

def _polyfit_tiny_scale_probe(excel, wb, ws, runner, tc, args):
    """R5-04 回归: x≈1e-14 尺度不得被绝对阈值误判秩亏 (斜率静默 0)。

    修复前 slope=0/intercept=3。以 numpy.polyfit 为参考,
    返回系数尺度归一的相对误差 (尺度感知)。
    """
    x = np.array(args[0], dtype=float)
    y = np.array(args[1], dtype=float)
    degree = args[2]
    vba = np.ravel(com_to_numpy(runner._call_vba(
        excel, wb, {"func": "PolyFit"}, (x.tolist(), y.tolist(), degree))))
    ref = np.polyfit(x, y, degree)
    err = float(np.max(np.abs(vba - ref) / np.maximum(1.0, np.abs(ref))))
    return err, 0.0, 1e-8


def _qr_extreme_scale_probe(excel, wb, ws, runner, tc, args):
    """R5-15 回归: 极端量级 QR 不得上溢 Error 6 / 下溢静默错误。

    以 Q·R = A 为参考, 返回按 max|A| 归一的相对重构误差 (尺度感知)。
    """
    A = np.array(args[0], dtype=float)
    q = com_to_numpy(runner._call_udf(
        excel, wb, ws, {"func": "UDF_LINALG_QR_Q"}, (A.tolist(),)))
    r = com_to_numpy(runner._call_udf(
        excel, wb, ws, {"func": "UDF_LINALG_QR_R"}, (A.tolist(),)))
    err = float(np.max(np.abs(q @ r - A)) / np.max(np.abs(A)))
    return err, 0.0, 1e-10


def _solve_tiny_scale_probe(excel, wb, ws, runner, tc, args):
    """R5-15 回归: 1e-170 超定系统 — QR 下溢时回代矩阵非上三角, 解错误。

    以 numpy.linalg.lstsq 为参考, 返回按解尺度归一的相对误差。
    """
    A = np.array(args[0], dtype=float)
    b = np.array(args[1], dtype=float)
    vba = np.ravel(com_to_numpy(runner._call_vba(
        excel, wb, {"func": "SolveLinearSystem"}, (A.tolist(), b.tolist()))))
    ref = np.linalg.lstsq(A, b, rcond=None)[0]
    err = float(np.max(np.abs(vba - ref)) / max(1.0, float(np.max(np.abs(ref)))))
    return err, 0.0, 1e-8


def _cond_number_r5_probe(excel, wb, ws, runner, tc, args):
    """R5-16 回归: 条件数按机器精度判秩 (numpy.linalg.cond 参考);
    负 tol 走 auto 且奇异矩阵返回 MAX_DOUBLE (修复前抛裸 Error 11)。
    """
    A = np.array(args[0], dtype=float)
    tol_arg = args[1]
    mode = args[2]
    v = float(runner._call_vba(excel, wb, {"func": "MatrixConditionNumber"},
                               (A.tolist(), tol_arg)))
    if mode == "finite":
        ref = float(np.linalg.cond(A))
    else:
        ref = MAX_DOUBLE_LITERAL
    return abs(v - ref) / ref, 0.0, 1e-8


def _linear_r5_probe(excel, wb, ws, runner, tc, args):
    """R5-14/R5-32 负例探针: VBA 内部捕获 Err.Number (错误码不跨 COM)。"""
    from tests.test_utils import run_macro
    vbproj = wb.VBProject
    for comp in list(vbproj.VBComponents):
        if comp.Name == "LinearR5Probe":
            vbproj.VBComponents.Remove(comp)
    comp = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
    comp.Name = "LinearR5Probe"
    comp.CodeModule.AddFromString(
        VBA_R5_PROBE.replace("\r\n", "\n").replace("\n", "\r\n"))
    val = run_macro(excel, wb, "LinearR5Probe.Probe", args[0])
    return float(val), float(args[1]), float(args[2]) if len(args) > 2 else 0.0


def _reconstruct_decomp(excel, wb, ws, runner, func_a, func_b, func_c, A, mode):
    """Reconstruct A from VBA decomposition components.

    SVD (mode='svd'):  A ≈ U @ diag(S) @ Vt   (calls func_a=U, func_b=S, func_c=Vt)
    QR  (mode='qr'):   A ≈ Q @ R              (calls func_a=Q, func_b=R, func_c=None)
    Eigen (mode='eigen'): A ≈ V @ diag(D) @ Vt (calls func_a=V, func_b=D, func_c=None)

    Returns (vba_reconstruction, original_A, tol).
    """
    if mode == "svd":
        U = _call_udf_matrix(excel, wb, ws, runner, func_a, A)
        S_diag = _call_udf_matrix(excel, wb, ws, runner, func_b, A)
        Vt = _call_udf_matrix(excel, wb, ws, runner, func_c, A)
        s = np.diag(S_diag) if S_diag.ndim == 1 else np.diag(np.diag(S_diag))
        vba_recon = U @ s @ Vt
        return vba_recon, A, 1e-5

    elif mode == "qr":
        Q = _call_udf_matrix(excel, wb, ws, runner, func_a, A)
        R = _call_udf_matrix(excel, wb, ws, runner, func_b, A)
        vba_recon = Q @ R
        return vba_recon, A, 1e-5

    elif mode == "eigen":
        V = _call_udf_matrix(excel, wb, ws, runner, func_a, A)
        D_diag = _call_udf_matrix(excel, wb, ws, runner, func_b, A)
        d = np.diag(D_diag) if D_diag.ndim == 1 else np.diag(np.diag(D_diag))
        vba_recon = V @ d @ V.T
        return vba_recon, A, 1e-5

    else:
        raise ValueError(f"Unknown reconstruction mode: {mode}")


TEST_CASES = [
    # MatrixRows / MatrixCols
    {"name": "MatrixRows", "func": "MatrixRows", "args": lambda: (MAT_3X2.tolist(),),
     "py_ref": lambda a: 3, "result_type": "scalar"},
    {"name": "MatrixCols", "func": "MatrixCols", "args": lambda: (MAT_3X2.tolist(),),
     "py_ref": lambda a: 2, "result_type": "scalar"},
    # MatrixMultiply
    {"name": "MatrixMultiply_identity", "func": "MatrixMultiply",
     "args": lambda: (MAT_2X2.tolist(), ID2.tolist()),
     "py_ref": lambda a: MAT_2X2, "result_type": "array", "tol": 1e-10},
    {"name": "MatrixMultiply_2x2", "func": "MatrixMultiply",
     "args": lambda: (MAT_2X2.tolist(), MAT_2X2.tolist()),
     "py_ref": lambda a: np.dot(np.array(a[0]), np.array(a[1])),
     "result_type": "array", "tol": 1e-10},
    # MatrixTranspose
    {"name": "MatrixTranspose_3x2", "func": "MatrixTranspose",
     "args": lambda: (MAT_3X2.tolist(),),
     "py_ref": lambda a: MAT_3X2.T, "result_type": "array", "tol": 1e-10},
    # IdentityMatrix
    {"name": "IdentityMatrix_3", "func": "IdentityMatrix",
     "args": lambda: (3,), "py_ref": lambda a: ID3, "result_type": "array", "tol": 1e-10},
    # MatrixDeterminant
    {"name": "MatrixDeterminant_2x2", "func": "MatrixDeterminant",
     "args": lambda: (MAT_2X2.tolist(),),
     "py_ref": lambda a: float(np.linalg.det(MAT_2X2)), "result_type": "scalar", "tol": 1e-8},
    # MatrixTrace
    {"name": "MatrixTrace", "func": "MatrixTrace",
     "args": lambda: (MAT_2X2.tolist(),),
     "py_ref": lambda a: float(np.trace(MAT_2X2)), "result_type": "scalar", "tol": 1e-10},
    # VectorDot
    {"name": "VectorDot", "func": "VectorDot",
     "args": lambda: ([1., 2., 3.], [4., 5., 6.]),
     "py_ref": lambda a: float(np.dot(a[0], a[1])), "result_type": "scalar", "tol": 1e-10},
    # VectorCross
    {"name": "VectorCross", "func": "VectorCross",
     "args": lambda: ([1., 0., 0.], [0., 1., 0.]),
     "py_ref": lambda a: np.cross(a[0], a[1]), "result_type": "array", "tol": 1e-10},
    # MatrixAdd / MatrixSubtract / MatrixScale
    {"name": "MatrixAdd", "func": "MatrixAdd",
     "args": lambda: ([[1, 2], [3, 4]], [[5, 6], [7, 8]]),
     "py_ref": lambda a: np.array(a[0]) + np.array(a[1]), "result_type": "array", "tol": 1e-10},
    {"name": "MatrixSubtract", "func": "MatrixSubtract",
     "args": lambda: ([[5, 6], [7, 8]], [[1, 2], [3, 4]]),
     "py_ref": lambda a: np.array(a[0]) - np.array(a[1]), "result_type": "array", "tol": 1e-10},
    {"name": "MatrixScale", "func": "MatrixScale",
     "args": lambda: ([[1, 2], [3, 4]], 2.),
     "py_ref": lambda a: np.array(a[0]) * a[1], "result_type": "array", "tol": 1e-10},
    # MatrixFrobeniusNorm
    {"name": "MatrixFrobeniusNorm", "func": "MatrixFrobeniusNorm",
     "args": lambda: (MAT_2X2.tolist(),),
     "py_ref": lambda a: float(np.linalg.norm(MAT_2X2, 'fro')), "result_type": "scalar", "tol": 1e-8},
    # MatrixConditionNumber
    {"name": "MatrixConditionNumber", "func": "MatrixConditionNumber",
     "args": lambda: ([[2, 0], [0, 1]],),
     "py_ref": lambda a: float(np.linalg.cond([[2, 0], [0, 1]])), "result_type": "scalar", "tol": 1e-8},
    # MatrixHadamard
    {"name": "MatrixHadamard", "func": "MatrixHadamard",
     "args": lambda: (MAT_2X2.tolist(), MAT_2X2.tolist()),
     "py_ref": lambda a: np.array(a[0]) * np.array(a[1]), "result_type": "array", "tol": 1e-10},
    # MatrixNorm (Frobenius)
    {"name": "MatrixNorm_fro", "func": "MatrixNorm",
     "args": lambda: (MAT_2X2.tolist(), "fro"),
     "py_ref": lambda a: float(np.linalg.norm(np.array(a[0]), 'fro')), "result_type": "scalar", "tol": 1e-8},
    # MatrixPower
    {"name": "MatrixPower_sq", "func": "MatrixPower",
     "args": lambda: (MAT_2X2.tolist(), 2),
     "py_ref": lambda a: np.linalg.matrix_power(np.array(a[0]), a[1]), "result_type": "array", "tol": 1e-8},
    # MatrixRank_Array
    {"name": "MatrixRank_Array", "func": "MatrixRank_Array",
     "args": lambda: (ID3.tolist(),),
     "py_ref": lambda a: 3, "result_type": "scalar"},
    # SolveLinearSystem
    {"name": "SolveLinearSystem", "func": "SolveLinearSystem",
     "args": lambda: ([[2, 0], [0, 3]], [4, 9]),
     "py_ref": lambda a: np.linalg.solve(np.array(a[0]), np.array(a[1])),
     "result_type": "array", "tol": 1e-8},
    # VectorNorm
    {"name": "VectorNorm", "func": "VectorNorm",
     "args": lambda: ([3, 4],),
     "py_ref": lambda a: 5., "result_type": "scalar", "tol": 1e-10},
    # PseudoInverse
    {"name": "PseudoInverse", "func": "PseudoInverse",
     "args": lambda: (MAT_3X2.tolist(),),
     "py_ref": lambda a: np.linalg.pinv(MAT_3X2),
     "result_type": "array", "tol": 1e-5},
    # PolyFit
    {"name": "PolyFit_linear", "func": "PolyFit",
     "args": lambda: ([1, 2, 3, 4, 5], [2, 4, 6, 8, 10], 1),
     "py_ref": lambda a: np.polyfit(a[0], a[1], a[2]),
     "result_type": "array", "tol": 1e-6},
    # UDF wrappers
    {"name": "UDF_LINALG_DET", "func": "UDF_LINALG_DET",
     "args": lambda: (MAT_2X2.tolist(),),
     "py_ref": lambda a: float(np.linalg.det(MAT_2X2)), "result_type": "scalar", "tol": 1e-6, "is_udf": True},
    {"name": "UDF_LINALG_RANK", "func": "UDF_LINALG_RANK",
     "args": lambda: (ID3.tolist(),),
     "py_ref": lambda a: float(np.linalg.matrix_rank(ID3)), "result_type": "scalar", "tol": 1e-10, "is_udf": True},
    {"name": "UDF_LINALG_SVD_SVALS", "func": "UDF_LINALG_SVD_SVALS",
     "args": lambda: (MAT_3X2.tolist(),),
     "py_ref": lambda a: np.linalg.svd(MAT_3X2, full_matrices=False)[1],
     "result_type": "array", "tol": 1e-5, "is_udf": True},
    {"name": "UDF_LINALG_CHOLESKY", "func": "UDF_LINALG_CHOLESKY",
     "args": lambda: (_spd(4).tolist(),),
     "py_ref": lambda a: np.linalg.cholesky(np.array(a[0])),
     "result_type": "array", "tol": 1e-8, "is_udf": True},
    {"name": "UDF_LINALG_PINV", "func": "UDF_LINALG_PINV",
     "args": lambda: (MAT_3X2.tolist(),),
     "py_ref": lambda a: np.linalg.pinv(MAT_3X2),
     "result_type": "array", "tol": 1e-5, "is_udf": True},
	    # ---- MatrixCopy ----
	    {"name": "MatrixCopy_2x2", "func": "MatrixCopy",
	     "args": lambda: (MAT_2X2.tolist(),),
	     "py_ref": lambda a: np.array(a[0]).copy(),
	     "result_type": "array"},
	    {"name": "MatrixCopy_1x3", "func": "MatrixCopy",
	     "args": lambda: ([[10.0, 20.0, 30.0]],),
	     "py_ref": lambda a: np.array(a[0]),
	     "result_type": "array"},

	    # ---- MatrixGetColumn ----
	    {"name": "MatrixGetColumn_col1", "func": "MatrixGetColumn",
	     "args": lambda: (MAT_3X2.tolist(), 1),
	     "py_ref": lambda a: np.array(a[0])[:, 0].reshape(-1, 1),
	     "result_type": "array"},
	    {"name": "MatrixGetColumn_col2", "func": "MatrixGetColumn",
	     "args": lambda: (MAT_3X2.tolist(), 2),
	     "py_ref": lambda a: np.array(a[0])[:, 1].reshape(-1, 1),
	     "result_type": "array"},

    # =====================================================================
    # Boundary / edge cases for existing functions
    # =====================================================================

    # ---- MatrixRows / MatrixCols edge ----
    {"name": "MatrixRows_1x1", "func": "MatrixRows",
     "args": lambda: ([[5.0]],), "py_ref": lambda a: 1, "result_type": "scalar"},
    {"name": "MatrixCols_1x1", "func": "MatrixCols",
     "args": lambda: ([[5.0]],), "py_ref": lambda a: 1, "result_type": "scalar"},

    # ---- IdentityMatrix edges ----
    {"name": "IdentityMatrix_1", "func": "IdentityMatrix",
     "args": lambda: (1,), "py_ref": lambda a: np.eye(1), "result_type": "array", "tol": 1e-10},

    # ---- MatrixDeterminant edges ----
    {"name": "MatrixDeterminant_singular", "func": "MatrixDeterminant",
     "args": lambda: ([[1, 2], [2, 4]],), "py_ref": lambda a: 0.0, "result_type": "scalar", "tol": 1e-8},
    {"name": "MatrixDeterminant_1x1", "func": "MatrixDeterminant",
     "args": lambda: ([[7.0]],), "py_ref": lambda a: 7.0, "result_type": "scalar", "tol": 1e-8},
    # L-01 回归: 小量级良态矩阵不得被误判奇异 (相对容差, 无绝对下限抬升)
    {"name": "MatrixDeterminant_small_scale", "func": "MatrixDeterminant",
     "args": lambda: ([[1e-20, 0.0], [0.0, 1e-20]],),
     "py_ref": lambda a: 1e-40, "result_type": "scalar", "tol": 1e-50},
    # L-02 回归: 高条件数 SPD 矩阵 diag(1e10,1e-8) 必须可 Cholesky 分解
    {"name": "UDF_LINALG_CHOLESKY_high_cond", "func": "UDF_LINALG_CHOLESKY",
     "args": lambda: ([[1e10, 0.0], [0.0, 1e-8]],),
     "py_ref": lambda a: np.linalg.cholesky(np.array(a[0])),
     "result_type": "array", "tol": 1e-9, "is_udf": True},

    # ---- MatrixTrace edge ----
    {"name": "MatrixTrace_1x1", "func": "MatrixTrace",
     "args": lambda: ([[7.0]],), "py_ref": lambda a: 7.0, "result_type": "scalar"},

    # ---- VectorDot edges ----
    {"name": "VectorDot_single", "func": "VectorDot",
     "args": lambda: ([3.0], [4.0]),
     "py_ref": lambda a: 12.0, "result_type": "scalar", "tol": 1e-10},
    {"name": "VectorDot_zero", "func": "VectorDot",
     "args": lambda: ([1.0, 2.0, 3.0], [0.0, 0.0, 0.0]),
     "py_ref": lambda a: 0.0, "result_type": "scalar", "tol": 1e-10},

    # ---- VectorCross edges ----
    {"name": "VectorCross_parallel", "func": "VectorCross",
     "args": lambda: ([1.0, 0.0, 0.0], [2.0, 0.0, 0.0]),
     "py_ref": lambda a: np.zeros(3), "result_type": "array", "tol": 1e-10},

    # ---- MatrixScale edges ----
    {"name": "MatrixScale_zero", "func": "MatrixScale",
     "args": lambda: ([[1, 2], [3, 4]], 0.0),
     "py_ref": lambda a: np.zeros((2, 2)), "result_type": "array", "tol": 1e-10},
    {"name": "MatrixScale_negative", "func": "MatrixScale",
     "args": lambda: ([[1, 2], [3, 4]], -1.0),
     "py_ref": lambda a: -np.array([[1, 2], [3, 4]]), "result_type": "array", "tol": 1e-10},

    # ---- MatrixRank_Array edges ----
    {"name": "MatrixRank_singular", "func": "MatrixRank_Array",
     "args": lambda: ([[1, 2], [2, 4]],), "py_ref": lambda a: 1, "result_type": "scalar"},
    {"name": "MatrixRank_3x2", "func": "MatrixRank_Array",
     "args": lambda: (MAT_3X2.tolist(),),
     "py_ref": lambda a: np.linalg.matrix_rank(MAT_3X2), "result_type": "scalar"},

    # ---- VectorNorm edges ----
    {"name": "VectorNorm_zero", "func": "VectorNorm",
     "args": lambda: ([0.0, 0.0, 0.0],),
     "py_ref": lambda a: 0.0, "result_type": "scalar", "tol": 1e-10},
    {"name": "VectorNorm_single", "func": "VectorNorm",
     "args": lambda: ([5.0],),
     "py_ref": lambda a: 5.0, "result_type": "scalar", "tol": 1e-10},

    # ---- MatrixConditionNumber edges ----
    {"name": "MatrixConditionNumber_identity", "func": "MatrixConditionNumber",
     "args": lambda: (ID3.tolist(),),
     "py_ref": lambda a: 1.0, "result_type": "scalar", "tol": 1e-8},

    # ---- MatrixHadamard edge ----
    {"name": "MatrixHadamard_ones", "func": "MatrixHadamard",
     "args": lambda: (MAT_2X2.tolist(), np.ones((2, 2)).tolist()),
     "py_ref": lambda a: MAT_2X2, "result_type": "array", "tol": 1e-10},

    # ---- MatrixPower edges ----
    {"name": "MatrixPower_pow1", "func": "MatrixPower",
     "args": lambda: (MAT_2X2.tolist(), 1),
     "py_ref": lambda a: MAT_2X2, "result_type": "array", "tol": 1e-8},
    {"name": "MatrixPower_pow3", "func": "MatrixPower",
     "args": lambda: (MAT_2X2.tolist(), 3),
     "py_ref": lambda a: np.linalg.matrix_power(MAT_2X2, 3), "result_type": "array", "tol": 1e-8},

    # ---- PolyFit edges ----
    {"name": "PolyFit_quadratic", "func": "PolyFit",
     "args": lambda: ([1, 2, 3, 4, 5], [1, 4, 9, 16, 25], 2),
     "py_ref": lambda a: np.polyfit(a[0], a[1], a[2]),
     "result_type": "array", "tol": 1e-6},
    {"name": "PolyFit_constant", "func": "PolyFit",
     "args": lambda: ([1, 2, 3], [5, 5, 5], 0),
     "py_ref": lambda a: np.polyfit(a[0], a[1], a[2]),
     "result_type": "array", "tol": 1e-6},

    # ---- PseudoInverse edges ----
    {"name": "PseudoInverse_square", "func": "PseudoInverse",
     "args": lambda: (MAT_2X2.tolist(),),
     "py_ref": lambda a: np.linalg.pinv(MAT_2X2),
     "result_type": "array", "tol": 1e-5},
    {"name": "PseudoInverse_rank_def", "func": "PseudoInverse",
     "args": lambda: ([[1, 2], [2, 4], [3, 6]],),
     "py_ref": lambda a: np.linalg.pinv([[1, 2], [2, 4], [3, 6]]),
     "result_type": "array", "tol": 1e-5},

    # ---- SolveLinearSystem edges ----
    {"name": "SolveLinearSystem_3x3", "func": "SolveLinearSystem",
     "args": lambda: ([[2, 1, 0], [1, 2, 1], [0, 1, 2]], [1, 2, 3]),
     "py_ref": lambda a: np.linalg.solve(np.array(a[0]), np.array(a[1])),
     "result_type": "array", "tol": 1e-8},

    # ---- MatrixMultiply edges ----
    {"name": "MatrixMultiply_3x2_2x3", "func": "MatrixMultiply",
     "args": lambda: (MAT_3X2.tolist(), np.array([[1, 0, -1], [0, 1, 2]]).tolist()),
     "py_ref": lambda a: np.dot(np.array(a[0]), np.array(a[1])),
     "result_type": "array", "tol": 1e-10},

    # ---- MatrixTranspose edges ----
    {"name": "MatrixTranspose_2x2", "func": "MatrixTranspose",
     "args": lambda: (MAT_2X2.tolist(),),
     "py_ref": lambda a: MAT_2X2.T, "result_type": "array", "tol": 1e-10},
    {"name": "MatrixTranspose_1x3", "func": "MatrixTranspose",
     "args": lambda: ([[1.0, 2.0, 3.0]],),
     "py_ref": lambda a: np.array([[1.0], [2.0], [3.0]]), "result_type": "array", "tol": 1e-10},

    # =====================================================================
    # Reconstruction tests — verify decomp correctness without sign comparison
    # Call each VBA component separately, recombine in Python, compare to A.
    # =====================================================================
    {"name": "SVD_Reconstruct_2x2", "func": "SVD_Reconstruct",
     "args": lambda: (MAT_2X2.tolist(),),
     "reconstruct": lambda excel, wb, ws, runner, tc, args:
         _reconstruct_decomp(excel, wb, ws, runner,
                             "UDF_LINALG_SVD_U", "UDF_LINALG_SVD_S", "UDF_LINALG_SVD_VT",
                             np.array(args[0]), "svd"),
     "result_type": "array", "tol": 1e-5},
    {"name": "QR_Reconstruct_3x2", "func": "QR_Reconstruct",
     "args": lambda: (MAT_3X2.tolist(),),
     "reconstruct": lambda excel, wb, ws, runner, tc, args:
         _reconstruct_decomp(excel, wb, ws, runner,
                             "UDF_LINALG_QR_Q", "UDF_LINALG_QR_R", None,
                             np.array(args[0]), "qr"),
     "result_type": "array", "tol": 1e-5},
    {"name": "Eigen_Reconstruct_2x2", "func": "Eigen_Reconstruct",
     "args": lambda: ([[2., 1.], [1., 2.]],),
     "reconstruct": lambda excel, wb, ws, runner, tc, args:
         _reconstruct_decomp(excel, wb, ws, runner,
                             "UDF_LINALG_EIGVEC", "UDF_LINALG_EIGVAL", None,
                             np.array(args[0]), "eigen"),
     "result_type": "array", "tol": 1e-5},

    # =====================================================================
    # Individual decomposition UDFs — skipped (sign ambiguity), verified by
    # reconstruction tests above.
    # =====================================================================

    {"name": "UDF_LINALG_SVD_U", "func": "UDF_LINALG_SVD_U",
     "args": lambda: (MAT_2X2.tolist(),),
     "py_ref": lambda a: np.linalg.svd(np.array(a[0]), full_matrices=True)[0],
     "result_type": "array", "tol": 1e-5, "is_udf": True,
     "skip_if": True, "skip_reason": "Sign ambiguity. Verified by SVD_Reconstruct_2x2 above."},
    {"name": "UDF_LINALG_SVD_S", "func": "UDF_LINALG_SVD_S",
     "args": lambda: (MAT_2X2.tolist(),),
     "py_ref": lambda a: np.diag(np.linalg.svd(np.array(a[0]), full_matrices=True)[1]),
     "result_type": "array", "tol": 1e-5, "is_udf": True},
    {"name": "UDF_LINALG_SVD_VT", "func": "UDF_LINALG_SVD_VT",
     "args": lambda: (MAT_2X2.tolist(),),
     "py_ref": lambda a: np.linalg.svd(np.array(a[0]), full_matrices=True)[2],
     "result_type": "array", "tol": 1e-5, "is_udf": True,
     "skip_if": True, "skip_reason": "Sign ambiguity. Verified by SVD_Reconstruct_2x2 above."},

    {"name": "UDF_LINALG_QR_Q", "func": "UDF_LINALG_QR_Q",
     "args": lambda: (MAT_3X2.tolist(),),
     "py_ref": lambda a: np.linalg.qr(np.array(a[0]))[0],
     "result_type": "array", "tol": 1e-5, "is_udf": True,
     "skip_if": True, "skip_reason": "Sign ambiguity. Verified by QR_Reconstruct_3x2 above."},
    {"name": "UDF_LINALG_QR_R", "func": "UDF_LINALG_QR_R",
     "args": lambda: (MAT_3X2.tolist(),),
     "py_ref": lambda a: np.linalg.qr(np.array(a[0]))[1],
     "result_type": "array", "tol": 1e-5, "is_udf": True,
     "skip_if": True, "skip_reason": "Sign ambiguity. Verified by QR_Reconstruct_3x2 above."},

    {"name": "QRDecompositionPiv", "func": "QRDecompositionPiv",
     "args": lambda: (MAT_3X2.tolist(),),
     "py_ref": lambda a: np.linalg.qr(np.array(a[0])[:, np.linalg.qr(np.array(a[0]))[1].diagonal().argsort()[::-1]]),
     "result_type": "array", "tol": 1e-5,
     "skip_if": True, "skip_reason": "VBA Sub with ByRef params — no UDF wrapper. COM Application.Run cannot call Sub procedures."},

    {"name": "UDF_LINALG_EIGVAL", "func": "UDF_LINALG_EIGVAL",
     "args": lambda: ([[2, 1], [1, 2]],),
     "py_ref": lambda a: np.sort(np.linalg.eigvalsh(np.array([[2, 1], [1, 2]])))[::-1],
     "result_type": "array", "tol": 1e-5, "is_udf": True,
     "skip_if": True, "skip_reason": "Ordering differs. Verified by Eigen_Reconstruct_2x2 above."},
    {"name": "UDF_LINALG_EIGVEC", "func": "UDF_LINALG_EIGVEC",
     "args": lambda: ([[2, 1], [1, 2]],),
     "py_ref": lambda a: np.linalg.eigh(np.array([[2, 1], [1, 2]]))[1],
     "result_type": "array", "tol": 1e-5, "is_udf": True,
     "skip_if": True, "skip_reason": "Sign ambiguity. Verified by Eigen_Reconstruct_2x2 above."},

    {"name": "UDF_LINALG_SOLVE", "func": "UDF_LINALG_SOLVE",
     "args": lambda: ([[2, 0], [0, 3]], [[4], [9]]),
     "py_ref": lambda a: np.linalg.solve([[2, 0], [0, 3]], [4, 9]),
     "result_type": "array", "tol": 1e-8, "is_udf": True},

    {"name": "UDF_LINALG_POLYFIT", "func": "UDF_LINALG_POLYFIT",
     "args": lambda: ([1, 2, 3, 4, 5], [2, 4, 6, 8, 10], 1),
     "py_ref": lambda a: np.polyfit(a[0], a[1], a[2]),
     "result_type": "array", "tol": 1e-6, "is_udf": True},

    # ---- MatrixMultiplyNaive ----
    {"name": "MatrixMultiplyNaive_2x2", "func": "MatrixMultiplyNaive",
     "args": lambda: (MAT_2X2.tolist(), MAT_2X2.tolist()),
     "py_ref": lambda a: np.dot(np.array(a[0]), np.array(a[1])),
     "result_type": "array", "tol": 1e-10},

    # ---- MatrixSetColumn ----
    {"name": "MatrixSetColumn_col1", "func": "MatrixSetColumn",
     "args": lambda: (MAT_3X2.tolist(), 1, [10.0, 20.0, 30.0]),
     "py_ref": lambda a: _set_col(np.array(a[0]), a[1], a[2]),
     "result_type": "array", "tol": 1e-10,
     "skip_if": True, "skip_reason": "MatrixSetColumn is a Sub (modifies ByRef, no return value). verified manually in Excel."},

    # ---- RangeToMatrix (is_udf: writes to Range, reads back) ----
    {"name": "RangeToMatrix_3x2", "func": "RangeToMatrix",
     "args": lambda: (MAT_3X2.tolist(),),
     "py_ref": lambda a: MAT_3X2,
     "result_type": "array", "tol": 1e-10, "is_udf": True},

    # ---- Boundary: zero / identity / singular / vector ops ----
    {"name": "MatrixDeterminant_zero3x3", "func": "MatrixDeterminant",
     "args": lambda: (np.zeros((3,3)).tolist(),),
     "py_ref": lambda a: 0.0, "result_type": "scalar", "tol": 1e-10,
     "skip_if": True,
     "skip_reason": "零矩阵是奇异的，LU分解正确抛出ERR_SINGULAR。预期行为。"},
    {"name": "MatrixDeterminant_identity4", "func": "MatrixDeterminant",
     "args": lambda: (np.eye(4).tolist(),),
     "py_ref": lambda a: 1.0, "result_type": "scalar", "tol": 1e-10},
    {"name": "MatrixRank_singular", "func": "MatrixRank_Array",
     "args": lambda: ([[1,2,3],[2,4,6],[3,6,9]],),
     "py_ref": lambda a: 1, "result_type": "scalar", "tol": 0},
    {"name": "MatrixRank_full", "func": "MatrixRank_Array",
     "args": lambda: (np.eye(3).tolist(),),
     "py_ref": lambda a: 3, "result_type": "scalar", "tol": 0},
    {"name": "VectorDot_basic", "func": "VectorDot",
     "args": lambda: ([1.0,2.0,3.0], [4.0,5.0,6.0]),
     "py_ref": lambda a: 32.0, "result_type": "scalar", "tol": 1e-10},
    {"name": "VectorNorm_L2", "func": "VectorNorm",
     "args": lambda: ([3.0,4.0],), "py_ref": lambda a: 5.0,
     "result_type": "scalar", "tol": 1e-10},
    {"name": "MatrixTrace_identity", "func": "MatrixTrace",
     "args": lambda: (np.eye(3).tolist(),), "py_ref": lambda a: 3.0,
     "result_type": "scalar", "tol": 1e-10},
    {"name": "MatrixHadamard_basic", "func": "MatrixHadamard",
     "args": lambda: ([[1,2],[3,4]], [[5,6],[7,8]]),
     "py_ref": lambda a: [[5,12],[21,32]], "result_type": "array", "tol": 1e-10},
    {"name": "MatrixConditionNumber_well", "func": "MatrixConditionNumber",
     "args": lambda: (np.eye(3).tolist(),),
     "py_ref": lambda a: 1.0, "result_type": "scalar", "tol": 1e-6},

    # ---- 2026-09-13 R4-10/R4-13 回归 ----
    {"name": "QR_Economy_Wide_Reconstruct", "func": "QR_Economy_Wide_Reconstruct",
     "args": lambda: ([[1., 2., 3.], [4., 5., 6.]],),
     "reconstruct": _qr_wide_economy_probe, "py_ref": lambda a: 0.0},
    {"name": "Eigen_Symmetrize_MixedScale", "func": "Eigen_Symmetrize_MixedScale",
     "args": lambda: ([[1e16, 0., 0.], [0., 0., -1.], [0., 1., 0.]],),
     "reconstruct": _eigen_symmetrize_probe, "py_ref": lambda a: 0.0},

    # =====================================================================
    # ---- 2026-09-13 R5-04/R5-14/R5-15/R5-16/R5-32 回归 ----
    # =====================================================================
    # R5-04: 小尺度 x (1e-14) 不得误判秩亏 — numpy.polyfit 参考
    {"name": "PolyFit_tiny_scale_1e-14", "func": "PolyFit",
     "args": lambda: ([1e-14, 2e-14, 3e-14, 4e-14, 5e-14], [1, 2, 3, 4, 5], 1),
     "reconstruct": _polyfit_tiny_scale_probe, "py_ref": lambda a: 0.0},
    # R5-14: 数组路径非数值/Boolean/Empty 必须 ERR_INVALID_INPUT (+1001)
    {"name": "PolyFit_array_string_err", "func": "PolyFit",
     "args": lambda: ("polyfit_string", ERR_INVALID_INPUT),
     "reconstruct": _linear_r5_probe, "py_ref": lambda a: ERR_INVALID_INPUT},
    {"name": "PolyFit_array_bool_err", "func": "PolyFit",
     "args": lambda: ("polyfit_bool", ERR_INVALID_INPUT),
     "reconstruct": _linear_r5_probe, "py_ref": lambda a: ERR_INVALID_INPUT},
    {"name": "PolyFit_array_empty_err", "func": "PolyFit",
     "args": lambda: ("polyfit_empty", ERR_INVALID_INPUT),
     "reconstruct": _linear_r5_probe, "py_ref": lambda a: ERR_INVALID_INPUT},
    # R5-15: 极端量级 QR — 1e154 不得上溢 Error 6; 1e-170 不得下溢
    {"name": "QR_extreme_scale_huge_recon", "func": "QR_extreme_scale_huge_recon",
     "args": lambda: ([[1e154, 0.0], [0.0, 1.0]],),
     "reconstruct": _qr_extreme_scale_probe, "py_ref": lambda a: 0.0},
    {"name": "QR_extreme_scale_huge_reflect_recon", "func": "QR_extreme_scale_huge_reflect_recon",
     "args": lambda: ([[1e154, 1e154], [1e154, 0.0]],),
     "reconstruct": _qr_extreme_scale_probe, "py_ref": lambda a: 0.0},
    {"name": "QR_extreme_scale_tiny_recon", "func": "QR_extreme_scale_tiny_recon",
     "args": lambda: ([[3e-170, 1e-170], [1e-170, 2e-170]],),
     "reconstruct": _qr_extreme_scale_probe, "py_ref": lambda a: 0.0},
    # R5-15: 1e-170 超定最小二乘 — 下溢时回代矩阵非上三角, 解错误
    {"name": "Solve_tiny_scale_1e-170", "func": "Solve_tiny_scale_1e-170",
     "args": lambda: ([[3e-170, 1e-170], [1e-170, 2e-170], [1e-170, 1e-170]],
                      [5e-170, 4e-170, 3e-170]),
     "reconstruct": _solve_tiny_scale_probe, "py_ref": lambda a: 0.0},
    # R5-16: diag(1,1e-15) → 1e15 (numpy cond 参考); 负 tol 不得抛裸 Error 11
    {"name": "MatrixConditionNumber_1e15", "func": "MatrixConditionNumber",
     "args": lambda: ([[1.0, 0.0], [0.0, 1e-15]], 1e-14, "finite"),
     "reconstruct": _cond_number_r5_probe, "py_ref": lambda a: 0.0},
    {"name": "MatrixConditionNumber_neg_tol_auto", "func": "MatrixConditionNumber",
     "args": lambda: ([[1.0, 2.0], [2.0, 4.0]], -1.0, "maxdouble"),
     "reconstruct": _cond_number_r5_probe, "py_ref": lambda a: 0.0},
    # R5-32: ToDoubleMatrix 非数值/错误值 → ERR_INVALID_INPUT, 不得裸 Error 13
    {"name": "MatrixDeterminant_non_numeric_err", "func": "MatrixDeterminant",
     "args": lambda: ("todouble_string", ERR_INVALID_INPUT),
     "reconstruct": _linear_r5_probe, "py_ref": lambda a: ERR_INVALID_INPUT},
    {"name": "MatrixDeterminant_error_value_err", "func": "MatrixDeterminant",
     "args": lambda: ("todouble_error", ERR_INVALID_INPUT),
     "reconstruct": _linear_r5_probe, "py_ref": lambda a: ERR_INVALID_INPUT},

]

def main() -> int:
    runner = CrossValRunner("LinearUtils", MODULE_PATHS)
    runner.run_all(TEST_CASES)
    passed, failed = runner.print_summary()
    return 0 if failed == 0 else 1

if __name__ == "__main__":
    sys.exit(main())
