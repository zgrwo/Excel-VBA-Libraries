"""Verify VBA decompositions on a highly rank-deficient matrix against numpy.

Matrix: 9x7, column 2-6 identical, column 7 all zeros, row 5 all zeros.
"""

import os, sys
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from tests.test_utils import (
    ensure_excel, teardown, create_workbook, run_macro, com_to_numpy,
    write_range, SRC_DIR, VBA_CORE_DIR, VBA_CORE_IMPORT_ORDER,
)

A = np.array([
    [49.45,  54.35,  54.35,  54.35,  54.35,  54.35,  0],
    [108.79, 119.57, 119.57, 119.57, 119.57, 119.57, 0],
    [17.6,   20,     20,     20,     20,     20,     0],
    [8448,   9600,   9600,   9600,   9600,   9600,   0],
    [0,      0,      0,      0,      0,      0,      0],
    [2358.4, 2680,   2680,   2680,   2680,   2680,   0],
    [880,    1000,   1000,   1000,   1000,   1000,   0],
    [4435.2, 5040,   5040,   5040,   5040,   5040,   0],
    [1460.8, 1660,   1660,   1660,   1660,   1660,   0],
], dtype=float)

print(f"Matrix shape: {A.shape}")
print(f"numpy rank:  {np.linalg.matrix_rank(A)}")
print(f"numpy cond:  {np.linalg.cond(A):.2e}")
print()

MODULE_PATHS = [
    os.path.join(VBA_CORE_DIR, name + ".cls") for name in VBA_CORE_IMPORT_ORDER
]
MODULE_PATHS.append(os.path.join(SRC_DIR, "LinearUtils.bas"))

print("Starting Excel...")
excel = ensure_excel()
wb_path = os.path.join(os.path.dirname(__file__), "_temp_rank_deficient.xlsm")
try:
    wb = create_workbook(excel, wb_path, MODULE_PATHS)
    ws = wb.Sheets("TestData")

    def call_udf(func_name, matrix):
        ws.UsedRange.ClearContents()
        arr = np.asarray(matrix, dtype=float)
        nr, nc = arr.shape
        write_range(ws, arr, 1, 1)
        rng = ws.Range(ws.Cells(1, 1), ws.Cells(nr, nc))
        raw = run_macro(excel, wb, func_name, rng)
        return com_to_numpy(raw)

    # ===== 1. PSEUDOINVERSE =====
    print("=" * 70)
    print("1. PSEUDOINVERSE")
    print("=" * 70)
    vba_pinv = call_udf("UDF_LINALG_PINV", A)
    py_pinv = np.linalg.pinv(A)
    pinv_err = np.max(np.abs(vba_pinv - py_pinv))
    print(f"  VBA shape: {vba_pinv.shape},  numpy shape: {py_pinv.shape}")
    print(f"  max|VBA - numpy| = {pinv_err:.6e}")
    print(f"  A @ pinv(A) @ A - A:       max|err| = {np.max(np.abs(A @ vba_pinv @ A - A)):.6e}")
    print(f"  pinv(A) @ A @ pinv(A) - pinv: max|err| = {np.max(np.abs(vba_pinv @ A @ vba_pinv - vba_pinv)):.6e}")
    print()

    # ===== 2. SVD =====
    print("=" * 70)
    print("2. SVD DECOMPOSITION")
    print("=" * 70)
    py_U, py_S, py_Vt = np.linalg.svd(A, full_matrices=False)
    print(f"  numpy U={py_U.shape}, S={py_S.shape}, Vt={py_Vt.shape}")
    print(f"  numpy S: {np.array2string(py_S, precision=4, max_line_width=120)}")

    vba_S_diag = call_udf("UDF_LINALG_SVD_S", A)
    vba_S_vec = np.diag(vba_S_diag)
    print(f"  VBA   S: {np.array2string(vba_S_vec, precision=4, max_line_width=120)}")
    S_err = np.max(np.abs(np.sort(vba_S_vec)[::-1] - np.sort(py_S)[::-1]))
    print(f"  max|S_VBA - S_numpy| (sorted) = {S_err:.6e}")

    vba_U = call_udf("UDF_LINALG_SVD_U", A)
    vba_Vt = call_udf("UDF_LINALG_SVD_VT", A)
    vba_svd_recon = vba_U @ np.diag(vba_S_vec) @ vba_Vt
    py_svd_recon = py_U @ np.diag(py_S) @ py_Vt
    print(f"  VBA  U*S*Vt ≈ A: max|err| = {np.max(np.abs(vba_svd_recon - A)):.6e}")
    print(f"  numpy U*S*Vt ≈ A: max|err| = {np.max(np.abs(py_svd_recon - A)):.6e}")
    print()

    # ===== 3. QR =====
    print("=" * 70)
    print("3. QR DECOMPOSITION")
    print("=" * 70)
    py_Q, py_R = np.linalg.qr(A)
    print(f"  numpy Q={py_Q.shape}, R={py_R.shape}")
    print(f"  numpy R diag: {np.array2string(np.diag(py_R), precision=4, max_line_width=120)}")

    vba_Q = call_udf("UDF_LINALG_QR_Q", A)
    vba_R = call_udf("UDF_LINALG_QR_R", A)
    print(f"  VBA   Q={vba_Q.shape}, R={vba_R.shape}")
    print(f"  VBA   R diag: {np.array2string(np.diag(vba_R), precision=4, max_line_width=120)}")

    vba_qr_recon = vba_Q @ vba_R
    py_qr_recon = py_Q @ py_R
    print(f"  VBA  Q*R ≈ A: max|err| = {np.max(np.abs(vba_qr_recon - A)):.6e}")
    print(f"  numpy Q*R ≈ A: max|err| = {np.max(np.abs(py_qr_recon - A)):.6e}")

    qtq_vba = vba_Q.T @ vba_Q
    qtq_py = py_Q.T @ py_Q
    print(f"  VBA  Q^T*Q ≈ I: max|err| = {np.max(np.abs(qtq_vba - np.eye(vba_Q.shape[1]))):.6e}")
    print(f"  numpy Q^T*Q ≈ I: max|err| = {np.max(np.abs(qtq_py - np.eye(py_Q.shape[1]))):.6e}")
    print()

    # ===== 4. LU =====
    print("=" * 70)
    print("4. LU DECOMPOSITION")
    print("=" * 70)
    try:
        from scipy.linalg import lu as scipy_lu
        py_P, py_L, py_U = scipy_lu(A)
        print(f"  scipy P={py_P.shape}, L={py_L.shape}, U={py_U.shape}")
        py_lu_err = np.max(np.abs(py_P @ py_L @ py_U - A))
        print(f"  scipy P*L*U ≈ A: max|err| = {py_lu_err:.6e}")
    except ImportError:
        print("  scipy not available — skipping LU numpy reference")

    print("  VBA LUDecomposition: Sub with ByRef params, cannot call via COM.")
    print("  Indirectly verified: SolveLinearSystem (uses LU) passes crossval.")
    print()

    # ===== SUMMARY =====
    print("=" * 70)
    print("SUMMARY")
    print("=" * 70)
    all_ok = True
    checks = [
        ("PseudoInverse vs numpy", pinv_err, 1e-3),
        ("SVD singular values", S_err, 1e-5),
        ("SVD reconstruction U*S*Vt ≈ A", np.max(np.abs(vba_svd_recon - A)), 1e-5),
        ("QR reconstruction Q*R ≈ A", np.max(np.abs(vba_qr_recon - A)), 1e-5),
        ("QR orthogonality Q^T*Q ≈ I", np.max(np.abs(qtq_vba - np.eye(vba_Q.shape[1]))), 1e-5),
    ]
    for name, err, tol in checks:
        status = "PASS" if err < tol else "FAIL"
        if err >= tol:
            all_ok = False
        print(f"  {status}: {name}: max|err| = {err:.6e}  (tol={tol:.0e})")

    if all_ok:
        print("\n  All checks PASSED.")
    else:
        print("\n  Some checks FAILED.")

finally:
    teardown(excel, wb)
    if os.path.exists(wb_path):
        try:
            os.remove(wb_path)
        except OSError:
            pass
