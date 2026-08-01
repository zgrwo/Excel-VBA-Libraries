"""Diagnose FitOLS R2=1.0 bug via COM Range path."""
import os, sys, numpy as np
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from tests.test_utils import (ensure_excel, teardown, create_workbook, run_macro, write_range,
                              SRC_DIR, VBA_CORE_DIR, VBA_CORE_IMPORT_ORDER)

np.random.seed(42)
n = 20
x1 = np.linspace(0, 10, n)
x2 = np.random.uniform(-2, 2, n)
y = 3.0 + 2.0 * x1 + 0.5 * x2 + np.random.normal(0, 0.5, n)
X = np.column_stack([x1, x2])
p = X.shape[1]

py_coef = np.linalg.lstsq(X, y, rcond=None)[0]
py_r2 = 1 - np.sum((y - X @ py_coef)**2) / np.sum((y - np.mean(y))**2)
print(f"Python: coef={py_coef}, R2={py_r2:.6f}")

MODULE_PATHS = [os.path.join(VBA_CORE_DIR, name + ".cls") for name in VBA_CORE_IMPORT_ORDER]
MODULE_PATHS += [os.path.join(SRC_DIR, f + ".bas") for f in ["LinearUtils", "StatsUtils", "RegressUtils"]]
excel = ensure_excel()
wb_path = os.path.join(os.path.dirname(__file__), "_debug_fitols.xlsm")
try:
    wb = create_workbook(excel, wb_path, MODULE_PATHS)
    ws = wb.Sheets("TestData")

    full = np.column_stack([X, y])
    nr, nc = full.shape
    write_range(ws, full, 1, 1)
    rng_X = ws.Range(ws.Cells(1, 1), ws.Cells(nr, p))
    rng_y = ws.Range(ws.Cells(1, p + 1), ws.Cells(nr, p + 1))

    # Test 1: Range path
    print("\n--- Range path ---")
    r = run_macro(excel, wb, "RegressUtils.FitOLS", rng_X, rng_y)
    vba_n = float(r.Item("n")); vba_p = float(r.Item("p"))
    vba_sse = float(r.Item("sse")); vba_r2 = float(r.Item("r_squared"))
    c = r.Item("coefficients")
    vba_coef = np.array([float(c[i]) for i in range(len(c))])
    print(f"  n={vba_n} p={vba_p} R2={vba_r2:.8f} SSE={vba_sse:.6f} coef={vba_coef}")

    # Test 2: Array path (Python lists)
    print("\n--- Array path ---")
    r2 = run_macro(excel, wb, "RegressUtils.FitOLS", X.tolist(), y.tolist())
    vba_r2a = float(r2.Item("r_squared"))
    vba_ssea = float(r2.Item("sse"))
    c2 = r2.Item("coefficients")
    vba_coef2 = np.array([float(c2[i]) for i in range(len(c2))])
    print(f"  R2={vba_r2a:.8f} SSE={vba_ssea:.6f} coef={vba_coef2}")

    if np.max(np.abs(vba_coef - vba_coef2)) < 1e-10:
        print("\nVerdict: Range == Array (IDENTICAL)")
    else:
        print(f"\nVerdict: Range != Array (coef diff = {np.max(np.abs(vba_coef - vba_coef2)):.2e})")
finally:
    teardown(excel, wb)
    if os.path.exists(wb_path):
        try: os.remove(wb_path)
        except OSError: pass
