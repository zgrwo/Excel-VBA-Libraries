"""Performance benchmarks for core VBA operations.

Measures end-to-end COM → VBA execution time for:
  - ArraySort (QuickSort, median-of-three pivot)
  - MatrixMultiply (blocked/tiled O(n³))
  - MatrixMultiplyNaive (triple-loop O(n³))
  - ArraySum, MatrixTranspose, MatrixDeterminant

Usage:
  python tests/benchmarks/bench_core_ops.py           # all benchmarks
  python tests/benchmarks/bench_core_ops.py --quick   # smaller sizes, fewer iterations
"""

import os
import sys
import time
import statistics
import json
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from tests.test_utils import (
    ensure_excel, teardown, create_workbook, inject_testrunner,
    write_range, run_macro,
    SRC_DIR, VBA_CORE_DIR, VBA_CORE_IMPORT_ORDER,
)

# Module paths for import order
CORE_MODULES = [os.path.join(VBA_CORE_DIR, n + ".cls") for n in VBA_CORE_IMPORT_ORDER]
BENCH_MODULES = CORE_MODULES + [
    os.path.join(SRC_DIR, "ArrayUtils.bas"),
    os.path.join(SRC_DIR, "LinearUtils.bas"),
]


# ═══════════════════════════════════════════════════════════════════════════════
# Benchmark Runner
# ═══════════════════════════════════════════════════════════════════════════════

class BenchRunner:
    """Times VBA function calls via COM with warmup + multiple iterations."""

    def __init__(self, excel, wb, ws):
        self.excel = excel
        self.wb = wb
        self.ws = ws
        self.results = []  # list of dicts

    # ── timing helpers ──────────────────────────────────────────────────────

    def time_call(self, macro, *args, warmup=2, iterations=5):
        """Time a direct COM → VBA call (Python objects → Variant args)."""
        for _ in range(warmup):
            run_macro(self.excel, self.wb, macro, *args)

        times = []
        for _ in range(iterations):
            t0 = time.perf_counter()
            run_macro(self.excel, self.wb, macro, *args)
            times.append(time.perf_counter() - t0)

        return self._stats(times, iterations)

    def time_range_call(self, macro, data_2d, *extra_args,
                        warmup=2, iterations=5):
        """Time a VBA call where the first arg is written to a Range."""
        self.ws.UsedRange.ClearContents()
        arr = np.atleast_2d(np.asarray(data_2d, dtype=object))
        nr, nc = arr.shape
        write_range(self.ws, arr, 1, 1)
        rng = self.ws.Range(self.ws.Cells(1, 1), self.ws.Cells(nr, nc))

        for _ in range(warmup):
            run_macro(self.excel, self.wb, macro, rng, *extra_args)

        times = []
        for _ in range(iterations):
            t0 = time.perf_counter()
            run_macro(self.excel, self.wb, macro, rng, *extra_args)
            times.append(time.perf_counter() - t0)

        return self._stats(times, iterations)

    def time_matrix_op(self, macro, a, b, *extra_args,
                       warmup=2, iterations=5):
        """Time a two-matrix VBA call (both args are 2D lists)."""
        for _ in range(warmup):
            run_macro(self.excel, self.wb, macro, a, b, *extra_args)

        times = []
        for _ in range(iterations):
            t0 = time.perf_counter()
            run_macro(self.excel, self.wb, macro, a, b, *extra_args)
            times.append(time.perf_counter() - t0)

        return self._stats(times, iterations)

    @staticmethod
    def _stats(times, iterations):
        return {
            "min_s": round(min(times), 6),
            "median_s": round(statistics.median(times), 6),
            "max_s": round(max(times), 6),
            "std_s": round(statistics.stdev(times), 6) if iterations > 1 else 0.0,
            "iterations": iterations,
        }

    # ── scale benchmarks ────────────────────────────────────────────────────

    def bench_scale(self, label, macro, gen_args_fn, sizes, *,
                    mode="direct", warmup=2, iterations=5):
        """Run same benchmark at multiple data sizes.

        Args:
            label: prefix for result labels
            macro: "Module.Function" name
            gen_args_fn: callable(n) → tuple of args
            sizes: list of sizes to test
            mode: "direct" (raw args) | "range" (Range arg) | "matrix" (two 2D lists)
        """
        for n in sizes:
            args = gen_args_fn(n)
            if mode == "range":
                stats = self.time_range_call(macro, args[0], *args[1:],
                                             warmup=warmup, iterations=iterations)
            elif mode == "matrix":
                stats = self.time_matrix_op(macro, args[0], args[1], *args[2:],
                                            warmup=warmup, iterations=iterations)
            else:
                stats = self.time_call(macro, *args,
                                       warmup=warmup, iterations=iterations)
            stats["label"] = f"{label}_n={n}"
            stats["size"] = n
            self.results.append(stats)
            print(f"  {label:30s} n={n:>6d}  median={stats['median_s']:>9.4f}s  "
                  f"min={stats['min_s']:>9.4f}s")

    # ── reporting ───────────────────────────────────────────────────────────

    def print_report(self):
        print("\n" + "=" * 78)
        print("PERFORMANCE BENCHMARK REPORT")
        print("=" * 78)
        print(f"{'Benchmark':<36s} {'Size':>6s}  {'Median(s)':>10s}  "
              f"{'Min(s)':>10s}  {'Std(s)':>8s}  {'N':>3s}")
        print("-" * 78)
        for r in self.results:
            print(f"{r['label']:<36s} {str(r.get('size','-')):>6s}  "
                  f"{r['median_s']:>10.4f}  {r['min_s']:>10.4f}  "
                  f"{r['std_s']:>8.4f}  {r['iterations']:>3d}")
        print("-" * 78)
        print(f"Total benchmarks: {len(self.results)}")
        print("=" * 78)

    def to_json(self, path=None):
        """Serialize results to JSON. If path is given, write to file."""
        data = {"benchmarks": self.results, "total": len(self.results)}
        if path:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        return json.dumps(data, indent=2)


# ═══════════════════════════════════════════════════════════════════════════════
# Data Generators
# ═══════════════════════════════════════════════════════════════════════════════

def gen_1d_random(n):
    """Generate 1D list of random floats for ArraySort/ArraySum."""
    rng = np.random.default_rng(42)
    return (rng.random(n).tolist(),)

def gen_1d_sorted_asc(n):
    """Already-sorted ascending list (best case for QuickSort)."""
    return (list(range(n)),)

def gen_1d_sorted_desc(n):
    """Reverse-sorted list (worst case without median-of-three)."""
    return (list(range(n, 0, -1)),)

def gen_square_matrices(n):
    """Generate two n×n random matrices for MatrixMultiply."""
    rng = np.random.default_rng(42)
    a = rng.random((n, n)).tolist()
    b = rng.random((n, n)).tolist()
    return (a, b)

def gen_single_square(n):
    """Generate one n×n random matrix for MatrixTranspose/MatrixDeterminant."""
    rng = np.random.default_rng(42)
    return (rng.random((n, n)).tolist(),)


# ═══════════════════════════════════════════════════════════════════════════════
# Main
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    quick = "--quick" in sys.argv

    # ── size config ─────────────────────────────────────────────────────────
    if quick:
        sort_sizes = [100, 500, 2000, 5000]
        mat_sizes = [10, 50, 100]
        mat_naive_sizes = [10, 50]
        large_1d_sizes = [1000, 10000, 50000]
        warmup, iters = 1, 3
    else:
        sort_sizes = [100, 500, 1000, 5000, 10000, 50000]
        mat_sizes = [10, 50, 100, 200]
        mat_naive_sizes = [10, 50, 100]
        large_1d_sizes = [1000, 10000, 100000, 500000]
        warmup, iters = 2, 5

    print("=" * 78)
    print("VBA Core Operations — Performance Benchmarks")
    print(f"Mode: {'QUICK' if quick else 'FULL'}  "
          f"warmup={warmup}  iterations={iters}")
    print("=" * 78)

    # ── setup ───────────────────────────────────────────────────────────────
    excel = ensure_excel(visible=False)
    wb = None
    try:
        import tempfile
        output = os.path.join(tempfile.gettempdir(), "vba_bench_core_ops.xlsm")
        wb = create_workbook(excel, output, BENCH_MODULES)
        inject_testrunner(wb)
        ws = wb.Sheets("TestData")
        runner = BenchRunner(excel, wb, ws)

        # ── ArraySort ───────────────────────────────────────────────────────
        print("\n-- ArraySort (QuickSort, median-of-three pivot) --")
        runner.bench_scale("ArraySort_random", "ArrayUtils.ArraySort",
                           gen_1d_random, sort_sizes, mode="direct",
                           warmup=warmup, iterations=iters)
        runner.bench_scale("ArraySort_asc", "ArrayUtils.ArraySort",
                           gen_1d_sorted_asc, sort_sizes, mode="direct",
                           warmup=warmup, iterations=iters)
        runner.bench_scale("ArraySort_desc", "ArrayUtils.ArraySort",
                           gen_1d_sorted_desc, sort_sizes, mode="direct",
                           warmup=warmup, iterations=iters)

        # ── ArraySum ────────────────────────────────────────────────────────
        print("\n-- ArraySum (linear scan) --")
        runner.bench_scale("ArraySum", "ArrayUtils.ArraySum",
                           gen_1d_random, large_1d_sizes, mode="direct",
                           warmup=warmup, iterations=iters)

        # ── MatrixMultiply (blocked) ────────────────────────────────────────
        print("\n-- MatrixMultiply (blocked O(n^3), blockSize=32) --")
        runner.bench_scale("MatMul_blocked", "LinearUtils.MatrixMultiply",
                           gen_square_matrices, mat_sizes, mode="matrix",
                           warmup=warmup, iterations=iters)

        # ── MatrixMultiplyNaive (triple loop) ───────────────────────────────
        print("\n-- MatrixMultiplyNaive (triple-loop O(n^3)) --")
        runner.bench_scale("MatMul_naive", "LinearUtils.MatrixMultiplyNaive",
                           gen_square_matrices, mat_naive_sizes, mode="matrix",
                           warmup=warmup, iterations=iters)

        # ── MatrixTranspose ─────────────────────────────────────────────────
        print("\n-- MatrixTranspose --")
        runner.bench_scale("MatTranspose", "LinearUtils.MatrixTranspose",
                           gen_single_square, mat_sizes, mode="direct",
                           warmup=warmup, iterations=iters)

        # ── MatrixDeterminant ───────────────────────────────────────────────
        print("\n-- MatrixDeterminant (LU decomposition) --")
        runner.bench_scale("MatDet", "LinearUtils.MatrixDeterminant",
                           gen_single_square, mat_sizes, mode="direct",
                           warmup=warmup, iterations=iters)

        # ── report ─────────────────────────────────────────────────────────
        runner.print_report()

        # Also write JSON for CI / trend tracking
        json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                 "bench_results.json")
        runner.to_json(json_path)
        print(f"\nResults saved to: {json_path}")

    finally:
        teardown(excel, wb)
        # Clean up temp file
        if wb is not None:
            import tempfile
            tmp = os.path.join(tempfile.gettempdir(), "vba_bench_core_ops.xlsm")
            if os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except OSError:
                    pass


if __name__ == "__main__":
    sys.exit(main())
