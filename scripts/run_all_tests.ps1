# run_all_tests.ps1 — One-click test runner for Excel-VBA-Libraries
# Integrates all 4 test layers into a single command.
#
# Usage:
#   .\scripts\run_all_tests.ps1           # Run all layers
#   .\scripts\run_all_tests.ps1 -Quick    # Only Layer 1 (no Excel needed)
#   .\scripts\run_all_tests.ps1 -Lint     # Only lint check

param(
    [switch]$Quick,
    [switch]$Lint
)

$ErrorActionPreference = "Continue"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

$passed = 0
$failed = 0

function Run-Step {
    param([string]$Name, [string]$Command)
    Write-Host "`n$('='*60)" -ForegroundColor Cyan
    Write-Host "  $Name" -ForegroundColor Cyan
    Write-Host "$('='*60)" -ForegroundColor Cyan
    Invoke-Expression $Command
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  [PASS] $Name" -ForegroundColor Green
        $script:passed++
    } else {
        Write-Host "  [FAIL] $Name (exit code: $LASTEXITCODE)" -ForegroundColor Red
        $script:failed++
    }
}

Write-Host "Excel-VBA-Libraries — Test Runner" -ForegroundColor White
Write-Host "Working directory: $Root" -ForegroundColor Gray

# Layer 0: Static Lint
if ($Lint) {
    Run-Step "Layer 0: VBA Lint" "python scripts/vba_lint.py"
    Write-Host "`nDone (lint only)."
    exit ($failed -gt 0 ? 1 : 0)
}

# Layer 1: Documentation Consistency (no Excel needed)
if ($Quick) {
    Run-Step "Layer 1: Validation (quick)" "python tests/run_all_validation.py --quick"
} else {
    Run-Step "Layer 1: Validation (full)" "python tests/run_all_validation.py"
}

# Layer 0: Lint (always run)
Run-Step "Layer 0: VBA Lint" "python scripts/vba_lint.py --summary"

if (-not $Quick) {
    # Layer 2: Cross-validation (requires Excel)
    Write-Host "`n  Note: Layers 2-3 require Excel. Skipping if not available." -ForegroundColor Yellow
    Run-Step "Layer 2: Cross-validation" "python tests/run_all_crossval.py"

    # Layer 3: Integration tests (requires Excel)
    Run-Step "Layer 3: Integration tests" "python tests/utils/integration_test_all_modules.py"
}

# Summary
Write-Host "`n$('='*60)" -ForegroundColor White
Write-Host "  SUMMARY: $passed passed, $failed failed" -ForegroundColor $(if ($failed -eq 0) { "Green" } else { "Red" })
Write-Host "$('='*60)" -ForegroundColor White

exit ($failed -gt 0 ? 1 : 0)
