# rebuild.ps1 — Rebuild VBA_Libraries.xlsm/.xlam from source
# Requires: Excel installed (COM automation)
#
# Usage:
#   .\scripts\rebuild.ps1                # Rebuild both .xlsm and .xlam
#   .\scripts\rebuild.ps1 -Version 1.2.0 # Inject version number
#   .\scripts\rebuild.ps1 -DryRun        # Show what would be done

param(
    [string]$Version = "",
    [switch]$DryRun
)

$Root = Split-Path -Parent $PSScriptRoot
$SrcDir = Join-Path $Root "src"
$CoreDir = Join-Path $Root "VBA-Core"
$DocsDir = Join-Path $Root "docs"
$OutputXlsm = Join-Path $DocsDir "VBA_Libraries.xlsm"
$OutputXlam = Join-Path $DocsDir "VBA_Libraries.xlam"

# Import order matters: VBA-Core first, then modules
$importOrder = @(
    # VBA-Core (classes)
    (Join-Path $CoreDir "VariantKit.cls"),
    (Join-Path $CoreDir "ArrayOps.cls"),
    (Join-Path $CoreDir "DictProxy.cls"),
    # src modules (order-independent except RegressUtils)
    (Join-Path $SrcDir "ArrayUtils.bas"),
    (Join-Path $SrcDir "DictSetUtils.bas"),
    (Join-Path $SrcDir "PivotUtils.bas"),
    (Join-Path $SrcDir "SqlUtils.bas"),
    (Join-Path $SrcDir "LinearUtils.bas"),
    (Join-Path $SrcDir "StatsUtils.bas"),
    (Join-Path $SrcDir "RegressUtils.bas"),  # depends on Linear + Stats
    (Join-Path $SrcDir "StringUtils.bas"),
    (Join-Path $SrcDir "RegexUtils.bas"),
    (Join-Path $SrcDir "JsonUtils.bas"),
    (Join-Path $SrcDir "XmlUtils.bas"),
    (Join-Path $SrcDir "DateTimeUtils.bas"),
    (Join-Path $SrcDir "RangeUtils.bas"),
    (Join-Path $SrcDir "FileSystemUtils.bas"),
    (Join-Path $SrcDir "PhyChemUtils.bas")
)

Write-Host "Excel-VBA-Libraries — Rebuild Script" -ForegroundColor White
Write-Host "Root: $Root" -ForegroundColor Gray

if ($Version) {
    Write-Host "Version: $Version" -ForegroundColor Cyan
}

if ($DryRun) {
    Write-Host "`n[DRY RUN] Would import these files:" -ForegroundColor Yellow
    $importOrder | ForEach-Object { Write-Host "  $_" }
    Write-Host "`n[DRY RUN] Output:" -ForegroundColor Yellow
    Write-Host "  $OutputXlsm"
    Write-Host "  $OutputXlam"
    exit 0
}

# Verify source files exist
$missing = $importOrder | Where-Object { -not (Test-Path $_) }
if ($missing) {
    Write-Host "ERROR: Missing source files:" -ForegroundColor Red
    $missing | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    exit 1
}

Write-Host "`nImporting $($importOrder.Count) modules into Excel..." -ForegroundColor Cyan

try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false

    # Create .xlsm
    $wb = $xl.Workbooks.Add()

    # Remove default sheets except one
    while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }

    # Import VBA modules
    foreach ($file in $importOrder) {
        $name = Split-Path $file -Leaf
        Write-Host "  Importing: $name" -ForegroundColor Gray
        $wb.VBProject.VBComponents.Import($file) | Out-Null
    }

    # Inject version if specified
    if ($Version) {
        $verModule = $wb.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
        $verModule.Name = "modVersion"
        $verModule.CodeModule.AddFromString("Public Const LIB_VERSION As String = `"$Version`"")
    }

    # Save as .xlsm
    if (Test-Path $OutputXlsm) { Remove-Item $OutputXlsm -Force }
    $wb.SaveAs($OutputXlsm, 52)  # xlOpenXMLWorkbookMacroEnabled
    Write-Host "`n  Saved: $OutputXlsm" -ForegroundColor Green

    # Save as .xlam
    if (Test-Path $OutputXlam) { Remove-Item $OutputXlam -Force }
    $wb.SaveAs($OutputXlam, 55)  # xlOpenXMLAddIn
    Write-Host "  Saved: $OutputXlam" -ForegroundColor Green

    $wb.Close($false)
    Write-Host "`nRebuild complete!" -ForegroundColor Green

} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Note: Ensure 'Trust access to VBA project object model' is enabled in Excel." -ForegroundColor Yellow
    exit 1
} finally {
    if ($xl) {
        $xl.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    }
}
