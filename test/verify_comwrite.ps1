# verify_comwrite.ps1 - Verify the COM-based .md write path (content preserved + UTF-8).
# Mirrors what writeMd's COM fallback will do: Documents.Add -> Content.Text = md -> SaveAs2(md, 2, ..., 65001).
$ErrorActionPreference = "Stop"
$root = "D:\Claude\WPS_WM_Tools\test"
$sample = Join-Path $root "md_sample.txt"
$mdOut = Join-Path $root "md_comwrite_out.md"
Remove-Item $mdOut -ErrorAction SilentlyContinue

# read the sample md content (UTF-8 file written by the Write tool)
$mdText = [System.IO.File]::ReadAllText($sample, [System.Text.Encoding]::UTF8)
Write-Host ("sample md length: " + $mdText.Length + " chars")

$app = New-Object -ComObject kwps.Application
try {
    try { $app.Visible = $false } catch {}
    try { $app.DisplayAlerts = 0 } catch {}
    try { $app.ScreenUpdating = $false } catch {}
    $newDoc = $app.Documents.Add()
    $newDoc.Content.Text = $mdText
    # SaveAs2: FileFormat=2 (wdFormatText), Encoding (12th arg) = 65001 (UTF-8)
    $newDoc.SaveAs2($mdOut, 2, $false, "", $false, "", $false, $false, $false, $false, $false, 65001)
    $newDoc.Close(0)
    try { $app.ScreenUpdating = $true } catch {}
} finally {
    try { $app.Quit() } catch {}
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
}

# read back as UTF-8
$readBack = [System.IO.File]::ReadAllText($mdOut, [System.Text.Encoding]::UTF8)
Write-Host ("=== .md read back: " + $readBack.Length + " chars ===")
$n = [Math]::Min(300, $readBack.Length)
Write-Host "=== first $n chars ==="
Write-Host $readBack.Substring(0, $n)
Write-Host "=== content checks (md syntax must survive Word untouched) ==="
Write-Host ("  has '# 测试标题'     : " + $readBack.Contains("# 测试标题"))
Write-Host ("  has '**加粗**'       : " + $readBack.Contains("**加粗**"))
Write-Host ("  has '- 列表项一'     : " + $readBack.Contains("- 列表项一"))
Write-Host ("  has '| 列A | 列B |'  : " + $readBack.Contains("| 列A | 列B |"))
$bytes = [System.IO.File]::ReadAllBytes($mdOut)
$hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
Write-Host ("=== has UTF-8 BOM: " + $hasBom + " ===")
