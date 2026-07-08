# install.ps1 —— 把 WM_TOOLS 加载项安装到本机 WPS（离线模式，无需服务器常驻）
#
# 用法：
#   双击「安装到WPS.bat」
#   或 PowerShell:  powershell -ExecutionPolicy Bypass -File install.ps1
#   卸载:           powershell -ExecutionPolicy Bypass -File install.ps1 -Uninstall
#
# 原理：把加载项文件复制到 %APPDATA%\kingsoft\wps\jsaddons\<name>_<version>\，
#       并在 publish.xml 写一条离线记录。装一次即常驻，不依赖任何后台服务。

param(
    [switch]$Uninstall
)

# 让 PowerShell 5.1 正确显示中文
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$ErrorActionPreference = 'Stop'
$ProjectRoot = $PSScriptRoot
if (-not $ProjectRoot) { $ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }

# ---- 读取插件信息（来自 package.json，避免硬编码）----
$pkgPath = Join-Path $ProjectRoot 'package.json'
if (-not (Test-Path $pkgPath)) {
    Write-Host "[错误] 未在 $ProjectRoot 找到 package.json，请确认脚本位于项目根目录。" -ForegroundColor Red
    exit 1
}
$pkg = Get-Content $pkgPath -Raw -Encoding UTF8 | ConvertFrom-Json
$AddonName    = $pkg.name                 # wm-tools
$AddonType    = $pkg.addonType            # wps
$Version      = $pkg.version              # 1.0.0
$AddonDirName = "${AddonName}_${Version}" # wm-tools_1.0.0

$JsaddonsRoot = Join-Path $env:APPDATA 'kingsoft\wps\jsaddons'
$PublishXml   = Join-Path $JsaddonsRoot 'publish.xml'
$AddonDest    = Join-Path $JsaddonsRoot $AddonDirName

function Write-Step($m) { Write-Host "-> $m" -ForegroundColor Cyan }
function Write-Ok($m)   { Write-Host "   [OK] $m" -ForegroundColor Green }
function Write-Warn($m) { Write-Host "   [!] $m" -ForegroundColor Yellow }

# 从 publish.xml 移除指定 name 的所有 jsplugin / jspluginonline 记录
function Remove-AddonRecords {
    param($XmlDoc, $Name)
    $root = $XmlDoc.DocumentElement
    if (-not $root) { return }
    foreach ($tag in @('jsplugin', 'jspluginonline')) {
        # 先固化成数组，避免边遍历边删除
        $nodes = @($root.GetElementsByTagName($tag))
        foreach ($n in $nodes) {
            if ($n.GetAttribute('name') -eq $Name) {
                $root.RemoveChild($n) | Out-Null
            }
        }
    }
}

# =====================================================================
# 卸载
# =====================================================================
if ($Uninstall) {
    Write-Host "卸载 $AddonName ..." -ForegroundColor Yellow

    if (Test-Path $PublishXml) {
        try {
            $xml = New-Object System.Xml.XmlDocument
            $xml.Load($PublishXml)
            Remove-AddonRecords -XmlDoc $xml -Name $AddonName
            $xml.Save($PublishXml)
            Write-Ok "已从 publish.xml 移除 $AddonName 记录"
        } catch {
            Write-Warn "更新 publish.xml 失败：$($_.Exception.Message)"
        }
    }

    if (Test-Path $AddonDest) {
        Remove-Item $AddonDest -Recurse -Force
        Write-Ok "已删除 $AddonDest"
    } else {
        Write-Warn "插件目录不存在：$AddonDest"
    }

    Write-Host "`n卸载完成。请【完全关闭所有 WPS 窗口】后重新打开 WPS。" -ForegroundColor Green
    exit 0
}

# =====================================================================
# 安装
# =====================================================================
Write-Host "安装 $AddonName (v$Version) 到 WPS —— 离线模式（装一次即常驻）`n" -ForegroundColor Cyan

Write-Step "准备目录 $JsaddonsRoot"
New-Item -ItemType Directory -Force -Path $JsaddonsRoot | Out-Null
Write-Ok "就绪"

# 1) 复制加载项文件（白名单，只拷运行所需文件，不含 node_modules 等）
Write-Step "复制加载项文件 -> $AddonDest"
if (Test-Path $AddonDest) { Remove-Item $AddonDest -Recurse -Force }
New-Item -ItemType Directory -Force -Path $AddonDest | Out-Null

$keep = @('index.html', 'main.js', 'manifest.xml', 'ribbon.xml', 'js', 'ui', 'images')
foreach ($item in $keep) {
    $src = Join-Path $ProjectRoot $item
    if (Test-Path $src) {
        Copy-Item -Path $src -Destination $AddonDest -Recurse -Force
    } else {
        Write-Warn "缺少 $item（可能影响加载项运行）"
    }
}
Write-Ok "文件已就位（$((Get-ChildItem $AddonDest -Recurse -File | Measure-Object).Count) 个文件）"

# 2) 更新 publish.xml：移除同名旧记录（含失效的在线调试记录），追加离线记录
Write-Step "更新 $PublishXml"
$xml = New-Object System.Xml.XmlDocument
$existed = $false
if (Test-Path $PublishXml) {
    try {
        $bak = "$PublishXml.bak"
        if (-not (Test-Path $bak)) { Copy-Item $PublishXml $bak -Force }
        $xml.Load($PublishXml)
        $existed = $true
    } catch {
        Write-Warn "现有 publish.xml 解析失败，将重建：$($_.Exception.Message)"
        $xml = New-Object System.Xml.XmlDocument
        $r = $xml.CreateElement('jsplugins'); $xml.AppendChild($r) | Out-Null
    }
} else {
    $r = $xml.CreateElement('jsplugins'); $xml.AppendChild($r) | Out-Null
}
if (-not $xml.DocumentElement) {
    $r = $xml.CreateElement('jsplugins'); $xml.AppendChild($r) | Out-Null
}

Remove-AddonRecords -XmlDoc $xml -Name $AddonName

$node = $xml.CreateElement('jsplugin')
$node.SetAttribute('name',         $AddonName)
$node.SetAttribute('type',         $AddonType)
$node.SetAttribute('url',          $AddonDirName)
$node.SetAttribute('version',      $Version)
$node.SetAttribute('enable',       'enable_dev')
$node.SetAttribute('install',      'null')
$node.SetAttribute('customDomain', '')
$xml.DocumentElement.AppendChild($node) | Out-Null

# 用 UTF-8（带 BOM）保存，WPS 能正确识别
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($PublishXml, $xml.OuterXml, $utf8Bom)
Write-Ok "publish.xml 已更新（离线记录 url=$AddonDirName）"

# 3) 完成
Write-Host ""
Write-Host "================================================" -ForegroundColor Green
Write-Host " 安装完成！" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green
Write-Host ""
Write-Host "下一步：请【完全关闭所有 WPS 窗口】，然后重新打开 WPS。" -ForegroundColor Cyan
Write-Host "      功能区应出现「WM_TOOLS」标签 -> 「清理并另存为…」按钮。" -ForegroundColor Cyan
Write-Host ""
Write-Host "若没出现，在 WPS 顶部菜单找「开发工具 -> WPS加载项」确认 $AddonName 已启用。" -ForegroundColor DarkGray
Write-Host ""
