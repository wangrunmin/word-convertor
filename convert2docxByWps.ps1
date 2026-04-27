<#
.SYNOPSIS
    批量将 Office 文档转换为 .docx（Worker Pool 并行版）

.PARAMETER InputDir      输入目录
.PARAMETER OutputDir     输出目录
.PARAMETER Recurse       是否递归（默认开启）
.PARAMETER ThrottleLimit 并行 Worker 数（默认 4）
.PARAMETER LogFile       日志路径
#>

param(
    [string]$InputPath,

    [string]$OutputDir,

    [switch]$Recurse = $true,

    [int]$ThrottleLimit = 4,

    # $null = 交互询问；$true = 强制清理；$false = 强制保留
    [Nullable[bool]]$CleanTemp = $null,

    [string]$LogFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ─── 加载原子计数器（C#）────────────────────────────────────────────────
Add-Type -TypeDefinition @"
using System.Threading;
public class AtomicCounter {
    private int _value;
    public int Value { get { return _value; } }
    public int Increment() { return Interlocked.Increment(ref _value); }
}
"@ -ErrorAction SilentlyContinue

# ─── PDF 类型识别（可选能力）─────────────────────────────────────────────

$script:PdfAnalysisAvailable = $false

function Initialize-PdfAnalysis {
    $libDir    = Join-Path $PSScriptRoot "lib"
    $marker    = Join-Path $libDir ".pdfpig_loaded"
    $nugetExe  = Join-Path $libDir "nuget.exe"

    # 缓存命中：直接加载本地 DLL
    if (Test-Path $marker) {
        try {
            Get-ChildItem -Path $libDir -Filter "*.dll" | ForEach-Object {
                Add-Type -Path $_.FullName -ErrorAction Stop
            }
            # 验证类型真的可用
            $null = [UglyToad.PdfPig.PdfDocument]
            $script:PdfAnalysisAvailable = $true
            return
        } catch {
            Remove-Item $marker -ErrorAction SilentlyContinue
            Get-ChildItem $libDir -Filter "*.dll" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        }
    }

    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor DarkGray
    Write-Host " 首次运行：尝试安装 PDF 类型识别库（可选，约 10MB）" -ForegroundColor DarkGray
    Write-Host " 成功后可区分 PDF 文字版/扫描版；失败不影响主流程" -ForegroundColor DarkGray
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor DarkGray

    try {
        New-Item -ItemType Directory -Force -Path $libDir | Out-Null

        # 1) 下载 nuget.exe 命令行工具（~5MB）
        if (-not (Test-Path $nugetExe)) {
            Write-Host "  [1/3] 下载 nuget.exe..." -NoNewline
            $url = "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe"
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri $url -OutFile $nugetExe -UseBasicParsing -ErrorAction Stop
            Write-Host " OK"
        } else {
            Write-Host "  [1/3] nuget.exe 已存在" -ForegroundColor DarkGray
        }

        # 2) 用 nuget.exe install 自动下载 PdfPig 及全部依赖
        Write-Host "  [2/3] 下载 PdfPig 及其依赖（nuget.exe 自动处理）..." -NoNewline
        $installDir = Join-Path $libDir "_packages"
        New-Item -ItemType Directory -Force -Path $installDir | Out-Null

        # -DependencyVersion Lowest 确保拿到稳定版本组合；-Framework net48 选择兼容 PS 5.1 的运行时
        $nugetArgs = @(
            "install", "PdfPig",
            "-OutputDirectory", $installDir,
            "-DependencyVersion", "Lowest",
            "-Source", "https://api.nuget.org/v3/index.json",
            "-NonInteractive",
            "-Verbosity", "quiet"
        )
        $proc = Start-Process -FilePath $nugetExe -ArgumentList $nugetArgs `
            -NoNewWindow -Wait -PassThru -RedirectStandardOutput "$libDir\nuget.log" -RedirectStandardError "$libDir\nuget.err"
        if ($proc.ExitCode -ne 0) {
            $errContent = Get-Content "$libDir\nuget.err" -Raw -ErrorAction SilentlyContinue
            throw "nuget.exe 退出码 $($proc.ExitCode): $errContent"
        }
        Write-Host " OK"

        # 3) 把所有包里最兼容的 DLL 拷贝出来
        Write-Host "  [3/3] 提取 DLL..." -NoNewline
        $preferredFrameworks = @("net48","net472","net471","net47","net462","net461","net46","net45","netstandard2.0","netstandard2.1","netstandard1.6","netstandard1.3")

        $pkgDirs = Get-ChildItem -Path $installDir -Directory
        $copiedCount = 0
        foreach ($pkg in $pkgDirs) {
            $libSub = Join-Path $pkg.FullName "lib"
            if (-not (Test-Path $libSub)) { continue }

            foreach ($fw in $preferredFrameworks) {
                $fwDir = Join-Path $libSub $fw
                if (Test-Path $fwDir) {
                    Get-ChildItem $fwDir -Filter "*.dll" | ForEach-Object {
                        Copy-Item $_.FullName $libDir -Force
                        $copiedCount++
                    }
                    break
                }
            }
        }
        Remove-Item $installDir -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item "$libDir\nuget.log", "$libDir\nuget.err" -ErrorAction SilentlyContinue
        Write-Host " OK ($copiedCount DLL)"

        if ($copiedCount -eq 0) { throw "未提取到任何 DLL" }

        # 加载并验证
        Get-ChildItem -Path $libDir -Filter "*.dll" | ForEach-Object {
            try { Add-Type -Path $_.FullName -ErrorAction Stop } catch {}
        }
        $null = [UglyToad.PdfPig.PdfDocument]   # 找不到会抛异常

        Set-Content -Path $marker -Value (Get-Date).ToString("s") -Encoding UTF8
        $script:PdfAnalysisAvailable = $true
        Write-Host "  ✓ PDF 类型识别已启用" -ForegroundColor Green
    } catch {
        Write-Host ""
        Write-Host "  ⚠ 安装失败：$($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "  继续运行（PDF 只会被标记，不做文字版/扫描版区分）" -ForegroundColor DarkGray
        $script:PdfAnalysisAvailable = $false
    }
    Write-Host ""
}

function Get-PdfType {
    param([string]$FilePath)
    if (-not $script:PdfAnalysisAvailable) { return "PDF" }

    try {
        $pdf = [UglyToad.PdfPig.PdfDocument]::Open($FilePath)
        try {
            $totalChars = 0
            $pageCount  = 0
            foreach ($page in $pdf.GetPages()) {
                $totalChars += $page.Text.Length
                $pageCount++
                if ($pageCount -ge 5) { break }   # 抽前 5 页足够判断
            }
            if ($pageCount -eq 0) { return "PDF" }
            $avg = $totalChars / $pageCount

            if ($avg -gt 100) { return "PDF(文字版)" }
            elseif ($avg -lt 10) { return "PDF(扫描版)" }
            else { return "PDF(混合)" }
        } finally {
            $pdf.Dispose()
        }
    } catch {
        return "PDF"
    }
}

# ─── PDF 文字版转换能力（pdf2docx，可选）─────────────────────────────────

$script:Pdf2DocxAvailable = $false
$script:PythonExe         = $null
$script:ConvertTextPdf    = $false

# ─── Excel 报告能力（ImportExcel，可选，首次使用自动尝试安装）───────────

$script:ExcelAvailable = $false

function Initialize-ExcelReport {
    # 已安装
    if (Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue) {
        try {
            Import-Module ImportExcel -ErrorAction Stop
            $script:ExcelAvailable = $true
            return
        } catch {
            # 装了但加载失败，静默降级
            return
        }
    }

    # 尝试静默安装（CurrentUser 作用域不需要管理员）
    Write-Host "  首次生成 Excel 报告，正在安装 ImportExcel 模块..." -NoNewline
    try {
        # 确保 PSGallery 可用
        if (-not (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue)) {
            Register-PSRepository -Default -ErrorAction SilentlyContinue
        }
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue

        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Import-Module ImportExcel -ErrorAction Stop
        $script:ExcelAvailable = $true
        Write-Host " OK" -ForegroundColor Green
    } catch {
        Write-Host " 失败（将回退到 CSV）" -ForegroundColor Yellow
    }
}

function Initialize-Pdf2Docx {
    # 找 Python
    $candidates = @("python", "python3", "py")
    foreach ($cmd in $candidates) {
        $py = Get-Command $cmd -ErrorAction SilentlyContinue
        if ($py) {
            $script:PythonExe = $py.Source
            break
        }
    }
    if (-not $script:PythonExe) {
        Write-Host "  ⚠ 未找到 Python 环境，跳过 PDF 文字提取" -ForegroundColor Yellow
        return
    }

    # 检查两个依赖：fitz (PyMuPDF) 和 docx (python-docx)
    & $script:PythonExe -c "import fitz, docx" 2>$null
    if ($LASTEXITCODE -eq 0) {
        $script:Pdf2DocxAvailable = $true
        Write-Host "  ✓ 检测到 PyMuPDF + python-docx，文字版 PDF 可转换" -ForegroundColor Green
        return
    }

    Write-Host ""
    Write-Host "  未检测到 PDF 文字提取所需的依赖 (PyMuPDF + python-docx)" -ForegroundColor DarkGray
    $ans = Read-Host "  是否现在安装？[Y/n]"
    if ($ans -match '^[nN]') {
        Write-Host "  已跳过（文字版 PDF 不会被转换）" -ForegroundColor DarkGray
        return
    }

    Write-Host "  正在安装 PyMuPDF 和 python-docx（约 30-60 秒）..." -NoNewline
    try {
        & $script:PythonExe -m pip install PyMuPDF python-docx --quiet 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) {
            & $script:PythonExe -c "import fitz, docx" 2>$null
            if ($LASTEXITCODE -eq 0) {
                $script:Pdf2DocxAvailable = $true
                Write-Host " OK" -ForegroundColor Green
            } else {
                Write-Host " 安装成功但导入失败" -ForegroundColor Yellow
            }
        } else {
            Write-Host " pip 安装失败" -ForegroundColor Yellow
        }
    } catch {
        Write-Host " 异常: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Convert-TextPdfToDocx {
    param([string]$PdfPath, [string]$DocxPath)
    if (-not $script:Pdf2DocxAvailable) {
        return @{ Success = $false; Error = "PyMuPDF + python-docx 不可用" }
    }

    $tmpScript = [System.IO.Path]::GetTempFileName() + ".py"
    $stderrFile = [System.IO.Path]::GetTempFileName()
    $pyCode = @"
# -*- coding: utf-8 -*-
import sys, traceback
pdf_path = sys.argv[1]
docx_path = sys.argv[2]
try:
    import fitz
    from docx import Document
    from docx.shared import Pt

    pdf = fitz.open(pdf_path)
    doc = Document()

    # 默认正文字体：宋体 + 小四号
    style = doc.styles["Normal"]
    style.font.name = "宋体"
    style.font.size = Pt(12)

    for page_num, page in enumerate(pdf, 1):
        # "text" 模式返回纯文本，按阅读顺序
        text = page.get_text("text")
        if not text.strip():
            # 该页可能是纯图（扫描页），加一行占位
            doc.add_paragraph(f"[第 {page_num} 页：未提取到文本]")
            continue
        for para in text.split("\n"):
            para = para.strip()
            if para:
                doc.add_paragraph(para)
        # 页间用空段落分隔
        if page_num < len(pdf):
            doc.add_paragraph("")

    pdf.close()
    doc.save(docx_path)
    sys.exit(0)
except Exception:
    traceback.print_exc(file=sys.stderr)
    sys.exit(1)
"@
    Set-Content -Path $tmpScript -Value $pyCode -Encoding UTF8

    $prevEAP = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        $cmdLine = '"{0}" "{1}" "{2}" "{3}" 2> "{4}"' -f $script:PythonExe, $tmpScript, $PdfPath, $DocxPath, $stderrFile
        cmd /c $cmdLine 2>&1 | Out-Null
        $exitCode = $LASTEXITCODE

        if ($exitCode -eq 0 -and (Test-Path $DocxPath)) {
            return @{ Success = $true; Error = "" }
        }

        $errRaw = Get-Content $stderrFile -Raw -ErrorAction SilentlyContinue
        $errMsg = if ($errRaw) { $errRaw.Trim() } else { "退出码 $exitCode" }
        if ($errMsg.Length -gt 400) { $errMsg = $errMsg.Substring(0, 400) + "..." }
        return @{ Success = $false; Error = $errMsg }
    } catch {
        return @{ Success = $false; Error = "PowerShell 异常: $($_.Exception.Message)" }
    } finally {
        $ErrorActionPreference = $prevEAP
        Remove-Item $tmpScript, $stderrFile -ErrorAction SilentlyContinue
    }
}

# ─── 友好启动引导 ─────────────────────────────────────────────────────────

function Show-Welcome {
    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║       WPS 批量转 docx 工具  (支持 doc/wps/ofd/rtf/txt)    ║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

function Read-PathFriendly {
    param([string]$Prompt, [switch]$MustExist, [switch]$AllowCreate)

    while ($true) {
        Write-Host ""
        Write-Host $Prompt -ForegroundColor Yellow
        Write-Host "  提示：可直接把文件/文件夹拖入窗口" -ForegroundColor DarkGray
        $input = Read-Host "  路径"
        if (-not $input) { continue }

        # 去掉拖入时常见的引号
        $input = $input.Trim().Trim('"').Trim("'")

        if ($MustExist) {
            if (Test-Path $input) { return $input }
            Write-Host "  路径不存在，请重新输入。" -ForegroundColor Red
            continue
        }
        if ($AllowCreate) {
            if (-not (Test-Path $input)) {
                $create = Read-Host "  目录不存在，是否创建？[Y/n]"
                if ($create -notmatch '^[nN]') {
                    New-Item -ItemType Directory -Force -Path $input | Out-Null
                    return $input
                }
                continue
            }
            return $input
        }
        return $input
    }
}

# 没传 InputPath/OutputDir 时走引导式输入
if (-not $InputPath) {
    Show-Welcome
    $InputPath = Read-PathFriendly -Prompt "请输入要转换的 [文件 / 压缩包(zip) / 文件夹] 路径：" -MustExist
}

# 尝试初始化 PDF 分析能力（可选）
Initialize-PdfAnalysis
Initialize-Pdf2Docx

if (-not $OutputDir) {
    $OutputDir = Read-PathFriendly -Prompt "请输入输出目录：" -AllowCreate
}

# ─── 输入归一化：自动识别文件/压缩包/目录 ─────────────────────────────────

Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

$InputPath    = (Resolve-Path $InputPath).Path
$inputItem    = Get-Item $InputPath
$tempExtract  = $null    # 如果是压缩包，记录解压目录，最后可能需要清理

if ($inputItem.PSIsContainer) {
    # 场景 A：目录，直接使用
    $InputDir = $InputPath
}
elseif ($inputItem.Extension -ieq ".zip") {
    # 场景 B：zip 压缩包 → 解压到临时目录
    $tempExtract = Join-Path $OutputDir ("_extract_" + [Guid]::NewGuid().ToString("N").Substring(0,8))
    New-Item -ItemType Directory -Force -Path $tempExtract | Out-Null
    Write-Host "检测到压缩包，正在解压到临时目录..." -NoNewline
    try {
        # 尝试按中文编码 (GBK/936) 解压，解决国内压缩包文件名乱码
        $zipEncoding = $null
        try {
            # .NET Framework 完整版支持；.NET Core/5+ 需要先注册 CodePagesEncodingProvider
            [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance)
        } catch {}
        try {
            $zipEncoding = [System.Text.Encoding]::GetEncoding(936)  # GBK
        } catch {
            $zipEncoding = [System.Text.Encoding]::Default
        }

        [System.IO.Compression.ZipFile]::ExtractToDirectory($InputPath, $tempExtract, $zipEncoding)
        Write-Host "`r解压完成：$tempExtract                        "
    } catch {
        Write-Host "`r解压失败：$_" -ForegroundColor Red
        Write-Host "提示：若压缩包使用非 GBK 编码（如 UTF-8 或其他），可先手动解压后使用解压目录作为输入。" -ForegroundColor Yellow
        exit 1
    }
    $InputDir = $tempExtract
    $Recurse  = [switch]$true   # 压缩包一律递归
}
else {
    # 场景 C：单个文档文件 → 伪造一个"目录"（把文件软链到一个临时目录，或直接用父目录 + 过滤）
    # 最简单：把它复制到临时目录当作单文件目录处理
    $tempExtract = Join-Path $OutputDir ("_single_" + [Guid]::NewGuid().ToString("N").Substring(0,8))
    New-Item -ItemType Directory -Force -Path $tempExtract | Out-Null
    Copy-Item $InputPath $tempExtract
    $InputDir = $tempExtract
    $Recurse  = [switch]$false  # 单文件没必要递归
    Write-Host "检测到单个文件，按单文件模式处理"
}

$extSet     = @(".doc", ".docx", ".wps", ".ofd", ".rtf", ".txt")
$startTime  = Get-Date
$reportTime = $startTime.ToString("yyyyMMdd_HHmmss")

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

if (-not $LogFile) { $LogFile = Join-Path $OutputDir "convert.log" }
$xlsxReport = Join-Path $OutputDir "convert_report_${reportTime}.xlsx"
$csvReport  = Join-Path $OutputDir "convert_report_${reportTime}.csv"
$htmlReport = Join-Path $OutputDir "convert_report_${reportTime}.html"
$failedList = Join-Path $OutputDir "failed_list_${reportTime}.txt"

$logMutex = New-Object System.Threading.Mutex($false, "WpsConvertLog")

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] [$Level] $Message"
    $logMutex.WaitOne() | Out-Null
    try { Add-Content -Path $LogFile -Value $line -Encoding UTF8 }
    finally { $logMutex.ReleaseMutex() }
}

# ─── 阶段 1：扫描文件 ─────────────────────────────────────────────────────

Write-Host "输入目录 : $InputDir"
Write-Host "输出目录 : $OutputDir"
Write-Host "并行数量 : $ThrottleLimit  |  递归: $($Recurse.IsPresent)"
Write-Host ""

$scanStart = Get-Date
Write-Host "正在扫描文件..." -NoNewline

$allFiles = [System.Collections.Generic.List[hashtable]]::new()
$seen     = @{}
$scanCount = 0

# PDF 单独列出，不进转换队列
$pdfFiles = [System.Collections.Generic.List[hashtable]]::new()

Get-ChildItem -Path $InputDir -File -Recurse:$Recurse.IsPresent -ErrorAction SilentlyContinue |
    Where-Object {
        $ext = $_.Extension.ToLower()
        (($extSet -contains $ext) -or ($ext -eq ".pdf")) -and ($_.Name -notmatch '^(\.~|~\$)')
    } |
    ForEach-Object {
        if ($Recurse.IsPresent) {
            $relativePath = $_.DirectoryName.Substring($InputDir.Length).TrimStart("\")
            $targetDir    = Join-Path $OutputDir $relativePath
        } else {
            $targetDir = $OutputDir
        }

        if ($_.Extension.ToLower() -eq ".pdf") {
            $pdfType = Get-PdfType -FilePath $_.FullName
            $pdfFiles.Add(@{
                FullName  = $_.FullName
                Name      = $_.Name
                BaseName  = $_.BaseName
                Extension = ".pdf"
                TargetDir = $targetDir
                PdfType   = $pdfType
            })
        } else {
            $outPath = Join-Path $targetDir ($_.BaseName + ".docx")
            if (-not $seen.ContainsKey($outPath)) {
                $seen[$outPath] = $true
                $allFiles.Add(@{
                    FullName  = $_.FullName
                    Name      = $_.Name
                    BaseName  = $_.BaseName
                    Extension = $_.Extension.ToLower()
                    TargetDir = $targetDir
                })
            }
        }

        $scanCount++
        if ($scanCount % 500 -eq 0) {
            Write-Host "`r正在扫描文件... 已找到 $($allFiles.Count) 个待转换 + $($pdfFiles.Count) 个 PDF（已扫描 $scanCount）" -NoNewline
        }
    }

$scanElapsed = (Get-Date) - $scanStart
$total = $allFiles.Count
Write-Host "`r扫描完成：$total 个待转换 + $($pdfFiles.Count) 个 PDF  (耗时 $([int]$scanElapsed.TotalSeconds)s)          "

# PDF 分类统计
$pdfTextCount    = @($pdfFiles | Where-Object { $_.PdfType -eq "PDF(文字版)" }).Count
$pdfScanCount    = @($pdfFiles | Where-Object { $_.PdfType -eq "PDF(扫描版)" }).Count
$pdfMixedCount   = @($pdfFiles | Where-Object { $_.PdfType -eq "PDF(混合)" }).Count
$pdfUnknownCount = @($pdfFiles | Where-Object { $_.PdfType -eq "PDF" }).Count

if ($pdfFiles.Count -gt 0) {
    Write-Host "  PDF 分类: 文字版 $pdfTextCount | 扫描版 $pdfScanCount | 混合 $pdfMixedCount | 未识别 $pdfUnknownCount" -ForegroundColor DarkGray
}

# 文字版 PDF 是否转换（交互询问）
if ($script:Pdf2DocxAvailable -and $pdfTextCount -gt 0) {
    Write-Host ""
    Write-Host "发现 $pdfTextCount 个文字版 PDF，可使用 pdf2docx 转换" -ForegroundColor Cyan
    Write-Host "  （扫描版和混合型不会转换，仅文字版）" -ForegroundColor DarkGray
    $ans = Read-Host "是否转换这些文字版 PDF？[Y/n]"
    if ($ans -notmatch '^[nN]') {
        $script:ConvertTextPdf = $true
        Write-Host "  将在 WPS 转换完成后处理文字版 PDF" -ForegroundColor DarkGray
    }
} elseif ($pdfFiles.Count -gt 0) {
    Write-Host "  注：PDF 文件不会被转换，仅记录在报告中" -ForegroundColor DarkGray
}
Write-Host ""

Write-Log "===== 开始转换 ====="
Write-Log "扫描完成: $total 个文件 ($([int]$scanElapsed.TotalSeconds)s)"

if ($total -eq 0) { Write-Host "没有需要转换的文件" -ForegroundColor Yellow; exit 0 }

# ─── 阶段 2：构建队列 + 原子计数器 ────────────────────────────────────────

$queue   = New-Object 'System.Collections.Concurrent.ConcurrentQueue[hashtable]'
$results = New-Object 'System.Collections.Concurrent.ConcurrentBag[hashtable]'
foreach ($f in $allFiles) { $queue.Enqueue($f) }

# 原子计数器（替代脆弱的 hashtable 自增）
$cntDone    = New-Object AtomicCounter
$cntSuccess = New-Object AtomicCounter
$cntSkipped = New-Object AtomicCounter
$cntFailed  = New-Object AtomicCounter

# 当前文件名用 Synchronized hashtable 存（直接赋值是原子的）
$syncCurrent = [System.Collections.Hashtable]::Synchronized(@{ Name = "" })

# ─── Worker 脚本块 ────────────────────────────────────────────────────────

$workerScript = {
    param($Queue, $Results, $SyncCurrent, $LogFile, $LogMutexName,
          $CntDone, $CntSuccess, $CntSkipped, $CntFailed)

    function Write-LogInner {
        param([string]$Message, [string]$Level = "INFO")
        $line = "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] [$Level] $Message"
        $m = [System.Threading.Mutex]::OpenExisting($LogMutexName)
        $m.WaitOne() | Out-Null
        try { Add-Content -Path $LogFile -Value $line -Encoding UTF8 }
        finally { $m.ReleaseMutex() }
    }

    function Get-RealFormatInner {
        param([string]$FilePath)
        try {
            $fs    = [System.IO.File]::OpenRead($FilePath)
            $bytes = New-Object byte[] 8
            $fs.Read($bytes, 0, 8) | Out-Null
            $fs.Close()
        } catch { return "UNKNOWN" }
        if ($bytes[0] -eq 0x50 -and $bytes[1] -eq 0x4B -and
            $bytes[2] -eq 0x03 -and $bytes[3] -eq 0x04) { return "OOXML" }
        if ($bytes[0] -eq 0xD0 -and $bytes[1] -eq 0xCF -and
            $bytes[2] -eq 0x11 -and $bytes[3] -eq 0xE0) { return "OLE2"  }
        if ($bytes[0] -eq 0x7B -and $bytes[1] -eq 0x5C -and
            $bytes[2] -eq 0x72 -and $bytes[3] -eq 0x74 -and
            $bytes[4] -eq 0x66)                          { return "RTF"   }
        return "UNKNOWN"
    }

    $wps = $null
    try {
        $type = [Type]::GetTypeFromProgID("KWps.Application")
        if ($null -eq $type) { $type = [Type]::GetTypeFromProgID("Wps.Application") }
        if ($null -eq $type) { throw "找不到 WPS COM 组件" }
        $wps = [Activator]::CreateInstance($type)
        $wps.Visible       = $false
        $wps.DisplayAlerts = 0
    } catch {
        Write-LogInner "Worker 启动失败: $_" "ERROR"
        return
    }

    try {
        $fileInfo = $null
        while ($Queue.TryDequeue([ref]$fileInfo)) {
            $fileStart = Get-Date
            $SyncCurrent['Name'] = $fileInfo.Name

            $outPath    = Join-Path $fileInfo.TargetDir ($fileInfo.BaseName + ".docx")
            $realFormat = Get-RealFormatInner -FilePath $fileInfo.FullName
            $extMismatch = ($fileInfo.Extension -ieq ".docx") -and ($realFormat -ne "OOXML")

            $result = @{
                FileName     = $fileInfo.Name
                SourcePath   = $fileInfo.FullName
                OutputPath   = $outPath
                SourceExt    = $fileInfo.Extension
                RealFormat   = $realFormat
                ExtMismatch  = $extMismatch
                Status       = ""
                ElapsedMs    = 0
                ErrorMessage = ""
            }

            try {
                if (Test-Path $outPath) {
                    $result.Status = "Skipped"
                    Write-LogInner "跳过（已存在）: $($fileInfo.Name)" "SKIP"
                    $CntSkipped.Increment() | Out-Null
                }
                elseif ($realFormat -eq "OOXML" -and -not $extMismatch) {
                    New-Item -ItemType Directory -Force -Path $fileInfo.TargetDir | Out-Null
                    Copy-Item $fileInfo.FullName $outPath -Force
                    $result.Status = "Skipped"
                    Write-LogInner "复制（已是 docx）: $($fileInfo.Name)" "SKIP"
                    $CntSkipped.Increment() | Out-Null
                }
                else {
                    if ($extMismatch) {
                        Write-LogInner "后缀.docx 实为 $realFormat，走 WPS 转换: $($fileInfo.Name)" "WARN"
                    }
                    New-Item -ItemType Directory -Force -Path $fileInfo.TargetDir | Out-Null

                    $doc = $null
                    try {
                        $doc = $wps.Documents.Open($fileInfo.FullName)
                        $doc.SaveAs($outPath, 12)
                        $result.Status = "Success"
                        Write-LogInner "OK: $($fileInfo.Name)" "OK"
                        $CntSuccess.Increment() | Out-Null
                    } finally {
                        if ($null -ne $doc) {
                            try { $doc.Close($false) } catch {}
                            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
                        }
                    }
                }
            } catch {
                $result.Status       = "Failed"
                $result.ErrorMessage = $_.ToString()
                Write-LogInner "失败: $($fileInfo.Name) | $_" "ERROR"
                if (Test-Path $outPath) {
                    Remove-Item $outPath -Force -ErrorAction SilentlyContinue
                }
                $CntFailed.Increment() | Out-Null
            } finally {
                $result.ElapsedMs = [int]((Get-Date) - $fileStart).TotalMilliseconds
                $Results.Add($result)
                $CntDone.Increment() | Out-Null
            }
        }
    } finally {
        if ($null -ne $wps) {
            try { $wps.Quit() } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wps) | Out-Null
        }
    }
}

# ─── 启动 Worker ──────────────────────────────────────────────────────────

$pool = [RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit)
$pool.ApartmentState = "STA"
$pool.Open()

$workers = 1..$ThrottleLimit | ForEach-Object {
    $ps = [PowerShell]::Create()
    $ps.RunspacePool = $pool
    $ps.AddScript($workerScript)     | Out-Null
    $ps.AddArgument($queue)          | Out-Null
    $ps.AddArgument($results)        | Out-Null
    $ps.AddArgument($syncCurrent)    | Out-Null
    $ps.AddArgument($LogFile)        | Out-Null
    $ps.AddArgument("WpsConvertLog") | Out-Null
    $ps.AddArgument($cntDone)        | Out-Null
    $ps.AddArgument($cntSuccess)     | Out-Null
    $ps.AddArgument($cntSkipped)     | Out-Null
    $ps.AddArgument($cntFailed)      | Out-Null
    @{ PS = $ps; Handle = $ps.BeginInvoke() }
}

# ─── 主线程：进度条 + 死锁检测 ────────────────────────────────────────────

$lastDone = 0
$stallSeconds = 0

while ($cntDone.Value -lt $total) {
    $done = $cntDone.Value
    $pct  = [Math]::Min(100, [int]($done / $total * 100))
    $elapsed = (Get-Date) - $startTime
    $speed = if ($elapsed.TotalSeconds -gt 1) { [int]($done / $elapsed.TotalSeconds * 60) } else { 0 }
    $eta = if ($speed -gt 0) {
        $s = [int](($total - $done) / $speed * 60)
        "{0:mm\:ss}" -f [TimeSpan]::FromSeconds($s)
    } else { "--:--" }

    $status = "完成 $done / $total  |  ✓$($cntSuccess.Value) ✗$($cntFailed.Value) ↷$($cntSkipped.Value)  |  $speed 文件/分  |  剩余 $eta"
    Write-Progress -Activity "转换中: $($syncCurrent.Name)" -Status $status -PercentComplete $pct

    # 死锁/停滞检测：60 秒没新进度说明所有 Worker 都卡住了
    if ($done -eq $lastDone) {
        $stallSeconds += 0.3
        if ($stallSeconds -gt 60 -and $done -lt $total) {
            Write-Progress -Activity "检测到停滞，强制退出" -Completed
            Write-Host ""
            Write-Host "⚠ 已 60 秒无新进度，所有 Worker 可能已卡死或崩溃。强制终止。" -ForegroundColor Yellow
            Write-Host "  已完成: $done / $total"
            break
        }
    } else {
        $stallSeconds = 0
        $lastDone = $done
    }

    # 所有 Worker 都退出了也要退出
    $aliveCount = ($workers | Where-Object { -not $_.Handle.IsCompleted }).Count
    if ($aliveCount -eq 0) {
        Write-Host ""
        Write-Host "所有 Worker 已退出，进入收尾阶段。" -ForegroundColor Cyan
        break
    }

    Start-Sleep -Milliseconds 300
}

Write-Progress -Activity "转换完成" -Completed

# 等 Worker 收尾
foreach ($w in $workers) {
    try { $w.PS.EndInvoke($w.Handle) | Out-Null } catch {}
    $w.PS.Dispose()
}
$pool.Close(); $pool.Dispose()

# ─── 汇总结果 ─────────────────────────────────────────────────────────────

# 如果所有 Worker 都没干活就退出，多半是 WPS 初始化失败
if ($cntDone.Value -eq 0 -and $total -gt 0) {
    Write-Host ""
    Write-Host "⚠ 没有任何文件被处理，所有 Worker 可能在启动 WPS 时失败了。" -ForegroundColor Yellow
    Write-Host "  请检查日志文件末尾的 ERROR 行: $LogFile" -ForegroundColor Yellow
    Write-Host "  常见原因：" -ForegroundColor Yellow
    Write-Host "    1. 上次运行的 WPS 进程残留 -> 任务管理器结束所有 wps.exe / wpp.exe / wpsoffice.exe" -ForegroundColor Yellow
    Write-Host "    2. WPS 被其他程序占用 -> 关闭所有 WPS 窗口" -ForegroundColor Yellow
    Write-Host "    3. COM 组件未注册 -> 以管理员身份启动一次 WPS 触发注册" -ForegroundColor Yellow
    Write-Host ""
}

Write-Host "正在生成报告..." -NoNewline

$allResults    = $results.ToArray()

# 文字版 PDF 转换
$pdfConvertedCount = 0
$pdfConvertFailedCount = 0
if ($script:ConvertTextPdf) {
    Write-Host ""
    Write-Host "正在转换文字版 PDF..." -ForegroundColor Cyan
    $pdfTextList = @($pdfFiles | Where-Object { $_.PdfType -eq "PDF(文字版)" })
    $idx = 0
    foreach ($pdf in $pdfTextList) {
        $idx++
        $pct = [int]($idx / $pdfTextList.Count * 100)
        Write-Progress -Activity "转换文字版 PDF" -Status "$idx / $($pdfTextList.Count) - $($pdf.Name)" -PercentComplete $pct

        New-Item -ItemType Directory -Force -Path $pdf.TargetDir | Out-Null
        $pdfOut = Join-Path $pdf.TargetDir ($pdf.BaseName + ".docx")

        if (Test-Path $pdfOut) {
            $pdf['ConvertStatus'] = "Skipped"
            $pdf['OutputPath']    = $pdfOut
            continue
        }

        $ret = Convert-TextPdfToDocx -PdfPath $pdf.FullName -DocxPath $pdfOut
        if ($ret.Success) {
            $pdf['ConvertStatus'] = "Success"
            $pdf['OutputPath']    = $pdfOut
            $pdf['ConvertError']  = ""
            $pdfConvertedCount++
        } else {
            $pdf['ConvertStatus'] = "Failed"
            $pdf['OutputPath']    = "(转换失败)"
            $pdf['ConvertError']  = $ret.Error
            $pdfConvertFailedCount++
            # 前 3 个失败直接打印到控制台，帮助定位问题
            if ($pdfConvertFailedCount -le 3) {
                Write-Host ""
                Write-Host "  ✗ 转换失败: $($pdf.Name)" -ForegroundColor Red
                Write-Host "    $($ret.Error)" -ForegroundColor DarkGray
            }
        }
    }
    Write-Progress -Activity "转换文字版 PDF" -Completed
    Write-Host "  文字版 PDF 转换完成: 成功 $pdfConvertedCount | 失败 $pdfConvertFailedCount" -ForegroundColor Cyan
}

# 把 PDF 作为特殊条目追加到结果
foreach ($pdf in $pdfFiles) {
    if ($pdf.ContainsKey('ConvertStatus')) {
        $pdfStatus = $pdf.ConvertStatus
        $pdfOutput = $pdf.OutputPath
    } else {
        $pdfStatus = "PDF"
        $pdfOutput = "(未转换)"
    }

    $allResults += @{
        FileName     = $pdf.Name
        SourcePath   = $pdf.FullName
        OutputPath   = $pdfOutput
        SourceExt    = ".pdf"
        RealFormat   = $pdf.PdfType
        ExtMismatch  = $false
        Status       = $pdfStatus
        ElapsedMs    = 0
        ErrorMessage = if ($pdf.ContainsKey('ConvertError')) { $pdf.ConvertError } else { "" }
    }
}
$pdfCount = $pdfFiles.Count
$successCount  = $cntSuccess.Value
$skippedCount  = $cntSkipped.Value
$failedCount   = $cntFailed.Value
$mismatchCount = @($allResults | Where-Object { $_.ExtMismatch }).Count
# hashtable 的 key 不是 .NET 属性，Measure-Object 无法识别，手动求和
$avgMs = 0
if ($allResults.Count -gt 0) {
    $sumMs = 0
    foreach ($r in $allResults) { $sumMs += [int]$r.ElapsedMs }
    $avgMs = [int]($sumMs / $allResults.Count)
}
$elapsed       = (Get-Date) - $startTime
$elapsedStr    = $elapsed.ToString('mm\:ss')
$generatedAt   = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

# 报告数据对象
$reportData = $allResults | ForEach-Object {
    [PSCustomObject]@{
        文件名    = $_.FileName
        原始路径  = $_.SourcePath
        输出路径  = $_.OutputPath
        后缀名    = $_.SourceExt
        真实格式  = $_.RealFormat
        后缀异常  = if ($_.ExtMismatch) { "是" } else { "否" }
        转换状态  = $_.Status
        耗时ms    = $_.ElapsedMs
        失败原因  = $_.ErrorMessage
    }
}

# 优先生成 xlsx（带筛选、冻结、自动列宽、状态颜色标记），失败回退 CSV
Initialize-ExcelReport
$xlsxOk = $false
if ($script:ExcelAvailable) {
    try {
        $excelPkg = $reportData | Export-Excel -Path $xlsxReport `
            -WorksheetName "转换报告" `
            -AutoFilter `
            -FreezeTopRow `
            -BoldTopRow `
            -AutoSize `
            -PassThru

        $ws = $excelPkg.Workbook.Worksheets["转换报告"]
        $rowCount = $reportData.Count

        # 状态列（第 7 列：转换状态）着色
        for ($row = 2; $row -le $rowCount + 1; $row++) {
            $status = $ws.Cells[$row, 7].Value
            $fillColor = switch ($status) {
                "Success" { [System.Drawing.Color]::FromArgb(235, 250, 235) }   # 淡绿
                "Failed"  { [System.Drawing.Color]::FromArgb(255, 230, 230) }   # 淡红
                "Skipped" { [System.Drawing.Color]::FromArgb(255, 252, 220) }   # 淡黄
                "PDF"     { [System.Drawing.Color]::FromArgb(220, 245, 240) }   # 淡青
                default   { $null }
            }
            if ($fillColor) {
                $range = $ws.Cells[$row, 1, $row, 9]
                $range.Style.Fill.PatternType = "Solid"
                $range.Style.Fill.BackgroundColor.SetColor($fillColor)
            }
        }

        Close-ExcelPackage $excelPkg
        $xlsxOk = $true
    } catch {
        Write-Host "  xlsx 生成失败，回退到 CSV：$($_.Exception.Message)" -ForegroundColor Yellow
    }
}

if (-not $xlsxOk) {
    $reportData | Export-Csv -Path $csvReport -Encoding UTF8 -NoTypeInformation
}

# failed_list.txt
$failedResults = @($allResults | Where-Object { $_.Status -eq "Failed" })
if ($failedResults) {
    $failedResults | ForEach-Object { $_.SourcePath } | Set-Content -Path $failedList -Encoding UTF8
}

# HTML
Add-Type -AssemblyName System.Web

$rowsHtml = ($allResults | ForEach-Object {
    $cls = switch ($_.Status) { "Success"{"success"} "Skipped"{"skipped"} "Failed"{"failed"} "PDF"{"pdf"} default{""} }
    $mismatchBadge = if ($_.ExtMismatch) { '<span class="badge warn">后缀异常</span>' } else { "" }
    $errShort = if ($_.ErrorMessage -and $_.ErrorMessage.Length -gt 60) {
        $_.ErrorMessage.Substring(0,60) + "…"
    } else { $_.ErrorMessage }
    $errCell = if ($_.ErrorMessage) {
        "<td class='error-msg' title='$([System.Web.HttpUtility]::HtmlEncode($_.ErrorMessage))'>$([System.Web.HttpUtility]::HtmlEncode($errShort))</td>"
    } else { "<td>—</td>" }
    "<tr class='$cls'>
        <td>$([System.Web.HttpUtility]::HtmlEncode($_.FileName)) $mismatchBadge</td>
        <td>$($_.SourceExt)</td>
        <td>$($_.RealFormat)</td>
        <td><span class='status-$cls'>$($_.Status)</span></td>
        <td>$($_.ElapsedMs) ms</td>
        $errCell
    </tr>"
}) -join "`n"

$html = @"
<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8">
<title>转换报告 — $generatedAt</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,"Microsoft YaHei",sans-serif;background:#f5f6fa;color:#333}
.header{background:#2c3e50;color:#fff;padding:24px 32px}
.header h1{font-size:20px;font-weight:600}
.header p{font-size:13px;opacity:.7;margin-top:4px}
.cards{display:flex;gap:16px;padding:24px 32px 8px;flex-wrap:wrap}
.card{background:#fff;border-radius:8px;padding:16px 24px;flex:1;min-width:140px;box-shadow:0 1px 4px rgba(0,0,0,.08)}
.card .val{font-size:32px;font-weight:700}
.card .lbl{font-size:12px;color:#888;margin-top:4px}
.card.c-success .val{color:#27ae60}
.card.c-skipped .val{color:#f39c12}
.card.c-failed .val{color:#e74c3c}
.card.c-warn .val{color:#e67e22}
.card.c-total .val{color:#2980b9}
.card.c-time .val{color:#8e44ad}
.card.c-pdf .val{color:#16a085}
.status-pdf{color:#16a085;font-weight:600}
tr.pdf td{background:#f0fffa}
.meta{padding:0 32px 16px;font-size:13px;color:#666}
.meta span{margin-right:24px}
.table-wrap{padding:0 32px 32px;overflow-x:auto}
table{width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.08);font-size:13px}
thead tr{background:#2c3e50;color:#fff}
th{padding:12px 14px;text-align:left;font-weight:500;white-space:nowrap}
td{padding:10px 14px;border-bottom:1px solid #f0f0f0;vertical-align:middle}
tr.failed td{background:#fff5f5}
tr.skipped td{background:#fffdf0}
.status-success{color:#27ae60;font-weight:600}
.status-skipped{color:#f39c12;font-weight:600}
.status-failed{color:#e74c3c;font-weight:600}
.badge{display:inline-block;font-size:11px;padding:1px 6px;border-radius:10px;margin-left:6px}
.badge.warn{background:#fef0cd;color:#b7690a;border:1px solid #f5c842}
.error-msg{color:#c0392b;font-size:12px;max-width:300px}
.filter-bar{padding:8px 32px 16px;display:flex;gap:10px;flex-wrap:wrap}
.filter-bar button,.filter-bar input{padding:6px 16px;border:1px solid #ddd;border-radius:20px;background:#fff;cursor:pointer;font-size:13px}
.filter-bar input{cursor:text;flex:1;min-width:200px}
.filter-bar button.active,.filter-bar button:hover{background:#2c3e50;color:#fff;border-color:#2c3e50}
tr.hidden{display:none}
</style></head><body>
<div class="header"><h1>文档批量转换报告</h1><p>生成时间：$generatedAt</p></div>
<div class="cards">
  <div class="card c-total"><div class="val">$total</div><div class="lbl">总文件数</div></div>
  <div class="card c-success"><div class="val">$successCount</div><div class="lbl">转换成功</div></div>
  <div class="card c-skipped"><div class="val">$skippedCount</div><div class="lbl">跳过</div></div>
  <div class="card c-failed"><div class="val">$failedCount</div><div class="lbl">转换失败</div></div>
  <div class="card c-warn"><div class="val">$mismatchCount</div><div class="lbl">后缀与内容不符</div></div>
  <div class="card c-pdf"><div class="val">$pdfCount</div><div class="lbl">发现 PDF（未转换）</div></div>
  <div class="card c-time"><div class="val">$elapsedStr</div><div class="lbl">总耗时 / 均${avgMs}ms</div></div>
</div>
<div class="meta">
  <span>输入目录：$InputDir</span>
  <span>输出目录：$OutputDir</span>
  <span>并行数：$ThrottleLimit</span>
</div>
<div class="filter-bar">
  <button class="active" onclick="filterTable('all',this)">全部</button>
  <button onclick="filterTable('success',this)">成功</button>
  <button onclick="filterTable('skipped',this)">跳过</button>
  <button onclick="filterTable('failed',this)">失败</button>
  <button onclick="filterTable('mismatch',this)">后缀异常</button>
  <button onclick="filterTable('pdf',this)">PDF 文件</button>
  <input type="text" placeholder="按文件名搜索…" oninput="searchTable(this.value)">
</div>
<div class="table-wrap"><table id="main-table"><thead><tr>
  <th>文件名</th><th>后缀名</th><th>真实格式</th><th>状态</th><th>耗时</th><th>失败原因</th>
</tr></thead><tbody>
$rowsHtml
</tbody></table></div>
<script>
let curFilter='all',curSearch='';
function applyFilter(){
  document.querySelectorAll('#main-table tbody tr').forEach(r=>{
    let show=true;
    if(curFilter==='mismatch') show=!!r.querySelector('.badge.warn');
    else if(curFilter!=='all') show=r.classList.contains(curFilter);
    if(show&&curSearch) show=r.cells[0].textContent.toLowerCase().includes(curSearch);
    r.classList.toggle('hidden',!show);
  });
}
function filterTable(t,b){
  document.querySelectorAll('.filter-bar button').forEach(x=>x.classList.remove('active'));
  b.classList.add('active'); curFilter=t; applyFilter();
}
function searchTable(v){curSearch=v.toLowerCase();applyFilter();}
</script></body></html>
"@

$html | Set-Content -Path $htmlReport -Encoding UTF8

Write-Host "`r报告生成完成                        "

# ─── 输出 ─────────────────────────────────────────────────────────────────

Write-Log "===== 转换完成 ====="
Write-Log "成功: $successCount  跳过: $skippedCount  失败: $failedCount  后缀异常: $mismatchCount  PDF: $pdfCount  耗时: $elapsedStr"

Write-Host ""
Write-Host "===== 转换完成 =====" -ForegroundColor Cyan
Write-Host "成功: $successCount  跳过: $skippedCount  失败: $failedCount  后缀异常: $mismatchCount  PDF: $pdfCount  耗时: $elapsedStr"
if ($xlsxOk) {
    Write-Host "Excel → $xlsxReport" -ForegroundColor Green
} else {
    Write-Host "CSV  → $csvReport"
}
Write-Host "HTML → $htmlReport"
if ($failedResults) { Write-Host "失败列表 → $failedList" -ForegroundColor Yellow }

# ─── 清理临时文件 ─────────────────────────────────────────────────────────

if ($tempExtract -and (Test-Path $tempExtract)) {
    $shouldClean = $false

    if ($null -ne $CleanTemp) {
        # 命令行参数已指定，直接按参数走
        $shouldClean = [bool]$CleanTemp
    } else {
        # 交互询问（非交互环境自动保留）
        if ([Environment]::UserInteractive -and $Host.UI.RawUI) {
            Write-Host ""
            Write-Host "临时解压目录: $tempExtract" -ForegroundColor Cyan
            $answer = Read-Host "是否删除该临时目录？[y/N]"
            $shouldClean = ($answer -match '^[yY]')
        } else {
            Write-Host "非交互环境，默认保留临时文件" -ForegroundColor DarkGray
        }
    }

    if ($shouldClean) {
        try {
            Remove-Item $tempExtract -Recurse -Force -ErrorAction Stop
            Write-Host "临时文件已清理"
        } catch {
            Write-Host "临时文件清理失败（可手动删除）: $tempExtract" -ForegroundColor Yellow
        }
    } else {
        Write-Host "临时文件已保留: $tempExtract" -ForegroundColor DarkGray
    }
}

# ─── 结束后友好提示 ───────────────────────────────────────────────────────

if ([Environment]::UserInteractive -and $Host.UI.RawUI) {
    Write-Host ""
    # 报告直接打开
    $openReport = Read-Host "是否立即打开 HTML 报告？[Y/n]"
    if ($openReport -notmatch '^[nN]') {
        try { Start-Process $htmlReport } catch { Write-Host "无法打开报告：$_" -ForegroundColor Yellow }
    }

    # 有失败文件时额外问
    if ($failedCount -gt 0 -and (Test-Path $failedList)) {
        $openFailed = Read-Host "共有 $failedCount 个文件失败，是否打开失败列表？[Y/n]"
        if ($openFailed -notmatch '^[nN]') {
            try { Start-Process notepad.exe $failedList } catch { Write-Host "无法打开失败列表：$_" -ForegroundColor Yellow }
        }
    }
}

if ($failedCount -gt 0) { exit 1 }