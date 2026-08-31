param(
    [Parameter(Mandatory = $true)]
    [string]$InputDocx,
    [Parameter(Mandatory = $true)]
    [string]$OutputPdf
)

$ErrorActionPreference = 'Stop'
$inputItem = Get-Item -LiteralPath $InputDocx -ErrorAction Stop
if ($inputItem.PSIsContainer -or $inputItem.Extension -ine '.docx') {
    throw "InputDocx 必须是存在的 .docx 文件：$InputDocx"
}

$inputPath = $inputItem.FullName
$outputPath = [System.IO.Path]::GetFullPath($OutputPdf)
if ([System.IO.Path]::GetExtension($outputPath) -ine '.pdf') {
    throw "OutputPdf 必须使用 .pdf 扩展名：$outputPath"
}
if ([System.StringComparer]::OrdinalIgnoreCase.Equals($inputPath, $outputPath)) {
    throw '输入 DOCX 与输出 PDF 不能是同一路径。'
}

$outputDirectory = Split-Path -Parent $outputPath
[System.IO.Directory]::CreateDirectory($outputDirectory) | Out-Null

$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try { $word.AutomationSecurity = 3 } catch {}
    try { $word.Options.UpdateLinksAtOpen = $false } catch {}
    try { $word.Options.UpdateFieldsAtPrint = $false } catch {}
    try { $word.Options.UpdateFieldsWithTrackedChangesAtPrint = $false } catch {}

    # ConfirmConversions=false, ReadOnly=true, AddToRecentFiles=false。
    # 四参数调用兼容不同 Word COM 版本；安全开关在应用层单独设置。
    $document = $word.Documents.Open($inputPath, $false, $true, $false)
    try { $document.UpdateStylesOnOpen = $false } catch {}
    [void]$document.ComputeStatistics(2)
    $document.ExportAsFixedFormat($outputPath, 17)
}
finally {
    if ($null -ne $document) {
        try { $document.Saved = $true } catch {}
        try { $document.Close(0) } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) } catch {}
    }
    if ($null -ne $word) {
        try { $word.Quit(0) } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$outputItem = Get-Item -LiteralPath $outputPath -ErrorAction Stop
if ($outputItem.PSIsContainer -or $outputItem.Length -lt 5) {
    throw "Word 未生成有效 PDF：$outputPath"
}
$stream = [System.IO.File]::OpenRead($outputPath)
try {
    $signatureBytes = New-Object byte[] 5
    [void]$stream.Read($signatureBytes, 0, 5)
    $signature = [System.Text.Encoding]::ASCII.GetString($signatureBytes)
}
finally {
    $stream.Dispose()
}
if ($signature -ne '%PDF-') {
    throw "输出文件不是可识别的 PDF：$outputPath"
}

Write-Output "Word PDF saved to $outputPath"
