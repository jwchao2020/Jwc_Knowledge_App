<#
.SYNOPSIS
    [最终修正版] Word 转 PDF 自动化脚本
    1. 修复了变量声明的语法错误 ($global)
    2. 针对复杂文档(公式多)增加了等待时间
    3. 完善了空对象检查，防止报错
#>

# === 配置路径 ===
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptPath
$SourceDir = Join-Path $ProjectRoot "source_word"
$WebRoot   = Join-Path $ProjectRoot "Documents"
$PdfOutDir = Join-Path $WebRoot "docs_pdf"
$DataJsPath = Join-Path $WebRoot "data.js"

# 修复点 1: 加上 $ 符号
$global:wordApp = $null

# === 辅助函数：获取或启动 Word ===
function Get-Or-Start-Word {
    try {
        # 检查进程是否存在且可用
        if ($null -eq $global:wordApp) {
            Write-Host "⚙️ 启动 Word 进程..." -ForegroundColor Cyan
            $global:wordApp = New-Object -ComObject Word.Application
            $global:wordApp.Visible = $false 
            $global:wordApp.DisplayAlerts = 0 
        }
        # 尝试访问属性以测试连接是否存活
        $test = $global:wordApp.Version
    } catch {
        Write-Warning "⚠️ Word 进程无响应或已断开，正在重启..."
        # 强制清理旧进程
        Stop-Process -Name "WINWORD" -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        
        # 重建
        $global:wordApp = New-Object -ComObject Word.Application
        $global:wordApp.Visible = $false
        $global:wordApp.DisplayAlerts = 0
    }
}

Write-Host "🚀 开始构建流程..." -ForegroundColor Cyan

# 准备输出目录
if (!(Test-Path $PdfOutDir)) { New-Item -ItemType Directory -Path $PdfOutDir | Out-Null }

$TreeData = @()

# === 递归处理函数 ===
function Process-Folder {
    param (
        [string]$CurrentSource,
        [string]$CurrentPdfOut,
        [string]$RelativeWebPath
    )

    $FolderNode = @{
        id = "dir_" + (Get-Random)
        title = (Split-Path $CurrentSource -Leaf)
        icon = "📂"
        children = @()
    }

    # A. 处理子文件夹
    $SubDirs = Get-ChildItem -Path $CurrentSource -Directory | Sort-Object { [regex]::Replace($_.Name, '\d+', { $args[0].Value.PadLeft(20, '0') }) }

    foreach ($dir in $SubDirs) {
        $NextSource = Join-Path $CurrentSource $dir.Name
        $NextPdfOut = Join-Path $CurrentPdfOut $dir.Name
        if (!(Test-Path $NextPdfOut)) { New-Item -ItemType Directory -Path $NextPdfOut | Out-Null }
        
        $ChildNode = Process-Folder -CurrentSource $NextSource -CurrentPdfOut $NextPdfOut -RelativeWebPath "$RelativeWebPath/$($dir.Name)"
        $FolderNode.children += $ChildNode
    }

    # B. 处理 Word 文件
    $Files = Get-ChildItem -Path $CurrentSource -Filter "*.docx" | Sort-Object { [regex]::Replace($_.Name, '\d+', { $args[0].Value.PadLeft(20, '0') }) }
    foreach ($file in $Files) {
        if ($file.Name.StartsWith("~")) { continue }

        $DocName = $file.BaseName
        $PdfName = "$DocName.pdf"
        $InputPath = $file.FullName
        $OutputPath = Join-Path $CurrentPdfOut $PdfName
        
        # 增量更新逻辑
        $NeedConvert = $true
        if (Test-Path $OutputPath) {
            $SrcTime = (Get-Item $InputPath).LastWriteTime
            $DstTime = (Get-Item $OutputPath).LastWriteTime
            if ($DstTime -gt $SrcTime) { $NeedConvert = $false }
        }

        if ($NeedConvert) {
            Write-Host "🔄 转换: $DocName" -NoNewline
            
            Get-Or-Start-Word

            $doc = $null
            try {
                # 打开文档 (只读)
                $doc = $global:wordApp.Documents.Open($InputPath, $false, $true)
                
                # 修复点 2: 对于复杂公式文档，打开可能需要时间，稍微等一下
                Start-Sleep -Milliseconds 500 

                if ($null -ne $doc) {
                    # 导出 PDF
                    $doc.ExportAsFixedFormat($OutputPath, 17)
                    $doc.Close($false)
                    Write-Host " [OK]" -ForegroundColor Green
                } else {
                    throw "文档打开失败 (对象为空)"
                }
            } catch {
                Write-Host " [失败]" -ForegroundColor Red
                Write-Host "   ❌ 原因: $($_.Exception.Message)" -ForegroundColor Red
                
                # 安全清理
                if ($doc) { try { $doc.Close($false) } catch {} }
                
                # 如果这个文件把 Word 搞崩了，标记 Word 为空，下次循环会自动重启
                try { $global:wordApp.Quit() } catch {}
                $global:wordApp = $null
                Stop-Process -Name "WINWORD" -ErrorAction SilentlyContinue
            }
        } else {
            Write-Host "⏩ 跳过: $DocName" -ForegroundColor DarkGray
        }

        # 添加到数据节点
        $WebUrl = "docs_pdf$RelativeWebPath/$PdfName"
        $FolderNode.children += @{
            id = "file_" + (Get-Random)
            title = $DocName
            pdf = $WebUrl
            icon = "📄"
        }
    }

    return $FolderNode
}

# === 执行 ===
# 修改点：这里增加了 Sort-Object 和正则表达式，强制按照数字顺序排序
$RootDirs = Get-ChildItem -Path $SourceDir -Directory | 
    Sort-Object { [regex]::Replace($_.Name, '\d+', { $args[0].Value.PadLeft(20, '0') }) }

foreach ($cat in $RootDirs) {
    $CatOut = Join-Path $PdfOutDir $cat.Name
    if (!(Test-Path $CatOut)) { New-Item -ItemType Directory -Path $CatOut | Out-Null }
    $Node = Process-Folder -CurrentSource $cat.FullName -CurrentPdfOut $CatOut -RelativeWebPath "/$($cat.Name)"
    $TreeData += $Node
}

# === 收尾 ===
try {
    if ($global:wordApp) {
        $global:wordApp.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($global:wordApp) | Out-Null
    }
} catch {}

# 生成数据
#$JsonStr = $TreeData | ConvertTo-Json -Depth 10 -Compress
#$JsContent = "const TREE = $JsonStr;"
#Set-Content -Path $DataJsPath -Value $JsContent -Encoding UTF8


# === 生成数据 (带版本号) ===
$Payload = @{
    version = (Get-Date -Format "yyyyMMddHHmmss") # 使用时间戳作为版本号
    tree = $TreeData
}

$JsonStr = $Payload | ConvertTo-Json -Depth 10 -Compress
# 注意：这里改为 const DATA，包含 version 和 tree
$JsContent = "const LOCAL_DATA = $JsonStr;" 
Set-Content -Path $DataJsPath -Value $JsContent -Encoding UTF8

Write-Host "`n✅ 构建流程结束！" -ForegroundColor Green
Read-Host "👉 按回车键退出..."