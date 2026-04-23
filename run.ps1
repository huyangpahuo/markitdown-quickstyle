param([string]$ScriptDir = $PWD.Path)
$ScriptDir = $ScriptDir.Trim("'", '"', ' ')
Set-Location -Path $ScriptDir
[console]::OutputEncoding = [System.Text.Encoding]::UTF8
$host.ui.RawUI.WindowTitle = "✨ MarkItDown 魔法工具箱 V2.0 ✨"

# ================= 首次安装引导逻辑 =================
if (!(Test-Path ".installed")) {
    Clear-Host
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host "       ✨ 欢迎来到 MarkItDown 魔法工具箱 ✨        " -ForegroundColor Cyan
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host "`n呀，检测到你是第一次来到这里呢！👋" -ForegroundColor Yellow
    Write-Host "我们需要先布置一下运行环境（只需要进行一次哦）~`n" -ForegroundColor White
    
    Write-Host "1. 🚀 开始安装环境 (冲鸭！)"
    Write-Host "2. 🏃 暂时退出 (下次一定)`n"
    
    $installChoice = Read-Host "👉 请输入选项并按回车"
    
    if ($installChoice -eq "1") {
        Write-Host "`n🌟 正在启动安装向导..." -ForegroundColor Cyan
        $installCmd = "& ([ScriptBlock]::Create((Get-Content -LiteralPath 'install.ps1' -Encoding UTF8 -Raw)))"
        powershell.exe -NoProfile -ExecutionPolicy Bypass -Command $installCmd
        
        New-Item -Path ".installed" -ItemType File -Force | Out-Null
        
        Write-Host "`n🎉 太棒啦，环境布置完成！正在带你进入主界面..." -ForegroundColor Green
        Start-Sleep -Seconds 2
    } else {
        Write-Host "`n拜拜啦，期待下次相见~ 👋" -ForegroundColor Cyan
        Start-Sleep -Seconds 2
        exit
    }
}
# ====================================================

$InputDir = Join-Path $ScriptDir "input"
$OutputDir = Join-Path $ScriptDir "output"
if (!(Test-Path $InputDir)) { New-Item -ItemType Directory -Force -Path $InputDir | Out-Null }
if (!(Test-Path $OutputDir)) { New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null }

$ScriptsDir = Join-Path $ScriptDir "python_env\Scripts"
$env:PATH = "$ScriptDir;$ScriptsDir;" + $env:PATH
$PythonExe = Join-Path $ScriptDir "python_env\python.exe"

function Show-Manual {
    Clear-Host
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host "            📖 魔法工具箱使用说明书 (V2.0) 📖      " -ForegroundColor Cyan
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host "`n 📌 【全能格式支持列表】" -ForegroundColor Yellow
    Write-Host "  • 📄 Office: 新版 Word (.docx), PPT (.pptx), Excel (.xlsx)"
    Write-Host "  • 📑 PDF & 电子书: PDF (.pdf), EPUB (.epub) (PDF不支持提取图片)"
    Write-Host "  • 🖼️ 静态/动态图: .jpg, .png, .gif, .webp 等 (批量转为单个 新建文件.md)"
    Write-Host "  • 🌐 网页与数据: .html, .csv, .json, .xml"
    Write-Host "  • 📦 高级压缩包: .zip (仅提取根目录文件，跳过嵌套文件夹)"
    Write-Host "  • ⚠️ [实验阶段不可用]: 🎧音视频(.wav/.mp3), 📧邮件(.msg)"
    Write-Host "`n 💡 【高级特性指南】" -ForegroundColor Cyan
    Write-Host "  1. 独立包裹: 转换后在 output 生成[独立同名文件夹]，杜绝混乱！"
    Write-Host "  2. Typora支持: Office提取的图片保存在 assets。推荐使用 Typora 软件打开"
    Write-Host "     生成的 Markdown 文件，已自动清洗乱码并写入 YAML 配置！"
    Write-Host "  3. PPT精准插图: PPT图片会自动插入到对应的幻灯片页码下方！"
    Write-Host "=================================================" -ForegroundColor Cyan
    pause
}

function Run-Python {
    param([string]$Mode, [string]$Target, [string]$Exts="")
    if ($Exts) {
        & $PythonExe "magic_convert.py" $Mode $Target $OutputDir $Exts
    } else {
        & $PythonExe "magic_convert.py" $Mode $Target $OutputDir
    }
}

function Show-BatchMenu {
    while ($true) {
        Clear-Host
        Write-Host "=================================================" -ForegroundColor Cyan
        Write-Host "           🎯 批量转换【input】中的文件          " -ForegroundColor Cyan
        Write-Host "=================================================" -ForegroundColor Cyan
        Write-Host "`n 请选择你要批量处理的类型：" -ForegroundColor Yellow
        Write-Host "  [ 1 ] 📄 Office 文档 (.docx / .pptx / .xlsx)"
        Write-Host "  [ 2 ] 📑 PDF 文档 (.pdf)"
        Write-Host "  [ 3 ] 📚 电子书 (.epub)"
        Write-Host "  [ 4 ] 🖼️ 批量合并图片 (将多图融合成 单个 新建文件.md)"
        Write-Host "  [ 5 ] 🎧 音频提取文本 [实验阶段不可用]"
        Write-Host "  [ 6 ] 🌐 网页与数据 (.html/.csv/.json/.xml)"
        Write-Host "  [ 7 ] 📦 压缩包解析 (.zip 深度1解析)"
        Write-Host "  [ 8 ] 📧 邮件文件 [实验阶段不可用]"
        Write-Host "  [ 0 ] ⬅️ 返回上级菜单`n"

        $bChoice = Read-Host "👉 告诉我你的选择"
        
        if ($bChoice -eq "5" -or $bChoice -eq "8") {
            Write-Host "`n❌ 此功能处于实验阶段，目前暂不可用！" -ForegroundColor Red
            pause
            continue
        }

        if ($bChoice -eq "0") { return }

        Write-Host "`n🚀 开始批量施法..." -ForegroundColor Cyan
        switch ($bChoice) {
            "1" { Run-Python "batch" $InputDir ".docx,.pptx,.xlsx"; break }
            "2" { Run-Python "batch" $InputDir ".pdf"; break }
            "3" { Run-Python "batch" $InputDir ".epub"; break }
            "4" { Run-Python "batch_images" $InputDir; break }
            "6" { Run-Python "batch" $InputDir ".html,.htm,.csv,.json,.xml"; break }
            "7" { Run-Python "batch" $InputDir ".zip"; break }
            default { Write-Host "❌ 无效选项" -ForegroundColor Red }
        }
        Write-Host "`n✅ 批量转换完成！请前往 output 文件夹查看。" -ForegroundColor Green
        pause
        return
    }
}

while ($true) {
    Clear-Host
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host "               ✨ MarkItDown 魔法工具箱 ✨        " -ForegroundColor Cyan
    Write-Host "=================================================" -ForegroundColor Cyan
    Write-Host " 📂 状态: 已挂载 input 和 output | 引擎: Python核心" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  [ 1 ] 🎯 转换单个文件 (支持任意格式，直接拖拽文件进来)"
    Write-Host "  [ 2 ] 🎯 批量转换【input】中的文件 (分类多选)"
    Write-Host "  [ 3 ] 📂 打开【output】文件夹"
    Write-Host "  [ 4 ] 📖 查看说明书"
    Write-Host "  [ 0 ] 🏃 退出工具箱"
    Write-Host ""

    $choice = Read-Host "👉 告诉我你的选择"
    switch ($choice) {
        "1" {
            $rawInput = Read-Host "📥 请输入或拖入文件路径"
            # 核心修复：拦截空输入
            if ([string]::IsNullOrWhiteSpace($rawInput)) {
                Write-Host "❌ 输入不能为空！" -ForegroundColor Red
                pause
                continue
            }
            
            $file = $rawInput.Trim('"').Trim("'")
            if (Test-Path $file) {
                Write-Host "`n✨ 正在对 $($file) 施加转换魔法..." -ForegroundColor Yellow
                Run-Python "single" $file
                Write-Host "`n✅ 搞定！已完美包裹并保存到 output 文件夹！ 🎈" -ForegroundColor Green
            } else {
                Write-Host "❌ 哎呀，没找到这个文件，是不是路径写错了呀？" -ForegroundColor Red
            }
            pause
        }
        "2" { Show-BatchMenu }
        "3" { Invoke-Item -Path $OutputDir }
        "4" { Show-Manual }
        "0" { exit }
    }
}