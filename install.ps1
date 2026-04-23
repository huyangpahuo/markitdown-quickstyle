# 修复终端中文乱码
[console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "  ✨ 正在为你施放 MarkItDown 安装魔法，请稍候... ✨" -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host ""

$ScriptDir = $PWD.Path
$PythonExe = Join-Path $ScriptDir "python_env\python.exe"
$GetPipScript = Join-Path $ScriptDir "get-pip.py"
$PipExe = Join-Path $ScriptDir "python_env\Scripts\pip.exe"

if (!(Test-Path $PythonExe)) {
    Write-Host "❌ 哎呀，没有找到内置的 Python 环境！" -ForegroundColor Red
    Write-Host "请确保你已经把便携版 Python 解压并重命名为了 'python_env' 文件夹。" -ForegroundColor Yellow
    pause
    exit
}

# 智能检测：如果没有 pip，就自动执行 get-pip.py (显示进度条，屏蔽黄字警告)
if (!(Test-Path $PipExe)) {
    if (!(Test-Path $GetPipScript)) {
        Write-Host "❌ 缺少核心安装组件：get-pip.py！请将该文件放入工具箱目录。" -ForegroundColor Red
        pause
        exit
    }
    Write-Host "💧 正在注入初始化魔力 (自动安装 Pip 工具)..." -ForegroundColor Yellow
    & $PythonExe $GetPipScript --no-warn-script-location
} else {
    Write-Host "💧 正在更新魔力源泉 (升级 pip)..." -ForegroundColor Yellow
    & $PythonExe -m pip install --upgrade pip --no-warn-script-location
}

Write-Host "`n📦 正在搬运格式转换法宝 (安装 markitdown 核心组件)..." -ForegroundColor Yellow
Write-Host "   (由于需要从云端下载，这可能需要 1~3 分钟，请欣赏魔法进度条 ☕)" -ForegroundColor Cyan

# 强制让便携版 Python 安装 markitdown (显示进度条，屏蔽黄字警告)
& $PythonExe -m pip install "markitdown[all]" --no-warn-script-location

Write-Host ""
Write-Host "✅ 魔法环境布置大功告成啦！准备起飞~ 🚀" -ForegroundColor Green