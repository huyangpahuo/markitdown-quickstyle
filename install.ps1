# 修复终端中文乱码
[console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "  ✨ 正在为你施放 MarkItDown 安装魔法，请稍候... ✨" -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host ""

$ScriptDir = $PWD.Path
$PythonExe = Join-Path $ScriptDir "python_env\python.exe"

if (!(Test-Path $PythonExe)) {
    Write-Host "❌ 哎呀，   有找到内置的 Python 环境！" -ForegroundColor Red
    Write-Host "请确保你已经把便携版 Python 解压并重命名为了 'python_env' 文件夹。" -ForegroundColor Yellow
    pause
    exit
}

Write-Host "💧 正在注入魔力 (初始化绿色环境)..." -ForegroundColor Yellow
& $PythonExe -m pip install --upgrade pip -q

Write-Host "📦 正在搬运格式转换法宝 (安装 markitdown 核心组件)..." -ForegroundColor Yellow
Write-Host "   (这需要一点时间，请耐心等待，不要关闭窗口 ☕)" -ForegroundColor Cyan
# 强制让便携版 Python 安装 markitdown
& $PythonExe -m pip install "markitdown[all]"

Write-Host ""
Write-Host "✅ 魔法环境布置大功告成啦！准备起飞~ 🚀" -ForegroundColor Green