@echo off
chcp 65001 >nul
title ✨ MarkItDown 魔法工具箱 ✨

:: 放弃使用 CMD 打印中文，将所有控制权直接无缝移交给 PowerShell 脚本
:: 使用 ScriptBlock 方式彻底解决编码问题，并传递当前绝对路径
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& ([ScriptBlock]::Create((Get-Content -LiteralPath '%~dp0run.ps1' -Encoding UTF8 -Raw))) -ScriptDir '%~dp0\'"

pause