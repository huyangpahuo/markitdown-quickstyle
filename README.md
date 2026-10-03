# markitdown-quickstyle

这是微软的 [markitdown](https://github.com/microsoft/markitdown) 的便捷使用版本!

把 Office / PDF / EPUB / 图片 / 网页 / ZIP 等文件一键转换成 Markdown,自动提取图片。

## 🖥️ 图形界面版(推荐)

小窗口工具 + 一只悬浮在桌面上的桌宠(会眨眼的圆 🟣):

```powershell
cd app
npm run tauri dev      # 开发运行
npm run tauri build    # 构建 exe(app\src-tauri\target\release\)
```

- 拖拽文件到窗口即可转换,也支持批量转换 `input/` 目录
- 单击桌宠唤起主窗口,按住可拖动
- 环境搭建与架构说明见 [TAURI_SETUP.md](TAURI_SETUP.md)

## 🖥️ 终端版(旧)

点击 **start.bat** 启动,首次需要下载环境。

## 🤖 MCP(AI Agent 接入)

程序内置标准 MCP server,AI 客户端(Claude Desktop / Cursor / ZCode 等)可直接调用转换能力:

```json
{ "mcpServers": { "markitdown-quickstyle": {
    "command": "D:\\markitdown-quickstyle\\app\\src-tauri\\target\\release\\markitdown-quickstyle.exe",
    "args": ["mcp"] } } }
```

提供 `convert_file` / `convert_batch` / `list_supported_formats` 三个工具,详见 [TAURI_SETUP.md](TAURI_SETUP.md)。

---

对于音频和邮件的识别和转录目前为实验性功能暂不可用
