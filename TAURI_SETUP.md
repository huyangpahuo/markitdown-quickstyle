# Tauri 环境搭建与开发指南(Windows)

本文档记录本项目从终端工具升级为 **Tauri 桌面应用** 的完整环境搭建步骤、项目结构与开发/构建/MCP 接入方法。

---

## 0. 前置条件

| 依赖 | 要求 | 本机验证结果 |
|---|---|---|
| Node.js | ≥ 20(推荐 LTS) | ✅ v24.14.0 |
| Rust(rustup,MSVC 工具链) | 稳定版 | ✅ cargo 1.94.1 |
| Visual Studio 2022 C++ 生成工具 | 含 "使用 C++ 的桌面开发" 工作负载(MSVC + Windows SDK) | ✅ VS Community 2022 |
| WebView2 Runtime | Win10/11 一般自带 | ✅ 154.0.4258.53 |

检查命令:

```powershell
node -v          # Node 版本
cargo -V         # Cargo 版本
rustup show      # 确认 default host 是 x86_64-pc-windows-msvc
# WebView2 运行时(有输出即已安装)
reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}" /v pv
```

> 若缺少 VS C++ 工具链:安装 **Visual Studio Installer → Visual Studio Community 2022 → 勾选"使用 C++ 的桌面开发"**。
> 若 `rustup show` 显示 gnu 工具链:`rustup default stable-x86_64-pc-windows-msvc`。

---

## 1. 架构设计

核心原则:**Python 转换引擎(markitdown)一行不动的逻辑照旧,Rust 只做调度和窗口**。

```
markitdown-quickstyle/
├── python_env/              # 内嵌 Python(markitdown 已装好)
├── magic_convert.py         # 转换核心(新增 --json 事件输出模式,向后兼容)
├── ffmpeg.exe
├── input/  output/          # 转换输入/输出目录
└── app/                     # ★ Tauri 应用
    ├── package.json         # 前端依赖(vite + @tauri-apps)
    ├── vite.config.js       # 双页面入口:index.html(主窗) + pet.html(桌宠)
    ├── index.html           # 主窗口页面(小窗 420×640)
    ├── pet.html             # 桌宠窗口页面(透明无边框置顶)
    ├── src/                 # 前端 JS/CSS
    └── src-tauri/
        ├── tauri.conf.json  # 窗口/打包配置
        ├── capabilities/    # Tauri v2 权限声明
        ├── icons/           # 应用图标(由脚本生成)
        └── src/
            ├── main.rs      # 入口:带 `mcp` 参数 → MCP server,否则启动 GUI
            ├── paths.rs     # 向上查找工具根目录(定位 python_env)
            ├── convert.rs   # 调度 python 转换,事件流推给前端
            └── mcp.rs       # stdio MCP server(initialize/tools/list/tools/call)
```

**根目录定位**:`paths.rs` 会从 cwd / exe 所在位置向上查找包含 `magic_convert.py`
的目录作为工具根目录。因此把编译出的 exe 放在仓库任意层级均可工作;找不到时可用
环境变量 `MARKITDOWN_ROOT` 强制指定根目录。

---

## 2. 环境搭建步骤(已完成,新机器照做)

```powershell
# 1. 进入前端工程目录
cd D:\markitdown-quickstyle\app

# 2. 安装 npm 依赖(vite、@tauri-apps/api、@tauri-apps/cli、dialog 插件)
npm install

# 3. 生成应用图标(纯 Python 标准库实现,不依赖网络)
..\python_env\python.exe scripts\make_icons.py

# 4. 首次构建前端产物(cargo 编译需要 dist/ 存在)
npm run build

# 5. 编译 Rust 后端(debug 版,用于 MCP 冒烟测试)
cd src-tauri && cargo build && cd ..
```

首次 `cargo build` 需要编译约 400 个 crate,耗时几分钟,之后有增量缓存。

---

## 3. 日常开发

```powershell
cd app
npm run tauri dev     # 热更新开发模式(会同时拉起主窗口 + 桌宠窗口)
```

- **主窗口**:拖拽文件转换、批量分类转换(input 目录)、实时日志、打开 output。
- **桌宠窗口**:独立悬浮在 Windows 桌面上的透明无边框置顶小窗(不占任务栏),
  一个会眨眼的圆。可按住拖动,单击唤起/聚焦主窗口;转换进行中会有"忙碌"表情。

## 4. 构建发布

```powershell
cd app
npm run tauri build
```

- 裸 exe:`app\src-tauri\target\release\markitdown-quickstyle.exe`
- 安装包:`app\src-tauri\target\release\bundle\nsis\*.exe`(首次会自动下载 NSIS)

> 本项目采用**绿色便携**分发模型:exe 可以单独拷走,只要它和 `python_env/`、
> `magic_convert.py` 处于同一目录(或其上级目录)即可运行。NSIS 安装包默认
> **不包含** python_env(数百 MB),更适合继续走"解压即用"的发布方式。

---

## 5. MCP 接入(让 AI Agent 使用本工具)

应用内置标准 **MCP(Model Context Protocol)stdio server**。任何 MCP 客户端
(Claude Desktop、Cursor、ZCode 等)都可以把本程序当作工具服务器:

```
markitdown-quickstyle.exe mcp
```

### 提供的工具

| 工具 | 参数 | 说明 |
|---|---|---|
| `convert_file` | `file_path`(必填)、`output_dir`(可选) | 单文件转 Markdown,支持 docx/pptx/xlsx/pdf/epub/html/csv/json/xml/图片/zip |
| `convert_batch` | `extensions`、`input_dir`、`output_dir`(均可选) | 批量转换目录;`extensions` 传空字符串则合并目录下所有图片 |
| `list_supported_formats` | 无 | 返回支持的格式与输出结构说明 |

### Claude Desktop 配置示例

`%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "markitdown-quickstyle": {
      "command": "D:\\markitdown-quickstyle\\app\\src-tauri\\target\\release\\markitdown-quickstyle.exe",
      "args": ["mcp"]
    }
  }
}
```

### Cursor 配置示例(`.cursor/mcp.json`)

```json
{
  "mcpServers": {
    "markitdown-quickstyle": {
      "command": "D:/markitdown-quickstyle/app/src-tauri/target/release/markitdown-quickstyle.exe",
      "args": ["mcp"]
    }
  }
}
```

### 命令行手动测试(管道模拟 MCP 客户端)

推荐用仓库自带的冒烟测试脚本(模拟完整 AI 客户端流程):

```bash
# 协议握手 + 工具列表
python_env/python.exe app/scripts/mcp_smoke.py

# 实际调用转换工具
python_env/python.exe app/scripts/mcp_smoke.py \
  '[["convert_file", {"file_path": "D:/markitdown-quickstyle/测试专用/office/MySQL数据类型.xlsx"}]]'
```

> 注意:在 Git Bash 里传 JSON 参数请用正斜杠路径,反斜杠会被 MSYS 转义吞掉。
> MCP server 走 stdio(换行分隔 JSON-RPC),由 AI 客户端用管道拉起,
> 不需要 GUI 在运行,也不占用端口。release 版 exe 是 Windows GUI 子系统,
> 但父进程通过管道 spawn 时 stdout/stdin 照常工作,这正是 MCP 的用法。

---

## 6. 常见问题

- **提示找不到 Python 引擎**:确认 exe 能向上找到 `magic_convert.py` 所在目录,
  或设置环境变量 `MARKITDOWN_ROOT=D:\markitdown-quickstyle`。
- **转换时不闪黑框**:Rust 侧已对所有子进程设置 `CREATE_NO_WINDOW`。
- **中文乱码**:子进程统一注入 `PYTHONUTF8=1` / `PYTHONIOENCODING=utf-8`。
- **`magic_convert.py` 兼容性**:不带 `--json` 参数时行为与旧版 `run.ps1`
  完全一致,终端版仍可继续使用。
