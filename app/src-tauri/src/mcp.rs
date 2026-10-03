//! 标准 MCP(Model Context Protocol)stdio server。
//!
//! 由 AI 客户端(Claude Desktop / Cursor / ZCode 等)通过管道拉起:
//! `markitdown-quickstyle.exe mcp`
//! 协议为换行分隔的 JSON-RPC 2.0(initialize / tools/list / tools/call / ping)。

use std::io::{BufRead, Write};

use serde_json::{json, Value};

use crate::convert;

const SERVER_NAME: &str = "markitdown-quickstyle";
const SERVER_VERSION: &str = env!("CARGO_PKG_VERSION");

const DEFAULT_EXTS: &str = ".docx,.pptx,.xlsx,.xls,.pdf,.epub,.html,.htm,.csv,.json,.xml,.zip";

const FORMAT_DOC: &str = "\
支持格式与输出结构:
- Office: .docx / .pptx / .xlsx / .xls(PPT 图片会插入对应幻灯片页码下方;Word/Office 图片保存在同名 assets 文件夹)
- 文档: .pdf(仅文本,图片不可提取)、.epub
- 网页与数据: .html / .htm / .csv / .json / .xml
- 图片: .jpg / .png / .gif / .webp / .svg / .bmp / .tiff(批量合并为单个「新建文件.md」)
- 压缩包: .zip(仅提取根目录深度 1 的文件)
- 输出: 每个文件生成独立同名文件夹,内含 .md 与提取出的图片;.md 头部已写入 Typora 的 YAML 配置
- 不支持(实验阶段): 音视频(.wav/.mp3)、旧版二进制 Office(.doc/.ppt)、邮件(.msg)";

pub fn run() {
    let _ = writeln!(std::io::stderr(), "[{SERVER_NAME}] MCP stdio server started (v{SERVER_VERSION})");

    let stdin = std::io::stdin();
    let stdout = std::io::stdout();
    let mut out = stdout.lock();

    for line in stdin.lock().lines() {
        let line = match line {
            Ok(l) => l,
            Err(_) => break, // 管道关闭,退出
        };
        let trimmed = line.trim();
        if trimmed.is_empty() {
            continue;
        }
        let msg: Value = match serde_json::from_str(trimmed) {
            Ok(v) => v,
            Err(e) => {
                let _ = writeln!(std::io::stderr(), "[{SERVER_NAME}] ignoring malformed line: {e}");
                continue;
            }
        };

        // 通知(无 id)不回应
        let Some(id) = msg.get("id").cloned() else {
            continue;
        };
        let method = msg
            .get("method")
            .and_then(Value::as_str)
            .unwrap_or("")
            .to_string();

        let response = match method.as_str() {
            "initialize" => {
                let requested = msg
                    .pointer("/params/protocolVersion")
                    .and_then(Value::as_str)
                    .unwrap_or("2025-03-26")
                    .to_string();
                json!({
                    "jsonrpc": "2.0",
                    "id": id,
                    "result": {
                        "protocolVersion": requested,
                        "capabilities": { "tools": { "listChanged": false } },
                        "serverInfo": { "name": SERVER_NAME, "version": SERVER_VERSION },
                        "instructions": "MarkItDown QuickStyle 本地转换工具:把 Office/PDF/EPUB/图片/网页/ZIP 等文件转为 Markdown(含图片提取),输出到独立文件夹。"
                    }
                })
            }
            "ping" => json!({ "jsonrpc": "2.0", "id": id, "result": {} }),
            "tools/list" => json!({ "jsonrpc": "2.0", "id": id, "result": { "tools": tools() } }),
            "tools/call" => {
                let params = msg.get("params").cloned().unwrap_or(Value::Null);
                match call_tool(&params) {
                    Ok(text) => json!({
                        "jsonrpc": "2.0", "id": id,
                        "result": { "content": [ { "type": "text", "text": text } ] }
                    }),
                    Err(e) => json!({
                        "jsonrpc": "2.0", "id": id,
                        "result": { "isError": true, "content": [ { "type": "text", "text": e } ] }
                    }),
                }
            }
            other => json!({
                "jsonrpc": "2.0", "id": id,
                "error": { "code": -32601, "message": format!("method not found: {other}") }
            }),
        };

        if writeln!(out, "{response}").is_err() {
            break;
        }
        let _ = out.flush();
    }
}

fn tools() -> Value {
    json!([
        {
            "name": "convert_file",
            "description": "转换单个文件为 Markdown(支持 docx/pptx/xlsx/pdf/epub/html/csv/json/xml/图片/zip)。在输出目录生成同名文件夹,内含 .md 与提取的图片。返回转换日志。",
            "inputSchema": {
                "type": "object",
                "properties": {
                    "file_path": { "type": "string", "description": "要转换的文件绝对路径" },
                    "output_dir": { "type": "string", "description": "可选,输出目录;默认 工具根目录/output" }
                },
                "required": ["file_path"]
            }
        },
        {
            "name": "convert_batch",
            "description": "批量转换目录下的文件。extensions 为逗号分隔的扩展名(如 \".docx,.pptx,.xlsx\"),不传则转换全部支持的文档格式;传空字符串 \"\" 表示把目录下所有图片合并为单个 Markdown。目录默认 工具根目录/input。",
            "inputSchema": {
                "type": "object",
                "properties": {
                    "extensions": { "type": "string", "description": "逗号分隔扩展名过滤,如 \".docx,.pptx,.xlsx\";空字符串 = 合并图片模式" },
                    "input_dir": { "type": "string", "description": "可选,输入目录;默认 input/" },
                    "output_dir": { "type": "string", "description": "可选,输出目录;默认 output/" }
                }
            }
        },
        {
            "name": "list_supported_formats",
            "description": "查看支持的全部格式、限制与输出目录结构说明。",
            "inputSchema": { "type": "object", "properties": {} }
        }
    ])
}

fn call_tool(params: &Value) -> Result<String, String> {
    let name = params.get("name").and_then(Value::as_str).unwrap_or("");
    let args = params.get("arguments").cloned().unwrap_or_else(|| json!({}));
    let get_str = |key: &str| {
        args.get(key)
            .and_then(Value::as_str)
            .map(str::to_string)
    };

    let root = crate::paths::find_root();
    let default_out = || root.join("output").to_string_lossy().into_owned();

    match name {
        "convert_file" => {
            let Some(file_path) = get_str("file_path") else {
                return Err("缺少必填参数 file_path".into());
            };
            if !std::path::Path::new(&file_path).exists() {
                return Err(format!("文件不存在: {file_path}"));
            }
            let out_dir = get_str("output_dir").unwrap_or_else(default_out);
            std::fs::create_dir_all(&out_dir).map_err(|e| format!("创建输出目录失败: {e}"))?;
            let argv = vec!["single".to_string(), file_path, out_dir];
            let outcome = convert::run_python(&argv, |_| {})?;
            Ok(format_report(&outcome.text, &root))
        }
        "convert_batch" => {
            let input_dir = get_str("input_dir")
                .unwrap_or_else(|| root.join("input").to_string_lossy().into_owned());
            if !std::path::Path::new(&input_dir).exists() {
                return Err(format!("输入目录不存在: {input_dir}"));
            }
            let out_dir = get_str("output_dir").unwrap_or_else(default_out);
            std::fs::create_dir_all(&out_dir).map_err(|e| format!("创建输出目录失败: {e}"))?;
            let exts = get_str("extensions").unwrap_or_else(|| DEFAULT_EXTS.to_string());
            let images_mode = exts.trim().is_empty();
            let mode = if images_mode { "batch_images" } else { "batch" };
            let ext_arg = if images_mode { "NONE".to_string() } else { exts };
            let argv = vec![mode.to_string(), input_dir, out_dir, ext_arg, "批量".to_string()];
            let outcome = convert::run_python(&argv, |_| {})?;
            Ok(format_report(&outcome.text, &root))
        }
        "list_supported_formats" => Ok(FORMAT_DOC.to_string()),
        _ => Err(format!("未知工具: {name}")),
    }
}

fn format_report(text: &str, root: &std::path::Path) -> String {
    let mut report = String::from(if text.trim().is_empty() {
        "(无输出)"
    } else {
        text
    });
    report.push_str(&format!("\n\n输出目录: {}", root.join("output").display()));
    report
}
