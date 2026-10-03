// 不带控制台的 GUI 程序;release 下 stdout/stdin 仍可通过父进程管道工作(MCP 模式依赖这一点)
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod convert;
mod mcp;
mod paths;

use serde_json::{json, Value};
use tauri::Manager;

fn main() {
    // `markitdown-quickstyle.exe mcp` → 作为 stdio MCP server 运行,不启动任何窗口
    let args: Vec<String> = std::env::args().collect();
    if args.iter().skip(1).any(|a| a == "mcp" || a == "--mcp") {
        mcp::run();
        return;
    }

    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            convert::convert_paths,
            convert::convert_batch,
            convert::open_output,
            show_main,
            get_status
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

/// 桌宠被单击时:唤起并聚焦主窗口
#[tauri::command]
fn show_main(app: tauri::AppHandle) {
    if let Some(w) = app.get_webview_window("main") {
        let _ = w.show();
        let _ = w.unminimize();
        let _ = w.set_focus();
    }
}

#[tauri::command]
fn get_status() -> Value {
    let root = paths::find_root();
    json!({
        "root": root.to_string_lossy(),
        "pythonOk": paths::python_exe(&root).is_some(),
        "inputExists": root.join("input").exists(),
        "outputExists": root.join("output").exists(),
        "version": env!("CARGO_PKG_VERSION"),
    })
}
