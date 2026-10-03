use std::io::{BufRead, BufReader};
use std::path::Path;
use std::process::{Command, Stdio};

use serde_json::{json, Value};
use tauri::{AppHandle, Emitter};

use crate::paths;

#[cfg(windows)]
use std::os::windows::process::CommandExt;
#[cfg(windows)]
const CREATE_NO_WINDOW: u32 = 0x0800_0000;

/// 一次 python 调用的结果
pub struct RunOutcome {
    pub ok: bool,
    pub text: String,
}

/// 调用 magic_convert.py --json,逐行解析事件并回调(GUI 实时上屏 / MCP 收集)
pub fn run_python<F>(args: &[String], mut on_event: F) -> Result<RunOutcome, String>
where
    F: FnMut(Value),
{
    let root = paths::find_root();
    let py = paths::python_exe(&root).ok_or_else(|| {
        format!(
            "未找到 Python 引擎: {}",
            root.join("python_env").join("python.exe").display()
        )
    })?;
    let script = root.join("magic_convert.py");
    if !script.exists() {
        return Err(format!("未找到转换脚本: {}", script.display()));
    }

    let mut cmd = Command::new(&py);
    cmd.arg(&script)
        .arg("--json")
        .args(args)
        .current_dir(&root)
        .env("PYTHONIOENCODING", "utf-8")
        .env("PYTHONUTF8", "1")
        .stdout(Stdio::piped())
        .stderr(Stdio::piped());
    #[cfg(windows)]
    cmd.creation_flags(CREATE_NO_WINDOW);

    let mut child = cmd.spawn().map_err(|e| format!("启动 Python 失败: {e}"))?;

    let mut text = String::new();
    let mut ok = true;

    if let Some(stdout) = child.stdout.take() {
        for line in BufReader::new(stdout).lines() {
            let line = match line {
                Ok(l) => l,
                Err(_) => break,
            };
            let trimmed = line.trim();
            if trimmed.is_empty() {
                continue;
            }
            let ev = serde_json::from_str::<Value>(trimmed)
                .unwrap_or_else(|_| json!({ "type": "info", "message": trimmed }));
            if ev.get("type").and_then(Value::as_str) == Some("error") {
                ok = false;
            }
            push_text(&mut text, &ev);
            on_event(ev);
        }
    }

    let status = child
        .wait()
        .map_err(|e| format!("等待 Python 结束失败: {e}"))?;
    if !status.success() && ok {
        ok = false;
        let ev = json!({ "type": "error", "message": format!("Python 进程异常退出,退出码 {}", status.code().unwrap_or(-1)) });
        push_text(&mut text, &ev);
        on_event(ev);
    }

    Ok(RunOutcome { ok, text })
}

/// 把事件渲染成人类可读的一行文本
fn push_text(text: &mut String, ev: &Value) {
    let kind = ev.get("type").and_then(Value::as_str).unwrap_or("info");
    let msg = ev.get("message").and_then(Value::as_str).unwrap_or("");
    let file = ev.get("file").and_then(Value::as_str);
    let line = match (kind, file) {
        ("start", Some(f)) => format!("🔄 {f}"),
        ("success", Some(f)) => format!("✅ {f}"),
        ("error", Some(f)) => format!("❌ {f}: {msg}"),
        _ => msg.to_string(),
    };
    if !line.is_empty() {
        if !text.is_empty() {
            text.push('\n');
        }
        text.push_str(&line);
    }
}

fn emit_convert_event(app: &AppHandle, ev: &Value) {
    let _ = app.emit("convert-event", ev);
}

fn emit_busy(app: &AppHandle, busy: bool) {
    let _ = app.emit("convert-state", json!({ "busy": busy }));
}

fn dir_str(p: &Path) -> String {
    p.to_string_lossy().into_owned()
}

fn ensure_output_dir(root: &Path) -> std::io::Result<std::path::PathBuf> {
    let out = root.join("output");
    std::fs::create_dir_all(&out)?;
    Ok(out)
}

/// 拖拽/选择的一批文件 → 逐个 single 转换
#[tauri::command]
pub async fn convert_paths(app: AppHandle, paths: Vec<String>) -> Result<Value, String> {
    if paths.is_empty() {
        return Err("没有收到任何文件路径".into());
    }
    emit_busy(&app, true);
    let result = (|| {
        let root = paths::find_root();
        let out_dir = dir_str(&ensure_output_dir(&root).map_err(|e| format!("创建 output 失败: {e}"))?);

        let mut converted = 0usize;
        let mut text = String::new();
        for p in &paths {
            let argv = vec!["single".to_string(), p.clone(), out_dir.clone()];
            let outcome = run_python(&argv, |ev| emit_convert_event(&app, &ev))?;
            if outcome.ok {
                converted += 1;
            }
            if !text.is_empty() {
                text.push('\n');
            }
            text.push_str(&outcome.text);
        }
        Ok(json!({ "ok": converted > 0, "converted": converted, "total": paths.len(), "detail": text }))
    })();
    emit_busy(&app, false);
    result
}

/// 批量转换 input/ 目录;exts 为空 → 图片合并模式
#[tauri::command]
pub async fn convert_batch(app: AppHandle, exts: String, cat: String) -> Result<Value, String> {
    emit_busy(&app, true);
    let result = (|| {
        let root = paths::find_root();
        let out_dir = dir_str(&ensure_output_dir(&root).map_err(|e| format!("创建 output 失败: {e}"))?);
        let input_dir = dir_str(&root.join("input"));

        let images_mode = exts.trim().is_empty();
        let mode = if images_mode { "batch_images" } else { "batch" };
        let ext_arg = if images_mode { "NONE".to_string() } else { exts.clone() };
        let argv = vec![mode.to_string(), input_dir, out_dir, ext_arg, cat];
        let outcome = run_python(&argv, |ev| emit_convert_event(&app, &ev))?;
        Ok(json!({ "ok": outcome.ok, "detail": outcome.text }))
    })();
    emit_busy(&app, false);
    result
}

#[tauri::command]
pub fn open_output() -> Result<(), String> {
    let root = paths::find_root();
    let out = ensure_output_dir(&root).map_err(|e| format!("创建 output 失败: {e}"))?;

    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        let _ = Command::new("explorer")
            .arg(out.as_os_str())
            .creation_flags(0x0800_0000)
            .spawn();
    }
    #[cfg(not(windows))]
    {
        let _ = Command::new("xdg-open").arg(out.as_os_str()).spawn();
    }
    Ok(())
}
