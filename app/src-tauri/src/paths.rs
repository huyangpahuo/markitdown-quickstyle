use std::path::{Path, PathBuf};

/// 定位工具根目录(包含 magic_convert.py 与 python_env/ 的那层)。
///
/// 依次尝试:环境变量 MARKITDOWN_ROOT → 当前工作目录向上找 → exe 所在目录向上找。
/// 因此无论 dev 模式(cwd 在 src-tauri)还是把 exe 拷到根目录,都能找到 Python 引擎。
pub fn find_root() -> PathBuf {
    if let Ok(p) = std::env::var("MARKITDOWN_ROOT") {
        let p = PathBuf::from(p);
        if p.join("magic_convert.py").exists() {
            return p;
        }
    }

    let mut bases: Vec<PathBuf> = Vec::new();
    if let Ok(cwd) = std::env::current_dir() {
        bases.push(cwd);
    }
    if let Ok(exe) = std::env::current_exe() {
        if let Some(dir) = exe.parent() {
            bases.push(dir.to_path_buf());
        }
    }

    for base in bases {
        let mut dir: Option<&Path> = Some(base.as_path());
        while let Some(d) = dir {
            if d.join("magic_convert.py").exists() {
                return d.to_path_buf();
            }
            dir = d.parent();
        }
    }

    std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."))
}

/// 内嵌 Python;不存在时退回系统 PATH 里的 python
pub fn python_exe(root: &Path) -> Option<PathBuf> {
    let bundled = root.join("python_env").join("python.exe");
    if bundled.exists() {
        Some(bundled)
    } else {
        None
    }
}
