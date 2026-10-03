import { invoke } from '@tauri-apps/api/core';
import { listen } from '@tauri-apps/api/event';
import { getCurrentWebview } from '@tauri-apps/api/webview';
import { open } from '@tauri-apps/plugin-dialog';

/** Python 转换脚本 --json 模式吐出的逐行事件 */
interface ConvertEvent {
  type: 'info' | 'start' | 'success' | 'error';
  message?: string;
  file?: string;
}

/** 日志行(比 ConvertEvent 多一个本地合成类型 done) */
interface LogLine {
  type: string;
  message?: string;
  file?: string;
}

interface ConvertPathsResult {
  ok: boolean;
  converted: number;
  total: number;
  detail: string;
}

interface BatchResult {
  ok: boolean;
  detail: string;
}

interface AppStatus {
  root: string;
  pythonOk: boolean;
  inputExists: boolean;
  outputExists: boolean;
  version: string;
}

const logEl = document.getElementById('log') as HTMLDivElement;
const dropEl = document.getElementById('drop') as HTMLDivElement;
const statusEl = document.getElementById('status') as HTMLDivElement;
const catBtns = document.querySelectorAll<HTMLButtonElement>('button.cat');
const openOutputBtn = document.getElementById('open-output') as HTMLButtonElement;
const clearLogBtn = document.getElementById('clear-log') as HTMLButtonElement;

let busy = false;

function appendLog(ev: LogLine): void {
  const kind = ev.type || 'info';
  let text = ev.message || ev.file || '';
  if (kind === 'start' && ev.file) text = '🔄 ' + ev.file;
  if (kind === 'success' && ev.file) text = '✅ ' + ev.file;
  if (kind === 'error') text = '❌ ' + (ev.file ? ev.file + ':' : '') + ' ' + (ev.message || '');
  if (!text) return;
  const div = document.createElement('div');
  div.className = 'line ' + kind;
  div.textContent = text;
  logEl.appendChild(div);
  while (logEl.children.length > 300 && logEl.firstChild) {
    logEl.removeChild(logEl.firstChild);
  }
  logEl.scrollTop = logEl.scrollHeight;
}

function setBusy(v: boolean): void {
  busy = v;
  document.body.classList.toggle('busy', v);
  catBtns.forEach((b) => (b.disabled = v));
  dropEl.classList.toggle('disabled', v);
}

async function refreshStatus(): Promise<void> {
  try {
    const s = await invoke<AppStatus>('get_status');
    statusEl.textContent = s.pythonOk ? `引擎就绪 · ${s.root}` : `⚠️ 未找到 Python 引擎 · ${s.root}`;
    statusEl.classList.toggle('bad', !s.pythonOk);
  } catch (e) {
    statusEl.textContent = '状态获取失败: ' + e;
  }
}

async function runSingle(paths: string[]): Promise<void> {
  if (busy || !paths.length) return;
  setBusy(true);
  try {
    const res = await invoke<ConvertPathsResult>('convert_paths', { paths });
    appendLog({ type: 'done', message: `—— 完成,成功 ${res.converted}/${res.total} 个 ——` });
  } catch (e) {
    appendLog({ type: 'error', message: String(e) });
  } finally {
    setBusy(false);
  }
}

async function runBatch(btn: HTMLButtonElement): Promise<void> {
  if (busy) return;
  setBusy(true);
  try {
    const res = await invoke<BatchResult>('convert_batch', {
      exts: btn.dataset.exts ?? '',
      cat: btn.dataset.cat ?? '',
    });
    appendLog({
      type: res.ok ? 'done' : 'error',
      message: res.ok
        ? `—— ${btn.textContent?.trim()} 批量完成 ——`
        : '—— 批量转换中存在失败项,详见上方日志 ——',
    });
  } catch (e) {
    appendLog({ type: 'error', message: String(e) });
  } finally {
    setBusy(false);
  }
}

async function init(): Promise<void> {
  await refreshStatus();

  await listen<ConvertEvent>('convert-event', (e) => appendLog(e.payload));

  // 拖拽:整个窗口区域接收文件
  await getCurrentWebview().onDragDropEvent((event) => {
    const t = event.payload.type;
    if (t === 'over' || t === 'enter') {
      dropEl.classList.add('over');
      return;
    }
    dropEl.classList.remove('over');
    if (t === 'drop' && event.payload.paths.length) {
      runSingle([...event.payload.paths]);
    }
  });

  dropEl.addEventListener('click', () => {
    if (busy) return;
    open({ multiple: true, title: '选择要转换的文件' })
      .then((files) => {
        if (Array.isArray(files) && files.length) runSingle(files);
      })
      .catch(() => {
        /* 用户取消选择 */
      });
  });

  catBtns.forEach((btn) => btn.addEventListener('click', () => runBatch(btn)));

  openOutputBtn.addEventListener('click', () => invoke('open_output'));
  clearLogBtn.addEventListener('click', () => (logEl.innerHTML = ''));
}

init();
