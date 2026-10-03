"""MCP stdio 冒烟测试: 模拟 AI 客户端调用 markitdown-quickstyle.exe mcp

用法:
    python_env\\python.exe app\\scripts\\mcp_smoke.py '[["convert_file", {"file_path": "..."}], ...]'
不带参数则只做 initialize / tools/list 检查。
"""
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent.parent
EXE = ROOT / "app" / "src-tauri" / "target" / "debug" / "markitdown-quickstyle.exe"


def rpc(method, msg_id=None, params=None):
    msg = {"jsonrpc": "2.0", "method": method}
    if params is not None:
        msg["params"] = params
    if msg_id is not None:
        msg["id"] = msg_id
    return json.dumps(msg, ensure_ascii=False)


def run(calls):
    requests = [
        rpc("initialize", 1, {"protocolVersion": "2025-03-26"}),
        rpc("notifications/initialized"),
        rpc("tools/list", 99),
    ]
    for i, (name, args) in enumerate(calls, start=100):
        requests.append(rpc("tools/call", i, {"name": name, "arguments": args}))

    proc = subprocess.run(
        [str(EXE), "mcp"],
        input="\n".join(requests) + "\n",
        capture_output=True,
        text=True,
        encoding="utf-8",
        cwd=str(ROOT),
        timeout=600,
    )
    replies = [json.loads(l) for l in proc.stdout.splitlines() if l.strip()]
    by_id = {r.get("id"): r for r in replies if "id" in r}

    init = by_id.get(1, {})
    server = init.get("result", {}).get("serverInfo", {})
    print(f"initialize : {server.get('name')} v{server.get('version')}")
    tools = [t["name"] for t in by_id.get(99, {}).get("result", {}).get("tools", [])]
    print(f"tools/list : {tools}")

    for (name, _), i in zip(calls, range(100, 100 + len(calls))):
        resp = by_id.get(i)
        if resp is None:
            print(f"\n[{name}] 无响应!")
            continue
        result = resp.get("result", {})
        text = result.get("content", [{}])[0].get("text", "")
        status = "ERROR" if result.get("isError") else "ok"
        print(f"\n[{name}] {status}\n{text[:500]}")

    if proc.stderr.strip():
        print(f"\n[stderr] {proc.stderr.strip()[:300]}")


if __name__ == "__main__":
    calls = [tuple(c) for c in json.loads(sys.argv[1])] if len(sys.argv) > 1 else []
    run(calls)
