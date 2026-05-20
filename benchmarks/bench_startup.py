"""Cold-start + tool-listing benchmark for the bundled MCP servers.

Measures:
  1. Time to import the server module (cold start)
  2. Number of tools registered
  3. Time to enumerate every tool via the FastMCP API

Run:
    python benchmarks/bench_startup.py
"""
import os, sys, time, json, statistics, subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

def _measure_import(label, cwd, module):
    """Spawn a fresh Python process to get a clean cold-start number."""
    code = f"import time;t=time.perf_counter();import {module};print(time.perf_counter()-t)"
    samples = []
    for _ in range(3):
        out = subprocess.run(
            [sys.executable, "-c", code],
            cwd=cwd, capture_output=True, text=True,
            env={**os.environ, "PYTHONPATH": cwd},
        )
        if out.returncode != 0:
            return label, None, out.stderr.strip().splitlines()[-1] if out.stderr else "unknown error"
        samples.append(float(out.stdout.strip().splitlines()[-1]))
    return label, statistics.median(samples), None

def _count_tools_excel():
    cwd = str(ROOT / "excel")
    sys.path.insert(0, cwd)
    os.environ.setdefault("EXCEL_SNAPSHOT_DIR", str(ROOT / "excel" / ".snapshots"))
    t = time.perf_counter()
    from excel_mcp.main import mcp
    elapsed = time.perf_counter() - t
    import asyncio
    tools = asyncio.run(mcp._list_tools())
    names = [t.name for t in tools]
    return len(names), elapsed, names[:10]

def _count_tools_word():
    cwd = str(ROOT / "word-mcp-live")
    sys.path.insert(0, cwd)
    t = time.perf_counter()
    from word_document_server.main import mcp
    from word_document_server.tool_registry import register_all_tools
    register_all_tools(mcp)
    elapsed = time.perf_counter() - t
    import asyncio
    tools = asyncio.run(mcp._list_tools())
    names = [t.name for t in tools]
    return len(names), elapsed, names[:10]

if __name__ == "__main__":
    print("=" * 64)
    print("AI Office MCP Suite — startup benchmark")
    print(f"Python: {sys.version.split()[0]}  Platform: {sys.platform}")
    print("=" * 64)

    print("\n[1] Cold-start import (median of 3 fresh subprocesses)\n")
    rows = []
    for label, cwd, mod in [
        ("excel.excel_mcp.main", str(ROOT / "excel"), "excel_mcp.main"),
        ("word_document_server.main", str(ROOT / "word-mcp-live"), "word_document_server.main"),
    ]:
        name, secs, err = _measure_import(label, cwd, mod)
        rows.append((name, secs, err))
        if err:
            print(f"  [FAIL] {name}  →  {err}")
        else:
            print(f"  [OK]   {name:32s}  {secs*1000:7.1f} ms")

    print("\n[2] Tool counts (loaded in-process, not subprocess)\n")
    try:
        n, secs, sample = _count_tools_excel()
        print(f"  [OK]   excel       tools={n:3d}   import={secs*1000:7.1f} ms")
        print(f"         sample: {', '.join(sample)}")
    except Exception as e:
        print(f"  [FAIL] excel       {e!r}")

    try:
        n, secs, sample = _count_tools_word()
        print(f"  [OK]   word        tools={n:3d}   import={secs*1000:7.1f} ms")
        print(f"         sample: {', '.join(sample)}")
    except Exception as e:
        print(f"  [FAIL] word        {e!r}")

    print("\n[3] PPT skill — workflow-based, not MCP server\n")
    skills_dir = ROOT / "ppt-master-main" / "skills" / "ppt-master"
    scripts = list((skills_dir / "scripts").glob("*.py"))
    workflows = list((skills_dir / "workflows").glob("*.md"))
    refs = list((skills_dir / "references").glob("*.md"))
    print(f"  scripts:    {len(scripts)} python utility scripts")
    print(f"  workflows:  {len(workflows)} workflow markdown files")
    print(f"  references: {len(refs)} reference docs")

    print("\n" + "=" * 64)
