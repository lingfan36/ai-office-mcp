# Benchmark Results

Single-run capture for the README's measured numbers. Re-run any time with the
commands shown — results are timestamped, not pinned.

## Environment

| Field | Value |
|---|---|
| Date | 2026-05-08 |
| OS | Windows 11 Home China 10.0.26200 |
| Python | 3.13.12 |
| Office COM | Excel 16.0 / Word 16.0 / PowerPoint 16.0 |
| FastMCP | 3.2.4 |

## Tool registration (measured, not grep-estimated)

```
$ python benchmarks/bench_startup.py
```

| Server | Cold-start (median of 3 fresh subprocesses) | In-process import | Tools registered |
|---|---:|---:|---:|
| `excel.excel_mcp.main`        | 1678.1 ms | 1461.1 ms | **152** |
| `word_document_server.main`   | 1214.1 ms |  219.0 ms | **205** |

PPT skill (workflow-based, not MCP server):

| Asset | Count |
|---|---:|
| Python utility scripts (`skills/ppt-master/scripts/`) | 12 |
| Workflow markdown files (`skills/ppt-master/workflows/`) | 3 |
| Reference docs (`skills/ppt-master/references/`) | 9 |

## End-to-end tests

```
$ cd excel
$ python test_excel_ops.py        # 32/32 passed in  5.8s
$ python test_accounting.py       # 197/197 passed in 22.1s
$ python test_incremental.py      # 56/56 passed in 13.4s
$ cd ../word-mcp-live
$ $env:PYTHONIOENCODING="utf-8"; $env:PYTHONUTF8="1"
$ python test_formatting.py       # ✅ all tests completed in 0.4s
$ pytest tests/                   # 1/1 passed in 10.8s
```

| Suite | Result | Time |
|---|---:|---:|
| `excel/test_excel_ops.py`            | 32 / 32  ✅ | 5.8s |
| `excel/test_accounting.py`           | 197 / 197 ✅ | 22.1s |
| `excel/test_incremental.py`          | 56 / 56  ✅ | 13.4s |
| `word-mcp-live/test_formatting.py`   | passed   ✅ | 0.4s |
| `word-mcp-live/tests/test_convert_to_pdf.py` | 1 / 1 ✅ | 10.8s |
| **Excel total**                      | **285 / 285** | **41.3s** |
| **Word total**                       | **2 / 2** | **11.2s** |
| **Grand total**                      | **287 / 287 (100%)** | **52.5s** |

## Notes

- Excel suite spawns COM automation against a real Excel.exe; window briefly appears.
- `test_excel_ops.py` exercises file/range/sheet/table/pivot/chart/undo/calc lifecycle.
- `test_accounting.py` is the largest single suite — 197 assertions covering number
  formats, formulas, totals row, currency, accounting symbol alignment, rounding.
- `test_incremental.py` validates snapshot undo across incremental writes.
- `test_convert_to_pdf.py` performs an actual `.docx` → `.pdf` conversion via Word COM.

## Reproducibility

This file is not regenerated automatically. Run the commands above and update the
table if you want a fresh capture. The benchmark script in `bench_startup.py` is
deterministic enough that import times will land within ±10% across runs on the
same machine.
