# Brita Yield Automation – Copilot Instructions

## Running the Tool

```bat
# Recommended: double-click or run from terminal
run_Brita_yield_gen.bat

# Direct Python
python Yield_analysis_gen.py
```

On first run, `Yield_analysis_gen.py` auto-creates a `venv/` with `--system-site-packages`, installs `requirements.txt`, then **re-launches itself inside the venv**. Subsequent runs skip installation and relaunch directly.

There are no automated tests or lint commands in this repo.

## Architecture

The pipeline has two Python entry points:

- **`Yield_analysis_gen.py`** – Orchestrator. Runs the full pipeline end-to-end.
- **`Yield_ppt_report.py`** – PowerPoint generator. Reads CSVs from `output/` and builds `output/Brita_yield_report.pptx`.

### Pipeline stages (in order)

1. **Prerequisites** – verify Python ≥ 3.10 and Brita+ install.
2. **Config validation** – load `input/config.json`; requires a `tasks` list.
3. **Venv bootstrap** – create/activate venv; re-launch inside it.
4. **GetXeusTestResult** – run `GetXeusTestResult/ApiSamples.GetXeusTestResult.exe` against `input/lot_wafer.csv` to obtain `tp_path`, `tp_name`, and `operations`. Retries every `check_period` minutes (up to `check_count` times; `null` = infinite).
5. **TP file validation** – confirm `.tpl` and `.stpl` exist at `tp_path`.
6. **DB query** – query Intel fab databases via **PyUber** (Intel-internal package; must be in system Python) using SQL against `A_Testing_Session` / `A_Lot_At_Operation` / `A_Device_Testing`.
7. **Per-task loop** – each task in `config.tasks` runs independently; one failure does not stop others:
   - `"type": "Brita"` → run Brita+ Batchmode with the task's `indicator_file`
   - `"type": "Plus"` → run AquaCmdLine.exe (network share) → TraceTestData.exe (network share)
   - Then run the task's `script` (`.jsl` launches JMP asynchronously; `.py` runs synchronously and receives `-output_dir` and optional `-filter_dict` JSON args)

All output lands in `output/<task_name>/`.

## Key Conventions

### `input/config.json` schema
```json
{
  "product": "...",
  "check_period": 10,
  "check_count": null,
  "tasks": [
    {
      "name": "TASK_NAME",
      "type": "Brita",
      "indicator_file": "my_indicator.xml",
      "script": "my_script.jsl",
      "filter_dict": {}
    }
  ]
}
```
- `tp_path`, `tp_name`, `operations` are **injected at runtime** by `run_get_xeus_test_result()` — do not put them in `config.json`.
- Task `name` becomes the output subdirectory name (`output/<name>/`).
- `indicator_file` path is resolved relative to `input/`.
- `script` must be `.jsl` (JMP, launched async) or `.py` (Python, run sync).

### Lot prefix → database routing
Lot ID first character determines the PyUber datasource:

| Prefix(es) | Datasource |
|---|---|
| 3, 4, D | D1D_PROD_XARIES |
| C, Q, H | F24_PROD_XEUS |
| L | F32_PROD_XEUS |
| N | F28_PROD_XEUS |
| Z | D1C_PROD_XEUS |

This mapping lives in `LOT_PREFIX_MAP` near the top of `Yield_analysis_gen.py`.

### Output file naming
Brita tasks produce `output/<task_name>/<task_name>_units_batch.csv` and `<task_name>_units_batch_summary.csv`. Plus tasks produce `output/<task_name>/aqua_result.csv` and `ttd_result.csv`.

### External tool paths (hard-coded constants)
```
BRITA_EXE  = C:\Program Files\Brita+\BRITA+.exe
JMP_EXE    = C:\Program Files\SAS\JMPPRO\17\jmp.exe
AQUA_EXE   = \\PGSAPP3301.gar.corp.intel.com\Installer\...\AquaCmdLine.exe   (network)
TTD_EXE    = \\amr.corp.intel.com\ec\proj\...\TraceTestData.exe              (network)
```

### PyUber
Intel-internal Oracle wrapper. Must be installed in the **system Python** (not the venv) because the venv is created with `--system-site-packages`. Import failures here mean PyUber is missing from system Python.

### `Yield_ppt_report.py` slide layout
Uses python-pptx with a fixed 13.33 × 7.5 inch slide (widescreen). Color constants `INTEL_BLUE`, `LIGHT_BLUE`, `WHITE`, `DARK_GREY`, `MID_GREY` are defined at module level. Every slide is blank layout (index 6) with a manually placed Intel-blue header bar and footer.
