# Brita Yield Automation Tool

This tool automates the end-to-end yield analysis pipeline for semiconductor wafer lots. Given a list of lots and wafers, it queries the test result database, runs Brita+ Batchmode to calculate yield, launches per-task scripts, and generates a consolidated PowerPoint report — all in a single command. It handles virtual environment setup and dependency installation automatically on first run.

```
GetXeusTestResult → DB Query → Brita+ Batchmode (per task) → Scripts → PowerPoint Report
```

## Prerequisites

| Software | Path |
|----------|------|
| Python 3.10+ | Installed on system |
| Brita+ | `C:\Program Files\Brita+\BRITA+.exe` |
| JMP Pro 17 | `C:\Program Files\SAS\JMPPRO\17\jmp.exe` |

## Repository Structure

```
Brita_yield_report/
├── Yield_analysis_gen.py    # Main orchestrator
├── Yield_ppt_report.py      # PowerPoint report generator
├── run_Brita_yield_gen.bat  # Windows launcher (double-click to run)
├── requirements.txt
├── GetXeusTestResult/       # GetXeusTestResult.exe and config files
├── input/
│   ├── config.json          # Run configuration (required)
│   ├── lot_wafer.csv        # Lot / wafer list to query
│   ├── *.xml                # Brita+ indicator files
│   └── *.jsl / *.py         # Per-task scripts
└── output/                  # Auto-created; all outputs saved here
    └── <task_name>/         # One subdirectory per task
```

## Configuration `input/config.json`

```json
{
    "product": "CRI",
    "check_period": 10,
    "check_count": null,
    "tasks": [
        {
            "name": "SCN_GT",
            "type": "Brita",
            "indicator_file": "MyIndicator.xml",
            "script": "my_script.jsl"
        },
        {
            "name": "MY_PLUS_TASK",
            "type": "Plus",
            "script": "post_process.py",
            "filter_dict": {"key": "value"}
        }
    ]
}
```

| Field | Description |
|-------|-------------|
| `product` | Product identifier (informational) |
| `check_period` | Minutes between retries when waiting for lots to appear in GetXeusTestResult (default: `10`) |
| `check_count` | Max number of retries; set to `null` for unlimited |
| `tasks` | List of analysis tasks to run (see below) |

> `tp_path`, `tp_name`, and `operations` are retrieved automatically from
> `GetXeusTestResult.exe` and do not need to be specified manually.

### Task fields

| Field | Required | Description |
|-------|----------|-------------|
| `name` | ✓ | Task name — also becomes the output subdirectory (`output/<name>/`) |
| `type` | ✓ | `"Brita"` (runs Brita+ Batchmode) or `"Plus"` (runs AquaCmdLine → TraceTestData) |
| `indicator_file` | Brita only | Indicator `.xml` filename, resolved from `input/` |
| `script` | | Post-processing script: `.jsl` (launched in JMP asynchronously) or `.py` (run synchronously); resolved from `input/` |
| `filter_dict` | `.py` scripts only | Optional JSON dict passed to the script via `-filter_dict` argument |

## Input File `input/lot_wafer.csv`

```csv
Lot,Wafer
Q503E4301,
Q503E4302,01
```

Leave the `Wafer` column empty to query all wafers for a given lot.

## Usage

### Windows (Recommended)

Double-click **`run_Brita_yield_gen.bat`** or run it from a terminal:

```bat
run_Brita_yield_gen.bat
```

The batch file clears the `output/` folder, launches `Yield_analysis_gen.py`, displays start/end timestamps and total duration, and pauses at the end so you can review the output before the window closes.

### Command Line

```bash
python Yield_analysis_gen.py
```

On first run the script will automatically:
1. Create a `venv` (inheriting the system Python environment so PyUber is accessible)
2. Install dependencies from `requirements.txt`
3. Re-launch itself inside the venv and continue execution

On subsequent runs, the existing venv is activated directly without reinstalling packages.

## Execution Steps

| Step | Description | Output |
|------|-------------|--------|
| 1. Prerequisites | Verify Brita+ installation and Python ≥ 3.10 | — |
| 2. Config validation | Check `tasks` list and required fields per task | — |
| 3. Virtual environment | Create venv and install dependencies | `venv/` |
| 4. GetXeusTestResult | Run exe to retrieve `tp_path`, `tp_name`, `operations` from lot data; retries until all lots found | `output/xeus_test_result.csv` |
| 5. Validate TP files | Confirm `.tpl` and `.stpl` files exist at `tp_path` | — |
| 6. DB query | Query lot/wafer data via PyUber | `output/lot_wafer_search_result.csv` |
| 7. Per-task loop | For each task: run Brita+ Batchmode (type=Brita) or AquaCmdLine→TraceTestData (type=Plus), then run optional script | `output/<task_name>/` |

Each task runs independently — a failure in one task does not stop the others. A summary of pass/fail per task is printed at the end.

## PowerPoint Report (`Yield_ppt_report.py`)

Run separately to generate `output/Brita_yield_report.pptx`. The report contains one slide per analysis type:

| Slide | Content |
|-------|---------|
| 1 | Test Run Summary — lot info, operation, part type, wafer list |
| 2 *(optional)* | SOC_SCAN Results |
| 3 *(optional)* | MBIST Results |
| 4 *(optional)* | ARRGT Results |
| 5 *(optional)* | MIO Results |
| 6 | SCN_GT Yield Analysis |
| 7 | SCN_GT_ATPG Yield Analysis |
| 8 | SCN_GT_TATPG Yield Analysis |

Each SCN_GT slide shows the JMP yield graph image on the left and an attachment panel (with the `.jrn` file icon) on the right, matching the SOC_SCAN layout.

## Logs

Logs are written to both the console and `output/yield_automation.log`.

## Lot Prefix → Database Mapping

| Lot Prefix | Data Source |
|------------|-------------|
| 3, 4, D | D1D_PROD_XARIES |
| C, H, Q | F24_PROD_XEUS |
| L | F32_PROD_XEUS |
| N | F28_PROD_XEUS |
| Z | D1C_PROD_XEUS |
