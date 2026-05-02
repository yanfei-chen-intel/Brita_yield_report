# Yield Automation Tool

Automates the full yield analysis pipeline:

```
GetXeusTestResult → DB Query → Brita+ Batchmode → JMP Yield Report
```

## Prerequisites

| Software | Path |
|----------|------|
| Python 3.10+ | Installed on system |
| Brita+ | `C:\Program Files\Brita+\BRITA+.exe` |
| JMP Pro 17 | `C:\Program Files\SAS\JMPPRO\17\jmp.exe` |

## Repository Structure

```
yield_automation/
├── main.py
├── requirements.txt
├── GetXeusTestResult/       # GetXeusTestResult.exe and config files
├── input/
│   ├── config.json          # Run configuration (required)
│   ├── lot_wafer.csv        # Lot / wafer list to query
│   ├── indicator/           # Brita+ indicator files
│   └── jmp_script/          # JMP JSL scripts
└── output/                  # Auto-created; all outputs saved here
```

## Configuration `input/config.json`

```json
{
    "indicator_file": "bmgg31_brita_type2_indicator_arraygt.xml",
    "jsl_file":       "ARRHRY_Script_Final.jsl"
}
```

| Field | Description |
|-------|-------------|
| `indicator_file` | Brita+ indicator filename, placed under `input/indicator/` |
| `jsl_file` | JMP JSL script filename, placed under `input/jmp_script/` |

> `tp_path`, `tp_name`, and `operations` are retrieved automatically from
> `GetXeusTestResult.exe` and do not need to be specified manually.

## Input File `input/lot_wafer.csv`

```csv
Lot,Wafer
Q503E4301,
Q503E4302,01
```

Leave the `Wafer` column empty to query all wafers for a given lot.

## Usage

```bash
python main.py
```

On first run the script will automatically:
1. Create a `venv` (inheriting the system Python environment)
2. Install dependencies from `requirements.txt` (including PyUber)
3. Re-launch itself inside the venv and continue execution

On subsequent runs, the existing venv is activated directly without reinstalling packages.

## Execution Steps

| Step | Description | Output |
|------|-------------|--------|
| 1. Prerequisites | Verify Brita+ installation and Python version | — |
| 2. Config validation | Check required fields in config.json | — |
| 3. Virtual environment | Create venv and install dependencies | `venv/` |
| 4. GetXeusTestResult | Run exe to retrieve `tp_path`, `tp_name`, `operations` from lot data | `output/xeus_test_result.csv` |
| 5. Validate TP files | Confirm `.tpl` and `.stpl` files exist | — |
| 6. DB query | Query lot/wafer data via PyUber | `output/lot_wafer_search_result.csv` |
| 7. Brita+ | Run Batchmode to calculate yield | `output/units_batch.csv` |
| 8. JMP report | Launch JMP asynchronously with JSL script | Yield analysis report |

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
