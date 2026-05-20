"""
Yield Automation Tool
Orchestrates: data search → Brita+ Batchmode → JMP report generation
"""

import os
import sys
import json
import logging
import subprocess
import pathlib

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
BRITA_EXE   = r"C:\Program Files\Brita+\BRITA+.exe"
JMP_EXE     = r"C:\Program Files\SAS\JMPPRO\17\jmp.exe"
XEUS_EXE    = pathlib.Path("GetXeusTestResult") / "ApiSamples.GetXeusTestResult.exe"
AQUA_EXE    = pathlib.Path(r"\\PGSAPP3301.gar.corp.intel.com\Installer\AquaHbase\AquaCMDClient\Client\AquaCmdLine.exe")
TTD_EXE     = pathlib.Path(r"\\amr.corp.intel.com\ec\proj\MDO\Global\YBSP\TraceTestData\Prod\TraceTestData.exe")
CONFIG_PATH = pathlib.Path("input") / "config.json"
LOT_WAFER_CSV = pathlib.Path("input") / "lot_wafer.csv"
VENV_DIR    = pathlib.Path("venv")
OUTPUT_DIR  = pathlib.Path("output")

REQUIRED_CONFIG_FIELDS = ["tasks"]
INPUT_DIR  = pathlib.Path("input")

# lot-prefix → database source mapping (Intel sites)
LOT_PREFIX_MAP = {
    ("4", "D", "3"): "D1D_PROD_XARIES",
    ("C", "Q", "H"): "F24_PROD_XEUS",
    ("L",):          "F32_PROD_XEUS",
    ("N",):          "F28_PROD_XEUS",
    ("Z",):          "D1C_PROD_XEUS",
}


def _lot_to_datasource(lot: str) -> str:
    for prefixes, src in LOT_PREFIX_MAP.items():
        if lot.startswith(prefixes):
            return src
    raise ValueError(
        f"Unknown lot prefix for lot '{lot}'. "
        f"Expected first character in: 3, 4, C, D, H, L, N, Q, Z"
    )


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
def setup_logging():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler(OUTPUT_DIR / "yield_automation.log", encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
        force=True,  # allow re-init on venv relaunch
    )


# ---------------------------------------------------------------------------
# Step 1 – Prerequisites
# ---------------------------------------------------------------------------
def check_prerequisites():
    errors = []

    # Python version
    v = sys.version_info
    if v < (3, 10):
        errors.append(
            f"Python 3.10+ required, found {v.major}.{v.minor}.{v.micro}"
        )
    else:
        logging.info(f"Python {v.major}.{v.minor}.{v.micro} ✓")

    # Brita+
    if not os.path.exists(BRITA_EXE):
        errors.append(f"Brita+ not found at: {BRITA_EXE}")
    else:
        logging.info(f"Brita+ found at {BRITA_EXE} ✓")

    if errors:
        for e in errors:
            logging.error(e)
        sys.exit(1)


# ---------------------------------------------------------------------------
# Step 2 – Config validation
# ---------------------------------------------------------------------------
def load_config() -> dict:
    if not CONFIG_PATH.exists():
        logging.error(f"Config file not found: {CONFIG_PATH}")
        sys.exit(1)

    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        config = json.load(f)

    # Validate top-level tasks list
    tasks = config.get("tasks")
    if not tasks or not isinstance(tasks, list):
        logging.error("Config must contain a non-empty 'tasks' list.")
        sys.exit(1)

    errors = []
    for i, task in enumerate(tasks):
        task_label = task.get("name") or f"tasks[{i}]"
        task_type  = task.get("type", "")
        if not task.get("name"):
            errors.append(f"Task {i}: missing required field 'name'")
        if task_type not in ("Brita", "Plus"):
            errors.append(
                f"Task '{task_label}': 'type' must be 'Brita' or 'Plus', got '{task_type}'"
            )
        if task_type == "Brita" and not task.get("indicator_file"):
            errors.append(f"Task '{task_label}': Brita task requires 'indicator_file'")
        script = task.get("script")
        if script and not script.endswith((".jsl", ".py")):
            errors.append(
                f"Task '{task_label}': unsupported script extension '{pathlib.Path(script).suffix}'. "
                "Expected .jsl or .py"
            )

    if errors:
        for e in errors:
            logging.error(e)
        sys.exit(1)

    logging.info(f"Config validation passed ✓  ({len(tasks)} task(s): {[t['name'] for t in tasks]})")
    return config


# ---------------------------------------------------------------------------
# Step 3 – Virtual environment
# ---------------------------------------------------------------------------
def ensure_venv():
    """
    Create venv with --system-site-packages so Intel-internal PyUber
    (installed in the system Python) is accessible inside the venv.
    Re-launch the script inside the venv if not already running there.
    """
    venv_python = VENV_DIR.absolute() / "Scripts" / "python.exe"
    already_in_venv = (
        os.path.normcase(os.path.abspath(sys.executable))
        == os.path.normcase(str(venv_python))
    )
    if already_in_venv:
        logging.info("Running inside virtual environment ✓")
        return

    if VENV_DIR.exists():
        # venv already exists – skip installation, just re-launch inside it
        logging.info(f"Virtual environment found at {VENV_DIR}, activating ...")
    else:
        # Create venv (system-site-packages lets us reach PyUber globally)
        logging.info(f"Creating virtual environment at {VENV_DIR} ...")
        subprocess.run(
            [sys.executable, "-m", "venv", "--system-site-packages", str(VENV_DIR)],
            check=True,
        )
        logging.info("Virtual environment created.")

        # Install requirements only on first-time creation
        pip_exe = VENV_DIR / "Scripts" / "pip.exe"
        req_file = pathlib.Path("requirements.txt")
        if req_file.exists():
            logging.info("Installing packages from requirements.txt ...")
            result = subprocess.run(
                [str(pip_exe), "install", "-r", str(req_file)],
                capture_output=True, text=True,
            )
            if result.returncode != 0:
                logging.error(f"Package installation failed:\n{result.stderr}")
                sys.exit(1)
            logging.info("Packages installed successfully.")
        else:
            logging.warning("requirements.txt not found – skipping package installation.")

    # Re-launch inside venv
    logging.info("Re-launching in virtual environment ...")
    script = os.path.abspath(sys.argv[0])
    result = subprocess.run([str(venv_python), script] + sys.argv[1:])
    sys.exit(result.returncode)


# ---------------------------------------------------------------------------
# Step 4 – Get tp_path / tp_name / operations via GetXeusTestResult.exe
# ---------------------------------------------------------------------------
def run_get_xeus_test_result(config: dict) -> dict:
    """
    Run GetXeusTestResult.exe with lot_wafer.csv to retrieve TplDirectory,
    ProgramName, and Operation. Retries every check_period minutes until all
    lots are found (up to check_count times; infinite if check_count is null).
    Returns a dict with keys: tp_path, tp_name, operations.
    """
    import pandas as pd
    import time

    if not XEUS_EXE.exists():
        logging.error(f"GetXeusTestResult.exe not found at: {XEUS_EXE}")
        sys.exit(1)

    check_period = float(config.get("check_period", 10))   # minutes
    check_count  = config.get("check_count")               # None = infinite
    max_attempts = None if not check_count else int(check_count)

    # Load expected lots from lot_wafer.csv
    df_lw = pd.read_csv(LOT_WAFER_CSV, dtype=str)
    expected_lots = set(df_lw["Lot"].dropna().str.strip().tolist())

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    xeus_output = OUTPUT_DIR / "xeus_test_result.csv"
    cmd = [str(XEUS_EXE), str(LOT_WAFER_CSV), str(xeus_output)]

    attempt = 0
    while True:
        attempt += 1
        logging.info(
            f"Running GetXeusTestResult (attempt {attempt}"
            + (f"/{max_attempts}" if max_attempts else "") + ") ..."
        )
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.stdout:
            logging.info(result.stdout)
        if result.returncode != 0:
            logging.error(
                f"GetXeusTestResult.exe failed (exit code {result.returncode}):\n{result.stderr}"
            )
            sys.exit(1)

        if not xeus_output.exists():
            logging.error(f"Expected output not found: {xeus_output}")
            sys.exit(1)

        df = pd.read_csv(xeus_output, dtype=str)

        # Validate expected columns exist
        required_cols = ["TplDirectory", "ProgramName", "Operation"]
        missing_cols = [c for c in required_cols if c not in df.columns]
        if missing_cols:
            logging.error(
                f"GetXeusTestResult output is missing expected columns: {missing_cols}\n"
                f"Available columns: {list(df.columns)}"
            )
            sys.exit(1)

        # Check if all expected lots are present in the result
        found_lots = set(df["Lot"].dropna().str.strip().tolist()) if "Lot" in df.columns else set()
        missing_lots = expected_lots - found_lots
        if not missing_lots:
            logging.info("All lots found in GetXeusTestResult output ✓")
            break

        logging.warning(f"Missing lots in result: {missing_lots}")

        if max_attempts and attempt >= max_attempts:
            logging.error(
                f"Reached maximum attempts ({max_attempts}). "
                f"Lots still missing: {missing_lots}"
            )
            sys.exit(1)

        logging.info(f"Retrying in {check_period} minute(s) ...")
        time.sleep(check_period * 60)

    tp_path    = df["TplDirectory"].dropna().iloc[0]
    tp_name    = df["ProgramName"].dropna().iloc[0]
    operations = df["Operation"].dropna().unique().tolist()

    if df["TplDirectory"].nunique() > 1:
        logging.warning(f"Multiple TplDirectory values found; using first: {tp_path}")
    if df["ProgramName"].nunique() > 1:
        logging.warning(f"Multiple ProgramName values found; using first: {tp_name}")

    logging.info(f"TplDirectory  → tp_path    : {tp_path}")
    logging.info(f"ProgramName   → tp_name    : {tp_name}")
    logging.info(f"Operation(s)  → operations : {operations}")

    return {"tp_path": tp_path, "tp_name": tp_name, "operations": operations}


def validate_tp_files(tp_path: str, tp_name: str):
    """Check that .tpl and .stpl files exist under tp_path (which already includes tp_name)."""
    base = pathlib.Path(tp_path)
    errors = []
    for ext in ("tpl", "stpl"):
        fp = base / f"{tp_name}.{ext}"
        if fp.exists():
            logging.info(f"{ext.upper()} file found: {fp} ✓")
        else:
            errors.append(f"{ext.upper()} file not found: {fp}")
    if errors:
        for e in errors:
            logging.error(e)
        sys.exit(1)


# ---------------------------------------------------------------------------
# Step 5 – Lot/wafer data search
# ---------------------------------------------------------------------------
def search_lot_wafer(config: dict):
    """Query DB to verify data exists; save results to output/lot_wafer_search_result.csv."""
    try:
        import PyUber
        import pandas as pd
    except ImportError as e:
        logging.error(
            f"Required package not importable: {e}\n"
            "PyUber is an Intel-internal package. Ensure it is installed in the "
            "system Python (the venv inherits it via --system-site-packages)."
        )
        sys.exit(1)

    if not LOT_WAFER_CSV.exists():
        logging.error(f"lot_wafer.csv not found: {LOT_WAFER_CSV}")
        sys.exit(1)

    df_lw = pd.read_csv(LOT_WAFER_CSV, dtype=str)
    lot_ids  = df_lw["Lot"].dropna().str.strip().tolist()
    wafer_ids = (
        df_lw["Wafer"]
        .apply(lambda w: str(w).strip() if pd.notna(w) and str(w).strip() not in ("", "nan") else None)
        .dropna()
        .tolist()
        if "Wafer" in df_lw.columns else []
    )

    if not lot_ids:
        logging.error("No lot IDs found in lot_wafer.csv")
        sys.exit(1)

    operations = config["operations"]
    tp_name    = config["tp_name"]

    def _in_clause(values):
        if len(values) == 1:
            return f"('{values[0]}')"
        return str(tuple(values))

    op_filter    = f"AND v0.operation In {_in_clause(operations)}" if operations else ""
    wafer_filter = f"AND v0.wafer_id In {_in_clause(wafer_ids)}"  if wafer_ids  else ""

    sql = f"""
SELECT /*+  use_nl (dt) */
         DISTINCT v0.lot            AS lot
                 ,v0.wafer_id       AS wafer_id
                 ,v0.operation      AS operation
                 ,v0.program_name   AS program_name
                 ,v0.devrevstep     AS devrevstep
FROM  A_Testing_Session  v0
INNER JOIN A_Lot_At_Operation  lao ON  v0.lot       = lao.lot
                                   AND v0.operation = lao.Operation
                                   AND v0.facility  = lao.facility
INNER JOIN A_Device_Testing    dt  ON  v0.lao_start_ww + 0 = dt.lao_start_ww
                                   AND v0.ts_id + 0        = dt.ts_id
WHERE 1=1
  AND v0.lot                   In {_in_clause(lot_ids)}
  AND v0.valid_flag             = 'Y'
  AND dt.within_lao_latest_flag = 'Y'
  AND v0.program_name           = '{tp_name}'
  AND v0.test_end_date_time    >= SYSDATE-180
  {op_filter}
  {wafer_filter}
"""

    # Group lots by database source
    lots_by_src: dict[str, list] = {}
    for lot in lot_ids:
        try:
            src = _lot_to_datasource(lot)
        except ValueError as e:
            logging.error(str(e))
            sys.exit(1)
        lots_by_src.setdefault(src, []).append(lot)

    all_rows, col_names = [], []
    for src, src_lots in lots_by_src.items():
        logging.info(f"Connecting to {src} for lots: {src_lots}")
        conn   = PyUber.connect(datasource=src)
        cursor = conn.cursor()
        cursor.execute(sql)
        if not col_names:
            col_names = [d[0] for d in cursor.description]
        rows = cursor.fetchall()
        all_rows.extend(rows)
        logging.info(f"  → {len(rows)} rows returned from {src}")

    df_result = pd.DataFrame(all_rows, columns=col_names)
    out_path  = OUTPUT_DIR / "lot_wafer_search_result.csv"
    df_result.to_csv(out_path, index=False)
    logging.info(f"Saved lot/wafer search results → {out_path} ({len(df_result)} rows)")

    if df_result.empty:
        logging.warning(
            "No data found for the specified lots/operations/program. "
            "Proceeding but Brita+ may also find no data."
        )

    return df_result


# ---------------------------------------------------------------------------
# Step 5 – Brita+ Batchmode  (per task)
# ---------------------------------------------------------------------------
def run_brita(config: dict, task: dict, task_out_dir: pathlib.Path):
    """Run Brita+ in Batchmode for a single task. Raises RuntimeError on failure."""
    tp_path        = pathlib.Path(config["tp_path"])
    tp_name        = config["tp_name"]
    operations     = config["operations"]
    indicator_file = task["indicator_file"]
    task_name      = task["name"]

    tpl_path       = tp_path / f"{tp_name}.tpl"
    stpl_path      = tp_path / f"{tp_name}.stpl"
    indicator_path  = INPUT_DIR / indicator_file
    units_out       = task_out_dir / f"{task_name}_units_batch.csv"
    units_summary_out = task_out_dir / f"{task_name}_units_batch_summary.csv"
    main_op         = operations[0]

    if not indicator_path.exists():
        raise RuntimeError(f"Indicator file not found: {indicator_path}")

    cmd = [
        BRITA_EXE, "Batchmode",
        "-test_program",             tp_name,
        "-tpl",                      str(tpl_path),
        "-stpl",                     str(stpl_path),
        "-lot_wafer_list",           str(LOT_WAFER_CSV),
        "-main_operation",           main_op,
        "-indicator_file",           str(indicator_path),
        "-unit_details_output_file", str(units_out),
        "-summary_output_file",      str(units_summary_out),
        "-brita_work",               str(task_out_dir),
        "-keep_logs",                "true",
    ]

    logging.info(f"[{task_name}] Running Brita+ Batchmode ...")
    logging.info("  " + " ".join(f'"{c}"' if " " in c else c for c in cmd))

    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.stdout:
        logging.info(result.stdout)
    if result.returncode != 0:
        raise RuntimeError(
            f"Brita+ failed (exit code {result.returncode}):\n{result.stderr}"
        )

    logging.info(f"[{task_name}] Brita+ Batchmode completed → {units_out} ✓")


# ---------------------------------------------------------------------------
# Step 6 – Script execution  (per task, .jsl → JMP, .py → Python)
# ---------------------------------------------------------------------------
def run_script(task: dict, task_out_dir: pathlib.Path):
    """Launch the task script. Raises RuntimeError on failure.
    For .py scripts, passes -filter_dict if the task defines one.
    """
    import json as _json

    script_name = task.get("script")
    task_name   = task["name"]

    if not script_name:
        logging.info(f"[{task_name}] No script specified – skipping report step.")
        return

    script_path = INPUT_DIR / script_name
    if not script_path.exists():
        raise RuntimeError(f"Script not found: {script_path}")

    suffix = script_path.suffix.lower()
    if suffix == ".jsl":
        cmd   = [JMP_EXE, str(script_path)]
        label = "JMP"
        proc  = subprocess.Popen(cmd)
        logging.info(f"[{task_name}] {label} script launched (PID={proc.pid}) ✓")
    elif suffix == ".py":
        cmd = [sys.executable, str(script_path)]
        cmd.extend(["-output_dir", str(task_out_dir)])
        filter_dict = task.get("filter_dict", {})
        if filter_dict:
            cmd.extend(["-filter_dict", _json.dumps(filter_dict)])
            logging.info(f"[{task_name}] Passing filter_dict with {len(filter_dict)} entries")
        label = "Python"
        logging.info(f"[{task_name}] Running {label} script: {script_path}")
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.stdout:
            logging.info(result.stdout)
        if result.returncode != 0:
            raise RuntimeError(
                f"[{task_name}] Script failed (exit {result.returncode}):\n{result.stderr}"
            )
        logging.info(f"[{task_name}] {label} script completed ✓")
    else:
        raise RuntimeError(
            f"Unsupported script extension '{suffix}' for task '{task_name}'. "
            "Expected .jsl or .py"
        )


# ---------------------------------------------------------------------------
# Step 5b – Plus pipeline  (AquaCmdLine → TraceTestData → post-script)
# ---------------------------------------------------------------------------
def run_plus(config: dict, task: dict, task_out_dir: pathlib.Path):
    """Run the Plus instance-level data collection pipeline for a single task."""
    import pandas as pd

    task_name = task["name"]

    # ── Read lot / wafer from lot_wafer.csv ──────────────────────────────────
    df_lw = pd.read_csv(LOT_WAFER_CSV, dtype=str)
    lots  = df_lw["Lot"].dropna().str.strip().tolist()
    if not lots:
        raise RuntimeError(f"[{task_name}] No lot IDs found in {LOT_WAFER_CSV}")
    lots_str = ",".join(lots)

    has_wafer = (
        "Wafer" in df_lw.columns
        and df_lw["Wafer"].apply(lambda w: str(w).strip() not in ("", "nan")).any()
    )
    ults_str = ""
    if has_wafer:
        ult_parts = []
        for _, row in df_lw.iterrows():
            lot   = str(row["Lot"]).strip()   if pd.notna(row.get("Lot"))   else ""
            wafer = str(row["Wafer"]).strip() if pd.notna(row.get("Wafer")) else ""
            if lot and wafer and wafer not in ("", "nan"):
                ult_parts.append(f"{lot},{wafer}")
        ults_str = ";".join(ult_parts)

    operations     = config.get("operations", [])
    operations_str = ",".join(str(op) for op in operations)

    # ── Step a: AquaCmdLine.exe ───────────────────────────────────────────────
    aqua_out      = task_out_dir / "aqua_result.csv"
    ttd_out       = task_out_dir / "ttd_result.csv"

    try:
        aqua_accessible = AQUA_EXE.exists()
    except OSError:
        aqua_accessible = False
    if not aqua_accessible:
        raise RuntimeError(
            f"[{task_name}] AquaCmdLine.exe not accessible: {AQUA_EXE}\n"
            "Please verify network connectivity to PGSAPP3301.gar.corp.intel.com"
        )

    report_config = INPUT_DIR / "PLUS_SORT_INSTANCE.txt"
    if not report_config.exists():
        raise RuntimeError(f"[{task_name}] Report config not found: {report_config}")

    cmd_aqua = [
        str(AQUA_EXE),
        "-aquaServer",    "AMR",
        "-reportConfig",  str(report_config),
        "-outputFileName", str(aqua_out),
        "-lots",           lots_str,
        "-operations",     operations_str,
    ]
    if has_wafer and ults_str:
        cmd_aqua.extend(["-ults", ults_str])

    logging.info(f"[{task_name}] Running AquaCmdLine ...")
    logging.info("  " + " ".join(f'"{c}"' if " " in c else c for c in cmd_aqua))
    result = subprocess.run(cmd_aqua, capture_output=True, text=True)
    if result.stdout:
        logging.info(result.stdout)
    if result.returncode != 0:
        raise RuntimeError(
            f"[{task_name}] AquaCmdLine failed (exit {result.returncode}):\n{result.stderr}"
        )
    logging.info(f"[{task_name}] AquaCmdLine completed → {aqua_out} ✓")

    # ── Step b: TraceTestData.exe ─────────────────────────────────────────────
    try:
        ttd_accessible = TTD_EXE.exists()
    except OSError:
        ttd_accessible = False
    if not ttd_accessible:
        raise RuntimeError(
            f"[{task_name}] TraceTestData.exe not accessible: {TTD_EXE}\n"
            "Please verify network connectivity to amr.corp.intel.com"
        )

    cmd_ttd = [
        str(TTD_EXE),
        "--data", "type", "=", "instance", "pass", "fail",
        "--input", "file", "=", str(aqua_out),
        "--output", "csv", "=", str(ttd_out),
        "--result", "columns", "=", "port,pass",
        "--material", "type", "=", "Xeus",
    ]

    logging.info(f"[{task_name}] Running TraceTestData ...")
    logging.info("  " + " ".join(f'"{c}"' if " " in c else c for c in cmd_ttd))
    result = subprocess.run(cmd_ttd, capture_output=True, text=True)
    if result.stdout:
        logging.info(result.stdout)
    if result.stderr:
        logging.warning(result.stderr)
    # TTD exits with 0 even on failure — verify output file was actually created
    if not ttd_out.exists():
        raise RuntimeError(
            f"[{task_name}] TraceTestData did not produce output file: {ttd_out}\n"
            f"stdout: {result.stdout.strip()}"
        )
    logging.info(f"[{task_name}] TraceTestData completed → {ttd_out} ✓")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    setup_logging()
    logging.info("=" * 60)
    logging.info("  Yield Automation Tool – Started")
    logging.info("=" * 60)

    check_prerequisites()
    config = load_config()
    ensure_venv()       # may re-launch; everything below runs inside venv

    # Shared steps – failure here aborts everything
    xeus_info = run_get_xeus_test_result(config)
    config.update(xeus_info)
    validate_tp_files(config["tp_path"], config["tp_name"])
    search_lot_wafer(config)

    # Per-task steps – one task failure does not stop the others
    tasks = config["tasks"]
    results: dict[str, str] = {}

    for task in tasks:
        task_name    = task["name"]
        task_out_dir = OUTPUT_DIR / task_name
        task_out_dir.mkdir(parents=True, exist_ok=True)
        logging.info("-" * 60)
        logging.info(f"  Task: {task_name}")
        logging.info("-" * 60)
        try:
            task_type = task.get("type", "Brita")
            if task_type == "Brita":
                run_brita(config, task, task_out_dir)
            elif task_type == "Plus":
                run_plus(config, task, task_out_dir)
            else:
                raise RuntimeError(f"Unknown task type: '{task_type}'")
            run_script(task, task_out_dir)
            results[task_name] = "SUCCESS"
        except Exception as e:
            logging.error(f"[{task_name}] Task failed: {e}")
            results[task_name] = f"FAILED: {e}"

    # Summary
    logging.info("=" * 60)
    logging.info("  Task Summary")
    logging.info("=" * 60)
    for name, status in results.items():
        icon = "✓" if status == "SUCCESS" else "✗"
        logging.info(f"  {icon}  {name}: {status}")

    failed = [n for n, s in results.items() if s != "SUCCESS"]
    if failed:
        logging.warning(f"{len(failed)} task(s) failed: {failed}")
    else:
        logging.info("All tasks completed successfully.")

    logging.info("=" * 60)
    logging.info("  Yield Automation Tool – Completed")
    logging.info("=" * 60)


if __name__ == "__main__":
    main()
