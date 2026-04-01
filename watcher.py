"""
Folder Watcher — drop any Excel file into the /inbox folder
and the agent processes it automatically.

Structure:
  excel_agent/
  ├── inbox/      ← drop files here
  ├── processed/  ← successfully processed files move here
  ├── failed/     ← files that errored move here
  └── output/     ← analyzed Excel files appear here
"""

import os
import time
import shutil
import traceback
from datetime import datetime
from pathlib import Path

from config_loader import load_config, describe_config
from agent import (
    load_and_translate, clean_sheets, build_monthly_summary,
    compute_kpis, detect_anomalies, event_impact_analysis,
    analyze_agents, build_context, write_output,
)

# ── Folders ────────────────────────────────────────────────────────
BASE_DIR      = Path(__file__).parent
INBOX_DIR     = BASE_DIR / "inbox"
PROCESSED_DIR = BASE_DIR / "processed"
FAILED_DIR    = BASE_DIR / "failed"
OUTPUT_DIR    = BASE_DIR / "output"

for d in [INBOX_DIR, PROCESSED_DIR, FAILED_DIR, OUTPUT_DIR]:
    d.mkdir(exist_ok=True)

SUPPORTED_EXTENSIONS = {".xlsx", ".xls"}
POLL_INTERVAL = 5  # seconds between checks

# ── Logging ────────────────────────────────────────────────────────
def log(msg, level="INFO"):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    prefix = {"INFO":"✅","WARN":"⚠️ ","ERROR":"❌","START":"🚀","FILE":"📂","WAIT":"⏳"}
    icon = prefix.get(level, "  ")
    print(f"[{ts}] {icon} {msg}")

# ── Process one file ───────────────────────────────────────────────
def process_file(file_path: Path, cfg: dict):
    log(f"Processing: {file_path.name}", "FILE")
    start = time.time()

    try:
        # Load and analyse
        sheets    = load_and_translate(str(file_path))
        sheets, action_log = clean_sheets(sheets)
        summary   = build_monthly_summary(sheets)
        kpis      = compute_kpis(summary)
        anomalies = detect_anomalies(sheets)
        impact    = event_impact_analysis(sheets)
        agents    = analyze_agents(sheets)

        # Build output filename
        ts_str      = datetime.now().strftime("%Y%m%d_%H%M%S")
        stem        = file_path.stem.replace(" ", "_")
        output_name = f"{stem}_analyzed_{ts_str}.xlsx"
        output_path = OUTPUT_DIR / output_name

        # Write output
        write_output(sheets, summary, kpis, anomalies, agents,
                     impact, action_log, str(output_path))

        elapsed = round(time.time() - start, 1)
        log(f"Done in {elapsed}s → {output_name}", "INFO")
        log(f"  Calls: {kpis.get('Grand Total Inbound Calls',0):,} | Anomalies: {len(anomalies)}", "INFO")

        # Move input file to processed
        dest = PROCESSED_DIR / f"{ts_str}_{file_path.name}"
        shutil.move(str(file_path), str(dest))
        log(f"Moved to processed/", "INFO")

        return True

    except Exception as e:
        log(f"Failed: {e}", "ERROR")
        traceback.print_exc()
        # Move to failed folder
        ts_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = FAILED_DIR / f"{ts_str}_{file_path.name}"
        shutil.move(str(file_path), str(dest))
        log(f"Moved to failed/", "WARN")
        return False

# ── Watch loop ─────────────────────────────────────────────────────
def watch():
    cfg = load_config()
    log(f"Folder Watcher started", "START")
    log(f"Client: {cfg['client_name']}", "INFO")
    log(f"Watching: {INBOX_DIR}", "INFO")
    log(f"Outputs:  {OUTPUT_DIR}", "INFO")
    log(f"Poll interval: every {POLL_INTERVAL} seconds", "INFO")
    print()

    processed_count = 0
    failed_count    = 0

    while True:
        # Find all Excel files in inbox
        files = [
            f for f in INBOX_DIR.iterdir()
            if f.is_file() and f.suffix.lower() in SUPPORTED_EXTENSIONS
            and not f.name.startswith("~$")  # skip Excel lock files
        ]

        if files:
            log(f"Found {len(files)} file(s) to process", "FILE")
            for file_path in files:
                success = process_file(file_path, cfg)
                if success:
                    processed_count += 1
                else:
                    failed_count += 1
            log(f"Total processed: {processed_count} | Failed: {failed_count}", "INFO")
            print()
        else:
            print(f"\r⏳ Watching inbox... (processed: {processed_count} | failed: {failed_count})", end="", flush=True)

        time.sleep(POLL_INTERVAL)

if __name__ == "__main__":
    watch()