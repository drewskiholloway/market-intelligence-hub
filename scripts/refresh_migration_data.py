#!/usr/bin/env python3
"""
Fetches Redfin's public "U.S. Migration Patterns" dataset (quarterly,
aggregated inflow/outflow by state and by top-100 metro) from Redfin's
public S3 bucket and writes flattened JSON for the dashboard.

Source: https://www.redfin.com/news/data-center/migration-patterns/
Underlying files (public, no key required):
  https://redfin-public-data.s3.us-west-2.amazonaws.com/redfin_data_center/migration_traffic/aggregated/all_states.csv
  https://redfin-public-data.s3.us-west-2.amazonaws.com/redfin_data_center/migration_traffic/aggregated/all_metros.csv
  https://redfin-public-data.s3.us-west-2.amazonaws.com/redfin_data_center/migration_traffic/od_pairs/all_states.csv

Redfin publishes this as a rolling monthly file with a 3-month trailing
window per row (columns are literally "Rolling 3-Month"); we always take
the most recent PERIOD BEGIN present in the file. Idempotent-ish: always
re-fetches (files are tiny) but only overwrites output if the latest
period actually advanced, so redundant runs are cheap no-ops.

The state-to-state OD-pairs CSV is written verbatim to
data/migration/od_pairs_states.csv (no reshaping) so the dashboard's
existing client-side aggregateMigrationCSV() JS parser -- originally built
for manual user CSV import -- can consume it directly with zero new
parsing logic and guaranteed format fidelity.
"""
import csv
import datetime
import io
import json
import subprocess
import sys
import tempfile
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = REPO_ROOT / "data" / "migration"
MANIFEST = DATA_DIR / "manifest.json"

BASE = "https://redfin-public-data.s3.us-west-2.amazonaws.com/redfin_data_center/migration_traffic/aggregated"
STATE_URL = f"{BASE}/all_states.csv"
METRO_URL = f"{BASE}/all_metros.csv"
OD_BASE = "https://redfin-public-data.s3.us-west-2.amazonaws.com/redfin_data_center/migration_traffic/od_pairs"
OD_STATES_URL = f"{OD_BASE}/all_states.csv"
UA = "Mozilla/5.0 (Market Intelligence Hub data refresh; contact: repo issue tracker)"


def log(msg):
    print(f"{datetime.datetime.now().isoformat(timespec='seconds')} {msg}")


def download_csv(url, retries=4):
    last_err = None
    for attempt in range(retries):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".csv") as tmp:
            tmp_path = tmp.name
        try:
            subprocess.run(
                ["curl", "-sfL", "--retry", "2", "--max-time", "60",
                 "-A", UA, "-o", tmp_path, url],
                check=True,
            )
            data = Path(tmp_path).read_text()
            if len(data) < 200:
                raise ValueError(f"Suspiciously small download from {url}")
            return data
        except Exception as e:
            last_err = e
            time.sleep(2 * (attempt + 1))
        finally:
            Path(tmp_path).unlink(missing_ok=True)
    raise last_err


def num(v):
    if v is None or v == "" or v == "NA":
        return None
    try:
        return float(v)
    except ValueError:
        return None


def latest_period(rows):
    periods = sorted(set(r["PERIOD BEGIN"] for r in rows if r.get("PERIOD BEGIN")))
    return periods[-1] if periods else None


def build_state_dataset(rows, period):
    out = {}
    for r in rows:
        if r["PERIOD BEGIN"] != period:
            continue
        name = r["REGION NAME"].strip()
        out[name] = {
            "inflow": num(r.get("INFLOW")),
            "outflow": num(r.get("OUTFLOW")),
            "netFlow": num(r.get("NET FLOW")),
            "pctSearchingIn": num(r.get("PCT SEARCHING IN (%)")),
            "pctSearchingOut": num(r.get("PCT SEARCHING OUT (%)")),
        }
    return out


def build_metro_dataset(rows, period):
    out = {}
    for r in rows:
        if r["PERIOD BEGIN"] != period:
            continue
        name = r["REGION NAME"].strip()
        out[name] = {
            "inflow": num(r.get("INFLOW")),
            "outflow": num(r.get("OUTFLOW")),
            "netFlow": num(r.get("NET FLOW")),
            "pctSearchingIn": num(r.get("PCT SEARCHING IN (%)")),
            "pctSearchingOut": num(r.get("PCT SEARCHING OUT (%)")),
        }
    return out


def load_manifest():
    if MANIFEST.exists():
        return json.loads(MANIFEST.read_text())
    return {"latestPeriod": None, "lastChecked": None}


def save_manifest(m):
    MANIFEST.write_text(json.dumps(m, indent=2))


def main():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    force = "--force" in sys.argv
    manifest = load_manifest()
    manifest["lastChecked"] = datetime.datetime.now().isoformat(timespec="seconds")

    log("Fetching Redfin migration_traffic aggregated state + metro CSVs ...")
    state_csv = download_csv(STATE_URL)
    metro_csv = download_csv(METRO_URL)

    state_rows = list(csv.DictReader(io.StringIO(state_csv)))
    metro_rows = list(csv.DictReader(io.StringIO(metro_csv)))

    period = latest_period(state_rows)
    if period is None:
        log("Could not determine latest period from Redfin state CSV; aborting.")
        save_manifest(manifest)
        return

    if period == manifest.get("latestPeriod") and not force:
        log(f"{period} already stored, nothing to do.")
        save_manifest(manifest)
        return

    state_data = build_state_dataset(state_rows, period)
    metro_data = build_metro_dataset(metro_rows, period)

    meta = {
        "period": period,
        "fetchedAt": datetime.datetime.now().isoformat(timespec="seconds"),
        "source": "Redfin U.S. Migration Patterns (quarterly, rolling 3-month window)",
        "sourceUrls": {"state": STATE_URL, "metro": METRO_URL},
        "unit": "searches (inflow/outflow), rolling 3-month",
    }

    (DATA_DIR / "state_migration.json").write_text(json.dumps({**meta, "data": state_data}, indent=2))
    (DATA_DIR / "metro_migration.json").write_text(json.dumps({**meta, "data": metro_data}, indent=2))

    log("Fetching Redfin migration_traffic OD-pairs (state-to-state) CSV ...")
    od_states_csv = download_csv(OD_STATES_URL)
    (DATA_DIR / "od_pairs_states.csv").write_text(od_states_csv)
    log(f"Wrote od_pairs_states.csv ({len(od_states_csv)} bytes) for the frontend's "
        f"client-side aggregateMigrationCSV() parser.")

    manifest["latestPeriod"] = period
    manifest["lastNewDataFetched"] = meta["fetchedAt"]
    save_manifest(manifest)

    log(f"Wrote {period}: {len(state_data)} states, {len(metro_data)} metros.")


if __name__ == "__main__":
    main()
