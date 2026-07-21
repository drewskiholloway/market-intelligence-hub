#!/usr/bin/env python3
"""
Fetches the latest Census Bureau Building Permits Survey (BPS) state-level
and metro/CBSA-level monthly permit files and stores them as JSON in this
repo, plus derives the flattened datasets the dashboard actually consumes:
  - data/permits/state_permits.json     {ABBR: ytd total permits, K}
  - data/permits/state_momentum.json    {ABBR: ytd single-family permits, K}
  - data/permits/metro_permits.json     {"City, ST": {permit, pctChg}}

Source: https://www.census.gov/construction/bps/statemonthly.html
        https://www.census.gov/construction/bps/msamonthly.html

Idempotent: skips re-downloading a period already cached on disk. Safe to
run on any schedule -- it only does new work when Census publishes a new
month, but always re-derives the flattened dashboard datasets so they stay
in sync if the derivation logic itself changes.
"""
import datetime
import json
import subprocess
import sys
import tempfile
import time
import urllib.request
from pathlib import Path

import xlrd

REPO_ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = REPO_ROOT / "data" / "permits"
RAW_STATE_DIR = DATA_DIR / "raw_state"
RAW_MSA_DIR = DATA_DIR / "raw_msa"
MANIFEST = DATA_DIR / "manifest.json"

BASE_URL = "https://www.census.gov/construction/bps/xls"
UA = "Mozilla/5.0 (Market Intelligence Hub data refresh; contact: repo issue tracker)"

STATE_UNITS_SHEET = "State Units"
MSA_UNITS_SHEET = "MSA Units"

STATE_FIELDS = ["total", "oneUnit", "twoUnit", "threeFourUnit", "fiveUnitPlus", "numStruct5Plus"]
MSA_FIELDS = STATE_FIELDS

NAME_TO_ABBR = {
    'Alabama':'AL','Alaska':'AK','Arizona':'AZ','Arkansas':'AR','California':'CA',
    'Colorado':'CO','Connecticut':'CT','Delaware':'DE','Florida':'FL','Georgia':'GA',
    'Hawaii':'HI','Idaho':'ID','Illinois':'IL','Indiana':'IN','Iowa':'IA',
    'Kansas':'KS','Kentucky':'KY','Louisiana':'LA','Maine':'ME','Maryland':'MD',
    'Massachusetts':'MA','Michigan':'MI','Minnesota':'MN','Mississippi':'MS','Missouri':'MO',
    'Montana':'MT','Nebraska':'NE','Nevada':'NV','New Hampshire':'NH','New Jersey':'NJ',
    'New Mexico':'NM','New York':'NY','North Carolina':'NC','North Dakota':'ND','Ohio':'OH',
    'Oklahoma':'OK','Oregon':'OR','Pennsylvania':'PA','Rhode Island':'RI','South Carolina':'SC',
    'South Dakota':'SD','Tennessee':'TN','Texas':'TX','Utah':'UT','Vermont':'VT',
    'Virginia':'VA','Washington':'WA','West Virginia':'WV','Wisconsin':'WI','Wyoming':'WY',
}

METRO_ALIASES = {
    'Urban Honolulu, HI': 'Honolulu, HI',
}


def log(msg):
    print(f"{datetime.datetime.now().isoformat(timespec='seconds')} {msg}")


def url_exists(url):
    req = urllib.request.Request(url, method="HEAD", headers={"User-Agent": UA})
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            return resp.status == 200
    except Exception:
        return False


def download(url, retries=4):
    """census.gov's chunked responses stall intermittently under urllib;
    curl -L with --retry handles it reliably."""
    last_err = None
    for attempt in range(retries):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp:
            tmp_path = tmp.name
        try:
            subprocess.run(
                ["curl", "-sfL", "--retry", "2", "--max-time", "90",
                 "-A", UA, "-o", tmp_path, url],
                check=True,
            )
            data = Path(tmp_path).read_bytes()
            if len(data) < 1000:
                raise ValueError(f"Suspiciously small download ({len(data)} bytes) from {url}")
            return data
        except Exception as e:
            last_err = e
            time.sleep(2 * (attempt + 1))
        finally:
            Path(tmp_path).unlink(missing_ok=True)
    raise last_err


def find_latest_available_period(max_back=4):
    today = datetime.date.today()
    y, m = today.year, today.month
    for _ in range(max_back):
        period = f"{y}{m:02d}"
        state_url = f"{BASE_URL}/statemonthly_{period}.xls"
        msa_url = f"{BASE_URL}/cbsamonthly_{period}.xls"
        if url_exists(state_url) and url_exists(msa_url):
            return period, state_url, msa_url
        m -= 1
        if m == 0:
            m = 12
            y -= 1
    return None, None, None


def parse_state_units(raw_bytes):
    wb = xlrd.open_workbook(file_contents=raw_bytes, ignore_workbook_corruption=True)
    sh = wb.sheet_by_name(STATE_UNITS_SHEET)
    start = None
    for r in range(sh.nrows):
        v = sh.cell_value(r, 0)
        if isinstance(v, str) and v.strip() == "United States":
            start = r
            break
    if start is None:
        raise ValueError("Could not locate 'United States' row in State Units sheet")
    records = []
    for r in range(start, sh.nrows):
        raw_name = sh.cell_value(r, 0)
        if not isinstance(raw_name, str) or not raw_name.strip():
            continue
        total = sh.cell_value(r, 1)
        if not isinstance(total, (int, float)):
            continue
        indent = len(raw_name) - len(raw_name.lstrip(" "))
        rec = {"name": raw_name.strip(), "indent": indent}
        for i, key in enumerate(STATE_FIELDS):
            rec[key] = sh.cell_value(r, 1 + i)
        for i, key in enumerate(STATE_FIELDS):
            rec["ytd" + key[0].upper() + key[1:]] = sh.cell_value(r, 8 + i)
        records.append(rec)
    return records


def parse_msa_units(raw_bytes):
    wb = xlrd.open_workbook(file_contents=raw_bytes, ignore_workbook_corruption=True)
    sh = wb.sheet_by_name(MSA_UNITS_SHEET)
    start = None
    for r in range(sh.nrows):
        cbsa = sh.cell_value(r, 1)
        name = sh.cell_value(r, 2)
        if isinstance(cbsa, (int, float)) and isinstance(name, str) and name.strip():
            start = r
            break
    if start is None:
        raise ValueError("Could not locate first data row in MSA Units sheet")
    records = []
    for r in range(start, sh.nrows):
        cbsa = sh.cell_value(r, 1)
        name = sh.cell_value(r, 2)
        if not isinstance(cbsa, (int, float)) or not isinstance(name, str) or not name.strip():
            continue
        rec = {
            "csa": sh.cell_value(r, 0),
            "cbsa": int(cbsa),
            "name": name.strip(),
            "metroMicroCode": sh.cell_value(r, 3),
        }
        for i, key in enumerate(MSA_FIELDS):
            rec[key] = sh.cell_value(r, 4 + i)
        for i, key in enumerate(MSA_FIELDS):
            rec["ytd" + key[0].upper() + key[1:]] = sh.cell_value(r, 11 + i)
        records.append(rec)
    return records


def metro_shortname(census_name):
    aliased = METRO_ALIASES.get(census_name, census_name)
    city_part, state_part = aliased.split(',', 1)
    city = city_part.split('-')[0].strip()
    state = state_part.strip().split('-')[0].strip()
    return f"{city}, {state}"


def period_to_label(period):
    y, m = int(period[:4]), int(period[4:])
    return f"{y}-{m:02d}"


def prior_year_period(period_label):
    y, m = period_label.split('-')
    return f"{int(y)-1}-{m}"


def load_manifest():
    if MANIFEST.exists():
        return json.loads(MANIFEST.read_text())
    return {"lastChecked": None, "lastNewDataFetched": None, "periods": []}


def save_manifest(manifest):
    MANIFEST.write_text(json.dumps(manifest, indent=2))


def ensure_period_cached(period_label):
    state_out = RAW_STATE_DIR / f"{period_label}.json"
    msa_out = RAW_MSA_DIR / f"{period_label}.json"
    if state_out.exists() and msa_out.exists():
        return (json.loads(state_out.read_text())["records"],
                json.loads(msa_out.read_text())["records"])

    y, m = period_label.split('-')
    period = f"{y}{m}"
    state_url = f"{BASE_URL}/statemonthly_{period}.xls"
    msa_url = f"{BASE_URL}/cbsamonthly_{period}.xls"
    if not (url_exists(state_url) and url_exists(msa_url)):
        log(f"prior-year period {period_label} not available from Census yet; YoY will fall back to 0%.")
        return None, None

    log(f"fetching prior-year period {period_label} for YoY comparisons ...")
    state_raw = download(state_url)
    msa_raw = download(msa_url)
    state_records = parse_state_units(state_raw)
    msa_records = parse_msa_units(msa_raw)
    payload_common = {
        "period": period_label,
        "sourceUrls": {"state": state_url, "msa": msa_url},
        "fetchedAt": datetime.datetime.now().isoformat(timespec="seconds"),
        "unit": "permits (count), YTD = year-to-date through this period",
    }
    RAW_STATE_DIR.mkdir(parents=True, exist_ok=True)
    RAW_MSA_DIR.mkdir(parents=True, exist_ok=True)
    state_out.write_text(json.dumps({**payload_common, "records": state_records}, indent=2))
    msa_out.write_text(json.dumps({**payload_common, "records": msa_records}, indent=2))
    return state_records, msa_records


def build_state_datasets(cur_state_records):
    states = {r["name"]: r for r in cur_state_records if r.get("indent") == 10}
    permits_by_abbr = {}
    momentum_by_abbr = {}
    for name, abbr in NAME_TO_ABBR.items():
        r = states.get(name)
        if not r:
            continue
        permits_by_abbr[abbr] = round(r["ytdTotal"] / 1000, 3)
        momentum_by_abbr[abbr] = round(r["ytdOneUnit"] / 1000, 3)
    return permits_by_abbr, momentum_by_abbr


def build_metro_dataset(cur_msa_records, prior_msa_records):
    """Top 100 metros by YTD single-family permit volume (matches the curated
    list size the dashboard has always used); output keyed by 'City, ST'."""
    prior_by_cbsa = {}
    if prior_msa_records:
        for r in prior_msa_records:
            prior_by_cbsa[r["cbsa"]] = r

    ranked = sorted(cur_msa_records, key=lambda r: r.get("ytdOneUnit") or 0, reverse=True)[:100]
    out = {}
    for r in ranked:
        key = metro_shortname(r["name"])
        cur_val = r["ytdOneUnit"]
        prior_r = prior_by_cbsa.get(r["cbsa"])
        if prior_r and prior_r.get("ytdOneUnit"):
            pct_chg = round((cur_val - prior_r["ytdOneUnit"]) / prior_r["ytdOneUnit"] * 100)
        else:
            pct_chg = 0
        out[key] = {"permit": round(cur_val / 1000, 3), "pctChg": pct_chg}
    return out


def main():
    force = "--force" in sys.argv
    RAW_STATE_DIR.mkdir(parents=True, exist_ok=True)
    RAW_MSA_DIR.mkdir(parents=True, exist_ok=True)

    manifest = load_manifest()
    manifest["lastChecked"] = datetime.datetime.now().isoformat(timespec="seconds")

    period, state_url, msa_url = find_latest_available_period()
    if period is None:
        log("No available BPS period found in lookback window -- Census may not have published yet.")
        save_manifest(manifest)
        return

    label = period_to_label(period)
    state_out = RAW_STATE_DIR / f"{label}.json"
    msa_out = RAW_MSA_DIR / f"{label}.json"

    is_new = force or not (state_out.exists() and msa_out.exists())
    if is_new:
        log(f"Fetching new period {label} ...")
        state_raw = download(state_url)
        msa_raw = download(msa_url)
        state_records = parse_state_units(state_raw)
        msa_records = parse_msa_units(msa_raw)
        payload_common = {
            "period": label,
            "sourceUrls": {"state": state_url, "msa": msa_url},
            "fetchedAt": datetime.datetime.now().isoformat(timespec="seconds"),
            "unit": "permits (count), YTD = year-to-date through this period",
        }
        state_out.write_text(json.dumps({**payload_common, "records": state_records}, indent=2))
        msa_out.write_text(json.dumps({**payload_common, "records": msa_records}, indent=2))
        if label not in manifest["periods"]:
            manifest["periods"].append(label)
            manifest["periods"].sort()
        manifest["lastNewDataFetched"] = payload_common["fetchedAt"]
    else:
        log(f"{label} already stored; re-deriving flattened datasets only.")
        state_records = json.loads(state_out.read_text())["records"]
        msa_records = json.loads(msa_out.read_text())["records"]

    manifest["latestPeriod"] = label
    save_manifest(manifest)

    prior_period = prior_year_period(label)
    prior_state_records, prior_msa_records = ensure_period_cached(prior_period)

    permits_by_abbr, momentum_by_abbr = build_state_datasets(state_records)
    metro_data = build_metro_dataset(msa_records, prior_msa_records)

    meta = {"period": label, "fetchedAt": datetime.datetime.now().isoformat(timespec="seconds"),
            "source": "Census Bureau Building Permits Survey (BPS)",
            "sourceUrls": {"state": state_url, "msa": msa_url}}

    (DATA_DIR / "state_permits.json").write_text(json.dumps({**meta, "data": permits_by_abbr}, indent=2))
    (DATA_DIR / "state_momentum.json").write_text(json.dumps({**meta, "data": momentum_by_abbr}, indent=2))
    (DATA_DIR / "metro_permits.json").write_text(json.dumps({**meta, "data": metro_data}, indent=2))

    log(f"Wrote {label}: {len(permits_by_abbr)} states, {len(metro_data)} metros.")


if __name__ == "__main__":
    main()
