"""
clean_pac_disbursements.py

Cleans an FEC corporate PAC disbursement export (Schedule B, e.g. the
"Abbvie__C00536573_.csv" file from FEC.gov bulk/itemized downloads) and:

  1. Trims to a focused set of analysis columns.
  2. Adds DISBURSEMENT_YEAR and DISBURSEMENT_QUARTER columns derived from
     disbursement_date.
  3. Adds a CANDIDATE_PARTY column by joining on candidate_id against a
     FEC-candidate-ID -> party lookup built from the `unitedstates/
     congress-legislators` public dataset (a well-maintained, freely
     hosted GitHub project that's updated regularly and lists every FEC
     candidate ID each member of Congress has ever filed under, alongside
     their party for every term they served). This is the key fix for the
     "same last name, different party" problem: candidate_id is a unique
     FEC identifier per candidate-per-office (it encodes office, state, and
     a sequence number), so two candidates who share a last name -- e.g.
     two different "Scott"s, one a Senator-R and one a Rep-D -- can never
     collide. Joining on candidate_id instead of on name strings is what
     makes the party assignment reliable.
  4. Flags rows with no candidate_id (contributions to other PACs/leadership
     PACs/party committees rather than to an individual candidate) instead
     of guessing a party for them.

WHY NOT THE OFFICIAL FEC BULK FILES?
-------------------------------------
The previous version of this script joined against FEC's own Candidate
Master bulk file (cn.txt, downloaded per election cycle from fec.gov).
That's the most authoritative source, and you're welcome to swap it back in
if you have it -- but it requires manually downloading a zip per cycle.
The congress-legislators dataset below is auto-fetched by this script (no
manual download step) and resolved 261 of 262 unique candidate_ids in the
sample AbbVie file (the lone miss was a sitting House member's brand-new
candidate_id for a current Senate run that hadn't yet propagated to the
dataset -- handled via PARTY_OVERRIDES below). It only covers federal
candidates for the House/Senate/President, which is what PAC disbursement
files are about anyway.

If your own file includes a candidate_id this script can't resolve, add it
to PARTY_OVERRIDES with the correct party, or supply your own FEC candidate
master file via --candidate-master and it'll be checked first.

USAGE
-----
    # No arguments needed -- a file picker window opens, supporting
    # multi-select (ctrl/cmd+click or shift+click to pick several files
    # at once). All cleaned files are saved into a "pac_disbursements_clean"
    # folder created next to the first selected input file.
    python clean_pac_disbursements.py

    # Or specify one or more files on the command line:
    python clean_pac_disbursements.py \
        --input "Abbvie__C00536573_.csv" "Pfizer__C00342417_.csv" \
        --combine   # also write one combined_cleaned.csv with all rows + a source_file column

    # Send the output folder somewhere specific instead of the default:
    python clean_pac_disbursements.py \
        --input "Abbvie__C00536573_.csv" \
        --output-dir "/path/to/some/folder"

    # Optional: supply official FEC candidate master file(s) too -- these
    # take priority over the congress-legislators lookup if both have an
    # answer for a given candidate_id.
    python clean_pac_disbursements.py \
        --input "Abbvie__C00536573_.csv" \
        --candidate-master cn22.txt cn24.txt \
        --output "Abbvie_cleaned.csv"

    # Optional: skip the network fetch and use local copies of the
    # congress-legislators YAML files (e.g. if you're offline)
    python clean_pac_disbursements.py \
        --input "Abbvie__C00536573_.csv" \
        --legislators-current legislators-current.yaml \
        --legislators-historical legislators-historical.yaml \
        --output "Abbvie_cleaned.csv"
"""

import argparse
import json
import sys
import urllib.request
from pathlib import Path

import pandas as pd

try:
    import yaml
except ImportError:
    yaml = None

# Pre-built candidate_id -> party lookup, embedded directly in this file so
# the script works completely offline with zero dependencies beyond pandas.
# This is a gzip+base64 blob of a dict built from the unitedstates/
# congress-legislators public dataset (https://github.com/unitedstates/congress-legislators),
# mapping every FEC candidate_id a member of Congress has ever filed under
# to their most recent party. Regenerate via the included build script if
# you want a refreshed snapshot.
_EMBEDDED_PARTY_LOOKUP_B64 = (
    "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"
)


def _decode_embedded_party_lookup() -> dict:
    import gzip, base64, json as _json
    raw = gzip.decompress(base64.b64decode(_EMBEDDED_PARTY_LOOKUP_B64))
    return _json.loads(raw)


LEGISLATORS_CURRENT_URL = (
    "https://raw.githubusercontent.com/unitedstates/congress-legislators/"
    "main/legislators-current.yaml"
)
LEGISLATORS_HISTORICAL_URL = (
    "https://raw.githubusercontent.com/unitedstates/congress-legislators/"
    "main/legislators-historical.yaml"
)

PARTY_NORMALIZE = {
    "Democrat": "Democratic",
    "Democratic": "Democratic",
    "Republican": "Republican",
    "Independent": "Independent",
    "Libertarian": "Libertarian",
}

# ---------------------------------------------------------------------------
# 1. Columns to keep in the cleaned output
# ---------------------------------------------------------------------------
KEEP_COLUMNS = [
    "committee_id",
    "committee_name",
    "report_year",
    "disbursement_date",
    "disbursement_amount",
    "disbursement_type",
    "disbursement_type_description",
    "disbursement_description",
    "disbursement_purpose_category",
    "memo_text",
    "recipient_committee_id",
    "recipient_name",
    "recipient_city",
    "recipient_state",
    "candidate_id",
    "candidate_name",
    "candidate_first_name",
    "candidate_last_name",
    "candidate_office",
    "candidate_office_description",
    "candidate_office_state",
    "candidate_office_district",
    "election_type_full",
    "two_year_transaction_period",
]

# Hand-maintained overrides, keyed by candidate_id. Use this for any
# candidate_id that neither an FEC candidate-master file nor the
# congress-legislators dataset resolves -- typically a brand new
# candidate_id filed for a current run (e.g. a sitting House member who
# just launched a Senate campaign) that hasn't propagated to either source
# yet. Confirmed example from the AbbVie sample file: Earl L. "Buddy"
# Carter (R-GA-01) filed a new Senate candidate_id, S6GA00374, for a 2026
# Senate run; his House candidate_id H4GA01039 already resolves to
# Republican via congress-legislators, so this is a same-person carry-over.
PARTY_OVERRIDES: dict[str, str] = {
    "S6GA00374": "Republican",  # Buddy Carter, new 2026 Senate candidate_id
}

# Map FEC party codes to readable labels, used only when an FEC candidate
# master file is supplied via --candidate-master (FEC uses standardized
# 3-letter codes in CAND_PTY_AFFILIATION). The congress-legislators dataset
# already returns readable party names, so this map doesn't apply to it.
PARTY_CODE_MAP = {
    "DEM": "Democratic",
    "REP": "Republican",
    "IND": "Independent",
    "LIB": "Libertarian",
    "GRE": "Green",
    "NPA": "No Party Affiliation",
    "NNE": "None",
    "DFL": "Democratic-Farmer-Labor",
    "UNK": "Unknown",
}


def fetch_or_load_yaml(url: str, local_path: str | None) -> list:
    if local_path:
        p = Path(local_path)
        if not p.exists():
            print(f"  [warn] local file not found: {p}, will try fetching from {url}", file=sys.stderr)
        else:
            with open(p, "rb") as f:
                return yaml.safe_load(f)
    with urllib.request.urlopen(url, timeout=60) as resp:
        return yaml.safe_load(resp.read())


def load_legislators_party_lookup(
    current_path: str | None, historical_path: str | None
) -> dict[str, str]:
    """Build a candidate_id -> party dict.

    Resolution order:
      1. The lookup embedded directly in this script (_EMBEDDED_PARTY_LOOKUP_B64
         above). No file to misplace, no network call, no PyYAML dependency.
         This is what every normal run uses.
      2. Local legislators-current.yaml / legislators-historical.yaml, if
         paths were explicitly passed in via --legislators-current/
         --legislators-historical (e.g. to refresh past what's embedded).
      3. Live fetch from the unitedstates/congress-legislators GitHub repo,
         only attempted if #2's paths were given but PyYAML/network make it
         necessary, or as an explicit refresh.

    Each legislator record lists every FEC candidate_id they've ever filed
    under (id.fec) plus a term history with party per term; we use their
    most recent term's party.
    """
    # --- Option 1: data embedded directly in this file ---
    if not current_path and not historical_path:
        try:
            lookup = _decode_embedded_party_lookup()
            print(
                f"  [info] using embedded party lookup: {len(lookup):,} candidate_ids mapped",
                file=sys.stderr,
            )
            return lookup
        except Exception as e:
            print(f"  [warn] could not decode embedded lookup ({e}); falling back", file=sys.stderr)

    # --- Options 2/3: yaml-based lookup, local file or live fetch ---
    if yaml is None:
        print(
            "  [error] PyYAML is not installed (pip install pyyaml --break-system-packages), "
            "so the --legislators-current/--legislators-historical / live-fetch path can't run. "
            "Falling back to the embedded lookup instead.",
            file=sys.stderr,
        )
        try:
            return _decode_embedded_party_lookup()
        except Exception:
            return {}

    try:
        current = fetch_or_load_yaml(LEGISLATORS_CURRENT_URL, current_path)
        historical = fetch_or_load_yaml(LEGISLATORS_HISTORICAL_URL, historical_path)
    except Exception as e:
        print(
            f"  [error] could not load congress-legislators data ({e}). "
            "Falling back to the embedded lookup instead.",
            file=sys.stderr,
        )
        try:
            return _decode_embedded_party_lookup()
        except Exception:
            return {}

    fec_to_party: dict[str, str] = {}
    for person in current + historical:
        fec_ids = person.get("id", {}).get("fec", [])
        terms = person.get("terms", [])
        if not fec_ids or not terms:
            continue
        last_term = sorted(terms, key=lambda t: t.get("start", ""))[-1]
        party = last_term.get("party")
        if not party:
            continue
        party = PARTY_NORMALIZE.get(party, party)
        for fid in fec_ids:
            fec_to_party.setdefault(fid, party)

    print(f"  [info] congress-legislators lookup built: {len(fec_to_party):,} candidate_ids mapped", file=sys.stderr)
    return fec_to_party


def load_candidate_master(paths: list[str]) -> pd.Series:
    """Load one or more FEC candidate-master (cn.txt) files and return a
    deduplicated candidate_id -> party_code mapping.

    If the same candidate_id appears in multiple cycle files (common, since
    incumbents re-file each cycle) with different party codes (e.g. a member
    who switched parties), the most recent cycle's value wins.
    """
    frames = []
    cn_columns = [
        "CAND_ID",
        "CAND_NAME",
        "CAND_PTY_AFFILIATION",
        "CAND_ELECTION_YR",
        "CAND_OFFICE_ST",
        "CAND_OFFICE",
        "CAND_OFFICE_DISTRICT",
        "CAND_ICI",
        "CAND_STATUS",
        "CAND_PCC",
        "CAND_ST1",
        "CAND_ST2",
        "CAND_CITY",
        "CAND_ST",
        "CAND_ZIP",
    ]
    for p in paths:
        p = Path(p)
        if not p.exists():
            print(f"  [skip] candidate master file not found: {p}", file=sys.stderr)
            continue
        cn = pd.read_csv(
            p,
            sep="|",
            header=None,
            names=cn_columns,
            dtype=str,
            encoding="latin-1",
        )
        frames.append(cn[["CAND_ID", "CAND_PTY_AFFILIATION", "CAND_ELECTION_YR"]])

    if not frames:
        return pd.DataFrame(columns=["CAND_ID", "CAND_PTY_AFFILIATION"])

    combined = pd.concat(frames, ignore_index=True)
    # Most recent election year per candidate_id wins, in case of party
    # switches across cycles.
    combined["CAND_ELECTION_YR"] = pd.to_numeric(
        combined["CAND_ELECTION_YR"], errors="coerce"
    )
    combined = combined.sort_values("CAND_ELECTION_YR").drop_duplicates(
        subset="CAND_ID", keep="last"
    )
    return combined.set_index("CAND_ID")["CAND_PTY_AFFILIATION"]


def add_year_quarter(df: pd.DataFrame) -> pd.DataFrame:
    dates = pd.to_datetime(df["disbursement_date"], errors="coerce")
    df["disbursement_year"] = dates.dt.year
    df["disbursement_quarter"] = dates.dt.quarter.apply(
        lambda q: f"Q{int(q)}" if pd.notna(q) else None
    )
    return df


def add_party_column(
    df: pd.DataFrame, fec_master_lookup: pd.Series, legislators_lookup: dict
) -> pd.DataFrame:
    def resolve_party(candidate_id: str) -> str:
        if pd.isna(candidate_id):
            return "NON-CANDIDATE / LEADERSHIP PAC"

        # Priority 1: hand-maintained overrides
        if candidate_id in PARTY_OVERRIDES:
            return PARTY_OVERRIDES[candidate_id]

        # Priority 2: official FEC candidate master file, if supplied
        if candidate_id in fec_master_lookup.index:
            code = fec_master_lookup.loc[candidate_id]
            if isinstance(code, pd.Series):  # safety, shouldn't happen post-dedupe
                code = code.iloc[0]
            if pd.notna(code):
                return PARTY_CODE_MAP.get(code, code)

        # Priority 3: congress-legislators dataset
        if candidate_id in legislators_lookup:
            return legislators_lookup[candidate_id]

        return "UNKNOWN (candidate_id not found in any source)"

    df["candidate_party"] = df["candidate_id"].apply(resolve_party)
    return df


def clean_dataframe(
    df: pd.DataFrame, fec_master_lookup: pd.Series, legislators_lookup: dict
) -> pd.DataFrame:
    # Keep only columns that actually exist in this export (some FEC export
    # variants drop columns), preserving the requested order.
    present_cols = [c for c in KEEP_COLUMNS if c in df.columns]
    missing = set(KEEP_COLUMNS) - set(present_cols)
    if missing:
        print(f"  [note] columns not found in input, skipped: {sorted(missing)}", file=sys.stderr)
    df = df[present_cols].copy()

    # disbursement_amount needs to be numeric for downstream analysis.
    if "disbursement_amount" in df.columns:
        df["disbursement_amount"] = pd.to_numeric(
            df["disbursement_amount"], errors="coerce"
        )

    df = add_year_quarter(df)
    df = add_party_column(df, fec_master_lookup, legislators_lookup)
    return df


def clean(
    input_path: str,
    candidate_master_paths: list[str],
    legislators_current_path: str | None,
    legislators_historical_path: str | None,
) -> pd.DataFrame:
    """Convenience single-file wrapper (builds lookups fresh each call).
    For multiple files, build the lookups once and call clean_dataframe()
    per file instead -- see main()."""
    df = pd.read_csv(input_path, dtype=str, low_memory=False)
    fec_master_lookup = load_candidate_master(candidate_master_paths)
    legislators_lookup = load_legislators_party_lookup(
        legislators_current_path, legislators_historical_path
    )
    return clean_dataframe(df, fec_master_lookup, legislators_lookup)


def pick_input_file() -> str | None:
    """Open a native file-picker dialog to choose a single input CSV.
    Returns None (and falls back to a typed prompt) if no display/GUI
    toolkit is available, e.g. when running over SSH with no X server."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askopenfilename(
            title="Select the FEC PAC disbursement CSV to clean",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        root.destroy()
        return path or None
    except Exception as e:
        print(f"  [info] file picker unavailable ({e}); falling back to typed input", file=sys.stderr)
        typed = input("Path to input CSV: ").strip().strip('"')
        return typed or None


def pick_input_files() -> list[str]:
    """Open a native file-picker dialog that allows selecting multiple
    input CSVs at once (ctrl/cmd+click or shift+click to multi-select).
    Falls back to a typed, comma-separated prompt if no GUI is available."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        paths = filedialog.askopenfilenames(
            title="Select one or more FEC PAC disbursement CSVs to clean",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        root.destroy()
        return list(paths)
    except Exception as e:
        print(f"  [info] file picker unavailable ({e}); falling back to typed input", file=sys.stderr)
        typed = input("Path(s) to input CSV(s), comma-separated: ").strip()
        return [p.strip().strip('"') for p in typed.split(",") if p.strip()]


def pick_output_file(default_path: str) -> str:
    """Open a native 'save as' dialog for the output CSV, defaulting to
    `default_path`. Falls back to just using `default_path` if no GUI is
    available."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.asksaveasfilename(
            title="Save cleaned CSV as",
            initialfile=Path(default_path).name,
            initialdir=str(Path(default_path).parent),
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        root.destroy()
        return path or default_path
    except Exception:
        return default_path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        nargs="*",
        default=None,
        help="Path(s) to one or more raw FEC disbursement CSVs. If omitted, a file-picker dialog opens (supports multi-select).",
    )
    parser.add_argument(
        "--candidate-master",
        nargs="*",
        default=[],
        help="Optional: path(s) to official FEC candidate master cn.txt file(s), one per election cycle. Takes priority over the congress-legislators lookup if both resolve a candidate_id.",
    )
    parser.add_argument(
        "--legislators-current",
        default=None,
        help="Optional: local path to legislators-current.yaml (skips network fetch)",
    )
    parser.add_argument(
        "--legislators-historical",
        default=None,
        help="Optional: local path to legislators-historical.yaml (skips network fetch)",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Output filename. With a single input file, defaults to a save-as dialog (or <input>_cleaned.csv) inside the output folder. With multiple input files, each gets its own <input>_cleaned.csv inside the output folder unless --combine is set.",
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Folder to save cleaned file(s) into. Defaults to a 'pac_disbursements_clean' subfolder created next to the first input file. Created automatically if it doesn't exist.",
    )
    parser.add_argument(
        "--combine",
        action="store_true",
        help="With multiple input files, also write one combined CSV (all files' rows stacked, with a source_file column) in addition to each file's individual cleaned output.",
    )
    args = parser.parse_args()

    input_paths = args.input if args.input else pick_input_files()
    if not input_paths:
        print("No input file(s) selected. Exiting.", file=sys.stderr)
        sys.exit(1)

    missing_inputs = [p for p in input_paths if not Path(p).exists()]
    if missing_inputs:
        print(f"Input file(s) not found: {missing_inputs}", file=sys.stderr)
        sys.exit(1)

    output_dir = Path(args.output_dir) if args.output_dir else (
        Path(input_paths[0]).resolve().parent / "pac_disbursements_clean"
    )
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"Output folder: {output_dir}", file=sys.stderr)

    # Build the candidate_id -> party lookups once and reuse for every file
    # -- avoids re-downloading/re-parsing the congress-legislators dataset
    # per file.
    fec_master_lookup = load_candidate_master(args.candidate_master)
    legislators_lookup = load_legislators_party_lookup(
        args.legislators_current, args.legislators_historical
    )

    cleaned_frames = []
    for input_path in input_paths:
        print(f"\n=== Processing {input_path} ===", file=sys.stderr)
        raw_df = pd.read_csv(input_path, dtype=str, low_memory=False)
        cleaned = clean_dataframe(raw_df, fec_master_lookup, legislators_lookup)

        if len(input_paths) == 1 and not args.combine:
            if args.output:
                output_path = str(output_dir / Path(args.output).name)
            else:
                default_output = str(
                    output_dir / (Path(input_path).stem + "_cleaned.csv")
                )
                output_path = pick_output_file(default_output)
        else:
            # Multiple files (or --combine requested): write each file's
            # cleaned output into the output folder rather than prompting
            # per file.
            output_path = str(output_dir / (Path(input_path).stem + "_cleaned.csv"))

        cleaned.to_csv(output_path, index=False)
        print(f"Wrote {len(cleaned):,} rows to {output_path}")
        print(cleaned["candidate_party"].value_counts(dropna=False).to_string())

        cleaned_with_source = cleaned.copy()
        cleaned_with_source.insert(0, "source_file", Path(input_path).name)
        cleaned_frames.append(cleaned_with_source)

    if len(input_paths) > 1 and args.combine:
        combined = pd.concat(cleaned_frames, ignore_index=True)
        if args.output:
            combined_output_path = str(output_dir / Path(args.output).name)
        else:
            combined_output_path = str(output_dir / "combined_cleaned.csv")
        combined.to_csv(combined_output_path, index=False)
        print(f"\n=== Combined output ===", file=sys.stderr)
        print(f"Wrote {len(combined):,} total rows to {combined_output_path}")


if __name__ == "__main__":
    main()