"""
FEC LPAC List Builder — Bulk Data Only
=======================================
Builds a complete list of federal candidates, their authorized campaign
committees, and their leadership PACs for each election cycle — using
only FEC bulk data files. No API calls, no rate limiting.

Bulk files used per cycle:
  cn{yy}.zip          — Candidate master
  ccl{yy}.zip         — Candidate-committee linkages (authorized committees)
  cm{yy}.zip          — Committee master (names, designations)
  leadership{yy}.csv  — Leadership PACs and sponsors

Requirements:
    pip install requests pandas tqdm openpyxl rapidfuzz

Output: fec_leadership_pacs.xlsx
"""

import io
import os
import re
import zipfile
import requests
import pandas as pd
from tqdm import tqdm
from rapidfuzz import process, fuzz

# ── Config ────────────────────────────────────────────────────────────────────
CYCLES          = [2022, 2024, 2026]
CYCLE_LABELS    = {2022: "2021-2022", 2024: "2023-2024", 2026: "2025-2026"}
FILTER_STATES   = []               # e.g. ["AK"] or [] for all
MANUAL_LPAC_CSV = "fec_lpac_wManual.csv"
OUTPUT          = "fec_leadership_pacs.xlsx"
DATA_DIR        = "fec_bulk_data"
MATCH_THRESHOLD = 85
# ─────────────────────────────────────────────────────────────────────────────

BASE_URL = "https://www.fec.gov/files/bulk-downloads/{cycle}/{file}"
os.makedirs(DATA_DIR, exist_ok=True)

# ── Column definitions ────────────────────────────────────────────────────────
CANDIDATE_MASTER_COLS = [
    "candidate_id", "candidate_name", "party_1", "party_2", "election_year",
    "candidate_state", "office", "district", "incumbent_challenger",
    "candidate_status", "principal_committee_id", "address_state",
    "zip", "city", "street1", "street2",
]

CCL_COLS = [
    "candidate_id", "candidate_election_year", "fec_election_year",
    "committee_id", "committee_type", "committee_designation", "linkage_id",
]

COMMITTEE_MASTER_COLS = [
    "committee_id", "committee_name", "treasurer_name", "street1", "street2",
    "city", "state", "zip", "filing_freq", "committee_type",
    "committee_designation", "org_type", "conn_org_name", "candidate_id",
]

LPACK_COLS = [
    "committee_id", "committee_name", "link_image", "sponsor_name",
    "total_receipts", "total_disbursements", "cash_on_hand", "cov_end_date",
]

PARTY_CODES = {
    "DEM": "Democratic Party", "REP": "Republican Party",
    "IND": "Independent",      "LIB": "Libertarian Party",
    "GRE": "Green Party",      "NNE": "No Party Affiliation",
}

OFFICE_CODES = {"H": "House", "S": "Senate", "P": "President"}


# ── Helper functions ──────────────────────────────────────────────────────────

def normalize(name):
    """Lowercase and strip punctuation for fuzzy name matching."""
    return re.sub(r"[^a-z\s]", "", str(name).lower()).strip()


def assign_columns(df, cols, label):
    actual, expected = len(df.columns), len(cols)
    if actual != expected:
        print(f"  Note: {label} has {actual} columns, expected {expected} — adjusting.")
    if actual <= expected:
        df.columns = cols[:actual]
    else:
        df.columns = cols + [f"extra_{i}" for i in range(actual - expected)]
    return df


def download_zip(filename, cycle):
    """Download and extract a zipped FEC bulk file, return as DataFrame."""
    local_zip = os.path.join(DATA_DIR, filename)
    url = BASE_URL.format(cycle=cycle, file=filename)

    if not os.path.exists(local_zip):
        print(f"  Downloading {filename}...")
        r = requests.get(url, stream=True, timeout=120)
        r.raise_for_status()
        total = int(r.headers.get("content-length", 0))
        with open(local_zip, "wb") as f, tqdm(
            total=total, unit="B", unit_scale=True, desc=filename
        ) as bar:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)
                bar.update(len(chunk))
    else:
        print(f"  Using cached {filename}")

    with zipfile.ZipFile(local_zip) as z:
        inner = [n for n in z.namelist() if not n.startswith("__")][0]
        with z.open(inner) as f:
            raw = f.read()

    return pd.read_csv(
        io.BytesIO(raw), sep="|", header=None,
        encoding="latin-1", low_memory=False, dtype=str,
    )


def download_leadership_csv(cycle):
    """Download the leadership PAC CSV (plain CSV, not zipped, different URL pattern)."""
    filename  = f"leadership{cycle}.csv"
    local_csv = os.path.join(DATA_DIR, filename)
    url = f"https://www.fec.gov/files/bulk-downloads/data.fec.gov/{filename}"

    if not os.path.exists(local_csv):
        print(f"  Downloading {filename}...")
        r = requests.get(url, stream=True, timeout=120)
        r.raise_for_status()
        total = int(r.headers.get("content-length", 0))
        with open(local_csv, "wb") as f, tqdm(
            total=total, unit="B", unit_scale=True, desc=filename
        ) as bar:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)
                bar.update(len(chunk))
    else:
        print(f"  Using cached {filename}")

    return pd.read_csv(local_csv, encoding="latin-1", dtype=str, low_memory=False)


# ── Main loop: one pass per cycle ─────────────────────────────────────────────
cycle_dfs = {}

for cycle in CYCLES:
    yy    = str(cycle)[2:]
    label = CYCLE_LABELS[cycle]
    print(f"\n{'='*60}")
    print(f"  Building LPAC list for {label}")
    print(f"{'='*60}")

    # ── Download bulk files ───────────────────────────────────────────────────
    df_cn = download_zip(f"cn{yy}.zip", cycle)
    df_cn = assign_columns(df_cn, CANDIDATE_MASTER_COLS, f"Candidate master {label}")

    df_ccl = download_zip(f"ccl{yy}.zip", cycle)
    df_ccl = assign_columns(df_ccl, CCL_COLS, f"CCL {label}")

    df_cm = download_zip(f"cm{yy}.zip", cycle)
    df_cm = assign_columns(df_cm, COMMITTEE_MASTER_COLS, f"Committee master {label}")

    df_lpack = download_leadership_csv(cycle)

    print(f"  Candidates:       {len(df_cn):,}")
    print(f"  CCL linkages:     {len(df_ccl):,}")
    print(f"  Committees:       {len(df_cm):,}")
    print(f"  Leadership PACs:  {len(df_lpack):,}")
    print(f"  LPAC columns:     {list(df_lpack.columns)}")

    # ── Filter candidates by state ────────────────────────────────────────────
    if FILTER_STATES:
        states_upper = [s.upper() for s in FILTER_STATES]
        df_cn = df_cn[df_cn["candidate_state"].str.upper().isin(states_upper)]
        print(f"  Candidates after state filter ({FILTER_STATES}): {len(df_cn):,}")

    cand_ids_in_scope = set(df_cn["candidate_id"].unique())

    # ── Authorized committees via CCL ─────────────────────────────────────────
    df_auth = df_ccl[df_ccl["committee_designation"].isin(["P", "A"])].copy()
    df_auth = df_auth[["candidate_id", "committee_id"]].drop_duplicates()

    df_cm_names = df_cm[["committee_id", "committee_name"]].drop_duplicates("committee_id")
    df_auth = df_auth.merge(df_cm_names, on="committee_id", how="left")

    df_auth_grouped = (
        df_auth.groupby("candidate_id")
        .agg(
            auth_committee_ids   = ("committee_id",   lambda x: "; ".join(x.dropna())),
            auth_committee_names = ("committee_name", lambda x: "; ".join(x.dropna())),
        )
        .reset_index()
    )

    # ── Match LPACs to candidates via sponsor name ────────────────────────────
    # Column names confirmed from actual file
    sponsor_col      = "Sponsor_Name"
    comm_id_col      = "Committee_Id"
    comm_name_col    = "Committee_Name"
    tot_receipts_col = "Total_Receipt"
    tot_disb_col     = "Total_Disbursement"
    cash_col         = "Cash_on_Hand"
    print(f"  Using sponsor column: '{sponsor_col}'")

    # Build candidate name lookup
    cand_name_lookup = {}
    for _, r in df_cn.iterrows():
        raw = str(r.get("candidate_name", ""))
        cand_name_lookup[normalize(raw)] = r["candidate_id"]
        if "," in raw:
            last, first = raw.split(",", 1)
            cand_name_lookup[normalize(f"{first.strip()} {last.strip()}")] = r["candidate_id"]
    name_keys = list(cand_name_lookup.keys())

    lpac_rows = []
    for _, lp in df_lpack.iterrows():
        comm_id      = str(lp.get(comm_id_col) or "").strip()
        comm_name    = str(lp.get(comm_name_col) or "").strip()
        sponsor      = str(lp.get(sponsor_col) or "").strip()
        tot_receipts = str(lp.get(tot_receipts_col) or "").strip()
        tot_disb     = str(lp.get(tot_disb_col) or "").strip()
        cash         = str(lp.get(cash_col) or "").strip()

        if not comm_id or not sponsor:
            continue

        matched_cid  = None
        match_method = None

        norm = normalize(sponsor)
        if norm in cand_name_lookup:
            cid = cand_name_lookup[norm]
            if cid in cand_ids_in_scope:
                matched_cid  = cid
                match_method = "exact_name"

        if not matched_cid:
            result = process.extractOne(
                norm, name_keys,
                scorer=fuzz.token_sort_ratio,
                score_cutoff=MATCH_THRESHOLD,
            )
            if result:
                matched_name, score, _ = result
                cid = cand_name_lookup[matched_name]
                if cid in cand_ids_in_scope:
                    matched_cid  = cid
                    match_method = f"fuzzy({score})"

        if matched_cid:
            lpac_rows.append({
                "candidate_id":        matched_cid,
                "leadership_pac_id":   comm_id,
                "leadership_pac_name": comm_name,
                "lpac_sponsor_name":   sponsor,
                "lpac_total_receipts": tot_receipts,
                "lpac_disbursements":  tot_disb,
                "lpac_cash_on_hand":   cash,
                "match_method":        match_method,
            })

    df_lpac_matched = pd.DataFrame(lpac_rows) if lpac_rows else pd.DataFrame(
        columns=["candidate_id", "leadership_pac_id", "leadership_pac_name",
                 "lpac_sponsor_name", "lpac_total_receipts", "lpac_disbursements",
                 "lpac_cash_on_hand", "match_method"]
    )

    # ── Supplementary pass: fill unmatched LPACs from manual CSV ─────────────
    if os.path.exists(MANUAL_LPAC_CSV):
        print(f"  Running supplementary match from {MANUAL_LPAC_CSV}...")
        df_manual = pd.read_csv(MANUAL_LPAC_CSV, dtype=str)
        df_manual.columns = [c.strip() for c in df_manual.columns]

        # IDs already matched in the primary pass
        matched_ids = set(df_lpac_matched["leadership_pac_id"].str.upper())

        manual_rows = []
        for _, lp in df_manual.iterrows():
            comm_id   = str(lp.get("Committee_Id") or "").strip()
            comm_name = str(lp.get("Committee_Name") or "").strip()
            sponsor   = str(lp.get("Sponsor_Name") or "").strip()

            # Skip if already matched or no sponsor name
            if not comm_id or not sponsor or comm_id.upper() in matched_ids:
                continue

            matched_cid  = None
            match_method = None

            norm = normalize(sponsor)
            if norm in cand_name_lookup:
                cid = cand_name_lookup[norm]
                if cid in cand_ids_in_scope:
                    matched_cid  = cid
                    match_method = "manual_exact"

            if not matched_cid:
                result = process.extractOne(
                    norm, name_keys,
                    scorer=fuzz.token_sort_ratio,
                    score_cutoff=MATCH_THRESHOLD,
                )
                if result:
                    matched_name, score, _ = result
                    cid = cand_name_lookup[matched_name]
                    if cid in cand_ids_in_scope:
                        matched_cid  = cid
                        match_method = f"manual_fuzzy({score})"

            if matched_cid:
                manual_rows.append({
                    "candidate_id":        matched_cid,
                    "leadership_pac_id":   comm_id,
                    "leadership_pac_name": comm_name,
                    "lpac_sponsor_name":   sponsor,
                    "lpac_total_receipts": str(lp.get("Total_Receipt") or "").strip(),
                    "lpac_disbursements":  str(lp.get("Total_Disbursement") or "").strip(),
                    "lpac_cash_on_hand":   str(lp.get("Cash_on_Hand") or "").strip(),
                    "match_method":        match_method,
                })

        if manual_rows:
            df_manual_matched = pd.DataFrame(manual_rows)
            df_lpac_matched   = pd.concat(
                [df_lpac_matched, df_manual_matched], ignore_index=True
            )
            print(f"  Supplementary matches added: {len(manual_rows):,}")
        else:
            print(f"  Supplementary pass: no new matches found.")
    else:
        print(f"  Supplementary CSV not found ({MANUAL_LPAC_CSV}) — skipping.")

    # ── Assemble final DataFrame ──────────────────────────────────────────────
    df_out = df_cn[[
        "candidate_id", "candidate_name", "party_1", "office",
        "candidate_state", "district", "candidate_status",
    ]].copy()
    df_out.rename(columns={"party_1": "party", "candidate_state": "state"}, inplace=True)
    df_out["party"]  = df_out["party"].map(PARTY_CODES).fillna(df_out["party"])
    df_out["office"] = df_out["office"].map(OFFICE_CODES).fillna(df_out["office"])
    df_out["fec_candidate_url"] = df_out["candidate_id"].apply(
        lambda x: f"https://www.fec.gov/data/candidate/{x}/"
    )
    df_out["election_cycle"] = label

    df_out = df_out.merge(df_auth_grouped, on="candidate_id", how="left")
    df_out = df_out.merge(df_lpac_matched, on="candidate_id", how="left")
    df_out["has_leadership_pac"] = df_out["leadership_pac_id"].notna()

    col_order = [
        "candidate_id", "candidate_name", "office", "state", "district",
        "party", "candidate_status", "election_cycle",
        "auth_committee_ids", "auth_committee_names",
        "leadership_pac_id", "leadership_pac_name",
        "lpac_sponsor_name", "lpac_total_receipts",
        "lpac_disbursements", "lpac_cash_on_hand",
        "match_method", "has_leadership_pac", "fec_candidate_url",
    ]
    df_out = df_out[[c for c in col_order if c in df_out.columns]]
    df_out.sort_values(["candidate_name"], ascending=True, inplace=True)

    cycle_dfs[cycle] = df_out
    n_lpac = df_out["has_leadership_pac"].sum()
    print(f"  Output: {len(df_out):,} rows, {n_lpac:,} with leadership PACs")


# ── Write to Excel ────────────────────────────────────────────────────────────
print(f"\nWriting to {OUTPUT}...")
with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:

    # All cycles combined into one tab
    df_all = pd.concat(cycle_dfs.values(), ignore_index=True)
    df_all.sort_values("candidate_name", inplace=True)
    df_all.to_excel(writer, index=False, sheet_name="All Cycles")

    for sheet in writer.sheets.values():
        for col in sheet.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=10)
            sheet.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)

print(f"\nDone! Saved to: {OUTPUT}")
total_lpac = df_all["has_leadership_pac"].sum()
print(f"  Total candidates across all cycles: {len(df_all):,}")
print(f"  With leadership PAC:                {total_lpac:,}")