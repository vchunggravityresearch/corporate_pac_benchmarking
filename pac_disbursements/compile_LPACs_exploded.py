"""
FEC LPAC Exploded Format Builder
==================================
Reads fec_leadership_pacs.xlsx (output from the LPAC builder script)
and produces an exploded version with one row per committee
(authorized campaign committees and LPACs), all cycles combined.

Output columns:
  candidate_id, candidate_name, office, state, district, party,
  candidate_status, fec_candidate_url, committee_id, committee_name,
  committee_type

Requirements:
    pip install pandas openpyxl

Input:  fec_leadership_pacs.xlsx
Output: fec_lpac_exploded.xlsx
"""

import pandas as pd

# ── Config ────────────────────────────────────────────────────────────────────
INPUT_FILE = "fec_leadership_pacs.xlsx"
OUTPUT     = "fec_lpac_exploded.xlsx"
# ─────────────────────────────────────────────────────────────────────────────

CAND_COLS = [
    "candidate_id", "candidate_name", "office", "state", "district",
    "party", "candidate_status", "fec_candidate_url",
]


def split(s):
    return [x.strip() for x in str(s or "").split(";") if x.strip()]


# ── Load all sheets and combine ───────────────────────────────────────────────
print(f"Reading {INPUT_FILE}...")
all_sheets = pd.read_excel(INPUT_FILE, sheet_name=None, dtype=str)

# Accept either a single "All Cycles" sheet or per-cycle "All YYYY-YYYY" sheets
sheets_to_use = {
    name: df for name, df in all_sheets.items()
    if name == "All Cycles" or name.startswith("All ")
}
if not sheets_to_use:
    sheets_to_use = all_sheets  # fallback: use everything

df_input = pd.concat(sheets_to_use.values(), ignore_index=True)
df_input = df_input.drop_duplicates()
print(f"  Loaded {len(df_input):,} rows from {list(sheets_to_use.keys())}")


# ── Explode to one row per committee ─────────────────────────────────────────
print("Exploding to one row per committee...")
rows = []

for _, row in df_input.iterrows():
    base = {c: row.get(c, "") for c in CAND_COLS if c in df_input.columns}

    # Authorized campaign committees
    auth_ids   = [x for x in split(row.get("auth_committee_ids",   "")) if x.lower() != "nan"]
    auth_names = [x for x in split(row.get("auth_committee_names", "")) if x.lower() != "nan"]
    if not auth_ids:
        continue
    for i, cid in enumerate(auth_ids):
        rows.append({
            **base,
            "committee_id":   cid,
            "committee_name": auth_names[i] if i < len(auth_names) else "",
            "committee_type": "Authorized Campaign Committee",
        })

    # Leadership PAC
    lpac_id   = str(row.get("leadership_pac_id",   "") or "").strip()
    lpac_name = str(row.get("leadership_pac_name", "") or "").strip()
    if lpac_id and lpac_id.lower() != "nan":
        rows.append({
            **base,
            "committee_id":   lpac_id,
            "committee_name": lpac_name,
            "committee_type": "Leadership PAC",
        })

df_out = pd.DataFrame(rows)

# Sort: alphabetically by candidate, LPACs before authorized committees
df_out.sort_values(
    ["candidate_name", "committee_type"],
    ascending=[True, False],
    inplace=True,
)

# Deduplicate in case the same committee appears across multiple cycles
df_out.drop_duplicates(subset=["candidate_id", "committee_id"], inplace=True)

print(f"  Total committee rows:          {len(df_out):,}")
print(f"  Authorized campaign committees:{(df_out['committee_type'] == 'Authorized Campaign Committee').sum():,}")
print(f"  Leadership PACs:               {(df_out['committee_type'] == 'Leadership PAC').sum():,}")


# ── Write output ──────────────────────────────────────────────────────────────
print(f"\nWriting to {OUTPUT}...")
with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:

    df_out.to_excel(writer, index=False, sheet_name="All Committees")

    df_out[df_out["committee_type"] == "Leadership PAC"].to_excel(
        writer, index=False, sheet_name="Leadership PACs Only"
    )

    df_out[df_out["committee_type"] == "Authorized Campaign Committee"].to_excel(
        writer, index=False, sheet_name="Auth Committees Only"
    )

    for sheet in writer.sheets.values():
        for col in sheet.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=10)
            sheet.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)

print(f"Done! Saved to: {OUTPUT}")