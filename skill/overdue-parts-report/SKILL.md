---
name: overdue-parts-report
description: Build the weekly "Overdue Parts Report" for Seawind catamaran production - joins PO tracking (OPD), inventory/forecast (InvMaster), open work-order material demand (GAB 6271), and the boat production schedule into one report that shows whether purchasing's ETAs will meet each boat's production need. Use this whenever the user asks to generate/update the overdue parts report, check if purchase orders will meet production schedule, track PO vs boat completion dates, or mentions files like "OPD export", "InvMaster export", "GAB 6271 open work orders", or the "overdue parts report" template - even if they just hand over a folder of exports and ask to "run the usual report." Always use this skill instead of manually re-deriving the join/filter logic from scratch.
---

# Overdue Parts Report

Turns four weekly ERP/GAB exports into the Overdue Parts Report used to check
whether current purchase order schedules will meet production's need dates.
The manual process this replaces is described in
`references/instructions.md` (translated/condensed from the user's own
instruction file) - read it once if you need to understand *why* a rule
exists, but the executable version of that process lives in
`scripts/build_report.py`.

## What you need from the user

Four input files (any filenames, matched by content not name) plus the
report template:

1. **OPD export** - PO tracking: PO number, supplier, part, ETA/ETD, in-transit
   status, ETA at port.
2. **InvMaster export** - inventory master: on-hand qty, qty on order, code
   sort (part category), inactive flag, vendor.
3. **GAB 6271 "Open Work Orders" export** - per-boat material demand: which
   boat needs which part, how much is still outstanding (`Qty Rem.`), and the
   two production due dates (`Order Required Date`, `WO Mat Due Date`).
4. **Production schedule** (e.g. "Seawind production schedule update ...") -
   the boat-by-boat schedule sheet with `SC #` and `Planned boat completion
   date`.
5. **Template** - a previous week's finished "OVERDUE PARTS REPORT ... PQ.xlsx"
   file. Its layout/styles/autofilter are reused for the new output - do not
   ask the user to redesign it.

If the user only gives some of the files, ask for the missing ones rather
than guessing - the join logic genuinely needs all four.

## Running it

```
python scripts/build_report.py \
  --opd "<OPD export path>" \
  --invmaster "<InvMaster export path>" \
  --workorders "<GAB 6271 export path>" \
  --schedule "<production schedule path>" \
  --template "<previous week's report, used as the style/layout template>" \
  --out "<output path, e.g. OVERDUE PARTS REPORT WK 28 (06.07.2026) PQ.xlsx>" \
  --config scripts/config.yaml
```

Requires `pandas`, `openpyxl`, `pyyaml` (`pip install -r scripts/requirements.txt`).
The GAB 6271 export is large (tens of MB, ~90k rows) - the script only reads
the columns it needs and typically finishes in under a minute.

The script prints, in order: input row counts, a breakdown of how many rows
were dropped by each *mechanical* rule, and the final row count with how many
are flagged for review. Read this back to the user after running - it is the
main way they sanity-check a given week's run before opening the spreadsheet.

## How the pipeline thinks (for when things need adjusting)

The base unit of the report is one **(boat, part)** row: a work order line
from the 6271 export that still has outstanding quantity (`Qty Rem. > 0`) and
an open job sequence (`Job Seq Closed == "N"`). The boat number is parsed out
of 6271's **`Customer PO`** column, which for Seawind jobs is formatted like
`SC625-SW1370` (`Customer PO`, not `Project`, despite the name - this was
confirmed against real data, don't "fix" it back). Non-Seawind brands sharing
the same GAB system (Corsair etc.) don't match `^SC\d+` and are naturally
excluded.

Each row is then enriched by joining:
- **InvMaster** on part number → on-hand qty, qty on order, code sort, vendor,
  inactive flag.
- **OPD** on part number → open PO lines are allocated to boats by need date
  (see stock/PO allocation below); the pegged line drives the primary
  ETA/PO/in-transit columns. Cancelled lines are removed first: `QTY_OPEN <=
  0.0001` (buyer's cancel sentinel) and lines whose `USER5`/`PO_LINE_TEXT` note
  contains `cancel` or `\bNCR\b`. All *surviving* open PO lines (not just the
  pegged one) are concatenated into the "GSS" backup columns (`PO Number GSS`,
  `ETA GSS`, ...), mirroring the manual cross-reference block in the template.
- **Production schedule** on boat number → boat completion date + delivery
  option (shown in the BOAT label column, `TBC` when blank).

`ETA REQUIRED` = 6271 `WO Mat Due Date`; `GAP` = `ETA REQUIRED - ETA` (matches
the template's own `=O-M` formula). A large positive gap means comfortable
buffer; negative or small means risk.

### Stock / PO allocation (per boat, by need date)

For each material the boats are sorted by need date (`WO Mat Due Date`) and
supply is consumed top-down: **on-hand stock first**, then open PO lines by
ETA. A boat whose demand begins while on-hand stock is still available is a
"using stock" boat and is **dropped**; the first boat past the on-hand quantity
is pegged to the earliest incoming PO, and so on. Boats beyond all supply get
no PO (a genuine, still-unordered shortage → flagged). This replaced an older
material-level `on-hand >= total demand` check that wrongly pegged a boat to an
incoming PO even when stock could have covered it.

### Mechanical drops vs. review flags

**Hard drops** (row removed, counted in the printed log): inactive parts,
**all** consumable / auxiliary-material code-sorts (no exception), non-critical
local-vendor parts, cancelled/zero-qty OPD lines, "using stock" boats (per the
allocation above), and the base filters baked into "open work order with
outstanding qty" itself.

Everything judgment-based is **kept and flagged** (yellow fill + a reason in
the `NOTE` column) instead of silently deleted: missing PO/schedule data, and
the gap/received/at-port criteria from `references/instructions.md`. Those
criteria read like a hard filter on paper, but a real finished report still had
~31% of rows with `GAP <= 7` - so in practice they're the buyer's final manual
judgment call, not a deterministic rule. Flag, don't delete. Promoting a rule
to a hard drop (or loosening a flag) is a one-line change in
`scripts/config.yaml` or a mask in `build_report.py`, not a redesign.

The output lands in the same ballpark as a hand-finished report (roughly a few
thousand rows for a typical week, vs. the ~7.6k-row WK27 reference). Point the
user at the preserved autofilter and the `NOTE`/fill-color column to slice
further to what they care about.

## Output

A copy of the template with fresh data written into `original (2)` (the
visible report) and `original` (the hidden per-material rollup sheet), styles
and autofilter preserved/resized. See `references/instructions.md` for the
full column-by-column mapping if a column needs to be re-sourced.

Config knobs (drop-list code sorts, critical-vendor/part keywords, gap
thresholds) live in `scripts/config.yaml` - prefer editing that file over
hardcoding new keywords in the script.
