# Business rules and column mapping (source of truth for build_report.py)

This condenses the user's own manual-process instruction file plus what was
verified against real exports. If a rule here and the code disagree, the code
comment at that rule explains why (usually: the plain reading of the
instruction didn't hold up against real data).

## Source files and what they're for

| File | Real content confirmed | Used for |
|---|---|---|
| input1 OPD export | 1 sheet, ~5k rows, 91 cols | PO number, ETA/ETD, in-transit, ETA port, supplier |
| input2 InvMaster export | 1 sheet, ~29k rows, 76 cols | On-hand/on-order qty, code sort, inactive flag, vendor, monthly forecast (T7_2026...) |
| input3 GAB 6271 "Open Work Orders" export | 1 sheet, ~88k rows, 81 cols | Per-boat material demand: `Qty Rem.`, `Order Required Date`, `WO Mat Due Date`, boat via `Customer PO` |
| input4 Seawind production schedule | 7 sheets; sheet1 "SW production Schedule 6W" is the one used | Boat's `Planned boat completion date`, header on row 3 (skiprows=2) |
| template | 2 sheets: `original` (hidden, 24 cols, per-material) and `original (2)` (visible, 32 cols A-AF, header row 3, data from row 4, autoFilter A3:AF{last}) | Layout/style/autofilter to reuse |

## Key join keys (verified against real data, not just column names)

- **Part number**: `PART` (OPD/InvMaster) / `Material` (6271), normalized via
  `.strip().upper()`. 100% of 6271 materials matched InvMaster; ~37% matched
  an OPD open PO line (the rest have no open PO placed yet, which is itself
  meaningful - flagged for review, not an error).
- **Boat number**: 6271's **`Customer PO`** column, not `Project` (which is
  populated on already-closed/legacy rows, not the current open-demand rows -
  don't trust its name). Format for Seawind rows: `SC625-SW1370`. Matched
  `^SC\d+` for 92% of actionable rows; the rest are other brands (Corsair
  etc.) sharing the same GAB system and are correctly excluded.
- **Boat completion date**: input4 sheet1, column `SC #` -> `Planned boat
  completion date`. Confirmed 100% of 6271's 44 distinct SC boats exist in
  this schedule.

## Base row filter (what counts as "still needed")

A 6271 row is a candidate row if:
- `Qty Rem.` > 0 (fully-issued lines are done, drop)
- `Job Seq Closed == "N"` (closed job sequences are done, drop)
- `Customer PO` matches `^SC\d+` (keeps Seawind boats only)

## ETA REQUIRED and GAP formula

- **ETA REQUIRED** (col O) = 6271 **`WO Mat Due Date`** - the date the material
  is actually needed for the work order. This drives both the GAP and the
  need-date ordering used to allocate stock/PO to boats.
- `GAP = ETA REQUIRED - ETA (matched OPD PO line)`. Positive = buffer before
  the part is needed; negative = it will arrive after it's needed.
- **WO ADJUST DATE** (col R) = when material was actually generated for the WO
  (`Actual Material Generation Date`), falling back to `WO Open Date`.

## Mechanical drop rules (safe to hard-delete, confirmed against real data)

- InvMaster `INACTIVE == True`.
- InvMaster `CODE_SORT` in the drop list - **exact real values**: `CONSUMABLE`,
  `LAM CONS`, `TOOL CONS`, `TOOL MAN`, `LAMINATE`, `FOAM`, `RESIN/GEL`,
  `PAINT/CHEM` (note singular `CONSUMABLE`, confirmed via `value_counts()`).
  These consumable / auxiliary-material groups are dropped **outright, with no
  critical exception** - the buyer confirmed they never belong in this report.
- Local vendor: `VENDOR_ADDRESS` contains `VIET NAM`/`VIETNAM`. This still keeps
  a critical-keyword exception (`config.yaml: critical_vendor_keywords` /
  `critical_part_keywords`).
- OPD lines dropped **before any PO matching** (so they are never pegged to a
  boat or shown in the GSS block):
  - `QTY_OPEN <= 0.0001` - sentinel for closed/cancelled/dropship/non-part
    lines the buyer has effectively cancelled.
  - note text (`USER5` / `PO_LINE_TEXT`) contains `cancel` or `NCR` (matched as
    a whole word, `\bNCR\b`, so `increase` isn't caught) - the buyer cancelled
    the line or raised an NCR.
- **Using stock (per boat)**: on-hand stock (`QTY_ONHAND`) is allocated to boats
  in need-date order (earliest `WO Mat Due Date` first). A boat whose demand
  begins while stock is still available is dropped as "using stock"; only boats
  past the on-hand quantity are pegged to incoming POs. This replaced the old
  material-level `QTY_ONHAND >= TOTAL_DEMAND` check, which wrongly pegged a boat
  to an incoming PO even when on-hand stock could have covered it.

## Review-flag rules (kept, never silently dropped)

- **No open PO found**: material has open demand but no matching OPD PO line
  at all - arguably the most actionable flag (nothing has been ordered yet).
- **Boat not found in schedule**: `Customer PO` boat number isn't in input4 -
  usually a data-entry mismatch worth a human look.
- **Already received/closed, already at port, GAP <= 7, GAP <= 14 with >= 30
  days of boat-completion buffer**: these read like hard removal criteria in
  the original instructions, but checking the actual finished WK27 report
  showed roughly a third of its rows still had `GAP <= 7` - so in practice a
  human made the final call rather than a formula. Flag, don't drop.

## Column mapping for `original (2)` (the visible report, A-AF)

| Col | Header | Source |
|---|---|---|
| A | BOAT | 6271 `Customer PO`, text before `-` |
| B | BOAT (label) | Constructed: `{SC#}-{Model}- planned_completion-{completion date}-{delivery option}, `. Boat number is the text before `-` in `Customer PO` (5-char `SC###` for Seawind, 6-char numeric e.g. `880067` for trimarans). Delivery option comes from input4; blank -> `TBC` (never `nan`). |
| C | Material | 6271 `Material` |
| D | Extra Desc | 6271 `Extra Desc` |
| E | Inv UM | InvMaster `UM_INVENTORY` |
| F | Code Sort | InvMaster `CODE_SORT` |
| G | BOM QTY | 6271 `Qty Rem.` for this boat/material row |
| H | Total | Sum of `Qty Rem.` across all open boats needing that material |
| I | On Hand | InvMaster `QTY_ONHAND` |
| J | Qty OnOrder | InvMaster `QTY_ONORDER_PO` |
| K | Default Vendor | InvMaster `VENDOR_NAME` |
| L | Buyer | OPD `ORIGINATOR` of the matched PO line |
| M | ETA | OPD `ETA_PO_LINE` of the matched (soonest-ETA) open PO line |
| N | (helper) | `Material + Boat` composite key, kept for parity with the template |
| O | ETA REQUIRED | 6271 `WO Mat Due Date` |
| P | BOAT COMPLETION DATE | input4 `Planned boat completion date` |
| Q | GAP | `O - M` |
| R | WO ADJUST DATE | 6271 `Actual Material Generation Date` (falls back to `WO Open Date`) |
| S | PO | OPD `PURCHASE_ORDER` of the matched line |
| T | IN TRANSIT | OPD `INTRANSIT` |
| U | ETA PORT | OPD `ETA_PORT` |
| V | Quantity Coming | OPD `QTY_OPEN` of the matched line |
| W, X | REMARK | left blank for manual buyer notes |
| Y | NOTE | auto-generated review reason(s), if any |
| Z-AE | *_GSS | all open OPD PO lines for that material (not just the matched one), pipe-joined |
| AF | PO_LINE_TEXT | 6271 `Job Seq Comment` (matches the template's actual content style, e.g. "MAST/ Electrical - navigation - ...") |

`original` (hidden) mirrors this at material level (one row per part, before
the boat join) using the same source columns, for parity with the template's
own two-sheet structure.
