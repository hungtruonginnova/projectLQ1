# projectLQ1 — Overdue Parts Report

Weekly procurement report for Seawind catamaran production. Joins four ERP/GAB
exports (OPD, InvMaster, GAB 6271 open work orders, production schedule) into
one Excel workbook that shows whether purchase-order ETAs will meet each boat's
material need dates.

## Repository layout

| Path | Purpose |
|------|---------|
| `skill/overdue-parts-report/` | Cursor agent skill + canonical `build_report.py` pipeline |
| `hf-space/` | FastAPI backend deployed to Hugging Face Spaces |
| `frontend/` | Next.js upload UI deployed to Vercel |
| `Overdue raw data/` | Sample inputs, templates, and generated reports (Git LFS) |
| `DEPLOYMENT.md` | HF Space + Vercel setup and URLs |

The skill script (`skill/.../scripts/build_report.py`) and the HF Space copy
(`hf-space/build_report.py`) implement the same logic. Prefer editing the skill
version first, then sync to `hf-space/` when deploying API changes.

## Inputs (four weekly exports + template)

1. **OPD** — PO tracking (ETA, in-transit, supplier, open qty)
2. **InvMaster** — on-hand qty, qty on order, code sort, vendor, inactive flag
3. **GAB 6271 Open Work Orders** — per-boat material demand (`Qty Rem.`, need dates)
4. **Production schedule** — boat `SC #` and planned completion date
5. **Template** — previous week's `OVERDUE PARTS REPORT ... PQ.xlsx` (layout/styles)

## Running locally

### CLI (skill script)

```bash
cd skill/overdue-parts-report
pip install -r scripts/requirements.txt
python scripts/build_report.py \
  --opd "<OPD path>" \
  --invmaster "<InvMaster path>" \
  --workorders "<GAB 6271 path>" \
  --schedule "<schedule path>" \
  --template "<template path>" \
  --out "<output path>" \
  --config scripts/config.yaml
```

### Frontend + API

```bash
# API (hf-space)
cd hf-space && pip install -r requirements.txt && uvicorn api:app --reload

# Frontend
cd frontend && cp .env.example .env.local && npm install && npm run dev
```

Set `NEXT_PUBLIC_API_URL` in `frontend/.env.local` to the HF Space URL.

## Agent skill

Use `skill/overdue-parts-report/SKILL.md` whenever the user asks to generate or
update the overdue parts report. Read `references/instructions.md` for business
rules; tune drop/flag thresholds in `scripts/config.yaml`.

Key pipeline concepts (see SKILL.md for detail):

- Base row = open 6271 work-order line with `Qty Rem. > 0`, boat parsed from
  `Customer PO` (`^SC\d+`, e.g. `SC625-SW1370`)
- Stock/PO allocation: boats sorted by `WO Mat Due Date`; on-hand stock consumed
  first, then open POs by ETA; "using stock" boats are dropped
- `ETA REQUIRED` = `WO Mat Due Date`; `GAP` = `ETA REQUIRED - ETA`
- Mechanical drops vs. review flags — ambiguous rules are flagged (yellow fill +
  `NOTE`), not silently deleted

## Deployment

| Service | URL |
|---------|-----|
| API | https://truonghung239810-lqvhspace.hf.space |
| GitHub | https://github.com/hungtruonginnova/projectLQ1 |

See `DEPLOYMENT.md` for Vercel/HF Space configuration, CORS, and env vars.

## Git LFS

All `*.xlsx` files are tracked via Git LFS (`.gitattributes`). After clone,
run `git lfs pull` to fetch binary content. Pointer-only files (~130 bytes) mean
LFS objects are not present locally.

## Conventions

- Conventional commits: `feat`, `fix`, `docs`, `chore`, `refactor`, `test`
- Do not commit secrets (`.env`, credentials)
- `frontend/` root for Vercel; `hf-space/` root for HF Space
