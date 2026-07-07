# CS Reports v2

A focused Streamlit dashboard for the US Customer Support queue. Upload one XL-Connector
`.xlsx` export and get seven metrics in two tabs, exportable to a styled Excel workbook.
No data is stored — everything is processed in memory per session.

## The seven metrics

**Group A — Update timing** (an *update* = an internal case comment, outbound email, or
internal status change; author role starts with `NA `; events within 60s collapse into one):

1. **Days since last update** — *open cases only.* Today − the latest of (Last Masternaut
   Update, Last Status Change, newest internal comment/email). Grouped + per CS Owner.
2. **Gap between consecutive internal updates** — *open + closed.* Average interval between
   updates during a case's life. Grouped + per CS Owner.
3. **Response time to a Partner User comment** — *open + closed.* Days from an installer
   (Partner User) comment to the next internal update. Unanswered inputs are counted
   separately, not averaged. Grouped + per CS Owner.
4. **Response time to an incoming customer email** — *open + closed.* Days from an incoming
   email to the next internal update. Same unanswered handling. Grouped + per CS Owner.

**Group B — Backlog & data-quality flags** (open cases; "older than N days" = case age from
Created Date):

5. **CS Queue cases older than 20 days**, excluding any case ever escalated to L2. The number
   removed for that reason is flagged.
6. **Cases older than 1 day without an Account.**
7. **Cases older than 1 day missing a Contact and/or CS Owner.**

## Run locally

```
pip install -r requirements.txt
streamlit run app.py
```

## Upload format

One `.xlsx` workbook with four sheets: `Case`, `Case Comment`, `Email Message`,
`Case History` (XL-Connector export). Column-name variants are normalized in `utils.py`.

## Configuration caveat

Queue detection compares the raw Case `Owner` field against the exact strings in `metrics.py`:

- `CS_QUEUE_NAME = "US Customer Support Queue"` — **unverified**; confirm against a real export.
- `L2_QUEUE_NAME = "US Level 2 Customer Support"`

If your org's queue names differ even slightly, metric 5 reads zero until these constants are
corrected.

## Structure

- `app.py` — Streamlit layout (2 tabs), upload, and download button
- `metrics.py` — all pandas metric logic
- `excel_export.py` — styled openpyxl workbook builder
- `utils.py` — export column-name normalization (reused from cs-reports, validated 2026-07-06)
