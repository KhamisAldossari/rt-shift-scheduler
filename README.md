# RT Shift Scheduler

Generate a safe, fair monthly Respiratory Therapy (RT) roster for **any month**, in
the browser, with **every employee working exactly 16 shifts** — and a downloadable,
color-coded Excel file at the end.

The engine models every hard rule as an explicit [CP-SAT](https://developers.google.com/optimization/cp/cp_solver)
constraint (Google OR-Tools), so it either returns a schedule that satisfies **all**
rules at once or proves the rule set is infeasible. A separate validator then
re-derives every rule from the finished grid — it never trusts the solver — and
reports each as **PASS/FAIL** with the measured value.

---

## Features

- **Any month / year** — day count, start weekday, and Fri/Sat weekends are computed
  from your chosen month.
- **Exactly 16 shifts per employee** (configurable, but exact — not a range).
- **Fixed 2-person night team** that works nights only (no day shifts, no circadian
  flipping).
- **Smart night coverage** — at least 1 night every day, with a 2nd overlapping night
  only on the days the math forces it (never a free-for-all).
- **Fatigue rules** — no Night→Day next-day flip, day runs ≤ 4, night runs ≤ 3,
  off runs 2–4, work runs ≥ 2.
- **Fairness** — even night split, even Fri/Sat duty per team, and at least one full
  Fri+Sat weekend off for everyone.
- **Independent validator** — every rule re-checked from the grid and shown PASS/FAIL.
- **Excel output** — a `Schedule` grid sheet + a `Summary`/validation sheet, color-coded.
- **Browser UI** (Streamlit) — configure everything; no code editing required.

---

## Requirements

- Python 3.9+
- Packages in [`requirements.txt`](requirements.txt): `ortools`, `openpyxl`, `streamlit`, `pandas`

```bash
.venv/bin/pip install -r requirements.txt
```

---

## Quick start

### Web app (recommended)

```bash
.venv/bin/streamlit run app.py
```

Your browser opens to the app. Set the month, staff, night team, and rule values in
the sidebar, click **Generate roster**, review the roster and the PASS/FAIL checks,
then **Download Excel (.xlsx)**.

### Command line

```bash
.venv/bin/python scheduler.py
```

Builds the default roster (August 2026, 7 staff), prints the per-employee counts and
the rule validation, and writes `schedule.xlsx`.

---

## How it works

```
ScheduleSettings ──▶ preflight() ──▶ build_and_solve() ──▶ validate() ──▶ Excel
   (one config        (cheap arith     (CP-SAT model,        (re-derives     (Schedule +
    object the UI      checks +         every hard rule       every rule       Summary
    populates)         suggestions)     as a constraint)      independently)   sheets)
```

1. **`ScheduleSettings`** — a single dataclass holds the month, the staff, the night
   team, the staffing bands, and every rule value. The CLI and the UI both build one;
   nothing is hard-coded.
2. **`preflight()`** — cheap arithmetic checks that catch an impossible rule set before
   the solver runs, each paired with a concrete suggestion (e.g. *"raise max nights/day
   to 2"*).
3. **`build_and_solve()`** — the CP-SAT model. Hard rules are constraints; fairness goals
   are minimized as a weighted objective. Deterministic (fixed seed → same roster each run).
4. **`validate()`** — re-derives every rule straight from the grid, independent of the
   solver, and returns PASS/FAIL + the measured value.
5. **`export()` / `export_bytes()`** — the color-coded workbook.

---

## Configuration (`ScheduleSettings`)

| Field | Default | Meaning |
|---|---|---|
| `year`, `month` | 2026, 8 | Calendar month to schedule |
| `employees` | Employee 1–7 | The staff list (any size) |
| `night_team` | Employee 6, 7 | The 2 people who work nights |
| `night_team_nights_only` | `True` | Night team works nights only |
| `day_min`, `day_max` | 2, 4 | Day staff required per day |
| `night_min`, `night_max` | 1, 2 | Night coverage band per day |
| `shifts_per_employee` | 16 | **Exact** monthly total per person |
| `max_consec_work` | 4 | Max consecutive day-shift run |
| `max_consec_night` | 3 | Max consecutive night run |
| `min_consec_work` | 2 | Min consecutive work (must be 2) |
| `min_consec_off` | 2 | Min consecutive off (must be 2) |
| `max_consec_off` | 4 | Max consecutive off run |
| `fairness_max_gap` | 1 | Allowed max−min gap for "even" rules |
| `weekend_days` | Fri, Sat | The weekend (Saudi convention) |
| `w_*` weights | — | Objective weights for the soft goals |

> **Note on exactly-16 + the night team.** Two nights-only members each working exactly
> 16 means the month always has **32 night-shifts**. Spread over an N-day month at ≥ 1/day,
> that forces exactly **(32 − N)** double-night days — `1` for a 31-day month, `2` for 30,
> `3` for 29, `4` for 28 (February). This is the minimum overlap, validated as
> *"2nd night only where needed."*

---

## Rules

**Hard** (modeled as CP-SAT constraints, can never be violated):
exactly-16 shifts · day staffing within `[day_min, day_max]` · night coverage within
`[night_min, night_max]` (never uncovered) · nights worked only by the night team ·
night team works nights only · no Night→Day next day · day runs ≤ `max_consec_work` ·
night runs ≤ `max_consec_night` · work runs ≥ 2 · off runs ≥ 2 and ≤ `max_consec_off`.

**Best-effort** (solver objectives, surfaced as PASS/FAIL by the validator):
even night split across the night team · even Fri/Sat duty within each team · at least
one full Fri+Sat weekend off for everyone. These PASS for typical months; for an
unusually tight configuration the validator will show them as FAIL rather than silently
dropping them.

---

## Output format

`schedule.xlsx` has two sheets:

- **Schedule** — rows = employees, columns = days (`D1 Sun … DN`), each cell is
  `D` (day), `N` (night), or `OFF`, color-coded; trailing Total / Day / Night / Fri-Sat
  columns. Weekends (Fri/Sat) are highlighted in the header.
- **Summary** — per-employee counts, the night-team split + double-night-day count, and
  the full rule-validation table (PASS/FAIL + measured).

---

## Files

| File | Purpose |
|---|---|
| `scheduler.py` | The engine: settings, preflight, CP-SAT model, validator, Excel export |
| `app.py` | Streamlit web UI (reuses the engine; no solver logic of its own) |
| `requirements.txt` | Python dependencies |
| `schedule.xlsx` | Most recent generated roster |
| `GUIDE-FOR-USE.md` | Plain-language guide for non-technical users |
| `CLAUDE.md` | Notes for AI assistants working on this repo |

---

## Infeasible inputs

If a rule combination can't produce a valid schedule, the app **explains the conflict
and suggests the nearest fix** (e.g. widen the day band, raise max consecutive off,
allow more night overlap) — it never silently relaxes a hard rule to force a result.
