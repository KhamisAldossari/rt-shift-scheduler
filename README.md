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
- **Two night-coverage models** — a **fixed night team of any size** that works nights only
  (no day shifts, no circadian flipping), **or** a **full rotation** where every employee is
  night-eligible and nights are spread across the whole roster.
- **Smart night coverage** — at least 1 night every day, with overlap above the floor only
  on the days the math forces it (never a free-for-all).
- **Fatigue rules** — no Night→Day next-day flip, day runs ≤ 4, night runs ≤ 3,
  off runs 2–4, work runs ≥ 2.
- **Fairness, four ways** — equal totals, balanced night load, balanced weekends (even
  Fri/Sat duty + a full weekend off each), and balanced "undesirable runs" (overlaps +
  max-length streaks) spread so no one absorbs the rough patterns. Each is reported
  PASS/FAIL **with its measured spread**, and all are tunable.
- **Independent validator** — every rule re-checked from the grid and shown PASS/FAIL.
- **Excel output** — a `Schedule` grid sheet + a `Summary`/validation sheet, color-coded.
- **Browser UI** (Streamlit) — configure everything; no code editing required.

---

## Requirements

- Python 3.9+
- Packages in [`requirements.txt`](requirements.txt): `ortools`, `openpyxl`, `streamlit`, `pandas`

One-time setup — create the virtual environment and install the packages:

```bash
# macOS / Linux
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

```powershell
# Windows (PowerShell)
py -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

---

## Quick start

### Web app (recommended)

```bash
# macOS / Linux
.venv/bin/streamlit run app.py
```

```powershell
# Windows (PowerShell)
.venv\Scripts\streamlit run app.py
```

Your browser opens to the app. Set the month, staff, night model, rules, and fairness in
the sidebar, click **Generate roster**, review the roster, the per-employee fairness
summary, and the PASS/FAIL checks, then **Download Excel (.xlsx)**.

### Command line

```bash
# macOS / Linux
.venv/bin/python scheduler.py
```

```powershell
# Windows (PowerShell)
.venv\Scripts\python scheduler.py
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
3. **`build_and_solve()`** — the CP-SAT model. Hard rules are constraints; the four fairness
   goals are minimized as an equal-weighted objective. Deterministic — fixed seed + a
   deterministic search budget → the same roster on any machine.
4. **`validate()`** — re-derives every rule straight from the grid, independent of the
   solver, and returns PASS/FAIL + the measured value.
5. **`export()` / `export_bytes()`** — the color-coded workbook.

---

## Configuration (`ScheduleSettings`)

| Field | Default | Meaning |
|---|---|---|
| `year`, `month` | 2026, 8 | Calendar month to schedule |
| `employees` | Employee 1–7 | The staff list (any size) |
| `rotation_mode` | `"fixed_team"` | `"fixed_team"` (dedicated team) or `"rotate"` (everyone) |
| `night_team` | Employee 6, 7 | Who works nights — **any size** (fixed_team mode) |
| `night_team_nights_only` | `True` | Night team works nights only |
| `day_min`, `day_max` | 2, 4 | Day staff required per day |
| `night_min`, `night_max` | 1, 2 | Night coverage band per day |
| `shifts_per_employee` | 16 | **Exact** monthly total per person |
| `max_consec_work` | 4 | Max consecutive day-shift run |
| `max_consec_night` | 3 | Max consecutive night run |
| `min_consec_work` | 2 | Min consecutive work (must be 2) |
| `min_consec_off` | 2 | Min consecutive off (must be 2) |
| `max_consec_off` | 4 | Max consecutive off run |
| `weekend_days` | Fri, Sat | The weekend (Saudi convention) |
| `alternating_weekends` | `False` | **Hard** toggle (opt-in): full Fri+Sat weekends alternate off/on for every employee |
| `fair_tol_total/night/weekend/runs` | 0 / 1 / 1 / 2 | Pass tolerance (max−min ≤) per fairness goal |
| `w_fair_total/night/weekend/runs` | 100 each | Equal objective weights for the four goals |
| `solver_det_time_limit` | 24 | Deterministic search budget (reproducible) |

> **Note on exactly-16 + a fixed night team.** A nights-only team of **N** members each
> working exactly 16 means the month always has **16·N night-shifts**. Spread over a
> `days`-day month at ≥ `night_min`/day, that forces exactly
> **max(0, 16·N − night_min·days)** extra "overlap" nights stacked onto some days — the
> unavoidable minimum, validated as *"Night overlap is minimal."* (For the default 2-person
> team at `night_min = 1` this is the classic `(32 − days)` double-night days.) In **rotate**
> mode the night total isn't pinned, so there's no forced overlap.

---

## Rules

**Hard** (modeled as CP-SAT constraints, can never be violated):
exactly-16 shifts · day staffing within `[day_min, day_max]` · night coverage within
`[night_min, night_max]` (never uncovered) · nights worked only by the night team, and the
night team works nights only *(fixed-team mode)* · no Night→Day next day · day runs ≤
`max_consec_work` · night runs ≤ `max_consec_night` · work runs ≥ 2 · off runs ≥ 2 and ≤
`max_consec_off` · alternating weekends *(when `alternating_weekends` is on — every
employee's full Fri+Sat weekends alternate off/on, one fully off of each consecutive pair)*.

**Best-effort fairness** (four equal-weighted soft objectives, surfaced as PASS/FAIL **plus
the measured spread** by the validator): equal totals · balanced night load across
night-eligible staff · balanced weekends (even Fri/Sat duty within each pool + a full
weekend off each) · balanced undesirable runs (overlaps + max-length streaks, balanced
within each pool). These PASS for typical months; for an unusually tight configuration
(e.g. a 3-person night team in a 28-day February) the validator shows a goal as FAIL rather
than silently dropping it — the roster stays valid because every hard rule still holds.

---

## Output format

`schedule.xlsx` has two sheets:

- **Schedule** — rows = employees, columns = days (`D1 Sun … DN`), each cell is
  `D` (day), `N` (night), or `OFF`, color-coded; trailing Total / Day / Night / Fri-Sat
  columns. Weekends (Fri/Sat) are highlighted in the header.
- **Summary** — a per-employee fairness summary (Total / Day / Night / Fri-Sat / Rough), the
  night split + overlap-night-day count, and the full rule-validation table (PASS/FAIL +
  measured).

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
