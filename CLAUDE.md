# CLAUDE.md — working notes for this repo

Guidance for AI assistants (and humans) editing this project. Read before changing the
scheduling logic.

## What this is

A monthly Respiratory Therapy (RT) shift scheduler. A CP-SAT model (Google OR-Tools)
produces a roster; an **independent** validator re-derives every rule from the finished
grid; results export to a color-coded Excel workbook. A Streamlit app (`app.py`) is a
thin front-end over the engine.

## Architecture (do not break this shape)

- **`scheduler.py` is the single source of truth.** All scheduling logic lives here.
- **`app.py` must reuse the engine** — it imports `scheduler` and calls
  `preflight → build_and_solve → validate → export_bytes`. It must **never** re-implement
  a constraint or recompute a rule. UI code only collects inputs and renders outputs.
- **All configuration flows through one `ScheduleSettings` dataclass.** Nothing about the
  month, the staff, the night team, or rule values may be hard-coded in the model, the
  validator, the export, or the UI. If you add a tunable, add it to `ScheduleSettings`
  and thread it through.

## The four stages

1. `ScheduleSettings` — config + derived calendar (`days`, `start_weekday`, `month_label`,
   `day_team`). Calendar facts come from `calendar`/`datetime` for the chosen month+year.
2. `preflight(settings) -> list[Problem]` — cheap arithmetic feasibility checks. Each
   `Problem` carries a `message` **and** a `suggestion` (the UI shows both).
3. `build_and_solve(settings) -> Solution` — the CP-SAT model. Deterministic.
4. `validate(settings, grid) -> list[RuleResult]` — re-derives every rule from the grid,
   independent of the solver. **Keep it independent** — never read solver internals here.

## Invariants / landmines

- **Exactly-16 is hard** (`sum(work) == shifts_per_employee`). Do not turn it back into a
  range without being asked.
- **Night team works nights only** (`work == night` for night-team indices) and is the
  **only** group eligible for nights. Two nights-only members on exactly-16 ⇒ a fixed
  **32 night-shifts/month** ⇒ exactly **(32 − days)** forced double-night days. The
  validator checks this as *"2nd night only where needed"* via `expected_double_nights()`.
  If you change `shifts_per_employee` or the night-team size, that helper and the
  `night_min/night_max` band must stay consistent or the model goes infeasible.
- **Night coverage is a band `[night_min, night_max]`**, never below `night_min` (a night
  is never uncovered). Default `[1, 2]`.
- **Fatigue caps:** day runs ≤ `max_consec_work` (4), night runs ≤ `max_consec_night` (3),
  **no Night→Day next day**. These are hard. Don't relax to "make it fit."
- **The min-run encoding only supports a minimum of 2.** `build_and_solve()` has an
  `assert min_consec_work == 2 and min_consec_off == 2`. If you ever need a different
  minimum, generalize the neighbour-encoding constraints first; `preflight()` also guards
  this with a `Problem`.
- **Determinism:** single worker + fixed `random_seed=42`. Keep it so the same inputs give
  the same roster (resume/reproducibility).
- **Hard vs soft:** even night split, even Fri/Sat spread, and full-weekend-off-for-all are
  **soft** (objective weights `w_*`), surfaced as PASS/FAIL by the validator. Exactly-16,
  coverage bands, the run caps, no-N→D, and nights-only are **hard**. Don't promote a soft
  goal to hard without checking it doesn't make common months infeasible.
- **Two places to update when adding a rule:** the constraint in `build_and_solve()` **and**
  an independent check in `validate()`. A rule that isn't in the validator isn't trusted.

## Weekend convention

"Weekend" = **Friday + Saturday** (Saudi). `weekend_days = ("Fri", "Sat")`. `D1` weekday is
derived from the real calendar, so don't assume the month starts on Sunday.

## Run / test

```bash
.venv/bin/python scheduler.py            # CLI: default Aug 2026 -> schedule.xlsx
.venv/bin/streamlit run app.py           # web UI
```

Smoke test across month lengths and a config change before claiming done:

```python
import scheduler as sch
for S in [
    sch.ScheduleSettings(year=2026, month=9),                       # 30 days
    sch.ScheduleSettings(year=2027, month=2),                       # 28 days (tightest)
    sch.ScheduleSettings(year=2028, month=2),                       # 29 days (leap)
    sch.ScheduleSettings(year=2027, month=1,                        # 8 staff, alt night pair
        employees=[f"Employee {i}" for i in range(1, 9)],
        night_team=["Employee 1", "Employee 2"]),
]:
    sol = sch.build_and_solve(S)
    res = sch.validate(S, sol.grid)
    assert sol.grid and all(r.passed for r in res), (S.month_label, [r.name for r in res if not r.passed])
```

All of the above should solve `OPTIMAL` with every rule `PASS`.

## Conventions

- Match the existing style: module-level config moved into the dataclass, helper functions
  take `settings`, clear section banners, comment density as in `scheduler.py`.
- Excel and the Streamlit grid share one color map (`WEB_COLORS` mirrors the openpyxl fills)
  so the two views look identical. Keep them in sync.
- Streamlit: prefer `width="stretch"` (not the deprecated `use_container_width`) and
  `Styler.map` (not `applymap`). Requires `pandas >= 2.1`.

## Dependencies

`ortools`, `openpyxl`, `streamlit`, `pandas` (see `requirements.txt`). Python 3.9+.
