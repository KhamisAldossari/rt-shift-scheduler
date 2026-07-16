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
- **Two rotation modes (`rotation_mode`).** *fixed_team* — a night team of **any size N**
  works nights only (`work == night` for night-team indices) and is the **only** group
  eligible for nights. *rotate* — no team; **every** employee is night-eligible and nights
  rotate across the roster, balanced softly (no hard per-person night cap). Role membership
  is resolved per mode by `night_eligible_indices()` / `day_capable_indices()`; never
  hard-code who works nights. **The web UI deliberately exposes only `fixed_team`** —
  `rotate` stays available engine/CLI-only (`rotation_mode=sch.ROTATE`). Likewise the
  fairness tolerances (`fair_tol_*`) and objective weights (`w_fair_*`) are deliberately
  not in the UI. Both are hidden by product decision — do not re-add either without being
  asked.
- **Forced night overlap (fixed_team).** N nights-only members on exactly-16 ⇒ a fixed
  **16·N night-shifts/month** ⇒ forced overlap above the floor =
  `max(0, 16·N − night_min·days)`. The validator checks this as *"Night overlap is minimal
  (only forced extra nights)"* via `expected_double_nights()` (it returns the forced
  surplus; **0** in rotate mode, where the night total isn't pinned). If you change
  `shifts_per_employee`, the night-team size, or `night_min`, that helper **and** the
  `night_min/night_max` band must stay consistent or the model goes infeasible.
- **Night coverage is a band `[night_min, night_max]`**, never below `night_min` (a night
  is never uncovered). Default `[1, 2]`.
- **Fatigue caps:** combined working runs (day or night, back to back) ≤ `max_consec_work`
  (4), night runs ≤ `max_consec_night` (3), **no Night→Day next day**. These are hard.
  Don't relax to "make it fit."
- **The min-run encoding only supports a minimum of 2.** `build_and_solve()` has an
  `assert min_consec_work == 2 and min_consec_off == 2`. If you ever need a different
  minimum, generalize the neighbour-encoding constraints first; `preflight()` also guards
  this with a `Problem`.
- **Determinism:** single worker + fixed `random_seed=42` + a **deterministic** time limit
  (`max_deterministic_time = solver_det_time_limit`). The fairness objective often can't be
  *proven* optimal in time, and a wall-clock cap would make the roster machine-dependent —
  the deterministic limit makes the stop point (and the roster) reproducible on any machine,
  even when the solve ends at `FEASIBLE`. `solver_time_limit` is only a wall-clock safety
  net, sized so it does **not** bind before the deterministic budget. Don't revert to a
  wall-clock-only limit.
- **Hard vs soft:** the **four fairness goals** are **soft** (equal-weighted objective terms
  `w_fair_*`), each surfaced by the validator as PASS/FAIL **plus the measured spread**
  against a tunable tolerance (`fair_tol_*`): (1) equal totals, (2) balanced night load,
  (3) balanced weekends — even Fri/Sat duty + a full weekend off each, (4) balanced
  undesirable runs — overlap nights + max-length streaks. Goal 4 is balanced **within each
  pool** (day-capable vs night-eligible): a fixed night team structurally absorbs all
  overlap, so a cross-pool comparison isn't a fairness signal. Exactly-16, coverage bands,
  the run caps, no-N→D, and nights-only eligibility are **hard**. Don't promote a soft goal
  to hard without checking it keeps common months feasible. Some tight months can *provably*
  soft-fail a goal (e.g. N=3 nights-only February can't give all three a full weekend off) —
  that's honestly reported; every hard rule still holds.
- **Alternating weekends (`alternating_weekends`) is a HARD toggle, not a soft goal.** It
  **defaults OFF (opt-in)** so the out-of-box roster stays bit-identical to pristine main (both
  the UI toggle and the CLI read this default). When on, every employee's full (Fri, Sat)
  weekends strictly alternate off/on: of every two consecutive full weekends from
  `weekend_pair_indices()`, exactly one is fully off (OFF on both days); the solver picks each
  person's phase. Encoded in `build_and_solve()`
  inside the goal-3 per-employee loop, reusing the same reified `offwknd_*` bits
  (`model.Add(bits[w] + bits[w+1] == 1)`); re-derived independently in `validate()` as the
  hard rule *"Weekends alternate off / on for every employee"* — same pair list, same
  off-both-days test, same gating, so constraint and check score one quantity. **OFF-mode
  bit-identity:** when the toggle is False, `build_and_solve()` adds no variable and no
  constraint, so the model is byte-identical to the pre-feature engine (grid hashes diff
  IDENTICAL against pristine main). `preflight()` adds a *necessary-not-sufficient* arithmetic
  guard (worst-case weekend availability is `floor(pool/2)` per pool — disjoint pools for a
  nights-only team, one shared floor otherwise). A tight config can be *provably INFEASIBLE*
  with the toggle on even when the preflight guard passes (verified: the **default** 7-staff /
  2-night-team config in February **2027** — `floor(5/2)=2 ≥ day_min` and `floor(2/2)=1 ≥
  night_min` both hold, yet the joint packing is infeasible; the same config in February 2026
  solves, so it depends on the weekday layout, not just month length) — the solver honestly
  returns INFEASIBLE and `relaxation_hints()` surfaces "turn off alternating weekends". Turning the
  toggle on can also push the **F3 "balanced weekends" soft goal to an honest FAIL** in some
  month/pool shapes: strict alternation pins *which* weekends each person is off, so a small
  night pool over an odd number of full weekends must split phases, forcing a weekend spread
  above `fair_tol_weekend` (e.g. a 2-person night pool over 5 weekends → spread 2 > tol 1).
  Every HARD rule still holds, so that is a contract-compliant soft FAIL, **not** the
  OPTIMAL-but-FAIL landmine. Keep alt-on configs out of any all-PASS smoke assertion the way
  the INFEASIBLE and tight soft-fail configs are kept out.
- **Leave (`ScheduleSettings.leave`) hard-blocks days and prorates targets.** Ranges are
  `(name, from, to)` inclusive date tuples; `leave_day_sets()` / `shift_targets()` /
  `available_segments()` are the single derivation all four stages share. Leave days are
  hard-blocked (`work == 0`) and shown as `V` (`LEAVE`); **grid extraction emits `LEAVE`
  only inside the `work == 0` branch** — a regressed hard-block surfaces as D/N and the
  *"Leave days are honored"* check FAILs loudly. Never label from settings alone: that
  would make the check score the labeling path instead of the constraint. **All five run
  rules apply per available segment** (the free stretches between leave blocks): max-off
  windows are enumerated per segment, min-work/min-off re-encode with segment edges playing
  the month-edge role, and a **singleton segment is forced OFF** (`work == 0`) — the
  validator mirrors this by exempting a whole-segment off-run shorter than `min_consec_off`
  and skipping target-0 employees (leave ≥ target), who are also exempt from the
  alternating-weekends chain **and** check; chain pairs touching a fully-leave weekend are
  skipped identically on both sides. **Night arithmetic uses summed targets:**
  `expected_double_nights()` and the preflight capacity checks sum `shift_targets()` over
  the team (identical to `16·N` with no leave), so with the default 2-member team **team
  leave-days beyond `32 − days` trip the night floor** (1 day of slack in a 31-day month,
  2 in a 30-day month, 4 in February) — never "any night-team leave is infeasible": short
  team leave inside the slack is feasible and must not be blocked. Leave that passes the
  floor can still be provably infeasible on run caps; the **pinned-crew preflight guard**
  catches that (a pool whose availability sits AT its floor for longer than its run cap
  forces someone over the cap — e.g. one member of the default team on 4 consecutive
  February leave days). **Weighted fairness is gated on leave:** with leave configured,
  G2/G3 switch to division-free target-weighted deviations (`T·x_e − tgt_e·X` per pool,
  via the **signed** gap helper — `minmax_gap`'s `[0, ub]` bounds go infeasible on negative
  deviations) and the raw G1/G4/no-full-weekend terms are scaled by the matching pool
  target sum so the four goals stay commensurate; `validate()` scores the same weighted
  quantity per pool (`spread ≤ fair_tol_* × T_pool`, F3 tested per pool since the two
  pools' T differ). With `leave=[]` every branch takes the legacy path and the model/grid
  are bit-identical to pristine main — verify by grid diff, not by eye. AM staff cannot
  take leave in v1 (preflight rejects the name) and their Hours cell stays blank.
- **OFF renders blank; nothing may borrow its old red.** `CELL_LABEL[OFF] = ""`,
  `CELL_FILL[OFF] = None`, `WEB_COLORS[OFF] = "#FFFFFF"`; the FFC7CE red now belongs to
  `LEAVE` (`FILL_LEAVE`). Anything styling FAIL/error states must point at the leave red
  (app.py's `color_result` does), never `WEB_COLORS[OFF]`. The Excel stat columns
  (Total/Day/Night/Fri-Sat/Hours) are live COUNTIF formulas over the day cells — they
  recount hand edits, while the Summary sheet and app validation reflect only the generated
  roster; `hours_per_shift` (default 12) lives in `ScheduleSettings`, never hard-coded.
- **Two places to update when adding a rule:** the constraint in `build_and_solve()` **and**
  an independent check in `validate()`. A rule that isn't in the validator isn't trusted.
  For a **fairness** goal, the objective term and the validator check must score the **same
  quantity** — otherwise the solver can report OPTIMAL while the validator reports FAIL.

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
    sch.ScheduleSettings(year=2027, month=3,                        # fixed night team of 3
        employees=[f"Employee {i}" for i in range(1, 10)],
        night_team=["Employee 7", "Employee 8", "Employee 9"], night_max=2),
    sch.ScheduleSettings(year=2026, month=9, rotation_mode=sch.ROTATE),  # everyone rotates
]:
    sol = sch.build_and_solve(S)
    res = sch.validate(S, sol.grid)
    assert sol.grid and all(r.passed for r in res), (S.month_label, [r.name for r in res if not r.passed])
```

The four fixed-team cases solve `OPTIMAL`; the rotate case may end at a **reproducible**
`FEASIBLE` within the deterministic budget — both with every rule `PASS`. (Determinism note:
a tight month + a larger night team can *provably* soft-fail the weekend goal, e.g. N=3
nights-only in February — that's a contract-compliant soft FAIL, not a broken smoke test, so
keep such configs out of the all-PASS assertion.) When you touch fairness goal 4 or the
solver limits, re-run an **all-12-months** sweep for the default, an N≥3 team, and rotate —
not just the months above — and confirm grids are identical across re-solves.

## Conventions

- Match the existing style: module-level config moved into the dataclass, helper functions
  take `settings`, clear section banners, comment density as in `scheduler.py`.
- Excel and the Streamlit grid share one color map (`WEB_COLORS` mirrors the openpyxl fills)
  so the two views look identical. Keep them in sync.
- Streamlit: prefer `width="stretch"` (not the deprecated `use_container_width`) and
  `Styler.map` (not `applymap`). Requires `pandas >= 2.1`.

## Dependencies

`ortools`, `openpyxl`, `streamlit`, `pandas` (see `requirements.txt`). Python 3.9+.
