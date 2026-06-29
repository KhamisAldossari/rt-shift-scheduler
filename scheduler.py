#!/usr/bin/env python3
"""
Monthly staff shift scheduler  (configurable engine).

Generates a valid roster for ANY month using a CP-SAT constraint solver
(Google OR-Tools), then independently validates every scheduling rule and
exports a color-coded Excel workbook (roster sheet + summary/validation sheet).

Why CP-SAT: every hard rule is modelled as an explicit constraint, so the
solver either returns a schedule that satisfies *all* of them simultaneously
or proves the rule set is infeasible. We never relax a hard rule to "make it
fit". After solving, a separate validator re-derives every rule from the
finished grid (it does not trust the solver) and reports PASS/FAIL.

All configuration lives in one `ScheduleSettings` object so the same engine
backs both the CLI (`python scheduler.py`) and the Streamlit app (`app.py`):
nothing about the month, the staff, or the rule values is hard-coded.

POLICY (this revision):
  * Every employee works EXACTLY `shifts_per_employee` shifts for the month.
  * Nights are covered under one of two selectable rotation modes:
      - "fixed_team": a dedicated night team of ANY size N works nights ONLY and
        is the only group eligible for nights. Coverage is the band
        [night_min, night_max]; overlap above the floor appears only on the days
        the exact arithmetic forces it, never as a free-for-all.
      - "rotate": there is no dedicated team -- every employee is night-eligible
        and nights rotate across the whole roster. The night load is balanced as
        evenly as possible (a soft fairness goal), not pinned to a fixed total.
  * No Night->Day next-calendar-day transition for anyone.
  * Day-shift runs are capped at `max_consec_work`; night runs at
    `max_consec_night`; off runs floored at `min_consec_off`, capped at
    `max_consec_off`; work runs floored at `min_consec_work`.
  * FAIRNESS is the top priority and is expressed as four soft, equally-weighted
    objective goals -- equal totals, balanced night load, balanced weekends, and
    balanced "undesirable runs" (forced overlaps + max-length streaks) -- each of
    which the independent validator re-derives and reports as PASS/FAIL plus the
    measured spread (max - min). Fairness goals are SOFT: they never block a
    solve, and never override a hard rule.

Run:  ./.venv/bin/python scheduler.py
Out:  schedule.xlsx  (+ a summary printed to stdout)
"""

from __future__ import annotations

import calendar
import datetime
import sys
from dataclasses import dataclass, field
from io import BytesIO

from ortools.sat.python import cp_model
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# Shift codes used throughout.
DAY, NIGHT, OFF = "D", "N", "O"

# Rotation modes (how nights are covered).
FIXED_TEAM, ROTATE = "fixed_team", "rotate"

# Calendar weekday names, indexed Sun..Sat to match the roster header style.
WEEKDAY_NAMES = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]


# ---------------------------------------------------------------------------
# SETTINGS  (the single source of truth the UI populates)
# ---------------------------------------------------------------------------

@dataclass
class ScheduleSettings:
    # --- Calendar -----------------------------------------------------------
    year: int = 2026
    month: int = 8                                  # 1-12

    # --- Staff --------------------------------------------------------------
    employees: list[str] = field(
        default_factory=lambda: [f"Employee {i}" for i in range(1, 8)])

    # --- Night model --------------------------------------------------------
    # "fixed_team": night_team (ANY size) works nights only and is the only
    #               group eligible for nights.
    # "rotate":     no dedicated team -- everyone is night-eligible.
    rotation_mode: str = FIXED_TEAM
    night_team: list[str] = field(
        default_factory=lambda: ["Employee 6", "Employee 7"])
    night_team_nights_only: bool = True             # fixed_team: no day shifts

    # --- Staffing bands -----------------------------------------------------
    day_min: int = 2                                # day staff required per day
    day_max: int = 4
    night_min: int = 1                              # >= 1 night every day
    night_max: int = 2                              # allow up to 2 overlapping

    # --- Workload -----------------------------------------------------------
    shifts_per_employee: int = 16                   # EXACT monthly total

    # --- Run lengths --------------------------------------------------------
    max_consec_work: int = 4                        # day-shift / working run cap
    max_consec_night: int = 3                       # night run cap (priority)
    min_consec_work: int = 2                        # no isolated single work day
    min_consec_off: int = 2                         # no isolated single off day
    max_consec_off: int = 4                         # no long idle block

    # --- Fairness -----------------------------------------------------------
    weekend_days: tuple = ("Fri", "Sat")            # the "weekend" for fairness

    # Tolerances: a fairness goal PASSES when its measured spread (max - min)
    # is within tolerance. Loose by default so common months stay green; tighten
    # from the UI to demand stricter fairness. Totals are pinned exact (spread 0).
    fair_tol_total: int = 0
    fair_tol_night: int = 1
    fair_tol_weekend: int = 1
    fair_tol_runs: int = 2

    # --- Objective weights (all minimized; equal => the four goals matter
    #     equally). Raising one weight prioritises that goal in trade-offs. -----
    w_fair_total: int = 100                         # equal total shifts
    w_fair_night: int = 100                         # balanced night load
    w_fair_weekend: int = 100                       # balanced weekends
    w_fair_runs: int = 100                          # balanced undesirable runs

    # Determinism: the solver is bounded by DETERMINISTIC time (reproducible
    # work units, machine-independent) so the same inputs always give the same
    # roster -- even when a hard fairness objective is not proven optimal in
    # time. Strict reproducibility relies on THIS limit binding first; the
    # wall-clock `solver_time_limit` is only a safety net, sized so it does not
    # bind before the deterministic budget on normal hardware. Fixed-team months
    # solve in well under a second; the larger budget is headroom for the harder
    # rotate mode, whose soft weekend goal needs more search to converge.
    solver_det_time_limit: float = 24.0
    solver_time_limit: float = 180.0

    # --- Derived calendar facts --------------------------------------------
    @property
    def days(self) -> int:
        return calendar.monthrange(self.year, self.month)[1]

    @property
    def start_weekday(self) -> str:
        wd = datetime.date(self.year, self.month, 1).weekday()   # Mon=0..Sun=6
        return WEEKDAY_NAMES[(wd + 1) % 7]

    @property
    def month_label(self) -> str:
        return f"{calendar.month_name[self.month]} {self.year}"

    @property
    def is_rotate(self) -> bool:
        return self.rotation_mode == ROTATE

    @property
    def day_team(self) -> list[str]:
        """Employees who may work day shifts (everyone, in rotate mode)."""
        if self.is_rotate or not self.night_team_nights_only:
            return list(self.employees)
        return [e for e in self.employees if e not in self.night_team]


# ---------------------------------------------------------------------------
# CALENDAR / ROLE HELPERS
# ---------------------------------------------------------------------------

def weekday_of(settings: ScheduleSettings, day_index: int) -> str:
    start = WEEKDAY_NAMES.index(settings.start_weekday)
    return WEEKDAY_NAMES[(start + day_index) % 7]


def weekend_day_indices(settings: ScheduleSettings) -> list[int]:
    return [d for d in range(settings.days)
            if weekday_of(settings, d) in settings.weekend_days]


def weekend_pair_indices(settings: ScheduleSettings) -> list[tuple[int, int]]:
    """(Fri, Sat) consecutive index pairs that make a full weekend."""
    return [(d, d + 1) for d in range(settings.days - 1)
            if weekday_of(settings, d) == "Fri" and weekday_of(settings, d + 1) == "Sat"]


def night_eligible_indices(settings: ScheduleSettings) -> list[int]:
    """Employee indices allowed to work nights (everyone, in rotate mode)."""
    if settings.is_rotate:
        return list(range(len(settings.employees)))
    return [settings.employees.index(name) for name in settings.night_team
            if name in settings.employees]


def day_capable_indices(settings: ScheduleSettings) -> list[int]:
    """Employee indices allowed to work day shifts.

    In fixed_team + nights-only mode the night team never works days; otherwise
    (rotate, or a fixed team that may also work days) everyone is day-capable.
    """
    n = len(settings.employees)
    if not settings.is_rotate and settings.night_team_nights_only:
        elig = set(night_eligible_indices(settings))
        return [i for i in range(n) if i not in elig]
    return list(range(n))


def is_night_member(settings: ScheduleSettings, e: int) -> bool:
    """True iff employee index `e` belongs to a dedicated nights-only team."""
    return (not settings.is_rotate
            and settings.employees[e] in settings.night_team)


def expected_double_nights(settings: ScheduleSettings) -> int:
    """Forced night-overlap (extra nights above the floor), given exact totals.

    A nights-only team of ANY size N, each working `shifts_per_employee`, places
    a fixed N * shifts night-shifts. With >= night_min covered every day, the
    surplus above the floor -- max(0, total_nights - night_min*days) -- must land
    as overlap (extra nights stacked onto some days). This is the minimum the
    arithmetic forces; nothing more. The validator checks the realised overlap
    equals it. For the default 2-person team on night_min=1 this is the classic
    "(32 - days) double-night days". In 'rotate' mode the night total is not
    pinned, so there is no forced overlap and this returns 0.
    """
    if settings.is_rotate or not settings.night_team_nights_only:
        return 0
    n_night = len(night_eligible_indices(settings))
    total_nights = settings.shifts_per_employee * n_night
    return max(0, total_nights - settings.night_min * settings.days)


def roster_caption(settings: ScheduleSettings) -> str:
    """One-line description of the roster, shared by Excel and the CLI."""
    S = settings
    head = f"Roster: {S.month_label} ({S.days} days)."
    if S.is_rotate:
        return f"{head} All staff rotate nights."
    names = ", ".join(S.night_team) if S.night_team else "(none)"
    only = " (nights only)" if S.night_team_nights_only else ""
    return f"{head} Night team: {names}{only}."


# ---------------------------------------------------------------------------
# FEASIBILITY PRE-FLIGHT  (catches a bad rule set before the solver, and -- per
# the UI choice -- pairs each problem with a concrete, non-binding suggestion)
# ---------------------------------------------------------------------------

@dataclass
class Problem:
    message: str
    suggestion: str


def _ceil(a: int, b: int) -> int:
    return -(-a // b)


def preflight(settings: ScheduleSettings) -> list[Problem]:
    S = settings
    problems: list[Problem] = []
    days = S.days
    n_emp = len(S.employees)
    w = S.shifts_per_employee

    # --- Basic shape (mode-independent) ------------------------------------
    if n_emp == 0:
        problems.append(Problem("No employees configured.",
                                "Add at least 3-4 employees."))
        return problems
    if S.min_consec_work != 2 or S.min_consec_off != 2:
        problems.append(Problem(
            "Minimum consecutive work/off must be 2 (the model encodes exactly 2).",
            "Keep 'min consecutive work' and 'min consecutive off' at 2."))
    if w > days:
        problems.append(Problem(
            f"{w} shifts/employee cannot fit in a {days}-day month.",
            f"Lower shifts/employee to at most {days}, or choose a longer month."))
        return problems
    if S.day_min + S.night_min > n_emp:
        problems.append(Problem(
            f"Each day needs at least {S.day_min} day + {S.night_min} night = "
            f"{S.day_min + S.night_min} people, but only {n_emp} employees exist.",
            f"Add staff, or lower the daily minimums to total <= {n_emp}."))

    # The run-length pattern check is shared by both modes.
    def pattern_problem(label, work_days, max_work, max_off):
        o = days - work_days
        bw_lo, bw_hi = _ceil(work_days, max_work), work_days // S.min_consec_work
        bo_lo = _ceil(o, max_off) if o > 0 else 0
        bo_hi = o // S.min_consec_off if o > 0 else 0
        if bw_lo > bw_hi:
            return Problem(
                f"{label}: {work_days} work days can't be split into blocks of "
                f"{S.min_consec_work}-{max_work}.",
                f"Raise the max {label.split()[0].lower()} run, or change shifts/employee.")
        if o > 0 and bo_lo > bo_hi:
            return Problem(
                f"{label}: {o} off days can't be split into blocks of "
                f"{S.min_consec_off}-{max_off}.",
                "Raise max consecutive off, or change shifts/employee.")
        if not any(abs(bw - bo) <= 1
                   for bw in range(bw_lo, bw_hi + 1)
                   for bo in range(bo_lo, bo_hi + 1)):
            return Problem(
                f"{label}: cannot alternate {work_days} work and {o} off days under "
                "the run-length rules.",
                "Widen a run-length cap, or change shifts/employee.")
        return None

    # --- Mode B: everyone rotates nights -----------------------------------
    if S.is_rotate:
        total_shifts = w * n_emp
        floor_need = (S.day_min + S.night_min) * days
        ceil_cap = (S.day_max + S.night_max) * days
        if total_shifts < floor_need:
            problems.append(Problem(
                f"{total_shifts} total shifts can't cover the daily floor of "
                f"{S.day_min + S.night_min}/day ({floor_need} needed).",
                "Add staff, raise shifts/employee, or lower the daily minimums."))
        if total_shifts > ceil_cap:
            problems.append(Problem(
                f"{total_shifts} total shifts exceed the {ceil_cap} the daily maxima "
                f"allow (<= {S.day_max + S.night_max}/day).",
                "Raise max day/night staffing, remove staff, or lower shifts/employee."))
        p = pattern_problem("Roster pattern", w, S.max_consec_work, S.max_consec_off)
        if p:
            problems.append(p)
        return problems

    # --- Mode A: fixed night team ------------------------------------------
    n_night = len(S.night_team)
    if not set(S.night_team).issubset(set(S.employees)):
        problems.append(Problem(
            "Night team includes someone who is not in the employee list.",
            "Choose night-team members from the current employee list."))
    if n_night < 1:
        problems.append(Problem(
            "No night-team members selected.",
            "Select at least 1 employee for the night team (any size)."))

    # Night coverage capacity (nights-only team).
    total_nights = w * n_night if S.night_team_nights_only else None
    if S.night_team_nights_only and n_night >= 1:
        if total_nights < S.night_min * days:
            problems.append(Problem(
                f"Night team supplies {total_nights} nights, but >= {S.night_min}/day "
                f"needs {S.night_min * days}.",
                f"Lower min nights/day, or give the night team more shifts."))
        if total_nights > S.night_max * days:
            need = _ceil(total_nights, days)
            problems.append(Problem(
                f"Night team must place {total_nights} nights, but <= {S.night_max}/day "
                f"only allows {S.night_max * days}.",
                f"Raise max nights/day to {need}, add a night-team member, or shrink "
                f"the night-team workload."))

    # Day coverage capacity.
    day_shifts = w * n_emp - (total_nights if total_nights is not None else 0)
    if day_shifts < S.day_min * days:
        problems.append(Problem(
            f"Only {day_shifts} day-shifts are available, but >= {S.day_min}/day "
            f"needs {S.day_min * days}.",
            f"Add day-team staff, or lower min day staffing to {day_shifts // days}."))
    if day_shifts > S.day_max * days:
        need = _ceil(day_shifts, days)
        problems.append(Problem(
            f"{day_shifts} day-shifts must be placed, but <= {S.day_max}/day only "
            f"allows {S.day_max * days}.",
            f"Raise max day staffing to {need}, or remove a day-team member."))

    # Per-employee run-length pattern feasibility.
    p = pattern_problem("Day-team pattern", w, S.max_consec_work, S.max_consec_off)
    if p:
        problems.append(p)
    if S.night_team_nights_only and n_night >= 1:
        p = pattern_problem("Night-team pattern", w, S.max_consec_night, S.max_consec_off)
        if p:
            problems.append(p)

    return problems


def relaxation_hints(settings: ScheduleSettings) -> list[str]:
    """Best-effort suggestions when the model passes preflight but the solver
    still proves no schedule exists (the binding conflict is a combination, not a
    single arithmetic check)."""
    S = settings
    return [
        f"Widen the day-staffing band (e.g. max day staff {S.day_max} -> {S.day_max + 1}).",
        f"Raise max consecutive off ({S.max_consec_off} -> {S.max_consec_off + 1}) "
        f"or max consecutive day run ({S.max_consec_work} -> {S.max_consec_work + 1}).",
        f"Allow more night overlap (max nights/day {S.night_max} -> {S.night_max + 1}).",
        "Reduce shifts/employee by 1 to loosen the packing.",
    ]


# ---------------------------------------------------------------------------
# MODEL + SOLVE
# ---------------------------------------------------------------------------

@dataclass
class Solution:
    grid: list[list[str]]      # grid[employee][day] in {D, N, O}
    status: str


def build_and_solve(settings: ScheduleSettings) -> Solution:
    S = settings
    days = S.days
    n = len(S.employees)
    elig = night_eligible_indices(S)            # may work nights
    elig_set = set(elig)
    day_cap = day_capable_indices(S)            # may work day shifts
    day_cap_set = set(day_cap)
    weekend_days = weekend_day_indices(S)
    weekend_pairs = weekend_pair_indices(S)

    model = cp_model.CpModel()

    work = {(e, d): model.NewBoolVar(f"work_{e}_{d}")
            for e in range(n) for d in range(days)}
    night = {(e, d): model.NewBoolVar(f"night_{e}_{d}")
             for e in range(n) for d in range(days)}

    for e in range(n):
        for d in range(days):
            model.Add(night[(e, d)] <= work[(e, d)])

    # Only night-eligible staff may work nights (rotate mode => everyone).
    for e in range(n):
        if e not in elig_set:
            for d in range(days):
                model.Add(night[(e, d)] == 0)

    # Fixed nights-only team: every working day is a night.
    if not S.is_rotate and S.night_team_nights_only:
        for e in elig_set:
            for d in range(days):
                model.Add(work[(e, d)] == night[(e, d)])

    # EXACT monthly workload per employee.
    for e in range(n):
        model.Add(sum(work[(e, d)] for d in range(days)) == S.shifts_per_employee)

    # Daily staffing bands. Night coverage is a band [night_min, night_max]:
    # overlap above the floor appears only where the exact totals force it (in
    # fixed_team mode), never below night_min (a night is never uncovered).
    night_count = {}
    for d in range(days):
        day_count = sum(work[(e, d)] - night[(e, d)] for e in range(n))
        model.Add(day_count >= S.day_min)
        model.Add(day_count <= S.day_max)
        ncv = model.NewIntVar(0, n, f"nightcount_{d}")
        model.Add(ncv == sum(night[(e, d)] for e in range(n)))
        model.Add(ncv >= S.night_min)
        model.Add(ncv <= S.night_max)
        night_count[d] = ncv

    # No Night -> Day the next calendar day (post-night recovery).
    for e in range(n):
        for d in range(days - 1):
            day_shift_next = work[(e, d + 1)] - night[(e, d + 1)]
            model.Add(night[(e, d)] + day_shift_next <= 1)

    # Run-length rules. The min-block encoding is specific to a minimum of 2.
    assert S.min_consec_work == 2 and S.min_consec_off == 2, (
        "build_and_solve() only encodes a minimum block length of 2.")
    for e in range(n):
        w = [work[(e, d)] for d in range(days)]

        # Max consecutive WORKING days (day-shift / working run cap).
        for start in range(days - S.max_consec_work):
            model.Add(sum(w[start:start + S.max_consec_work + 1]) <= S.max_consec_work)

        # Max consecutive NIGHTS (priority cap; binds night workers to <= 3).
        for start in range(days - S.max_consec_night):
            model.Add(sum(night[(e, d)] for d in range(start, start + S.max_consec_night + 1))
                      <= S.max_consec_night)

        # Max consecutive OFF days: any window of (MAX_OFF+1) days has >= 1 work.
        for start in range(days - S.max_consec_off):
            model.Add(sum(w[start:start + S.max_consec_off + 1]) >= 1)

        # Min consecutive working days = 2 (no isolated single work day).
        model.Add(w[0] <= w[1])
        model.Add(w[days - 1] <= w[days - 2])
        for d in range(1, days - 1):
            model.Add(w[d] <= w[d - 1] + w[d + 1])

        # Min consecutive off days = 2 (no isolated single off day).
        model.Add(w[1] <= w[0])
        model.Add(w[days - 2] <= w[days - 1])
        for d in range(1, days - 1):
            model.Add(w[d - 1] + w[d + 1] - w[d] <= 1)

    # ---- Soft fairness objectives (four equally-weighted goals) ----------
    # Each goal minimises a spread (max - min) so the burden lands evenly. They
    # are SOFT: they shape the chosen roster but never make a feasible rule set
    # infeasible. The validator re-derives each from the finished grid.

    def minmax_gap(values, name, ub):
        hi = model.NewIntVar(0, ub, name + "_hi")
        lo = model.NewIntVar(0, ub, name + "_lo")
        for c in values:
            model.Add(hi >= c)
            model.Add(lo <= c)
        return hi - lo

    # Goal 1: equal total shifts (already pinned exact; kept for robustness).
    totals = [sum(work[(e, d)] for d in range(days)) for e in range(n)]
    gap_total = minmax_gap(totals, "tot", days)

    # Goal 2: balanced night load across night-eligible staff.
    night_per_emp = [sum(night[(e, d)] for d in range(days)) for e in elig] or [0]
    gap_night = minmax_gap(night_per_emp, "night", days)

    # Goal 3: balanced weekends -- even Fri/Sat duty within each pool, plus a
    # full Fri+Sat weekend off for everyone (best effort).
    wknd = max(1, len(weekend_days))
    wknd_day = [sum(work[(e, d)] - night[(e, d)] for d in weekend_days) for e in day_cap] or [0]
    wknd_night = [sum(night[(e, d)] for d in weekend_days) for e in elig] or [0]
    gap_wknd = minmax_gap(wknd_day, "wkday", wknd) + minmax_gap(wknd_night, "wknight", wknd)
    has_full_off = []
    for e in range(n):
        bits = []
        for (f, s) in weekend_pairs:
            b = model.NewBoolVar(f"offwknd_{e}_{f}")
            model.Add(work[(e, f)] + work[(e, s)] == 0).OnlyEnforceIf(b)
            model.Add(work[(e, f)] + work[(e, s)] >= 1).OnlyEnforceIf(b.Not())
            bits.append(b)
        h = model.NewBoolVar(f"hasoffwknd_{e}")
        if bits:
            model.AddMaxEquality(h, bits)
        else:
            model.Add(h == 0)
        has_full_off.append(h)
    num_no_full_off = n - sum(has_full_off)
    weekend_term = gap_wknd + num_no_full_off

    # Goal 4: balanced "undesirable runs", spread evenly *within* each pool (a
    # fixed night team structurally carries all overlap, so a cross-pool
    # comparison is not a fairness signal -- mirror the weekend goal). Rough load
    # per person = overlap nights worked + count of max-length day runs + count
    # of max-length night runs. This is the EXACT quantity validate() re-derives
    # from the grid, so optimising it actually moves the validated metric (a run
    # that hits the cap can only be length == cap, so one window == one run).
    overlap = {}
    for d in range(days):
        b = model.NewBoolVar(f"overlap_{d}")          # a stacked night above floor
        model.Add(night_count[d] >= S.night_min + 1).OnlyEnforceIf(b)
        model.Add(night_count[d] <= S.night_min).OnlyEnforceIf(b.Not())
        overlap[d] = b

    Lw, Ln = S.max_consec_work, S.max_consec_night
    rough_terms = [0] * n
    for e in range(n):
        parts = []
        if e in elig_set:                              # overlap nights worked
            for d in range(days):
                ov = model.NewBoolVar(f"ovn_{e}_{d}")
                model.Add(ov <= night[(e, d)])
                model.Add(ov <= overlap[d])
                model.Add(ov >= night[(e, d)] + overlap[d] - 1)
                parts.append(ov)
        if e in day_cap_set:                           # max-length day runs
            for start in range(days - Lw + 1):
                seg = [work[(e, d)] - night[(e, d)] for d in range(start, start + Lw)]
                wmax = model.NewBoolVar(f"wmax_{e}_{start}")
                model.Add(sum(seg) >= Lw).OnlyEnforceIf(wmax)
                model.Add(sum(seg) <= Lw - 1).OnlyEnforceIf(wmax.Not())
                parts.append(wmax)
        if e in elig_set:                              # max-length night runs
            for start in range(days - Ln + 1):
                seg = [night[(e, d)] for d in range(start, start + Ln)]
                nmax = model.NewBoolVar(f"nmax_{e}_{start}")
                model.Add(sum(seg) >= Ln).OnlyEnforceIf(nmax)
                model.Add(sum(seg) <= Ln - 1).OnlyEnforceIf(nmax.Not())
                parts.append(nmax)
        rough_terms[e] = sum(parts) if parts else 0
    # Sum the per-pool spreads -- an upper bound on the max-pool spread the
    # validator checks, so minimising it drives the validated metric down.
    pools = [elig] + ([day_cap] if set(day_cap) != elig_set else [])
    gap_runs = sum(minmax_gap([rough_terms[e] for e in pool] or [0], f"rough{k}", 2 * days)
                   for k, pool in enumerate(pools))

    model.Minimize(
        S.w_fair_total * gap_total
        + S.w_fair_night * gap_night
        + S.w_fair_weekend * weekend_term
        + S.w_fair_runs * gap_runs
    )

    solver = cp_model.CpSolver()
    solver.parameters.max_deterministic_time = S.solver_det_time_limit
    solver.parameters.max_time_in_seconds = S.solver_time_limit
    solver.parameters.num_search_workers = 1
    solver.parameters.random_seed = 42

    status = solver.Solve(model)
    status_name = solver.StatusName(status)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return Solution(grid=[], status=status_name)

    grid: list[list[str]] = []
    for e in range(n):
        row = []
        for d in range(days):
            if solver.Value(work[(e, d)]) == 0:
                row.append(OFF)
            elif solver.Value(night[(e, d)]) == 1:
                row.append(NIGHT)
            else:
                row.append(DAY)
        grid.append(row)
    return Solution(grid=grid, status=status_name)


# ---------------------------------------------------------------------------
# INDEPENDENT VALIDATOR
# ---------------------------------------------------------------------------

@dataclass
class RuleResult:
    name: str
    passed: bool
    measured: str


def _runs(seq: list[bool]) -> list[int]:
    lengths, current = [], 0
    for v in seq:
        if v:
            current += 1
        elif current:
            lengths.append(current)
            current = 0
    if current:
        lengths.append(current)
    return lengths


def employee_loads(settings: ScheduleSettings, grid: list[list[str]]) -> list[dict]:
    """Per-employee fairness facts, re-derived from the grid (used by the UI,
    the Excel summary, and the fairness validator so they all agree)."""
    S = settings
    days = S.days
    weekend_days = weekend_day_indices(S)
    weekend_pairs = weekend_pair_indices(S)
    night_counts = [sum(1 for e in range(len(S.employees)) if grid[e][d] == NIGHT)
                    for d in range(days)]
    loads = []
    for e in range(len(S.employees)):
        row = grid[e]
        day_runs = _runs([row[d] == DAY for d in range(days)])
        night_runs = _runs([row[d] == NIGHT for d in range(days)])
        overlaps = sum(1 for d in range(days)
                       if row[d] == NIGHT and night_counts[d] >= S.night_min + 1)
        max_runs = (sum(1 for r in day_runs if r >= S.max_consec_work)
                    + sum(1 for r in night_runs if r >= S.max_consec_night))
        loads.append({
            "name": S.employees[e],
            "night_member": is_night_member(S, e),
            "total": sum(1 for d in range(days) if row[d] != OFF),
            "day": sum(1 for d in range(days) if row[d] == DAY),
            "night": sum(1 for d in range(days) if row[d] == NIGHT),
            "weekend": sum(1 for d in weekend_days if row[d] != OFF),
            "overlaps": overlaps,
            "max_runs": max_runs,
            "rough": overlaps + max_runs,
            "weekends_off": sum(1 for (f, s) in weekend_pairs
                                if row[f] == OFF and row[s] == OFF),
        })
    return loads


def validate(settings: ScheduleSettings, grid: list[list[str]]) -> list[RuleResult]:
    S = settings
    days = S.days
    n = len(S.employees)
    rotate = S.is_rotate
    elig = night_eligible_indices(S)
    elig_set = set(elig)
    day_cap = day_capable_indices(S)
    weekend_days = weekend_day_indices(S)
    weekend_pairs = weekend_pair_indices(S)
    results: list[RuleResult] = []

    day_counts = [sum(1 for e in range(n) if grid[e][d] == DAY) for d in range(days)]
    night_counts = [sum(1 for e in range(n) if grid[e][d] == NIGHT) for d in range(days)]

    # ===================== HARD RULES =====================================
    results.append(RuleResult(
        f"Day staffing per day within [{S.day_min}, {S.day_max}]",
        all(S.day_min <= c <= S.day_max for c in day_counts),
        f"min={min(day_counts)}, max={max(day_counts)}"))
    results.append(RuleResult(
        f"Night coverage every day >= {S.night_min} (<= {S.night_max})",
        all(S.night_min <= c <= S.night_max for c in night_counts),
        f"min={min(night_counts)}, max={max(night_counts)}"))

    # Forced-overlap minimality only has a fixed target with a nights-only team.
    if not rotate and S.night_team_nights_only:
        expected = expected_double_nights(S)
        actual = sum(max(0, c - S.night_min) for c in night_counts)
        results.append(RuleResult(
            "Night overlap is minimal (only forced extra nights)",
            actual == expected,
            f"overlap-nights above floor={actual} (forced minimum={expected})"))

    results.append(RuleResult(
        "At most one shift per employee per day",
        all(grid[e][d] in (DAY, NIGHT, OFF) for e in range(n) for d in range(days)),
        "one code per cell"))

    totals = [sum(1 for d in range(days) if grid[e][d] != OFF) for e in range(n)]
    results.append(RuleResult(
        f"Exactly {S.shifts_per_employee} shifts per employee",
        all(t == S.shifts_per_employee for t in totals),
        f"totals={totals}"))

    night_workers = {S.employees[e] for e in range(n)
                     if any(grid[e][d] == NIGHT for d in range(days))}
    if rotate:
        # No dedicated team: nights may land on anyone -- just confirm that.
        results.append(RuleResult(
            "All staff are night-eligible (rotation)",
            night_workers.issubset(set(S.employees)),
            f"{len(night_workers)} of {n} staff worked >=1 night"))
    else:
        results.append(RuleResult(
            "Only the night team works nights",
            night_workers.issubset(set(S.night_team)),
            f"night workers={sorted(night_workers)}"))
        if S.night_team_nights_only:
            # With a nights-only team, every member works only nights, so all of
            # them necessarily carry >=1 night. (If the team could also work days
            # a member legitimately taking zero nights must NOT count as a fault,
            # so this check only applies in the nights-only case.)
            results.append(RuleResult(
                f"Exactly {len(S.night_team)} people carry nights for the month",
                len(night_workers) == len(S.night_team),
                f"{len(night_workers)} have >=1 night: {sorted(night_workers)}"))
            nightteam_dayshifts = [S.employees[e] for e in elig_set
                                   if any(grid[e][d] == DAY for d in range(days))]
            results.append(RuleResult(
                "Night team works nights only (no day shifts)",
                not nightteam_dayshifts,
                "0 day shifts by night team" if not nightteam_dayshifts
                else f"violations: {nightteam_dayshifts}"))

    nd = [(S.employees[e], f"D{d + 1}->D{d + 2}") for e in range(n)
          for d in range(days - 1) if grid[e][d] == NIGHT and grid[e][d + 1] == DAY]
    results.append(RuleResult(
        "No Night->Day next-day transition",
        not nd,
        f"{len(nd)} transitions" + ("" if not nd else f": {nd[:6]}")))

    max_day_run = max((max(_runs([grid[e][d] == DAY for d in range(days)]), default=0)
                       for e in range(n)), default=0)
    results.append(RuleResult(
        f"Max {S.max_consec_work} consecutive day shifts",
        max_day_run <= S.max_consec_work,
        f"longest day streak={max_day_run}"))

    max_night_run = max((max(_runs([grid[e][d] == NIGHT for d in range(days)]), default=0)
                         for e in range(n)), default=0)
    results.append(RuleResult(
        f"Max {S.max_consec_night} consecutive nights",
        max_night_run <= S.max_consec_night,
        f"longest night streak={max_night_run}"))

    min_work_run = min((min(_runs([grid[e][d] != OFF for d in range(days)]), default=S.min_consec_work)
                        for e in range(n)), default=S.min_consec_work)
    results.append(RuleResult(
        f"Min {S.min_consec_work} consecutive working days",
        min_work_run >= S.min_consec_work,
        f"shortest work block={min_work_run}"))

    min_off_run = min((min(_runs([grid[e][d] == OFF for d in range(days)]), default=S.min_consec_off)
                       for e in range(n)), default=S.min_consec_off)
    results.append(RuleResult(
        f"Min {S.min_consec_off} consecutive off days",
        min_off_run >= S.min_consec_off,
        f"shortest off block={min_off_run}"))

    max_off_run = max((max(_runs([grid[e][d] == OFF for d in range(days)]), default=0)
                       for e in range(n)), default=0)
    results.append(RuleResult(
        f"Max {S.max_consec_off} consecutive off days",
        max_off_run <= S.max_consec_off,
        f"longest off streak={max_off_run}"))

    # ===================== FAIRNESS (4 soft goals) ========================
    # Each reports PASS/FAIL against a tunable tolerance, plus the spread.
    loads = employee_loads(S, grid)

    def spread(values):
        vals = list(values) or [0]
        return max(vals) - min(vals)

    # F1 -- equal total shifts.
    sp_total = spread(ld["total"] for ld in loads)
    results.append(RuleResult(
        f"Fairness - equal total shifts (spread <= {S.fair_tol_total})",
        sp_total <= S.fair_tol_total,
        f"spread={sp_total} (every total = {S.shifts_per_employee})"))

    # F2 -- balanced night load across night-eligible staff.
    night_by = {loads[e]["name"]: loads[e]["night"] for e in elig}
    sp_night = spread(night_by.values())
    results.append(RuleResult(
        f"Fairness - balanced night load (spread <= {S.fair_tol_night})",
        sp_night <= S.fair_tol_night,
        f"spread={sp_night}  nights/eligible={night_by}"))

    # F3 -- balanced weekends: even Fri/Sat duty in each pool + a weekend off each.
    sp_wknd_day = spread(sum(1 for d in weekend_days if grid[e][d] == DAY) for e in day_cap)
    sp_wknd_night = spread(sum(1 for d in weekend_days if grid[e][d] == NIGHT) for e in elig)
    no_weekend_off = [loads[e]["name"] for e in range(n)
                      if weekend_pairs and loads[e]["weekends_off"] == 0]
    passed_wknd = (max(sp_wknd_day, sp_wknd_night) <= S.fair_tol_weekend
                   and not no_weekend_off)
    results.append(RuleResult(
        f"Fairness - balanced weekends (spread <= {S.fair_tol_weekend}, a weekend off each)",
        passed_wknd,
        f"day spread={sp_wknd_day}, night spread={sp_wknd_night}, "
        f"no weekend-off for: {no_weekend_off or 'none'}"))

    # F4 -- balanced undesirable runs *within* each pool (a fixed night team
    # structurally carries all overlap; comparing it to the day team is not a
    # fairness signal -- balance is what matters inside each group).
    rough_by = {ld["name"]: ld["rough"] for ld in loads}
    pools = [elig] + ([day_cap] if set(day_cap) != elig_set else [])
    sp_rough = max((spread(loads[e]["rough"] for e in pool) for pool in pools), default=0)
    results.append(RuleResult(
        f"Fairness - balanced undesirable runs (within-pool spread <= {S.fair_tol_runs})",
        sp_rough <= S.fair_tol_runs,
        f"within-pool spread={sp_rough}  rough-load(overlap+max-runs)={rough_by}"))

    return results


# ---------------------------------------------------------------------------
# EXCEL EXPORT
# ---------------------------------------------------------------------------

FILL_DAY = PatternFill("solid", fgColor="C6EFCE")
FILL_NIGHT = PatternFill("solid", fgColor="BDD7EE")
FILL_OFF = PatternFill("solid", fgColor="FFC7CE")
FILL_PASS = PatternFill("solid", fgColor="C6EFCE")
FILL_FAIL = PatternFill("solid", fgColor="FFC7CE")
FILL_HEADER = PatternFill("solid", fgColor="305496")
FILL_WEEKEND = PatternFill("solid", fgColor="FFF2CC")

FONT_HEADER = Font(bold=True, color="FFFFFF")
FONT_BOLD = Font(bold=True)
CENTER = Alignment(horizontal="center", vertical="center")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CELL_LABEL = {DAY: "D", NIGHT: "N", OFF: "OFF"}
CELL_FILL = {DAY: FILL_DAY, NIGHT: FILL_NIGHT, OFF: FILL_OFF}

# Hex colors reused by the Streamlit grid so the two views match.
WEB_COLORS = {DAY: "#C6EFCE", NIGHT: "#BDD7EE", OFF: "#FFC7CE"}


def _style(cell, value="", fill=None, font=None, align=CENTER, border=True):
    cell.value = value
    if fill:
        cell.fill = fill
    if font:
        cell.font = font
    cell.alignment = align
    if border:
        cell.border = BORDER


def build_workbook(settings: ScheduleSettings, grid: list[list[str]],
                   results: list[RuleResult]) -> Workbook:
    wb = Workbook()
    _write_roster(settings, wb.active, grid)
    _write_summary(settings, wb.create_sheet("Summary"), grid, results)
    return wb


def export(settings: ScheduleSettings, grid: list[list[str]],
           results: list[RuleResult], path: str = "schedule.xlsx") -> None:
    build_workbook(settings, grid, results).save(path)


def export_bytes(settings: ScheduleSettings, grid: list[list[str]],
                 results: list[RuleResult]) -> bytes:
    buf = BytesIO()
    build_workbook(settings, grid, results).save(buf)
    return buf.getvalue()


def _write_roster(settings: ScheduleSettings, ws, grid: list[list[str]]) -> None:
    S = settings
    days = S.days
    ws.title = "Schedule"
    weekend = set(weekend_day_indices(S))

    c = ws.cell(1, 1)
    _style(c, f"Employee\n{S.month_label}", FILL_HEADER, FONT_HEADER)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for d in range(days):
        cell = ws.cell(1, 2 + d)
        is_wknd = d in weekend
        _style(cell, f"D{d + 1}\n{weekday_of(S, d)}",
               FILL_WEEKEND if is_wknd else FILL_HEADER,
               FONT_BOLD if is_wknd else FONT_HEADER)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for j, label in enumerate(["Total", "Day", "Night", "Fri/Sat"]):
        _style(ws.cell(1, 2 + days + j), label, FILL_HEADER, FONT_HEADER)

    for e in range(len(S.employees)):
        label = S.employees[e] + ("  (night team)" if is_night_member(S, e) else "")
        _style(ws.cell(2 + e, 1), label, font=FONT_BOLD,
               align=Alignment(horizontal="left", vertical="center"))
        for d in range(days):
            code = grid[e][d]
            _style(ws.cell(2 + e, 2 + d), CELL_LABEL[code], CELL_FILL[code])
        total = sum(1 for d in range(days) if grid[e][d] != OFF)
        day_n = sum(1 for d in range(days) if grid[e][d] == DAY)
        night_n = sum(1 for d in range(days) if grid[e][d] == NIGHT)
        wknd_n = sum(1 for d in weekend_day_indices(S) if grid[e][d] != OFF)
        for j, val in enumerate([total, day_n, night_n, wknd_n]):
            _style(ws.cell(2 + e, 2 + days + j), val, font=FONT_BOLD)

    legend_row = len(S.employees) + 4
    _style(ws.cell(legend_row, 1), "Legend:", font=FONT_BOLD, border=False)
    for j, (code, text) in enumerate([(DAY, "Day shift"), (NIGHT, "Night shift"), (OFF, "Off")]):
        _style(ws.cell(legend_row, 2 + j * 2), CELL_LABEL[code], CELL_FILL[code])
        _style(ws.cell(legend_row, 3 + j * 2), text, border=False,
               align=Alignment(horizontal="left", vertical="center"))
    _style(ws.cell(legend_row + 1, 1), roster_caption(S),
           font=FONT_BOLD, border=False, align=Alignment(horizontal="left", vertical="center"))

    ws.column_dimensions["A"].width = 18
    for d in range(days):
        ws.column_dimensions[get_column_letter(2 + d)].width = 6
    for j in range(4):
        ws.column_dimensions[get_column_letter(2 + days + j)].width = 8
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "B2"


def _write_summary(settings: ScheduleSettings, ws, grid: list[list[str]],
                   results: list[RuleResult]) -> None:
    S = settings
    days = S.days
    elig = night_eligible_indices(S)
    loads = employee_loads(S, grid)
    row = 1

    def header(text):
        nonlocal row
        _style(ws.cell(row, 1), text, FILL_HEADER, FONT_HEADER,
               align=Alignment(horizontal="left", vertical="center"))
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        row += 1

    header(f"Per-employee fairness summary - {S.month_label}")
    for j, h in enumerate(["Employee", "Total", "Day", "Night", "Fri/Sat", "Rough"]):
        _style(ws.cell(row, 1 + j), h, FILL_HEADER, FONT_HEADER)
    row += 1
    for e, ld in enumerate(loads):
        name = ld["name"] + ("  (night team)" if ld["night_member"] else "")
        cells = [name, ld["total"], ld["day"], ld["night"], ld["weekend"], ld["rough"]]
        for j, val in enumerate(cells):
            _style(ws.cell(row, 1 + j), val,
                   align=Alignment(horizontal="left" if j == 0 else "center", vertical="center"))
        row += 1
    row += 1

    header("Night coverage")
    split = {loads[e]["name"]: loads[e]["night"] for e in elig}
    for name, nval in split.items():
        _style(ws.cell(row, 1), name, align=Alignment(horizontal="left", vertical="center"))
        _style(ws.cell(row, 2), f"{nval} nights")
        row += 1
    if split:
        _style(ws.cell(row, 1), "Gap", font=FONT_BOLD, align=Alignment(horizontal="left", vertical="center"))
        _style(ws.cell(row, 2), max(split.values()) - min(split.values()), font=FONT_BOLD)
        row += 1
    overlap_days = sum(1 for d in range(days)
                       if sum(1 for e in range(len(S.employees)) if grid[e][d] == NIGHT)
                       >= S.night_min + 1)
    _style(ws.cell(row, 1), "Overlap-night days",
           font=FONT_BOLD, align=Alignment(horizontal="left", vertical="center"))
    _style(ws.cell(row, 2), overlap_days, font=FONT_BOLD)
    row += 2

    header("Rule validation")
    for j, h in enumerate(["Rule", "Result", "Measured"]):
        _style(ws.cell(row, 1 + j), h, FILL_HEADER, FONT_HEADER)
    ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=6)
    row += 1
    for r in results:
        _style(ws.cell(row, 1), r.name, align=Alignment(horizontal="left", vertical="center"))
        _style(ws.cell(row, 2), "PASS" if r.passed else "FAIL",
               FILL_PASS if r.passed else FILL_FAIL, FONT_BOLD)
        _style(ws.cell(row, 3), r.measured, align=Alignment(horizontal="left", vertical="center"))
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=6)
        row += 1

    ws.column_dimensions["A"].width = 52
    for col in ("B", "C", "D", "E", "F"):
        ws.column_dimensions[col].width = 14


# ---------------------------------------------------------------------------
# MAIN (CLI)
# ---------------------------------------------------------------------------

def print_summary(settings: ScheduleSettings, grid: list[list[str]],
                  results: list[RuleResult]) -> bool:
    S = settings
    loads = employee_loads(S, grid)
    print(f"\n{roster_caption(S)}")
    print("\nPer-employee counts:")
    print(f"  {'Employee':<12} {'Total':>5} {'Day':>4} {'Night':>6} {'Fri/Sat':>8} {'Rough':>6}")
    for ld in loads:
        tag = "  <- night team" if ld["night_member"] else ""
        print(f"  {ld['name']:<12} {ld['total']:>5} {ld['day']:>4} {ld['night']:>6} "
              f"{ld['weekend']:>8} {ld['rough']:>6}{tag}")

    print("\nRule validation:")
    all_pass = True
    for r in results:
        all_pass &= r.passed
        print(f"  [{'PASS' if r.passed else 'FAIL'}] {r.name:<58} | {r.measured}")
    return all_pass


def main() -> int:
    settings = ScheduleSettings()

    problems = preflight(settings)
    if problems:
        print("INFEASIBLE rule set - stopping before scheduling:")
        for p in problems:
            print(f"  - {p.message}\n    suggestion: {p.suggestion}")
        return 2

    solution = build_and_solve(settings)
    if not solution.grid:
        print(f"No schedule found (solver status: {solution.status}).")
        print("The constraints are jointly infeasible. Try:")
        for hint in relaxation_hints(settings):
            print(f"  - {hint}")
        return 2

    results = validate(settings, solution.grid)
    all_pass = print_summary(settings, solution.grid, results)
    export(settings, solution.grid, results)
    print(f"\nWrote schedule.xlsx (solver status: {solution.status}).")

    if not all_pass:
        print("\nWARNING: at least one rule FAILED validation - do not use this roster.")
        return 1
    print("\nAll rules PASS.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
