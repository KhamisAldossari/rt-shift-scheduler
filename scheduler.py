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
  * Nights are covered by a fixed night team that works nights ONLY. To let
    every night-team member reach the exact target, night coverage is "at least
    1, at most `night_max`" per day -- a 2nd overlapping night appears only on
    the days arithmetic forces it, never as a free-for-all, and a night is never
    left uncovered.
  * No Night->Day next-calendar-day transition for anyone.
  * Day-shift runs are capped at `max_consec_work`; night runs at
    `max_consec_night`; off runs floored at `min_consec_off`, capped at
    `max_consec_off`; work runs floored at `min_consec_work`.
  * The solver tries to give every employee >= 1 full Fri+Sat weekend off and to
    spread Fri/Sat duty evenly within each team.

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
    # The fixed night team works nights ONLY (no day shifts).
    night_team: list[str] = field(
        default_factory=lambda: ["Employee 6", "Employee 7"])
    night_team_nights_only: bool = True

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
    fairness_max_gap: int = 1
    weekend_days: tuple = ("Fri", "Sat")            # the "weekend" for fairness

    # --- Objective weights (all minimized; lower = better) ------------------
    w_night_split: int = 100                        # even night split (team)
    w_workload: int = 80                            # even total workload
    w_weekend_off: int = 40                         # >=1 full weekend off each
    w_weekend_split: int = 10                       # even Fri/Sat duty per team

    solver_time_limit: float = 30.0

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
    def day_team(self) -> list[str]:
        return [e for e in self.employees if e not in self.night_team]


# ---------------------------------------------------------------------------
# CALENDAR HELPERS
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


def expected_double_nights(settings: ScheduleSettings) -> int:
    """How many days must carry a 2nd overlapping night, given exact totals.

    With a nights-only team each on `shifts_per_employee`, the month has a fixed
    number of night-shifts; spread over `days` at >= 1/day, the surplus lands as
    double-night days. This is the minimum the arithmetic forces -- nothing more.
    """
    if not settings.night_team_nights_only:
        return 0
    total_nights = settings.shifts_per_employee * len(settings.night_team)
    return max(0, total_nights - settings.days)


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
    n_night = len(S.night_team)
    n_day = n_emp - n_night
    w = S.shifts_per_employee

    # --- Basic shape --------------------------------------------------------
    if n_emp == 0:
        problems.append(Problem("No employees configured.",
                                "Add at least 3-4 employees."))
        return problems
    if not set(S.night_team).issubset(set(S.employees)):
        problems.append(Problem(
            "Night team includes someone who is not in the employee list.",
            "Choose night-team members from the current employee list."))
    if n_night < 1:
        problems.append(Problem(
            "No night-team members selected.",
            "Select exactly 2 employees as the night team."))
    if S.min_consec_work != 2 or S.min_consec_off != 2:
        problems.append(Problem(
            "Minimum consecutive work/off must be 2 (the model encodes exactly 2).",
            "Keep 'min consecutive work' and 'min consecutive off' at 2."))
    if w > days:
        problems.append(Problem(
            f"{w} shifts/employee cannot fit in a {days}-day month.",
            f"Lower shifts/employee to at most {days}, or choose a longer month."))
        return problems

    # --- Night coverage capacity (nights-only team) -------------------------
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
                f"Raise max nights/day to {need}, or shrink the night-team workload."))

    # --- Day coverage capacity ----------------------------------------------
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

    # --- Per-employee run-length pattern feasibility ------------------------
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
    eligible_idx = [S.employees.index(name) for name in S.night_team
                    if name in S.employees]
    night_team_idx = set(eligible_idx)
    day_team_idx = [i for i in range(n) if i not in night_team_idx]
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

    # Only the night team may work nights.
    for e in range(n):
        if e not in night_team_idx:
            for d in range(days):
                model.Add(night[(e, d)] == 0)

    # Night team works nights ONLY (every working day is a night).
    if S.night_team_nights_only:
        for e in night_team_idx:
            for d in range(days):
                model.Add(work[(e, d)] == night[(e, d)])

    # EXACT monthly workload per employee.
    for e in range(n):
        model.Add(sum(work[(e, d)] for d in range(days)) == S.shifts_per_employee)

    # Daily staffing bands. Night coverage is a band [night_min, night_max]: a
    # 2nd night appears only where the exact totals force it (see below), never
    # below night_min (a night is never uncovered).
    for d in range(days):
        day_count = sum(work[(e, d)] - night[(e, d)] for e in range(n))
        model.Add(day_count >= S.day_min)
        model.Add(day_count <= S.day_max)
        night_count = sum(night[(e, d)] for e in range(n))
        model.Add(night_count >= S.night_min)
        model.Add(night_count <= S.night_max)

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

        # Max consecutive WORKING days (day-shift run cap for the day team).
        for start in range(days - S.max_consec_work):
            model.Add(sum(w[start:start + S.max_consec_work + 1]) <= S.max_consec_work)

        # Max consecutive NIGHTS (priority cap; binds the night team to <= 3).
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

    # ---- Soft objectives -------------------------------------------------

    def minmax_gap(values, hi_name, lo_name, ub):
        hi = model.NewIntVar(0, ub, hi_name)
        lo = model.NewIntVar(0, ub, lo_name)
        for c in values:
            model.Add(hi >= c)
            model.Add(lo <= c)
        return hi - lo

    # (a) Even night split across the night team.
    night_counts = [sum(night[(e, d)] for d in range(days)) for e in eligible_idx] or [0]
    night_split_gap = minmax_gap(night_counts, "night_hi", "night_lo", days)

    # (b) Even total workload (already pinned to exact, but kept for robustness).
    totals = [sum(work[(e, d)] for d in range(days)) for e in range(n)]
    workload_gap = minmax_gap(totals, "tot_hi", "tot_lo", days)

    # (c) Even Fri/Sat duty within each team.
    wknd = max(1, len(weekend_days))
    day_wknd = [sum(work[(e, d)] - night[(e, d)] for d in weekend_days) for e in day_team_idx] or [0]
    night_wknd = [sum(night[(e, d)] for d in weekend_days) for e in eligible_idx] or [0]
    wknd_day_gap = minmax_gap(day_wknd, "wkday_hi", "wkday_lo", wknd)
    wknd_night_gap = minmax_gap(night_wknd, "wknight_hi", "wknight_lo", wknd)

    # (d) Every employee gets >= 1 full Fri+Sat weekend off (best effort).
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

    model.Minimize(
        S.w_night_split * night_split_gap
        + S.w_workload * workload_gap
        + S.w_weekend_off * num_no_full_off
        + S.w_weekend_split * (wknd_day_gap + wknd_night_gap)
    )

    solver = cp_model.CpSolver()
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


def validate(settings: ScheduleSettings, grid: list[list[str]]) -> list[RuleResult]:
    S = settings
    days = S.days
    n = len(S.employees)
    night_team_idx = {S.employees.index(x) for x in S.night_team if x in S.employees}
    day_team_idx = {i for i in range(n) if i not in night_team_idx}
    weekend_days = weekend_day_indices(S)
    weekend_pairs = weekend_pair_indices(S)
    results: list[RuleResult] = []

    day_counts = [sum(1 for e in range(n) if grid[e][d] == DAY) for d in range(days)]
    night_counts = [sum(1 for e in range(n) if grid[e][d] == NIGHT) for d in range(days)]

    results.append(RuleResult(
        f"Day staffing per day within [{S.day_min}, {S.day_max}]",
        all(S.day_min <= c <= S.day_max for c in day_counts),
        f"min={min(day_counts)}, max={max(day_counts)}"))
    results.append(RuleResult(
        f"Night coverage every day >= {S.night_min} (<= {S.night_max})",
        all(S.night_min <= c <= S.night_max for c in night_counts),
        f"min={min(night_counts)}, max={max(night_counts)}"))

    expected_doubles = expected_double_nights(S)
    actual_doubles = sum(1 for c in night_counts if c >= 2)
    results.append(RuleResult(
        "2nd night only where needed (overlap is minimal)",
        actual_doubles == expected_doubles,
        f"double-night days={actual_doubles} (forced minimum={expected_doubles})"))

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
    results.append(RuleResult(
        "Only the night team works nights",
        night_workers.issubset(set(S.night_team)),
        f"night workers={sorted(night_workers)}"))
    results.append(RuleResult(
        f"Exactly {len(S.night_team)} people carry nights for the month",
        len(night_workers) == len(S.night_team),
        f"{len(night_workers)} have >=1 night: {sorted(night_workers)}"))

    nightteam_dayshifts = [S.employees[e] for e in night_team_idx
                           if any(grid[e][d] == DAY for d in range(days))]
    results.append(RuleResult(
        "Night team works nights only (no day shifts)",
        not nightteam_dayshifts,
        "0 day shifts by night team" if not nightteam_dayshifts else f"violations: {nightteam_dayshifts}"))

    night_by_emp = {S.employees[e]: sum(1 for d in range(days) if grid[e][d] == NIGHT)
                    for e in sorted(night_team_idx)}
    nvals = list(night_by_emp.values()) or [0]
    results.append(RuleResult(
        f"Nights split evenly across the night team (gap <= {S.fairness_max_gap})",
        (max(nvals) - min(nvals)) <= S.fairness_max_gap and min(nvals) >= 1,
        f"{night_by_emp} -> gap={max(nvals) - min(nvals)}"))

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

    day_wknd = [sum(1 for d in weekend_days if grid[e][d] == DAY) for e in sorted(day_team_idx)] or [0]
    results.append(RuleResult(
        f"Fri/Sat day shifts spread evenly across the day team (gap <= {S.fairness_max_gap})",
        (max(day_wknd) - min(day_wknd)) <= S.fairness_max_gap,
        f"counts={day_wknd} -> gap={max(day_wknd) - min(day_wknd)}"))

    night_wknd = [sum(1 for d in weekend_days if grid[e][d] == NIGHT) for e in sorted(night_team_idx)] or [0]
    results.append(RuleResult(
        f"Fri/Sat nights spread evenly across the night team (gap <= {S.fairness_max_gap})",
        (max(night_wknd) - min(night_wknd)) <= S.fairness_max_gap,
        f"counts={night_wknd} -> gap={max(night_wknd) - min(night_wknd)}"))

    if weekend_pairs:
        full_off = {S.employees[e]: sum(1 for (f, s) in weekend_pairs
                                        if grid[e][f] == OFF and grid[e][s] == OFF)
                    for e in range(n)}
        no_off = [name for name, c in full_off.items() if c == 0]
        results.append(RuleResult(
            "Every employee gets >= 1 full Fri+Sat weekend off",
            not no_off,
            "all have >=1" if not no_off else f"none for: {no_off}"))

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
        is_night = S.employees[e] in S.night_team
        label = S.employees[e] + ("  (night team)" if is_night else "")
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
    night_names = ", ".join(S.night_team) if S.night_team else "(none)"
    _style(ws.cell(legend_row + 1, 1),
           f"Roster: {S.month_label} ({days} days). Night team: {night_names} (nights only).",
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
    night_team_idx = {S.employees.index(x) for x in S.night_team if x in S.employees}
    weekend_days = weekend_day_indices(S)
    row = 1

    def header(text):
        nonlocal row
        _style(ws.cell(row, 1), text, FILL_HEADER, FONT_HEADER,
               align=Alignment(horizontal="left", vertical="center"))
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        row += 1

    header(f"Per-employee shift counts - {S.month_label}")
    for j, h in enumerate(["Employee", "Total", "Day", "Night", "Fri/Sat"]):
        _style(ws.cell(row, 1 + j), h, FILL_HEADER, FONT_HEADER)
    row += 1
    for e in range(len(S.employees)):
        total = sum(1 for d in range(days) if grid[e][d] != OFF)
        day_n = sum(1 for d in range(days) if grid[e][d] == DAY)
        night_n = sum(1 for d in range(days) if grid[e][d] == NIGHT)
        wknd_n = sum(1 for d in weekend_days if grid[e][d] != OFF)
        name = S.employees[e] + ("  (night team)" if S.employees[e] in S.night_team else "")
        for j, val in enumerate([name, total, day_n, night_n, wknd_n]):
            _style(ws.cell(row, 1 + j), val,
                   align=Alignment(horizontal="left" if j == 0 else "center", vertical="center"))
        row += 1
    row += 1

    header("Night coverage - night team")
    split = {S.employees[e]: sum(1 for d in range(days) if grid[e][d] == NIGHT)
             for e in sorted(night_team_idx)}
    for name, nval in split.items():
        _style(ws.cell(row, 1), name, align=Alignment(horizontal="left", vertical="center"))
        _style(ws.cell(row, 2), f"{nval} nights")
        row += 1
    if split:
        _style(ws.cell(row, 1), "Gap", font=FONT_BOLD, align=Alignment(horizontal="left", vertical="center"))
        _style(ws.cell(row, 2), max(split.values()) - min(split.values()), font=FONT_BOLD)
        row += 1
    _style(ws.cell(row, 1), "Double-night days",
           font=FONT_BOLD, align=Alignment(horizontal="left", vertical="center"))
    _style(ws.cell(row, 2), sum(1 for d in range(days)
                                if sum(1 for e in range(len(S.employees)) if grid[e][d] == NIGHT) >= 2),
           font=FONT_BOLD)
    row += 2

    header("Rule validation")
    for j, h in enumerate(["Rule", "Result", "Measured"]):
        _style(ws.cell(row, 1 + j), h, FILL_HEADER, FONT_HEADER)
    ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=5)
    row += 1
    for r in results:
        _style(ws.cell(row, 1), r.name, align=Alignment(horizontal="left", vertical="center"))
        _style(ws.cell(row, 2), "PASS" if r.passed else "FAIL",
               FILL_PASS if r.passed else FILL_FAIL, FONT_BOLD)
        _style(ws.cell(row, 3), r.measured, align=Alignment(horizontal="left", vertical="center"))
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=5)
        row += 1

    ws.column_dimensions["A"].width = 52
    for col in ("B", "C", "D", "E"):
        ws.column_dimensions[col].width = 16


# ---------------------------------------------------------------------------
# MAIN (CLI)
# ---------------------------------------------------------------------------

def print_summary(settings: ScheduleSettings, grid: list[list[str]],
                  results: list[RuleResult]) -> bool:
    S = settings
    days = S.days
    night_team_idx = {S.employees.index(x) for x in S.night_team if x in S.employees}
    night_names = ", ".join(S.night_team) if S.night_team else "(none)"
    print(f"\nRoster: {S.month_label} ({days} days). Night team: {night_names} (nights only).")
    print("\nPer-employee counts:")
    print(f"  {'Employee':<12} {'Total':>5} {'Day':>4} {'Night':>6} {'Fri/Sat':>8}")
    for e in range(len(S.employees)):
        total = sum(1 for d in range(days) if grid[e][d] != OFF)
        day_n = sum(1 for d in range(days) if grid[e][d] == DAY)
        night_n = sum(1 for d in range(days) if grid[e][d] == NIGHT)
        wknd_n = sum(1 for d in weekend_day_indices(S) if grid[e][d] != OFF)
        tag = "  <- night team" if e in night_team_idx else ""
        print(f"  {S.employees[e]:<12} {total:>5} {day_n:>4} {night_n:>6} {wknd_n:>8}{tag}")

    print("\nRule validation:")
    all_pass = True
    for r in results:
        all_pass &= r.passed
        print(f"  [{'PASS' if r.passed else 'FAIL'}] {r.name:<55} | {r.measured}")
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
