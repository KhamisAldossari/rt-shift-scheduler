#!/usr/bin/env python3
"""
Streamlit front-end for the RT shift scheduler  (a 4-step wizard).

A THIN UI over scheduler.py. It only collects configuration into one
ScheduleSettings object and renders the engine's outputs -- it never
re-implements a constraint or recomputes a rule. Every number, count, spread,
and PASS/FAIL on screen comes straight from a scheduler function/field
(preflight -> build_and_solve -> validate -> export_bytes).

Layout -- one screen at a time, so a non-technical charge RT can build and
export a month's roster unaided:

  Persistent header : the app title + a custom stepper (1 Setup - 2 Check -
                      3 Roster - 4 Export) with completed / current / upcoming
                      states. Every step past the first has a quiet Back control
                      and exactly one accent primary action.

  Step 1 - Setup    : builds ScheduleSettings -- month, staff editor, night-team
                      picker, AM staff editor + AM working days, a leave editor
                      (per-person vacation ranges inside the month; the engine
                      blocks those days and prorates the shift target), and one
                      collapsed "Adjust rules" expander (staffing bands,
                      shifts/employee, run-length caps, the alternating-weekends
                      toggle, and a plain-language "Scheduling effort" level --
                      raising it scales the engine's deterministic search budget
                      AND its wall-clock safety net together, so the net never
                      binds first). Duplicate staff names are flagged inline and
                      block Continue -- the engine's preflight doesn't check for
                      those, so this screen does. Primary: Continue.
  Step 2 - Check    : runs preflight automatically. Problems render calmly as
                      message + suggestion; a clean setup shows a plain recap.
                      If a roster already exists for these exact settings (e.g.
                      Back landed here after a solve), the primary is "View
                      roster" -- no re-solve -- with "Generate again" alongside
                      it; otherwise the primary is "Generate roster", which
                      solves. No valid roster -> a plain explanation + the
                      engine's things-to-try (never the solver status). Solved ->
                      results stored, advance to step 3.
  Step 3 - Roster   : a 3-way status banner (hard fail / soft-fairness miss /
                      all clear), the color-coded grid as the visual hero with a
                      legend, the per-day coverage strip, the required-rule
                      PASS/FAIL table (Rule + Result; the raw Measured values sit
                      in a "Technical detail" expander), the four fairness goals
                      as cards (spread vs tolerance), and the per-employee
                      summary. On a hard-rule failure the primary is "Back to
                      setup" instead of Continue -- this roster must not reach
                      Export.
  Step 4 - Export   : a summary card + the one-click Excel download
                      (schedule_YYYY_MM.xlsx). Re-checks the stored result for a
                      hard-rule failure (defense in depth, in case this screen is
                      reached some other way) and shows the same do-not-publish
                      warning if so. A secondary "Start a new month" clears the
                      roster and advances the mirrored month by one (December
                      wraps to January of the next year); staff and rules survive.

Wizard state lives in st.session_state: `step` (1-4) and `setup` (a plain, non-
widget dict that mirrors every step-1 input so custom entries survive Back
navigation -- widget keys are garbage-collected when their step isn't on screen).
The staff, AM, and leave data_editor widgets are seeded from DataFrames that are
rebuilt only when (re)entering step 1, not on every rerun -- rebuilding a seed
from the mirror every rerun would change it after each edit and flip the
editor's widget identity, silently reverting the second of two consecutive edits.
A generated `result` snapshots the settings signature it was built from and is
dropped when the setup changes, so steps 3-4 never show a stale roster.

The grid colors come from scheduler.WEB_COLORS (mirrors the Excel fills) and the
chrome from .streamlit/config.toml, so the web view and the workbook agree. All
user-facing copy stays in plain operational language -- no solver/engine jargon
(OPTIMAL, FEASIBLE, CP-SAT, deterministic) reaches the screen; rule names shown
on screen get a cosmetic glyph substitution ("->", ">=", "<=" render as their
arrow/comparator symbols) purely for readability, without touching the
underlying RuleResult data.

Run:  ./.venv/bin/streamlit run app.py
"""

from __future__ import annotations

import calendar
import copy
import datetime

import pandas as pd
import streamlit as st

import scheduler as sch


# ---------------------------------------------------------------------------
# PAGE + LIGHT GLOBAL POLISH
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="RT Shift Scheduler",
    page_icon=":material/calendar_month:",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# A small amount of CSS, scoped to OUR OWN markup only (the app bar, the wizard
# stepper, kicker labels, and the grid legend) so we never fight Streamlit's
# internals across versions. Palette and typography come from
# .streamlit/config.toml; grid colors from WEB_COLORS.
st.markdown(
    """
    <style>
      .appbar{text-align:center;margin:.1rem 0 0}
      .brand-kicker{font-size:.72rem;letter-spacing:.18em;font-weight:600;
        color:#0E7569;text-transform:uppercase;margin-bottom:.15rem}
      .brand-title{font-size:1.55rem;font-weight:700;color:#1C2B2A;letter-spacing:-.01em}

      .stepper{display:flex;align-items:center;gap:.55rem;max-width:660px;
        margin:1rem auto 1.7rem;padding:0 .4rem}
      .stepper .step{display:inline-flex;align-items:center;gap:.5rem;white-space:nowrap}
      .stepper .dot{width:1.72rem;height:1.72rem;border-radius:50%;display:inline-flex;
        align-items:center;justify-content:center;font-weight:600;font-size:.82rem;
        border:1.5px solid #D9D6CE;color:#6E6F68;background:#FBFAF8;flex:none}
      .stepper .lbl{font-size:.86rem;font-weight:500;color:#6E6F68}
      .stepper .step.done .dot{background:#0E7569;border-color:#0E7569;color:#FBFAF8}
      .stepper .step.done .lbl{color:#3D4B49}
      .stepper .step.now .dot{border-color:#0E7569;color:#0E7569;background:#E4EFEC;
        border-width:2px}
      .stepper .step.now .lbl{color:#0E7569;font-weight:600}
      .stepper .bar{flex:1 1 auto;min-width:.8rem;height:2px;background:#E4E2DC;
        border-radius:2px}
      .stepper .bar.done{background:#0E7569}

      .kicker{font-size:.7rem;letter-spacing:.15em;font-weight:600;color:#6E6F68;
        text-transform:uppercase;margin:.25rem 0 .1rem}

      .legend{display:flex;flex-wrap:wrap;gap:.4rem 1.15rem;align-items:center;
        font-size:.85rem;color:#5C6360;margin:.15rem 0 .5rem}
      .legend-item{display:inline-flex;align-items:center;gap:.42rem}
      .legend-sw{width:.95rem;height:.95rem;border-radius:.28rem;
        border:1px solid rgba(28,43,42,.14);display:inline-block}
      .legend-code{font-weight:700;color:#1C2B2A}
      .legend-star{color:#0E7569;font-weight:700;font-size:1.05rem;line-height:1}
    </style>
    """,
    unsafe_allow_html=True,
)

# EVERY UI default is pulled from here -- no scheduling value is hard-coded.
defaults = sch.ScheduleSettings()

# Plain-language "Scheduling effort" levels. The base level is the engine's own
# default budget, so the out-of-box roster is bit-identical to the pristine
# engine; the higher levels simply spend more (deterministic) search time. This
# is an effort knob, not a scheduling rule -- the words solver/deterministic/
# CP-SAT never reach the user-facing copy.
#
# Both the deterministic limit AND its wall-clock safety net scale together by
# the same multiplier, so the safety net never creeps closer to the
# deterministic budget as effort rises (it must stay a net, not a binding
# limit, or the roster stops being reproducible -- see ScheduleSettings).
EFFORT_LEVELS = ["Standard", "Thorough", "Maximum"]
EFFORT_MULTIPLIER = {"Standard": 1, "Thorough": 2.5, "Maximum": 5}
EFFORT_DET_LIMIT = {lvl: defaults.solver_det_time_limit * m
                    for lvl, m in EFFORT_MULTIPLIER.items()}
EFFORT_TIME_LIMIT = {lvl: defaults.solver_time_limit * m
                     for lvl, m in EFFORT_MULTIPLIER.items()}


# ---------------------------------------------------------------------------
# WIZARD STATE  (step index + a plain mirror of every step-1 input)
# ---------------------------------------------------------------------------

def _init_setup() -> dict:
    """A plain (non-widget) dict seeded entirely from ScheduleSettings defaults.

    Widget keys are garbage-collected on reruns where their step isn't rendered,
    so step-1 inputs are mirrored here on every step-1 render and widgets re-seed
    from here on return -- this is what makes custom staff survive Back navigation.
    """
    return {
        "year": defaults.year,
        "month": defaults.month,
        "employees": list(defaults.employees),
        "night_team": list(defaults.night_team),
        "am_team": list(defaults.am_team),
        "am_days": list(defaults.am_days),
        "leave": list(defaults.leave),
        "day_min": defaults.day_min,
        "day_max": defaults.day_max,
        "night_min": defaults.night_min,
        "night_max": defaults.night_max,
        "shifts": defaults.shifts_per_employee,
        "max_day": defaults.max_consec_work,
        "max_night": defaults.max_consec_night,
        "min_work": defaults.min_consec_work,
        "min_off": defaults.min_consec_off,
        "max_off": defaults.max_consec_off,
        "alt_weekends": defaults.alternating_weekends,
        "effort": EFFORT_LEVELS[0],
    }


def _seed_editors() -> None:
    """(Re)build the three data_editor seed frames from the setup mirror.

    Called only when (re)entering step 1 -- NOT on every step-1 rerun -- so the
    same DataFrame instance backs each editor across reruns while the user
    stays on the screen. Rebuilding the seed from the mirror on every rerun
    (the mirror changes after each edit) flips the editor's widget identity,
    which silently reverts the second of two consecutive edits.
    """
    cfg = st.session_state.setup
    st.session_state["_emp_seed"] = pd.DataFrame(
        {"Employee": pd.Series(list(cfg["employees"]), dtype="object")})
    st.session_state["_am_seed"] = pd.DataFrame(
        {"AM staff": pd.Series(list(cfg["am_team"]), dtype="object")})
    st.session_state["_leave_seed"] = pd.DataFrame(
        {"Employee": pd.Series([r[0] for r in cfg["leave"]], dtype="object"),
         "From": pd.Series([r[1] for r in cfg["leave"]], dtype="datetime64[ns]"),
         "To": pd.Series([r[2] for r in cfg["leave"]], dtype="datetime64[ns]")})


if "step" not in st.session_state:
    st.session_state.step = 1
if "setup" not in st.session_state:
    st.session_state.setup = _init_setup()
if "_emp_seed" not in st.session_state:
    _seed_editors()


def go_to(step: int) -> None:
    """Switch screens. Leaving for Setup clears any transient solve outcome
    and re-seeds the staff/AM/leave editors from the (possibly just-updated)
    mirror."""
    st.session_state.step = step
    if step == 1:
        st.session_state.pop("gen_failed", None)
        _seed_editors()
    st.rerun()


def make_settings(cfg: dict) -> sch.ScheduleSettings:
    """Collect the setup dict into the single ScheduleSettings object.

    The UI is fixed-team only -- rotation_mode is left at its ScheduleSettings
    default (fixed_team); rotate mode stays engine/CLI-only. Fairness tuning
    (tolerances + objective weights) is intentionally not exposed either. No
    scheduling value is hard-coded here -- defaults all come from
    sch.ScheduleSettings(); the effort level maps to a deterministic budget.
    """
    return sch.ScheduleSettings(
        year=int(cfg["year"]), month=int(cfg["month"]),
        employees=cfg["employees"], night_team=cfg["night_team"],
        am_team=cfg["am_team"], am_days=tuple(cfg["am_days"]),
        leave=cfg["leave"],
        day_min=int(cfg["day_min"]), day_max=int(cfg["day_max"]),
        night_min=int(cfg["night_min"]), night_max=int(cfg["night_max"]),
        shifts_per_employee=int(cfg["shifts"]),
        max_consec_work=int(cfg["max_day"]), max_consec_night=int(cfg["max_night"]),
        min_consec_work=int(cfg["min_work"]), min_consec_off=int(cfg["min_off"]),
        max_consec_off=int(cfg["max_off"]),
        alternating_weekends=bool(cfg["alt_weekends"]),
        solver_det_time_limit=float(EFFORT_DET_LIMIT.get(cfg["effort"],
                                                         defaults.solver_det_time_limit)),
        solver_time_limit=float(EFFORT_TIME_LIMIT.get(cfg["effort"],
                                                      defaults.solver_time_limit)),
    )


def _sig(cfg: dict):
    """A value-comparable snapshot of the setup, used to invalidate a stale
    roster when the user edits settings after generating."""
    return copy.deepcopy(cfg)


# ---------------------------------------------------------------------------
# DISPLAY HELPERS  (display only -- never recompute a rule or a scheduling fact)
# ---------------------------------------------------------------------------

def grid_dataframe(S: sch.ScheduleSettings, grid: list[list[str]]) -> pd.DataFrame:
    """The roster as a DataFrame: rows = employees (+ any AM staff), columns = D1..Dn."""
    cols = [f"D{d + 1} {sch.weekday_of(S, d)}" for d in range(S.days)]
    rows = [S.employees[e] + (" ∗" if sch.is_night_member(S, e) else "")
            for e in range(len(S.employees))]
    data = [[sch.CELL_LABEL[grid[e][d]] for d in range(S.days)]
            for e in range(len(S.employees))]
    # AM staff live entirely outside the solver: their fixed-pattern rows come
    # straight from sch.am_rows() (never recomputed here) and are appended below.
    for a, arow in enumerate(sch.am_rows(S)):
        rows.append(S.am_team[a] + " (AM)")
        data.append([sch.CELL_LABEL[arow[d]] for d in range(S.days)])
    return pd.DataFrame(data, index=rows, columns=cols)


def color_cell(val: str) -> str:
    """Mirror the openpyxl fills so the web grid matches the Excel exactly."""
    code = {"D": sch.DAY, "N": sch.NIGHT, "V": sch.LEAVE, "": sch.OFF,
            "AM": sch.AM}.get(val)
    bg = sch.WEB_COLORS.get(code, "#FFFFFF")
    return f"background-color: {bg}; text-align: center; color: #1C2B2A; font-weight: 600;"


def color_result(val: str) -> str:
    """PASS/FAIL cell color, reusing the grid's green/red so the views agree.
    FAIL uses the LEAVE red: OFF is now a blank white cell, so borrowing it
    would silently blank every FAIL badge."""
    if val == "PASS":
        return f"background-color: {sch.WEB_COLORS[sch.DAY]}; color: #1C2B2A; font-weight: 600;"
    if val == "FAIL":
        return f"background-color: {sch.WEB_COLORS[sch.LEAVE]}; color: #1C2B2A; font-weight: 600;"
    return ""


def legend(show_am: bool = False, show_leave: bool = False) -> None:
    """A compact key whose swatches are the EXACT grid colors (WEB_COLORS)."""
    items = [(sch.WEB_COLORS[sch.DAY], "D", "Day"),
             (sch.WEB_COLORS[sch.NIGHT], "N", "Night"),
             (sch.WEB_COLORS[sch.OFF], "", "Off")]
    if show_leave:
        items.append((sch.WEB_COLORS[sch.LEAVE], "V", "Leave"))
    if show_am:
        items.append((sch.WEB_COLORS[sch.AM], "AM", "AM"))
    html = ['<div class="legend">']
    for bg, code, label in items:
        html.append(
            f'<span class="legend-item"><span class="legend-sw" '
            f'style="background:{bg}"></span><span class="legend-code">{code}</span>'
            f'&nbsp;{label}</span>')
    html.append('<span class="legend-item"><span class="legend-star">∗</span>'
                '&nbsp;Night-team member</span>')
    html.append('</div>')
    st.markdown("".join(html), unsafe_allow_html=True)


def render_grid(S: sch.ScheduleSettings, grid: list[list[str]]) -> None:
    """The hero: the full color-coded roster. Comfortable rows, all employees in
    view (it scrolls only for very large teams / long months)."""
    gdf = grid_dataframe(S, grid)
    row_h = 39
    height = min(48 + row_h * len(gdf), 660)
    st.dataframe(gdf.style.map(color_cell), width="stretch",
                 height=height, row_height=row_h)


def render_coverage(S: sch.ScheduleSettings, grid: list[list[str]]) -> None:
    """Per-day coverage strip from coverage_per_day(): a compact day/night table
    that rhymes with the roster's columns. Weekend (Fri/Sat) columns are washed
    amber and forced-overlap nights (over_floor) are emphasized -- all read
    straight from the engine's per-day facts, nothing recomputed here."""
    cov = sch.coverage_per_day(S, grid)
    cols = [c["label"] for c in cov]
    is_wknd = [c["is_weekend"] for c in cov]
    over = [c["over_floor"] for c in cov]
    df = pd.DataFrame(
        [[c["day"] for c in cov], [c["night"] for c in cov]],
        index=["Day staff", "Night staff"], columns=cols)

    def _styles(_df: pd.DataFrame) -> pd.DataFrame:
        css = pd.DataFrame("", index=_df.index, columns=_df.columns)
        for j in range(len(_df.columns)):
            wk, ov = is_wknd[j], over[j]
            # Row 0 -- day staff: soft warm green, amber wash on weekends.
            css.iloc[0, j] = (
                f"background-color: {'#F3EDDA' if wk else '#EAF1EC'}; "
                "text-align: center; color: #26332C; font-weight: 500;")
            # Row 1 -- night staff: soft teal; deeper teal + bold on overlap days.
            if ov:
                nbg, ncol, nwt = "#CFE1DD", "#1C4A44", 700
            elif wk:
                nbg, ncol, nwt = "#F3EDDA", "#2B3B39", 500
            else:
                nbg, ncol, nwt = "#E7EFEE", "#2B3B39", 500
            css.iloc[1, j] = (f"background-color: {nbg}; text-align: center; "
                              f"color: {ncol}; font-weight: {nwt};")
        return css

    st.dataframe(df.style.apply(_styles, axis=None), width="stretch",
                 height=36 + 36 * len(df), row_height=36)
    st.caption("Staff on duty each day. Amber columns are the Fri/Sat weekend; a "
               "deeper-teal night cell is a forced extra (overlap) night above the floor.")


def _pretty_rule_name(name: str) -> str:
    """Cosmetic glyph substitution for display only -- operates on a COPY of
    the string; the underlying RuleResult.name is never touched or re-derived."""
    return name.replace("->", "→").replace(">=", "≥").replace("<=", "≤")


def render_checks(rule_results, full: bool = True) -> None:
    """A scannable PASS/FAIL table for a list of RuleResults (display only).

    `full=False` drops the Measured column for a quick-scan view -- the same
    numbers stay reachable in a paired "Technical detail" expander.
    """
    data = {
        "Rule": [_pretty_rule_name(r.name) for r in rule_results],
        "Result": ["PASS" if r.passed else "FAIL" for r in rule_results],
    }
    column_config = {
        "Rule": st.column_config.TextColumn(width="large"),
        "Result": st.column_config.TextColumn(width="small"),
    }
    if full:
        data["Measured"] = [r.measured for r in rule_results]
        column_config["Measured"] = st.column_config.TextColumn(width="medium")
    checks = pd.DataFrame(data)
    st.dataframe(
        checks.style.map(color_result, subset=["Result"]),
        width="stretch", hide_index=True,
        height=min(80 + 35 * len(checks), 600),
        column_config=column_config)


def status_banner(hard, fairness) -> None:
    """Positive, next-step status. Preserves the prior three-way logic exactly:
    hard fail -> error; all hard pass but a soft goal misses -> info; all -> success."""
    hard_fail = [r for r in hard if not r.passed]
    soft_fail = [r for r in fairness if not r.passed]
    if not hard_fail and not soft_fail:
        st.success("All rules pass — this roster is ready to publish.")
    elif not hard_fail:
        # Fairness goals are soft: a miss means the roster is still valid (every
        # hard rule holds), just not perfectly balanced this month.
        st.info(
            f"All required safety rules pass — this roster is valid and ready to "
            f"publish. {len(soft_fail)} fairness goal(s) couldn't be fully balanced "
            f"this month (see the fairness goals below). Some months can't be "
            f"balanced perfectly. Raising the scheduling effort in Setup and "
            f"generating again sometimes improves the balance.")
    else:
        st.error(
            f"{len(hard_fail)} required rule(s) did not pass — do not publish. See the "
            f"failing rule in the required-rules table below.")


def short_goal(name: str) -> str:
    """'Fairness - balanced night load (spread <= 1)' -> 'Balanced Night Load'."""
    return name.replace("Fairness - ", "").split(" (")[0].title()


def render_problems(problems) -> None:
    """Surface each preflight Problem calmly: its message (what's wrong) and its
    suggestion (what to do), one bordered card each -- no wall of red."""
    st.warning("These settings need attention before the month can be scheduled.")
    for p in problems:
        with st.container(border=True):
            st.markdown(f"**{p.message}**")
            st.caption(f"Try: {p.suggestion}")


def fairness_cards(fairness) -> None:
    """The four soft goals as PASS/FAIL cards with spread vs tolerance."""
    cards = st.columns(len(fairness))
    for col, r in zip(cards, fairness):
        with col.container(border=True):
            st.markdown(f"**{short_goal(r.name)}**")
            if r.passed:
                st.badge("Pass", color="green", icon=":material/check_circle:")
            else:
                st.badge("Fail", color="red", icon=":material/cancel:")
            st.caption(f"spread {r.spread} · tolerance {r.tolerance}")
            # A goal can sit within tolerance yet still FAIL on a non-spread
            # condition (e.g. the weekends goal also wants a full Fri+Sat off for
            # everyone). Say so, so the card never looks self-contradictory; the
            # full reason is in the measured details below.
            if (not r.passed and r.spread is not None
                    and r.tolerance is not None and r.spread <= r.tolerance):
                st.caption("Within tolerance, but another part of this goal isn't "
                           "fully met this month — see the technical detail below.")
    with st.expander("Technical detail"):
        render_checks(fairness)


def employee_table(S: sch.ScheduleSettings, grid: list[list[str]]) -> None:
    """The per-employee fairness summary, straight from employee_loads()."""
    loads = sch.employee_loads(S, grid)
    df = pd.DataFrame([{
        "Employee": ld["name"] + (" ∗" if ld["night_member"] else ""),
        "Total": ld["total"], "Leave": ld["leave"], "Target": ld["target"],
        "Day": ld["day"], "Night": ld["night"],
        "Fri/Sat": ld["weekend"], "Weekends off": ld["weekends_off"],
        "Overlaps": ld["overlaps"], "Max-runs": ld["max_runs"],
    } for ld in loads])
    st.dataframe(df, width="stretch", hide_index=True,
                 height=min(80 + 35 * len(df), 460))
    st.caption("∗ night-team member. Overlaps = forced extra-night days worked · "
               "Max-runs = max-length day/night streaks absorbed · both are balanced "
               "within each pool (day-capable vs night-eligible).")


def recap(S: sch.ScheduleSettings) -> None:
    """A short, plain-language summary of what Generate will produce."""
    st.markdown(f"**{S.month_label}** · {S.days} days")
    st.markdown(f"- **{len(S.employees)}** staff scheduled")
    st.markdown(f"- Night team (nights only): {', '.join(S.night_team) or '(none selected)'}")
    if S.am_team:
        st.markdown(f"- AM shift: {', '.join(S.am_team)} on {', '.join(S.am_days)}")
    if S.leave:
        st.markdown(f"- Leave: {len(S.leave)} range(s) entered — those days are "
                    f"blocked and each person's shift target drops to match")
    st.markdown(f"- Each day: {S.day_min}–{S.day_max} day staff, "
                f"{S.night_min}–{S.night_max} night staff")
    st.markdown(f"- {S.shifts_per_employee} shifts per person; at most "
                f"{S.max_consec_work} day / {S.max_consec_night} night shifts in a row")
    if S.alternating_weekends:
        st.markdown("- Alternating weekends: on")


# ---------------------------------------------------------------------------
# PERSISTENT HEADER + STEPPER
# ---------------------------------------------------------------------------

def render_header(step: int) -> None:
    st.markdown(
        '<div class="appbar">'
        '<div class="brand-kicker">Respiratory Therapy Workforce</div>'
        '<div class="brand-title">Monthly Shift Scheduler</div>'
        '</div>',
        unsafe_allow_html=True)
    labels = ["Setup", "Check", "Roster", "Export"]
    html = ['<div class="stepper">']
    for i, label in enumerate(labels, start=1):
        cls = "step done" if i < step else "step now" if i == step else "step"
        dot = "&#10003;" if i < step else str(i)          # check-mark for completed
        html.append(f'<div class="{cls}"><span class="dot">{dot}</span>'
                    f'<span class="lbl">{label}</span></div>')
        if i < len(labels):
            html.append(f'<div class="{"bar done" if i < step else "bar"}"></div>')
    html.append('</div>')
    st.markdown("".join(html), unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# STEP 1 -- SETUP  (builds ScheduleSettings; mirrors inputs into session_state)
# ---------------------------------------------------------------------------

def screen_setup() -> None:
    cfg = st.session_state.setup
    _, mid, _ = st.columns([1, 2.2, 1])
    with mid:
        st.subheader("Set up the month")
        st.caption("Choose the month, list the staff, and pick the night team. "
                   "The rule settings are pre-filled — open Adjust rules only if "
                   "your department's rules differ.")

        # -- Month --------------------------------------------------------
        st.markdown('<div class="kicker">Month</div>', unsafe_allow_html=True)
        cy, cm = st.columns(2)
        year = cy.number_input("Year", 2000, 2100, int(cfg["year"]), step=1, key="w_year")
        month = cm.selectbox("Month", options=list(range(1, 13)),
                             index=int(cfg["month"]) - 1,
                             format_func=lambda m: calendar.month_name[m], key="w_month")

        # -- Staff --------------------------------------------------------
        st.markdown('<div class="kicker">Staff</div>', unsafe_allow_html=True)
        st.caption("Add, remove, or rename employees — one per row.")
        emp_df = st.data_editor(
            st.session_state["_emp_seed"],
            num_rows="dynamic", width="stretch", hide_index=True, key="w_emp",
            column_config={"Employee": st.column_config.TextColumn(
                "Employee", width="large", help="One row per person.")})
        employees = [str(x).strip() for x in emp_df["Employee"].dropna().tolist()
                     if str(x).strip()]

        # Duplicate names crash the roster grid downstream (a non-unique index)
        # and could produce a nonsense roster -- input hygiene, not a scheduling
        # rule, so it's caught here rather than in the engine's preflight. Never
        # silently deduped: the user must fix the list.
        seen, dupe_names = set(), []
        for name in employees:
            if name in seen and name not in dupe_names:
                dupe_names.append(name)
            seen.add(name)
        if dupe_names:
            word = "name" if len(dupe_names) == 1 else "names"
            st.error(f"Two rows have the same {word}: {', '.join(dupe_names)}. "
                     "Give each person a unique name.")

        # -- Night team ---------------------------------------------------
        st.markdown('<div class="kicker">Night team</div>', unsafe_allow_html=True)
        night_team = st.multiselect(
            "Night-team members — they work nights only",
            options=employees,
            default=[e for e in cfg["night_team"] if e in employees],
            key="w_night")
        if not night_team:
            st.caption("Pick at least one night-team member — they work nights only.")
        else:
            st.caption(f"{len(night_team)} of {len(employees)} staff on the night team.")

        # -- AM shift -----------------------------------------------------
        st.markdown('<div class="kicker">AM shift</div>', unsafe_allow_html=True)
        st.caption("Staff who work the same weekdays every week. They appear on the "
                   "roster as extra rows and don't count toward the daily staffing "
                   "numbers. Leave empty if unused.")
        am_df = st.data_editor(
            st.session_state["_am_seed"],
            num_rows="dynamic", width="stretch", hide_index=True, key="w_am",
            column_config={"AM staff": st.column_config.TextColumn(
                "AM staff", width="large", help="One row per AM-shift person.")})
        am_team = [str(x).strip() for x in am_df["AM staff"].dropna().tolist()
                   if str(x).strip()]
        am_days = st.multiselect(
            "AM working days", options=sch.WEEKDAY_NAMES,
            default=[d for d in cfg["am_days"] if d in sch.WEEKDAY_NAMES], key="w_amdays",
            help="The weekdays AM staff work every week (default Sun–Thu).")

        # -- Leave ----------------------------------------------------------
        st.markdown('<div class="kicker">Leave</div>', unsafe_allow_html=True)
        st.caption("Vacation inside the month: those days are blocked ('V') and "
                   "the person's shift target drops by one per leave day. One row "
                   "per range; several rows per person are fine. Rows missing a "
                   "name or a date are ignored.")
        month_first = datetime.date(int(year), int(month), 1)
        month_last = datetime.date(int(year), int(month),
                                   calendar.monthrange(int(year), int(month))[1])
        leave_df = st.data_editor(
            st.session_state["_leave_seed"],
            num_rows="dynamic", width="stretch", hide_index=True, key="w_leave",
            column_config={
                "Employee": st.column_config.SelectboxColumn(
                    "Employee", options=employees, width="medium",
                    help="Scheduled staff only — AM staff cannot take rostered leave."),
                "From": st.column_config.DateColumn(
                    "From", min_value=month_first, max_value=month_last),
                "To": st.column_config.DateColumn(
                    "To", min_value=month_first, max_value=month_last),
            })
        # Collect complete rows only (name + both dates); every rule-level check
        # -- unknown names, reversed or out-of-month ranges -- stays in preflight.
        leave = []
        for _, lrow in leave_df.iterrows():
            lname, d_from, d_to = lrow.get("Employee"), lrow.get("From"), lrow.get("To")
            if pd.isna(lname) or not str(lname).strip() or pd.isna(d_from) or pd.isna(d_to):
                continue
            leave.append((str(lname).strip(),
                          pd.Timestamp(d_from).date(), pd.Timestamp(d_to).date()))

        # -- Adjust rules (collapsed; pre-filled from ScheduleSettings) ----
        with st.expander("Adjust rules"):
            st.caption("Pre-filled with the department defaults. Change these only if "
                       "your rules differ.")

            st.markdown("**Staffing per day**")
            d1, d2 = st.columns(2)
            day_min = d1.number_input("Min day staff", 0, 50, int(cfg["day_min"]),
                                      key="w_daymin")
            day_max = d2.number_input("Max day staff", 1, 50, int(cfg["day_max"]),
                                      key="w_daymax")
            n1, n2 = st.columns(2)
            night_min = n1.number_input("Min night staff", 0, 10, int(cfg["night_min"]),
                                        key="w_nmin")
            night_max = n2.number_input("Max night staff (overlap)", 1, 10,
                                        int(cfg["night_max"]), key="w_nmax")

            st.markdown("**Workload**")
            shifts = st.number_input("Shifts per employee (exact)", 1, 31,
                                     int(cfg["shifts"]), key="w_shifts")

            st.markdown("**Run lengths**")
            r1, r2 = st.columns(2)
            max_day = r1.number_input("Max day shifts in a row", 1, 14,
                                      int(cfg["max_day"]), key="w_maxday")
            max_night = r2.number_input("Max nights in a row", 1, 14,
                                        int(cfg["max_night"]), key="w_maxnight")
            r3, r4 = st.columns(2)
            min_work = r3.number_input(
                "Min work days in a row", 2, 7, int(cfg["min_work"]), disabled=True,
                key="w_minwork",
                help="Fixed at 2 — the schedule never leaves a single isolated work day.")
            min_off = r4.number_input(
                "Min days off in a row", 2, 7, int(cfg["min_off"]), disabled=True,
                key="w_minoff",
                help="Fixed at 2 — the schedule never leaves a single isolated day off.")
            max_off = st.number_input("Max days off in a row", 1, 14,
                                      int(cfg["max_off"]), key="w_maxoff")

            st.markdown("**Weekends**")
            alt_weekends = st.toggle(
                "Alternating weekends", value=bool(cfg["alt_weekends"]), key="w_alt",
                help="On: every employee's full Fri+Sat weekends strictly alternate "
                     "off / on / off — of every two full weekends, exactly one is fully "
                     "off. This is a firm rule, and it can make very tight months "
                     "impossible to schedule; if that happens, turn it off and generate "
                     "again.")

            st.markdown("**Scheduling effort**")
            effort = st.select_slider(
                "Scheduling effort", options=EFFORT_LEVELS,
                value=cfg["effort"] if cfg["effort"] in EFFORT_LEVELS else EFFORT_LEVELS[0],
                key="w_effort", label_visibility="collapsed",
                help="Higher effort searches longer for a better-balanced roster. With "
                     "the same settings you always get the same roster.")

        # Mirror every input into plain (non-widget) session_state so it survives
        # Back navigation, when this step's widget keys are garbage-collected.
        st.session_state.setup = {
            "year": int(year), "month": int(month),
            "employees": employees, "night_team": night_team,
            "am_team": am_team, "am_days": list(am_days),
            "leave": leave,
            "day_min": int(day_min), "day_max": int(day_max),
            "night_min": int(night_min), "night_max": int(night_max),
            "shifts": int(shifts),
            "max_day": int(max_day), "max_night": int(max_night),
            "min_work": int(min_work), "min_off": int(min_off),
            "max_off": int(max_off),
            "alt_weekends": bool(alt_weekends),
            "effort": effort,
        }

        st.divider()
        _, col_next = st.columns([2, 1])
        with col_next:
            if st.button("Continue to check", type="primary", width="stretch",
                         icon=":material/arrow_forward:", key="nav_1_next",
                         disabled=bool(dupe_names)):
                # Any change to the setup invalidates a previously generated roster.
                if ("result" in st.session_state
                        and st.session_state["result"]["sig"] != _sig(st.session_state.setup)):
                    st.session_state.pop("result", None)
                st.session_state.pop("gen_failed", None)
                go_to(2)


# ---------------------------------------------------------------------------
# STEP 2 -- CHECK  (preflight runs automatically; Generate solves)
# ---------------------------------------------------------------------------

def screen_check() -> None:
    cfg = st.session_state.setup
    S = make_settings(cfg)
    problems = sch.preflight(S)

    _, mid, _ = st.columns([1, 2.2, 1])
    with mid:
        if st.button("Back", icon=":material/arrow_back:", key="nav_2_back"):
            go_to(1)
        st.subheader("Check the setup")

        # (a) Preflight found problems -- show each message + suggestion, calmly.
        if problems:
            render_problems(problems)
            st.divider()
            if st.button("Back to setup", type="primary", width="stretch",
                         icon=":material/arrow_back:", key="nav_2_fix"):
                go_to(1)
            return

        # (b) A previous Generate on these settings found no valid roster.
        if "gen_failed" in st.session_state:
            st.error("No roster fits these rules together.")
            st.markdown(
                "Each rule is fine on its own, but together they leave no valid "
                "schedule for this month — it's the combination, not a single wrong "
                "value. Things to try:")
            for hint in st.session_state["gen_failed"]["hints"]:
                st.markdown(f"- {hint}")
            st.divider()
            if st.button("Back to setup", type="primary", width="stretch",
                         icon=":material/arrow_back:", key="nav_2_infeasible"):
                go_to(1)
            return

        # (c) Clean -- a stored roster for these exact settings means Back
        # landed here after a Generate; offer to view it instead of stranding
        # it behind a full re-solve. Otherwise, recap what Generate will build.
        has_current_result = ("result" in st.session_state
                              and st.session_state["result"]["sig"] == _sig(cfg))
        if has_current_result:
            st.success("This month's roster is ready. View it, or generate again "
                       "if you'd like to reshuffle within the same settings.")
        else:
            st.success("Everything checks out. Here's what will be generated.")
        with st.container(border=True):
            recap(S)
        st.divider()

        if has_current_result:
            col_view, col_regen = st.columns([2, 1])
            with col_view:
                if st.button("View roster", type="primary", width="stretch",
                             icon=":material/table_chart:", key="nav_2_view"):
                    go_to(3)
            with col_regen:
                generate_clicked = st.button(
                    "Generate again", width="stretch",
                    icon=":material/refresh:", key="nav_2_generate")
        else:
            generate_clicked = st.button(
                "Generate roster", type="primary", width="stretch",
                icon=":material/bolt:", key="nav_2_generate")

        if generate_clicked:
            with st.spinner("Building the roster…"):
                sol = sch.build_and_solve(S)
            if not sol.grid:
                st.session_state["gen_failed"] = {"hints": sch.relaxation_hints(S)}
                st.rerun()
            else:
                results = sch.validate(S, sol.grid)
                st.session_state["result"] = {
                    "settings": S, "grid": sol.grid, "results": results,
                    "xlsx": sch.export_bytes(S, sol.grid, results),
                    "sig": _sig(cfg),
                }
                st.session_state.pop("gen_failed", None)
                go_to(3)


# ---------------------------------------------------------------------------
# STEP 3 -- ROSTER  (renders validate() results + the grid)
# ---------------------------------------------------------------------------

def screen_roster() -> None:
    if "result" not in st.session_state:
        _, mid, _ = st.columns([1, 2.2, 1])
        with mid:
            st.info("No roster has been generated yet.")
            if st.button("Back to check", key="nav_3_empty"):
                go_to(2)
        return

    R = st.session_state["result"]
    S = R["settings"]
    results = R["results"]
    # Split hard vs fairness via the structured `kind` field (not by name).
    fairness = [r for r in results if r.kind == "fairness"]
    hard = [r for r in results if r.kind != "fairness"]
    hard_fail = [r for r in hard if not r.passed]

    _, mid, _ = st.columns([0.5, 11, 0.5])          # wide, so the grid is the hero
    with mid:
        if st.button("Back", icon=":material/arrow_back:", key="nav_3_back"):
            go_to(2)
        st.markdown('<div class="kicker">Monthly roster</div>', unsafe_allow_html=True)
        st.subheader(f"{S.month_label} · {S.days} days")
        st.caption(sch.roster_caption(S))

        # -- 3-way status banner (exact prior logic) --------------------
        status_banner(hard, fairness)

        # -- The color grid: the visual hero, with the legend -----------
        st.markdown("##### Roster grid")
        legend(show_am=bool(S.am_team), show_leave=bool(S.leave))
        render_grid(S, R["grid"])

        st.divider()
        st.markdown("##### Daily coverage")
        render_coverage(S, R["grid"])

        st.divider()
        st.markdown("##### Required rules")
        st.caption("Every required rule re-checked independently from the finished "
                   "grid. These are non-negotiable — all must pass to publish.")
        render_checks(hard, full=False)
        with st.expander("Technical detail"):
            render_checks(hard)

        st.divider()
        st.markdown("##### Fairness goals")
        st.caption("Four goals, weighted equally. Each passes when its measured spread "
                   "(max − min) is within tolerance; the weekends goal also needs every "
                   "person to get at least one full Fri–Sat weekend off.")
        fairness_cards(fairness)

        st.divider()
        st.markdown("##### Per-employee summary")
        employee_table(S, R["grid"])

        st.divider()
        _, col_next = st.columns([2, 1])
        with col_next:
            if hard_fail:
                if st.button("Back to setup", type="primary", width="stretch",
                             icon=":material/arrow_back:", key="nav_3_fix"):
                    go_to(1)
            else:
                if st.button("Continue to export", type="primary", width="stretch",
                             icon=":material/arrow_forward:", key="nav_3_next"):
                    go_to(4)


# ---------------------------------------------------------------------------
# STEP 4 -- EXPORT  (summary card + the Excel download)
# ---------------------------------------------------------------------------

def screen_export() -> None:
    if "result" not in st.session_state:
        _, mid, _ = st.columns([1, 2.2, 1])
        with mid:
            st.info("No roster has been generated yet.")
            if st.button("Back to check", key="nav_4_empty"):
                go_to(2)
        return

    R = st.session_state["result"]
    S = R["settings"]
    results = R["results"]
    passed = sum(1 for r in results if r.passed)
    # Defense in depth: the Roster step already blocks forward navigation on a
    # hard-rule failure, but re-check here too in case this screen is reached
    # some other way -- reading the stored kind/passed fields is display
    # aggregation, not re-deriving a rule.
    hard_fail = [r for r in results if r.kind != "fairness" and not r.passed]

    _, mid, _ = st.columns([1, 2.2, 1])
    with mid:
        if st.button("Back", icon=":material/arrow_back:", key="nav_4_back"):
            go_to(3)
        st.subheader("Export the roster")
        st.caption("Download the color-coded Excel workbook — the roster sheet plus a "
                   "summary and rule-check sheet, matching what's on screen.")

        if hard_fail:
            st.error(
                f"{len(hard_fail)} required rule(s) did not pass — do not publish "
                f"this roster. Go back to the roster to review, or return to Setup "
                f"to adjust and generate again.")

        with st.container(border=True):
            m1, m2, m3 = st.columns(3)
            m1.metric("Month", S.month_label)
            m2.metric("Staff", len(S.employees))
            m3.metric("Checks passed", f"{passed}/{len(results)}")

        st.write("")
        st.download_button(
            "Download Excel (.xlsx)", data=R["xlsx"],
            file_name=f"schedule_{S.year}_{S.month:02d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", icon=":material/download:", width="stretch", key="dl_xlsx")

        st.divider()
        if st.button("Start a new month", icon=":material/restart_alt:", key="nav_4_new"):
            # Advance the mirrored month by one (December wraps to January of
            # the next year) so Setup doesn't reopen on the month just
            # exported. Everything else (staff, rules) intentionally survives.
            next_cfg = dict(st.session_state.setup)
            if next_cfg["month"] == 12:
                next_cfg["month"], next_cfg["year"] = 1, next_cfg["year"] + 1
            else:
                next_cfg["month"] += 1
            st.session_state.setup = next_cfg
            st.session_state.pop("result", None)
            st.session_state.pop("gen_failed", None)
            go_to(1)


# ---------------------------------------------------------------------------
# ROUTER
# ---------------------------------------------------------------------------

render_header(st.session_state.step)

if st.session_state.step == 1:
    screen_setup()
elif st.session_state.step == 2:
    screen_check()
elif st.session_state.step == 3:
    screen_roster()
elif st.session_state.step == 4:
    screen_export()
