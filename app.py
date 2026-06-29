#!/usr/bin/env python3
"""
Streamlit front-end for the RT shift scheduler.

Drives `scheduler.py` entirely through a single `ScheduleSettings` object: pick a
month + year, edit the staff, choose the night-coverage model, set every rule and
fairness value, click Generate, then review the roster and the independent
PASS/FAIL checks and download the color-coded Excel. The solver logic lives in
`scheduler.py` and is reused as-is -- this file never re-implements a constraint;
it only collects inputs (config) and renders outputs
(preflight -> solve -> validate -> export).

Run:  ./.venv/bin/streamlit run app.py
"""

from __future__ import annotations

import calendar

import pandas as pd
import streamlit as st

import scheduler as sch

st.set_page_config(page_title="RT Shift Scheduler", layout="wide")
st.title("Respiratory Therapy — Monthly Shift Scheduler")
st.caption("Configure any month and rule set, generate a fair roster, and download the Excel. "
           "Every hard rule is enforced by the CP-SAT solver and re-checked by an independent "
           "validator; fairness goals are reported with their measured spread.")

defaults = sch.ScheduleSettings()


# ---------------------------------------------------------------------------
# SIDEBAR — every input that populates ScheduleSettings (config only)
# ---------------------------------------------------------------------------

with st.sidebar:
    st.header("1 · Month")
    col_y, col_m = st.columns(2)
    year = col_y.number_input("Year", min_value=2000, max_value=2100,
                              value=defaults.year, step=1)
    month = col_m.selectbox("Month", options=list(range(1, 13)),
                            index=defaults.month - 1,
                            format_func=lambda m: calendar.month_name[m])

    st.header("2 · Staff")
    st.caption("Add, remove, or rename employees.")
    emp_df = st.data_editor(
        pd.DataFrame({"Employee": list(defaults.employees)}),
        num_rows="dynamic", width="stretch", hide_index=True, key="emp_editor")
    employees = [str(x).strip() for x in emp_df["Employee"].tolist() if str(x).strip()]

    st.header("3 · Night coverage")
    mode_label = st.radio(
        "How are nights covered?",
        ["Fixed night team (nights only)", "Everyone rotates nights"],
        help="Fixed: a dedicated team of any size works nights only and is the only "
             "group eligible for nights. Rotate: all staff are night-eligible and "
             "nights are spread across the whole roster.")
    rotation_mode = sch.FIXED_TEAM if mode_label.startswith("Fixed") else sch.ROTATE

    if rotation_mode == sch.FIXED_TEAM:
        night_team = st.multiselect(
            "Night team (any size ≥ 1 — they work nights only)",
            options=employees,
            default=[e for e in defaults.night_team if e in employees])
        if not night_team:
            st.warning("Pick at least one night-team member.")
        else:
            st.caption(f"{len(night_team)} night-team member(s) selected.")
    else:
        night_team = []
        st.caption("All staff are night-eligible; nights rotate across everyone.")

    st.header("4 · Staffing per day")
    c1, c2 = st.columns(2)
    day_min = c1.number_input("Min day staff", 0, 50, defaults.day_min)
    day_max = c2.number_input("Max day staff", 1, 50, defaults.day_max)
    c3, c4 = st.columns(2)
    night_min = c3.number_input("Min nights/day", 0, 10, defaults.night_min)
    night_max = c4.number_input("Max nights/day (overlap)", 1, 10, defaults.night_max)

    st.header("5 · Rules")
    shifts = st.number_input("Shifts per employee (exact)", 1, 31, defaults.shifts_per_employee)
    c5, c6 = st.columns(2)
    max_day = c5.number_input("Max consecutive day shifts", 1, 14, defaults.max_consec_work)
    max_night = c6.number_input("Max consecutive nights", 1, 14, defaults.max_consec_night)
    c7, c8 = st.columns(2)
    min_work = c7.number_input("Min consecutive work", 2, 7, defaults.min_consec_work)
    min_off = c8.number_input("Min consecutive off", 2, 7, defaults.min_consec_off)
    max_off = st.number_input("Max consecutive off", 1, 14, defaults.max_consec_off)

    st.header("6 · Fairness")
    st.caption("All four fairness goals are soft. A goal PASSES when its spread "
               "(max − min) is within tolerance.")
    c9, c10 = st.columns(2)
    tol_nw = c9.number_input("Night/weekend tolerance", 0, 10, defaults.fair_tol_night)
    tol_runs = c10.number_input("Runs tolerance", 0, 10, defaults.fair_tol_runs)
    with st.expander("Objective weights (equal ⇒ the goals matter equally)"):
        w_total = st.number_input("Equal total shifts", 0, 1000, defaults.w_fair_total, step=10)
        w_night = st.number_input("Balanced night load", 0, 1000, defaults.w_fair_night, step=10)
        w_weekend = st.number_input("Balanced weekends", 0, 1000, defaults.w_fair_weekend, step=10)
        w_runs = st.number_input("Balanced undesirable runs", 0, 1000, defaults.w_fair_runs, step=10)

    with st.expander("Advanced · solver"):
        det_budget = st.number_input(
            "Solver budget (deterministic units)", 4.0, 240.0,
            float(defaults.solver_det_time_limit), step=4.0,
            help="Higher = more search = better fairness convergence on hard months "
                 "(especially rotate mode), at the cost of solve time. The budget is "
                 "deterministic, so the roster is reproducible regardless of machine speed.")

    generate = st.button("Generate roster", type="primary", width="stretch")


def make_settings() -> sch.ScheduleSettings:
    """Collect every sidebar input into the single ScheduleSettings object."""
    return sch.ScheduleSettings(
        year=int(year), month=int(month),
        employees=employees, rotation_mode=rotation_mode, night_team=night_team,
        day_min=int(day_min), day_max=int(day_max),
        night_min=int(night_min), night_max=int(night_max),
        shifts_per_employee=int(shifts),
        max_consec_work=int(max_day), max_consec_night=int(max_night),
        min_consec_work=int(min_work), min_consec_off=int(min_off),
        max_consec_off=int(max_off),
        fair_tol_night=int(tol_nw), fair_tol_weekend=int(tol_nw),
        fair_tol_runs=int(tol_runs),
        w_fair_total=int(w_total), w_fair_night=int(w_night),
        w_fair_weekend=int(w_weekend), w_fair_runs=int(w_runs),
        solver_det_time_limit=float(det_budget),
    )


# ---------------------------------------------------------------------------
# RENDER HELPERS  (display only — never recompute a rule)
# ---------------------------------------------------------------------------

def grid_dataframe(S: sch.ScheduleSettings, grid: list[list[str]]) -> pd.DataFrame:
    cols = [f"D{d + 1} {sch.weekday_of(S, d)}" for d in range(S.days)]
    rows = [S.employees[e] + (" *" if sch.is_night_member(S, e) else "")
            for e in range(len(S.employees))]
    data = [[sch.CELL_LABEL[grid[e][d]] for d in range(S.days)]
            for e in range(len(S.employees))]
    return pd.DataFrame(data, index=rows, columns=cols)


def color_cell(val: str) -> str:
    """Mirror the openpyxl fills so the web grid matches the Excel exactly."""
    code = {"D": sch.DAY, "N": sch.NIGHT, "OFF": sch.OFF}.get(val)
    bg = sch.WEB_COLORS.get(code, "#FFFFFF")
    return f"background-color: {bg}; text-align: center; color: #111;"


def color_result(val: str) -> str:
    if val == "PASS":
        return f"background-color: {sch.WEB_COLORS[sch.DAY]}; color: #111;"
    if val == "FAIL":
        return f"background-color: {sch.WEB_COLORS[sch.OFF]}; color: #111;"
    return ""


def show_problems(title: str, problems) -> None:
    """Surface each preflight Problem's message AND its suggestion."""
    st.error(f"**{title}**")
    for p in problems:
        st.markdown(f"- {p.message}  \n  💡 *{p.suggestion}*")


def render_checks(rule_results) -> None:
    checks = pd.DataFrame({
        "Rule": [r.name for r in rule_results],
        "Result": ["PASS" if r.passed else "FAIL" for r in rule_results],
        "Measured": [r.measured for r in rule_results],
    })
    st.dataframe(
        checks.style.map(color_result, subset=["Result"]),
        width="stretch", hide_index=True,
        height=min(80 + 35 * len(checks), 600))


# ---------------------------------------------------------------------------
# GENERATE  (config -> preflight -> solve -> validate -> export)
# ---------------------------------------------------------------------------

if generate:
    S = make_settings()
    problems = sch.preflight(S)
    if problems:
        st.session_state.pop("result", None)
        st.session_state["blocked"] = problems
    else:
        st.session_state.pop("blocked", None)
        with st.spinner("Solving…"):
            sol = sch.build_and_solve(S)
        if not sol.grid:
            st.session_state.pop("result", None)
            st.session_state["infeasible"] = {"status": sol.status, "hints": sch.relaxation_hints(S)}
        else:
            st.session_state.pop("infeasible", None)
            results = sch.validate(S, sol.grid)
            st.session_state["result"] = {
                "settings": S, "grid": sol.grid, "results": results,
                "status": sol.status,
                "xlsx": sch.export_bytes(S, sol.grid, results),
            }


# ---------------------------------------------------------------------------
# RENDER  (top-down: blocked? infeasible? else roster -> fairness -> rules)
# ---------------------------------------------------------------------------

if "blocked" in st.session_state:
    show_problems("This rule combination has no valid schedule — preflight stopped "
                  "before solving.", st.session_state["blocked"])

if "infeasible" in st.session_state:
    info = st.session_state["infeasible"]
    st.error(f"The solver could not satisfy every hard rule (status: {info['status']}). "
             "The conflict is a combination of rules, not a single value. Try one of:")
    for hint in info["hints"]:
        st.markdown(f"- 💡 {hint}")

if "result" in st.session_state:
    R = st.session_state["result"]
    S = R["settings"]
    results = R["results"]
    fairness = [r for r in results if r.name.startswith("Fairness")]
    hard = [r for r in results if not r.name.startswith("Fairness")]
    all_pass = all(r.passed for r in results)

    # --- Headline -------------------------------------------------------
    top = st.columns([2, 1, 1, 1])
    top[0].subheader(f"{S.month_label} · {S.days} days")
    top[1].metric("Employees", len(S.employees))
    top[2].metric("Rules passed", f"{sum(r.passed for r in results)}/{len(results)}")
    top[3].metric("Solver", R["status"])
    st.caption(sch.roster_caption(S))

    hard_fail = [r for r in hard if not r.passed]
    soft_fail = [r for r in fairness if not r.passed]
    if all_pass:
        st.success("All rules PASS — safe to publish.")
    elif not hard_fail:
        # Fairness goals are soft/best-effort: a miss means the roster is still
        # valid (every hard rule holds), just not perfectly balanced this month.
        st.info(f"All hard rules pass — the roster is valid and safe to publish. "
                f"{len(soft_fail)} soft fairness goal(s) couldn't be fully met this "
                f"month (see **Fairness goals** below for the spread). Loosen a "
                f"tolerance or raise the solver budget to push further.")
    else:
        st.error(f"{len(hard_fail)} HARD rule(s) FAILED — do not publish. "
                 f"Review the hard-rule checks below.")

    st.download_button(
        "⬇ Download Excel (.xlsx)", data=R["xlsx"],
        file_name=f"schedule_{S.year}_{S.month:02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary")

    # --- Roster ---------------------------------------------------------
    st.markdown("#### Roster  · `D` day · `N` night · `OFF` off · `*` night team")
    gdf = grid_dataframe(S, R["grid"])
    st.dataframe(gdf.style.map(color_cell), width="stretch",
                 height=min(80 + 35 * len(gdf), 500))

    # --- Per-employee fairness summary ----------------------------------
    st.markdown("#### Per-employee fairness summary")
    loads = sch.employee_loads(S, R["grid"])
    fair_df = pd.DataFrame([{
        "Employee": ld["name"] + (" *" if ld["night_member"] else ""),
        "Total": ld["total"], "Day": ld["day"], "Night": ld["night"],
        "Fri/Sat": ld["weekend"], "Weekends off": ld["weekends_off"],
        "Overlaps": ld["overlaps"], "Max-runs": ld["max_runs"],
    } for ld in loads])
    st.dataframe(fair_df, width="stretch", hide_index=True)
    st.caption("Overlaps = forced extra-night days worked · Max-runs = max-length "
               "day/night streaks absorbed · these spread evenly within each pool.")

    # --- Fairness goals (the business priority) -------------------------
    st.markdown("#### Fairness goals")
    fcols = st.columns(len(fairness))
    for col, r in zip(fcols, fairness):
        short = r.name.replace("Fairness - ", "").split(" (")[0]
        col.metric(short.title(), "PASS" if r.passed else "FAIL",
                   delta=r.measured.split("  ")[0], delta_color="off")
    render_checks(fairness)

    # --- All hard rules -------------------------------------------------
    st.markdown("#### Hard-rule validation")
    render_checks(hard)

elif "blocked" not in st.session_state and "infeasible" not in st.session_state:
    st.info("Set your month, staff, night model, rules, and fairness in the sidebar, "
            "then click **Generate roster**.")
