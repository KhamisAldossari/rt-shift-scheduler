#!/usr/bin/env python3
"""
Streamlit front-end for the RT shift scheduler.

Drives `scheduler.py` entirely through a single `ScheduleSettings` object: pick a
month + year, edit the staff and the night team, set every rule value, click
Generate, review the roster and the independent PASS/FAIL checks, and download
the color-coded Excel. The solver logic lives in `scheduler.py` and is reused as
is -- this file never re-implements a constraint.

Run:  ./.venv/bin/streamlit run app.py
"""

from __future__ import annotations

import calendar

import pandas as pd
import streamlit as st

import scheduler as sch

st.set_page_config(page_title="RT Shift Scheduler", layout="wide")
st.title("Respiratory Therapy — Monthly Shift Scheduler")
st.caption("Configure any month and rule set, generate a roster, and download the Excel. "
           "Every hard rule is enforced by the CP-SAT solver and re-checked by an independent validator.")


# ---------------------------------------------------------------------------
# SIDEBAR — all inputs that populate ScheduleSettings
# ---------------------------------------------------------------------------

defaults = sch.ScheduleSettings()

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
        num_rows="dynamic", width="stretch", hide_index=True,
        key="emp_editor")
    employees = [str(x).strip() for x in emp_df["Employee"].tolist() if str(x).strip()]

    night_team = st.multiselect(
        "Night team (pick exactly 2 — they work nights only)",
        options=employees,
        default=[e for e in defaults.night_team if e in employees])
    if len(night_team) != 2:
        st.warning("The fixed night team should have exactly 2 members.")

    st.header("3 · Staffing per day")
    c1, c2 = st.columns(2)
    day_min = c1.number_input("Min day staff", 0, 50, defaults.day_min)
    day_max = c2.number_input("Max day staff", 1, 50, defaults.day_max)
    c3, c4 = st.columns(2)
    night_min = c3.number_input("Min nights/day", 0, 10, defaults.night_min)
    night_max = c4.number_input("Max nights/day (overlap)", 1, 10, defaults.night_max)

    st.header("4 · Rules")
    shifts = st.number_input("Shifts per employee (exact)", 1, 31, defaults.shifts_per_employee)
    c5, c6 = st.columns(2)
    max_day = c5.number_input("Max consecutive day shifts", 1, 14, defaults.max_consec_work)
    max_night = c6.number_input("Max consecutive nights", 1, 14, defaults.max_consec_night)
    c7, c8 = st.columns(2)
    min_work = c7.number_input("Min consecutive work", 2, 7, defaults.min_consec_work)
    min_off = c8.number_input("Min consecutive off", 2, 7, defaults.min_consec_off)
    c9, c10 = st.columns(2)
    max_off = c9.number_input("Max consecutive off", 1, 14, defaults.max_consec_off)
    gap = c10.number_input("Fairness gap (<=)", 0, 5, defaults.fairness_max_gap)

    generate = st.button("Generate roster", type="primary", width="stretch")


def make_settings() -> sch.ScheduleSettings:
    return sch.ScheduleSettings(
        year=int(year), month=int(month),
        employees=employees, night_team=night_team,
        day_min=int(day_min), day_max=int(day_max),
        night_min=int(night_min), night_max=int(night_max),
        shifts_per_employee=int(shifts),
        max_consec_work=int(max_day), max_consec_night=int(max_night),
        min_consec_work=int(min_work), min_consec_off=int(min_off),
        max_consec_off=int(max_off), fairness_max_gap=int(gap),
    )


# ---------------------------------------------------------------------------
# RENDER HELPERS
# ---------------------------------------------------------------------------

def grid_dataframe(S: sch.ScheduleSettings, grid: list[list[str]]) -> pd.DataFrame:
    cols = [f"D{d + 1} {sch.weekday_of(S, d)}" for d in range(S.days)]
    rows = [S.employees[e] + (" *" if S.employees[e] in S.night_team else "")
            for e in range(len(S.employees))]
    data = [[sch.CELL_LABEL[grid[e][d]] for d in range(S.days)]
            for e in range(len(S.employees))]
    return pd.DataFrame(data, index=rows, columns=cols)


def color_cell(val: str) -> str:
    code = {"D": sch.DAY, "N": sch.NIGHT, "OFF": sch.OFF}.get(val)
    bg = sch.WEB_COLORS.get(code, "#FFFFFF")
    return f"background-color: {bg}; text-align: center; color: #111;"


def show_blocked(title: str, items, suggestion_of=None):
    st.error(f"**{title}**")
    for it in items:
        if suggestion_of:
            st.markdown(f"- {it.message}\n  \n  💡 *{it.suggestion}*")
        else:
            st.markdown(f"- {it}")


# ---------------------------------------------------------------------------
# GENERATE
# ---------------------------------------------------------------------------

if generate:
    S = make_settings()
    problems = sch.preflight(S)
    if problems:
        st.session_state.pop("result", None)
        show_blocked("This rule combination has no valid schedule — preflight stopped before solving.",
                     problems, suggestion_of=True)
    else:
        with st.spinner("Solving…"):
            sol = sch.build_and_solve(S)
        if not sol.grid:
            st.session_state.pop("result", None)
            st.error(f"The solver could not satisfy every hard rule (status: {sol.status}). "
                     "The conflict is a combination of rules, not a single value. Try one of:")
            for hint in sch.relaxation_hints(S):
                st.markdown(f"- 💡 {hint}")
        else:
            results = sch.validate(S, sol.grid)
            st.session_state["result"] = {
                "settings": S, "grid": sol.grid, "results": results,
                "status": sol.status,
                "xlsx": sch.export_bytes(S, sol.grid, results),
            }

# ---------------------------------------------------------------------------
# RESULTS
# ---------------------------------------------------------------------------

if "result" in st.session_state:
    R = st.session_state["result"]
    S = R["settings"]
    results = R["results"]
    all_pass = all(r.passed for r in results)

    top = st.columns([2, 1, 1, 1])
    top[0].subheader(f"{S.month_label} · {S.days} days")
    top[1].metric("Employees", len(S.employees))
    top[2].metric("Rules passed", f"{sum(r.passed for r in results)}/{len(results)}")
    top[3].metric("Solver", R["status"])

    if all_pass:
        st.success("All rules PASS — safe to publish.")
    else:
        st.warning("Some rules FAILED — review the checks below before publishing.")

    st.download_button(
        "⬇ Download Excel (.xlsx)", data=R["xlsx"],
        file_name=f"schedule_{S.year}_{S.month:02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary")

    st.markdown("#### Roster  · `D` day · `N` night · `OFF` off · `*` night team")
    gdf = grid_dataframe(S, R["grid"])
    st.dataframe(gdf.style.map(color_cell), width="stretch", height=min(80 + 35 * len(gdf), 500))

    st.markdown("#### Per-employee totals")
    totals = pd.DataFrame({
        "Employee": [S.employees[e] for e in range(len(S.employees))],
        "Total": [sum(1 for d in range(S.days) if R["grid"][e][d] != sch.OFF) for e in range(len(S.employees))],
        "Day": [sum(1 for d in range(S.days) if R["grid"][e][d] == sch.DAY) for e in range(len(S.employees))],
        "Night": [sum(1 for d in range(S.days) if R["grid"][e][d] == sch.NIGHT) for e in range(len(S.employees))],
        "Fri/Sat": [sum(1 for d in sch.weekend_day_indices(S) if R["grid"][e][d] != sch.OFF) for e in range(len(S.employees))],
    })
    st.dataframe(totals, width="stretch", hide_index=True)

    st.markdown("#### Independent rule validation")
    checks = pd.DataFrame({
        "Rule": [r.name for r in results],
        "Result": ["PASS" if r.passed else "FAIL" for r in results],
        "Measured": [r.measured for r in results],
    })
    st.dataframe(
        checks.style.map(
            lambda v: "background-color: #C6EFCE" if v == "PASS"
            else ("background-color: #FFC7CE" if v == "FAIL" else ""),
            subset=["Result"]),
        width="stretch", hide_index=True, height=min(80 + 35 * len(checks), 700))
else:
    st.info("Set your month and rules in the sidebar, then click **Generate roster**.")
