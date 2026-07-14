#!/usr/bin/env python3
"""
Streamlit front-end for the RT shift scheduler  (calm, grid-first redesign).

A THIN UI over scheduler.py. It only collects configuration into one
ScheduleSettings object and renders the engine's outputs -- it never
re-implements a constraint or recomputes a rule. Every number, count, spread,
and PASS/FAIL on screen comes straight from a scheduler function/field
(preflight -> build_and_solve -> validate -> export_bytes).

Layout
  Sidebar : configuration only -- month, staff, night team, staffing bands, and
            rules (everything that builds ScheduleSettings). The rarely-touched
            solver budget lives in an expander, with one primary "Generate
            roster" button.
  Main    : results, with the color-coded grid as the visual hero --
              * a month header + compact metrics + a subtle solver status,
              * a positive, next-step status banner (3-way: success / info / error),
              * a prominent Excel download, then three tabs:
                  Roster     -- the big color grid (hero), a legend, and a per-day
                                coverage strip built from coverage_per_day().
                  Fairness   -- the four soft goals as PASS/FAIL cards (spread vs
                                tolerance), the per-employee summary, and the checks.
                  Hard rules -- the independent hard-rule validation table.

Three render states persist across reruns via st.session_state: the initial
hint, a preflight "blocked" state (each Problem's message + suggestion), and an
"infeasible" state (a positive error + the engine's relaxation hints).

The grid colors come from scheduler.WEB_COLORS (mirrors the Excel fills) and the
chrome from .streamlit/config.toml, so the web view and the workbook agree.

Run:  ./.venv/bin/streamlit run app.py
"""

from __future__ import annotations

import calendar

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
    initial_sidebar_state="expanded",
)

# A small amount of CSS, scoped to OUR OWN markup only (a kicker label and the
# grid legend) so we never fight Streamlit's internals across versions. Palette
# and typography come from .streamlit/config.toml; grid colors from WEB_COLORS.
st.markdown(
    """
    <style>
      .kicker{font-size:.72rem;letter-spacing:.16em;font-weight:600;
              color:#7A8194;text-transform:uppercase;margin:0 0 .15rem}
      .legend{display:flex;flex-wrap:wrap;gap:.45rem 1.2rem;align-items:center;
              font-size:.85rem;color:#5A6273;margin:.1rem 0 .35rem}
      .legend-item{display:inline-flex;align-items:center;gap:.42rem}
      .legend-sw{width:.95rem;height:.95rem;border-radius:.28rem;
                 border:1px solid rgba(20,30,55,.12);display:inline-block}
      .legend-code{font-weight:700;color:#1B2330}
      .legend-star{color:#305496;font-weight:700;font-size:1.05rem;line-height:1}
    </style>
    """,
    unsafe_allow_html=True,
)

defaults = sch.ScheduleSettings()       # EVERY UI default is pulled from here


# ---------------------------------------------------------------------------
# APP HEADER  (constant identity; the month-specific header lives in results)
# ---------------------------------------------------------------------------

st.markdown('<div class="kicker">Respiratory Therapy &middot; Workforce</div>',
            unsafe_allow_html=True)
st.title("Monthly Shift Scheduler")
st.caption(
    "Configure a month, generate a fair roster, and download the color-coded "
    "Excel. Every hard rule is enforced by the solver and independently "
    "re-checked; the four fairness goals are reported with their measured spread.")
st.divider()


# ---------------------------------------------------------------------------
# SIDEBAR  -- configuration only (every input that populates ScheduleSettings)
# ---------------------------------------------------------------------------

with st.sidebar:
    st.markdown('<div class="kicker">Configuration</div>', unsafe_allow_html=True)

    st.subheader("1 · Month")
    col_y, col_m = st.columns(2)
    year = col_y.number_input("Year", min_value=2000, max_value=2100,
                              value=defaults.year, step=1)
    month = col_m.selectbox("Month", options=list(range(1, 13)),
                            index=defaults.month - 1,
                            format_func=lambda m: calendar.month_name[m])

    st.subheader("2 · Staff")
    st.caption("Add, remove, or rename employees.")
    emp_df = st.data_editor(
        pd.DataFrame({"Employee": list(defaults.employees)}),
        num_rows="dynamic", width="stretch", hide_index=True, key="emp_editor",
        column_config={"Employee": st.column_config.TextColumn(
            "Employee", width="large", help="One row per person.")})
    employees = [str(x).strip() for x in emp_df["Employee"].tolist() if str(x).strip()]

    st.subheader("3 · Night team")
    night_team = st.multiselect(
        "Night team (any size ≥ 1 — they work nights only)",
        options=employees,
        default=[e for e in defaults.night_team if e in employees])
    if not night_team:
        st.warning("Pick at least one night-team member.")
    else:
        st.caption(f"{len(night_team)} night-team member(s) selected.")

    st.subheader("4 · AM shift")
    st.caption("Fixed-pattern staff rostered outside the solver — appended as extra "
               "rows, not part of coverage or fairness. Leave empty if unused.")
    am_df = st.data_editor(
        pd.DataFrame({"AM staff": pd.Series(list(defaults.am_team), dtype="object")}),
        num_rows="dynamic", width="stretch", hide_index=True, key="am_editor",
        column_config={"AM staff": st.column_config.TextColumn(
            "AM staff", width="large", help="One row per AM-shift person.")})
    am_team = [str(x).strip() for x in am_df["AM staff"].tolist() if str(x).strip()]
    am_days = st.multiselect(
        "AM working days", options=sch.WEEKDAY_NAMES,
        default=list(defaults.am_days),
        help="The weekdays AM staff work every week (default Sun–Thu).")

    st.subheader("5 · Staffing per day")
    c1, c2 = st.columns(2)
    day_min = c1.number_input("Min day staff", 0, 50, defaults.day_min)
    day_max = c2.number_input("Max day staff", 1, 50, defaults.day_max)
    c3, c4 = st.columns(2)
    night_min = c3.number_input("Min nights/day", 0, 10, defaults.night_min)
    night_max = c4.number_input("Max nights/day (overlap)", 1, 10, defaults.night_max)

    st.subheader("6 · Rules")
    shifts = st.number_input("Shifts per employee (exact)", 1, 31,
                             defaults.shifts_per_employee)
    c5, c6 = st.columns(2)
    max_day = c5.number_input("Max consecutive day shifts", 1, 14, defaults.max_consec_work)
    max_night = c6.number_input("Max consecutive nights", 1, 14, defaults.max_consec_night)
    c7, c8 = st.columns(2)
    min_work = c7.number_input("Min consecutive work", 2, 7, defaults.min_consec_work,
                               help="The model encodes a minimum of 2; other values are "
                                    "reported as infeasible by preflight.")
    min_off = c8.number_input("Min consecutive off", 2, 7, defaults.min_consec_off,
                              help="The model encodes a minimum of 2; other values are "
                                   "reported as infeasible by preflight.")
    max_off = st.number_input("Max consecutive off", 1, 14, defaults.max_consec_off)
    alt_weekends = st.toggle(
        "Alternating weekends (hard)", value=defaults.alternating_weekends,
        help="On: every employee's full Fri+Sat weekends strictly alternate "
             "off / on / off — of every two consecutive full weekends, exactly one "
             "is fully off. A HARD rule (not a soft fairness goal); it can make very "
             "tight months infeasible, in which case turn it off and re-generate.")

    with st.expander("Advanced · solver budget"):
        det_budget = st.number_input(
            "Solver budget (deterministic units)", 4.0, 240.0,
            float(defaults.solver_det_time_limit), step=4.0,
            help="Higher = more search = better fairness convergence on hard months, "
                 "at the cost of solve time. The budget is deterministic, so the "
                 "roster is reproducible regardless of machine speed.")

    st.divider()
    generate = st.button("Generate roster", type="primary", width="stretch",
                         icon=":material/bolt:")


def make_settings() -> sch.ScheduleSettings:
    """Collect every sidebar input into the single ScheduleSettings object.

    The UI is fixed-team only -- rotation_mode is left at its ScheduleSettings
    default (fixed_team); rotate mode remains available in the engine/CLI via
    rotation_mode=sch.ROTATE. Fairness tuning (tolerances + objective weights)
    is intentionally not exposed here either -- the engine defaults apply. No
    scheduling value is hard-coded here -- defaults all come from
    sch.ScheduleSettings().
    """
    return sch.ScheduleSettings(
        year=int(year), month=int(month),
        employees=employees, night_team=night_team,
        am_team=am_team, am_days=tuple(am_days),
        day_min=int(day_min), day_max=int(day_max),
        night_min=int(night_min), night_max=int(night_max),
        shifts_per_employee=int(shifts),
        max_consec_work=int(max_day), max_consec_night=int(max_night),
        min_consec_work=int(min_work), min_consec_off=int(min_off),
        max_consec_off=int(max_off),
        alternating_weekends=bool(alt_weekends),
        solver_det_time_limit=float(det_budget),
    )


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
    code = {"D": sch.DAY, "N": sch.NIGHT, "OFF": sch.OFF, "AM": sch.AM}.get(val)
    bg = sch.WEB_COLORS.get(code, "#FFFFFF")
    return f"background-color: {bg}; text-align: center; color: #13233B; font-weight: 600;"


def color_result(val: str) -> str:
    """PASS/FAIL cell color, reusing the grid's green/pink so the views agree."""
    if val == "PASS":
        return f"background-color: {sch.WEB_COLORS[sch.DAY]}; color: #13233B; font-weight: 600;"
    if val == "FAIL":
        return f"background-color: {sch.WEB_COLORS[sch.OFF]}; color: #13233B; font-weight: 600;"
    return ""


def legend(show_am: bool = False) -> None:
    """A compact key whose swatches are the EXACT grid colors (WEB_COLORS)."""
    items = [(sch.WEB_COLORS[sch.DAY], "D", "Day"),
             (sch.WEB_COLORS[sch.NIGHT], "N", "Night"),
             (sch.WEB_COLORS[sch.OFF], "OFF", "Off")]
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
            # Row 0 -- day staff: light green, amber on weekends.
            css.iloc[0, j] = (
                f"background-color: {'#F6F1DD' if wk else '#F1F8F3'}; "
                "text-align: center; color: #2C3A30; font-weight: 500;")
            # Row 1 -- night staff: light blue; deeper blue + bold on overlap days.
            if ov:
                nbg, ncol, nwt = "#D6E2F4", "#1F3A66", 700
            elif wk:
                nbg, ncol, nwt = "#F6F1DD", "#2B3550", 500
            else:
                nbg, ncol, nwt = "#EEF3FA", "#2B3550", 500
            css.iloc[1, j] = (f"background-color: {nbg}; text-align: center; "
                              f"color: {ncol}; font-weight: {nwt};")
        return css

    st.dataframe(df.style.apply(_styles, axis=None), width="stretch",
                 height=36 + 36 * len(df), row_height=36)
    st.caption("Staff on duty each day. Amber columns are the Fri/Sat weekend; a "
               "deeper-blue night cell is a forced extra (overlap) night above the floor.")


def render_checks(rule_results) -> None:
    """A scannable PASS/FAIL table for a list of RuleResults (display only)."""
    checks = pd.DataFrame({
        "Rule": [r.name for r in rule_results],
        "Result": ["PASS" if r.passed else "FAIL" for r in rule_results],
        "Measured": [r.measured for r in rule_results],
    })
    st.dataframe(
        checks.style.map(color_result, subset=["Result"]),
        width="stretch", hide_index=True,
        height=min(80 + 35 * len(checks), 600),
        column_config={
            "Rule": st.column_config.TextColumn(width="large"),
            "Result": st.column_config.TextColumn(width="small"),
            "Measured": st.column_config.TextColumn(width="medium"),
        })


def render_problems(title: str, problems) -> None:
    """Surface each preflight Problem's message AND its suggestion."""
    st.error(f"**{title}**")
    for p in problems:
        st.markdown(f"- {p.message}  \n  💡 *{p.suggestion}*")


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
            f"All hard safety rules pass — this roster is valid and ready to publish. "
            f"{len(soft_fail)} soft fairness goal(s) couldn't be fully balanced this "
            f"month (see the spread in the **Fairness** tab). Raise the solver "
            f"budget (Advanced expander) to push the balance further.")
    else:
        st.error(
            f"{len(hard_fail)} hard rule(s) failed — do not publish. Open the "
            f"**Hard rules** tab to see which constraint was violated.")


def short_goal(name: str) -> str:
    """'Fairness - balanced night load (spread <= 1)' -> 'Balanced Night Load'."""
    return name.replace("Fairness - ", "").split(" (")[0].title()


# ---------------------------------------------------------------------------
# GENERATE  (config -> preflight -> solve -> validate -> export); results persist
# in st.session_state so they survive Streamlit reruns.
# ---------------------------------------------------------------------------

if generate:
    S = make_settings()
    problems = sch.preflight(S)
    if problems:
        st.session_state.pop("result", None)
        st.session_state.pop("infeasible", None)
        st.session_state["blocked"] = problems
    else:
        st.session_state.pop("blocked", None)
        with st.spinner("Solving…"):
            sol = sch.build_and_solve(S)
        if not sol.grid:
            st.session_state.pop("result", None)
            st.session_state["infeasible"] = {
                "status": sol.status, "hints": sch.relaxation_hints(S)}
        else:
            st.session_state.pop("infeasible", None)
            results = sch.validate(S, sol.grid)
            st.session_state["result"] = {
                "settings": S, "grid": sol.grid, "results": results,
                "status": sol.status,
                "xlsx": sch.export_bytes(S, sol.grid, results),
            }


# ---------------------------------------------------------------------------
# RENDER  (blocked? infeasible? else the result; otherwise the initial hint)
# ---------------------------------------------------------------------------

if "blocked" in st.session_state:
    render_problems(
        "This rule combination has no valid schedule — preflight stopped before "
        "solving.", st.session_state["blocked"])

if "infeasible" in st.session_state:
    info = st.session_state["infeasible"]
    st.error(
        f"The solver could not satisfy every hard rule (status: {info['status']}). "
        "The conflict is a combination of rules, not a single value. Try one of:")
    for hint in info["hints"]:
        st.markdown(f"- 💡 {hint}")

if "result" in st.session_state:
    R = st.session_state["result"]
    S = R["settings"]
    results = R["results"]
    # Split hard vs fairness via the structured `kind` field (not by name).
    fairness = [r for r in results if r.kind == "fairness"]
    hard = [r for r in results if r.kind != "fairness"]
    passed = sum(1 for r in results if r.passed)

    # --- Month header + compact metrics + a subtle solver status ------------
    head_l, head_r = st.columns([2.3, 1.4])
    with head_l:
        st.markdown('<div class="kicker">Monthly Roster</div>', unsafe_allow_html=True)
        st.subheader(f"{S.month_label} · {S.days} days")
        st.caption(sch.roster_caption(S))
        st.badge(f"Solver · {R['status']}", color="gray",
                 icon=":material/check_small:")
    with head_r:
        m1, m2 = st.columns(2)
        m1.metric("Employees", len(S.employees), border=True)
        m2.metric("Rules passed", f"{passed}/{len(results)}", border=True)

    # --- Positive, next-step status banner ----------------------------------
    status_banner(hard, fairness)

    # --- Prominent export, near the top of the results ----------------------
    dl, _ = st.columns([1.4, 3])
    dl.download_button(
        "Download Excel (.xlsx)", data=R["xlsx"],
        file_name=f"schedule_{S.year}_{S.month:02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary", icon=":material/download:", width="stretch")

    st.write("")
    tab_roster, tab_fair, tab_hard = st.tabs(["Roster", "Fairness", "Hard rules"])

    # ======================= ROSTER (the hero) ==============================
    with tab_roster:
        legend(show_am=bool(S.am_team))
        render_grid(S, R["grid"])
        st.divider()
        st.markdown("##### Daily coverage")
        render_coverage(S, R["grid"])

    # ======================= FAIRNESS =======================================
    with tab_fair:
        st.markdown("##### Fairness goals")
        st.caption("Four soft, equally-weighted goals. Each PASSES when its measured "
                   "spread (max − min) is within tolerance.")
        cards = st.columns(len(fairness))
        for col, r in zip(cards, fairness):
            with col.container(border=True):
                st.markdown(f"**{short_goal(r.name)}**")
                if r.passed:
                    st.badge("PASS", color="green", icon=":material/check_circle:")
                else:
                    st.badge("FAIL", color="red", icon=":material/cancel:")
                st.caption(f"spread {r.spread} · tol {r.tolerance}")
                # A goal can sit within tolerance yet still FAIL on a non-spread
                # condition (e.g. the weekends goal also wants a full Fri+Sat off
                # for everyone). Say so, so the card never looks self-contradictory;
                # the full reason is in the checks table below.
                if (not r.passed and r.spread is not None
                        and r.tolerance is not None and r.spread <= r.tolerance):
                    st.caption("Within tolerance, but another fairness condition "
                               "isn't met — see the checks below.")

        st.divider()
        st.markdown("##### Per-employee summary")
        loads = sch.employee_loads(S, R["grid"])
        fair_df = pd.DataFrame([{
            "Employee": ld["name"] + (" ∗" if ld["night_member"] else ""),
            "Total": ld["total"], "Day": ld["day"], "Night": ld["night"],
            "Fri/Sat": ld["weekend"], "Weekends off": ld["weekends_off"],
            "Overlaps": ld["overlaps"], "Max-runs": ld["max_runs"],
        } for ld in loads])
        st.dataframe(fair_df, width="stretch", hide_index=True,
                     height=min(80 + 35 * len(fair_df), 460))
        st.caption("Overlaps = forced extra-night days worked · Max-runs = max-length "
                   "day/night streaks absorbed · both are balanced within each pool "
                   "(day-capable vs night-eligible).")

        st.divider()
        st.markdown("##### Fairness checks")
        render_checks(fairness)

    # ======================= HARD RULES =====================================
    with tab_hard:
        st.markdown("##### Hard-rule validation")
        st.caption("Every hard rule re-derived independently from the finished grid. "
                   "These are non-negotiable — all must PASS to publish.")
        render_checks(hard)

elif "blocked" not in st.session_state and "infeasible" not in st.session_state:
    # --- Initial / empty state: a calm hint -------------------------------
    with st.container(border=True):
        st.markdown("##### Build a monthly roster")
        st.markdown(
            "Set the month, staff, night team, and rules in the sidebar, "
            "then press **Generate roster**. You'll get a color-coded grid, an "
            "independent PASS/FAIL check of every rule, and a one-click Excel export.")
        legend()
