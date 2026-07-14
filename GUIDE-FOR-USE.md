# Guide for Mishary 👋

Hey Mishary — Khamis here. This is a little tool that builds the monthly RT shift
schedule for you. You pick the month and a few numbers, click one button, and it gives
you a finished, color-coded Excel roster where the shifts are already fair and safe.

You don't need to know any coding. Just follow the steps below.

---

## What it does (in plain words)

- Splits the month into **Day (D)**, **Night (N)**, and **Off (OFF)** shifts for each person.
- Makes sure **everyone works exactly 16 shifts** — nobody does more, nobody does less.
- Covers nights with a **dedicated night team** (any number of people, who never do
  day shifts).
- Follows safety rules automatically:
  - never a **night shift followed by a day shift** the next morning,
  - never more than **4 day shifts** in a row, or **3 nights** in a row,
  - it does its best to give everyone a **full Friday + Saturday weekend off** (and it
    tells you honestly if a tight month can't manage it),
  - rest days are spread out (never more than 4 days off in a row, never a single lonely off day).
- Checks **every rule itself** and shows you a green **PASS** or red **FAIL** for each one,
  so you can trust it before you share it.

---

## Starting the app (every time)

> **Before the very first start on a new computer:** get the `shift-scheduler` folder from
> Khamis (he'll send a zip or a link) and put it in your home folder inside a folder called
> `Projects` — so it ends up at `Projects/shift-scheduler`. Everything below assumes it's there.

### On a Mac

1. Open the **Terminal** app (press **Cmd + Space**, type `terminal`, press **Enter**).
2. Type this and press **Enter** (this goes to the project folder):
   ```
   cd ~/Projects/shift-scheduler
   ```
3. Type this and press **Enter** (this starts the app):
   ```
   .venv/bin/streamlit run app.py
   ```
4. Your web browser opens the tool automatically. If it doesn't, the Terminal shows a
   link like `http://localhost:8501` — hold **Cmd** and click it.
5. Leave the Terminal window open while you use the tool — closing it stops the app.

When you're done, go back to Terminal and press **Ctrl + C** to stop it (yes, **Ctrl** —
not Cmd — that's just how Terminal works). That's it.

> **First time only (Mac):** if step 3 gives an error about something "not found," run these
> two lines once (from the project folder), then try step 3 again:
> ```
> python3 -m venv .venv
> .venv/bin/pip install -r requirements.txt
> ```
> (If anything looks scary, just send Khamis a screenshot.)

### On Windows

1. Open **PowerShell** (press the **Windows key**, type `powershell`, press **Enter**).
2. Type this and press **Enter** (this goes to the project folder):
   ```
   cd ~\Projects\shift-scheduler
   ```
3. Type this and press **Enter** (this starts the app):
   ```
   .venv\Scripts\streamlit run app.py
   ```
4. Your web browser opens the tool automatically. If it doesn't, PowerShell shows a
   link like `http://localhost:8501` — hold **Ctrl** and click it (or copy it into
   your browser).
5. Leave the PowerShell window open while you use the tool — closing it stops the app.

When you're done, go back to PowerShell and press **Ctrl + C** to stop it. That's it.

> **First time only (Windows):** you need Python once — get it from
> [python.org/downloads](https://www.python.org/downloads/) and, on the installer's first
> screen, tick **"Add python.exe to PATH"**. Then run these two lines in PowerShell (from
> the project folder), and try step 3 again:
> ```
> py -m venv .venv
> .venv\Scripts\pip install -r requirements.txt
> ```
> (Same rule as always: anything scary → screenshot → Khamis.)

---

## Using the tool (the fun part)

Everything you change is on the **left sidebar**. The results show on the right.

### 1 · Pick the month
Choose the **Year** and the **Month**. The tool figures out how many days the month has
and which days are Fri/Sat — you don't have to.

### 2 · Set your staff
You'll see a small table of names (Employee 1, Employee 2, …).
- **Rename** anyone by clicking a name and typing the real person's name.
- **Add** a person with the **+** at the bottom of the table.
- **Remove** a person by ticking the box at the left of their row, then pressing the
  **Delete** key (or the small trash icon at the top of the table).

Then, under **Night team**, pick the people who cover nights that month (any number; they
work nights only).

If you have **AM staff** (people on a fixed weekly pattern — for example Sun–Thu), add
their names under **AM shift** and pick their working days. They show up on the roster too,
but they're kept separate from the shift math above.

### 3 · Staffing per day
- **Min / Max day staff** — how many people you want on days each day (default 2–4).
- **Min / Max nights per day** — just leave these at **1 and 2**. On a few days two people
  will both be working the same night (extra coverage — not a split shift); that's normal,
  and the tool keeps it to as few days as possible.
  (If you pick a bigger night team and a number needs to change, the tool tells you.)

### 4 · Rules
These are pre-set to safe values. You usually don't need to touch them:
- Shifts per employee: **16**
- Max day shifts in a row: **4** · Max nights in a row: **3**
- Min work/off in a row: **2** · Max off in a row: **4**
- **Alternating weekends** (off unless you switch it on): everyone's weekends go
  **off, on, off, on…** — every other Fri+Sat fully off, guaranteed. Two things to know:
  on a very tight month the tool may say no schedule is possible (just switch it off and
  Generate again), and the "balanced weekends" check (a separate, best-effort evenness
  check — not this switch) can show red — the roster is still safe; it only means the
  weekend counts aren't perfectly even that month.

### 5 · Generate
Click the big **Generate roster** button.

---

## Reading the result

- A green banner saying **all rules pass — ready to publish** means you're good to go. 🎉
- A **blue** banner means **all the safety rules pass and the roster is valid**, but the
  workload split couldn't be made perfectly even this month (only happens on unusually
  tight months) — it's still safe to publish. If it bothers you, tell Khamis and he'll
  tune it.
- A **red** banner would mean a **safety rule** failed — don't publish (this shouldn't happen).
- The **Roster** table shows everyone's month:
  - 🟩 **D** = Day shift
  - 🟦 **N** = Night shift
  - 🟥 **OFF** = Day off
  - 🟨 **AM** = AM shift (a fixed set of weekdays)
  - a small **\*** next to a name = night-team member
- The **Per-employee fairness summary** shows each person's Total / Day / Night / Fri-Sat and
  how many rough patterns (overlap nights, long streaks) they got — so you can see the load
  is shared evenly.
- The **Fairness goals** and **Hard-rule validation** tables list every rule with a green
  **PASS** or red **FAIL**, plus the actual numbers.

### Download it
Click **⬇ Download Excel (.xlsx)** to save the schedule. It's the same colored table plus
a summary page with all the checks — ready to print or share.

---

## "It couldn't make a schedule" — what now?

Sometimes the numbers you picked just can't fit together (for example, asking for too many
day staff in a short month). The tool will **explain what clashed** and **suggest a fix**,
like *"Raise max day staff to 5"* or *"Allow 2 nights per day."*

Just make that one change in the sidebar and click **Generate** again. It will never give
you a schedule that quietly breaks a rule — it would rather stop and tell you.

---

## Why some rules are the way they are

- **Exactly 16 shifts** keeps the workload equal for everyone.
- **A dedicated night team** means rested night staff instead of everyone flipping between
  days and nights.
- **No night → day next morning** protects against the most tiring, error-prone pattern.
- **Max 3 nights / 4 days in a row** keeps fatigue down.
- **A full weekend off for everyone** (whenever the month allows it) keeps it fair.
- **Alternating weekends** (when on) spreads the weekends evenly over time — nobody is
  stuck working several weekends in a row while someone else always gets them off.

The tool checks all of these for you and shows the green PASS marks as proof.

---

## Quick help

| Problem | Try this |
|---|---|
| Browser didn't open | Click the `http://localhost:8501` link in Terminal / PowerShell |
| "command not found" / "not recognized" on start | Run the first-time install lines above (Mac or Windows) |
| It says it can't schedule | Read the suggestion, change that one number, Generate again |
| A **hard rule** shows FAIL in red | Don't publish — tell Khamis which rule (this shouldn't happen) |
| A **fairness** goal shows FAIL | Still safe to publish — mention it to Khamis only if you want it evened out |
| Anything confusing | Screenshot it and send it to Khamis 🙂 |

You've got this, Mishary. — Khamis
