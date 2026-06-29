# Guide for Mishary 👋

Hey Mishary — Khamis here. This is a little tool that builds the monthly RT shift
schedule for you. You pick the month and a few numbers, click one button, and it gives
you a finished, color-coded Excel roster where the shifts are already fair and safe.

You don't need to know any coding. Just follow the steps below.

---

## What it does (in plain words)

- Splits the month into **Day (D)**, **Night (N)**, and **Off (OFF)** shifts for each person.
- Makes sure **everyone works exactly 16 shifts** — nobody does more, nobody does less.
- Keeps **2 people on nights** for the whole month (they never do day shifts).
- Follows safety rules automatically:
  - never a **night shift followed by a day shift** the next morning,
  - never more than **4 day shifts** in a row, or **3 nights** in a row,
  - everyone gets a **full Friday + Saturday weekend off** at least once,
  - rest days are spread out (never more than 4 days off in a row, never a single lonely off day).
- Checks **every rule itself** and shows you a green **PASS** or red **FAIL** for each one,
  so you can trust it before you share it.

---

## Starting the app (every time)

1. Open the **Terminal** app on the Mac.
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

When you're done, go back to Terminal and press **Ctrl + C** to stop it. That's it.

> **First time only:** if step 3 gives an error about something "not found," run this once,
> then try step 3 again:
> ```
> .venv/bin/pip install -r requirements.txt
> ```
> (If anything looks scary, just send Khamis a screenshot.)

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
- **Remove** a person by selecting their row and deleting it.

Then, under **Night team**, pick the **2 people** who will cover nights that month.
(Pick exactly 2 — the tool reminds you if you don't.)

### 3 · Staffing per day
- **Min / Max day staff** — how many people you want on days each day (default 2–4).
- **Min / Max nights per day** — leave at **1 and 2**. The "2" lets the two night people
  finish their 16 shifts; the tool only uses a 2nd night on the few days it has to.

### 4 · Rules
These are pre-set to safe values. You usually don't need to touch them:
- Shifts per employee: **16**
- Max day shifts in a row: **4** · Max nights in a row: **3**
- Min work/off in a row: **2** · Max off in a row: **4**

### 5 · Generate
Click the big **Generate roster** button.

---

## Reading the result

- A green banner **"All rules PASS — safe to publish"** means you're good to go. 🎉
- The **Roster** table shows everyone's month:
  - 🟩 **D** = Day shift
  - 🟦 **N** = Night shift
  - 🟥 **OFF** = Day off
  - a small **\*** next to a name = night-team member
- The **Per-employee totals** table shows each person's Total / Day / Night / Fri-Sat counts.
- The **Rule validation** table lists every safety & fairness rule with a green **PASS**
  or red **FAIL**.

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
- **Only 2 night people** means dedicated, rested night staff instead of everyone flipping
  between days and nights.
- **No night → day next morning** protects against the most tiring, error-prone pattern.
- **Max 3 nights / 4 days in a row** keeps fatigue down.
- **A full weekend off for everyone** keeps it fair.

The tool checks all of these for you and shows the green PASS marks as proof.

---

## Quick help

| Problem | Try this |
|---|---|
| Browser didn't open | Click the `http://localhost:8501` link in Terminal |
| "command not found" on start | Run the first-time install line above |
| It says it can't schedule | Read the suggestion, change that one number, Generate again |
| A rule shows **FAIL** in red | Don't publish yet — tell Khamis which rule failed |
| Anything confusing | Screenshot it and send it to Khamis 🙂 |

You've got this, Mishary. — Khamis
