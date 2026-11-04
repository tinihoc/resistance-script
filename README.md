# Resistance Workout Script

An automated workout management and exercise selection engine for Google Sheets, built with Google Apps Script.

This system acts as an intelligent training partner: tracking exercises across individual sheets, curating daily workout routines based on muscle fatigue and training intensity, synchronizing a master dashboard in real time, and maintaining a snapshot undo history for seamless workout logging.

---

## 🚀 Key Features

* **Intelligent Daily Exercise Selector**:
  * Automatically recommends workouts based on Day Type (**Heavy** vs. **Light**), desired muscle volume, and recovery constraints.
  * Prioritizes missed sessions and exercises that haven't been trained for the longest duration.
  * Prevents consecutive lower-body fatigue (restricts lower-body muscles trained the day before).
  * Automatically relaxes constraints in a 5-stage graceful fallback hierarchy if quota cannot be strictly met.
* **Interactive Selection Dialog & In-Place Editor**:
  * Preview suggested exercises, toggle exclusions, and trigger re-rolls with a single click.
  * In-place spreadsheet editor modal that dynamically simulates and recalculates spreadsheet formulas (`SUM`, `ROUND`, cell references) before writing to the sheet.
* **Master Dashboard Synchronization**:
  * Real-time bidirectional synchronization between the central `dashboard` sheet and individual exercise sheets.
  * Dynamic hyperlinking via `=LET(...)` formulas that jump directly to the target exercise and active date block.
  * Color-coded status tracking (`planned`, `completed`, `missed`, `unscheduled`).
  * Stable sorting by date with index preservation.
* **Automated Block Duplication via Triggers**:
  * Ticking the workout date checkbox automatically clones the workout block for the next session.
  * Removing the checkbox deletes the subsequent block cleanly.
* **Snapshot & Reversal (Undo) System**:
  * Maintains an in-memory / `DocumentProperties` stack (up to 5 snapshots) of confirmed workout generations.
  * Full rollback capability: delete generated rows, restore previous dates, statuses, and custom edits.

---

## 🧠 Selection Algorithm & Muscle Rules

When generating a workout via **🏋️ Workout > Select Exercises...**, the algorithm enforces the following:

### 1. Pre-Filtering & Prioritization
* **Excluded Exercises & Muscles**: Skips any exercise or muscle manually checked as excluded in the UI.
* **Date Bounds**: Exercises with dates logged after the planned workout date are excluded.
* **Queue Order**:
  1. Exercises marked as `missed` are placed first.
  2. Exercises sorted by oldest `lastSession` (longest duration since last worked).

### 2. Core Workout Restrictions
* **Category Match**:
  * **Heavy Day**: Curated from `heavy` and `moderate` exercises.
  * **Light Day**: Curated from `light` and `moderate` exercises.
* **Lower Body Frequency**: At most **one** lower-body exercise per routine (`calves`, `hammies`, `quads`, `inner thighs`, `glutes`, `lower back`).
* **Day-Before Lower Body Fatigue**: An exercise cannot target a lower-body muscle if that exact muscle was trained (`completed`) the **day before** the planned date (`date - 1 day`). If yesterday was a rest day, all lower-body muscles are eligible.
* **Heavy Day Back Conflict**: Exercises targeting `lower back` and `upper back` cannot be scheduled on the same heavy day.
* **Muscle Overlap**: In strict mode, every selected exercise must target unique, unaddressed muscle groups.
* **Target Quota Ceiling**: Exercises whose muscle count would push the routine past the requested target muscles are skipped.

### 3. Graceful Fallback Hierarchy (Relaxation Stages)
If the strict constraints cannot satisfy the requested muscle quota, the algorithm progressively relaxes constraints:
1. **Stage 1 (Strict)**: All restrictions strictly enforced.
2. **Stage 2 (Allow Muscle Overlap)**: Allows exercises that share muscles with already selected exercises, provided they contribute at least one new muscle group.
3. **Stage 3 (Allow Lower Muscles from Day Before)**: Lifts the day-before lower-body restriction.
4. **Stage 4 (Allow Multiple Lower Body & Back Overlap)**: Allows multiple lower-body exercises and combines upper/lower back on heavy days.
5. **Stage 5 (Allow All Categories)**: Opens up all intensity tiers (`heavy`, `moderate`, and `light`).

---

## 🎨 Status Color System

| Status | Color | Description |
| :--- | :--- | :--- |
| **`planned`** | Blue (`#1b73e8`) | Scheduled for an upcoming workout. |
| **`completed`** | Green (`#137333`) | Successfully executed session. |
| **`missed`** | Red (`#d93025`) | Scheduled session that was skipped. |
| **`unscheduled`** | Purple (`#7c02ff`) | Newly cloned workout block awaiting scheduling. |

---

## 📂 Project Architecture

```
├── .clasp.json                 # Clasp project configuration and Apps Script ID
├── appsscript.json             # Google Apps Script manifest (timeZone, dependencies)
├── Config.js                   # Global configurations and status color palettes
├── DailyExerciseSelector.js    # Core selection algorithm, snapshot stack & confirmation engine
├── DashboardService.js         # Block boundary detection, section extractors, dashboard sync
├── Dialog.html                 # Interactive selection modal UI and client-side formula engine
├── Triggers.js                 # Spreadsheet onOpen menu and onEdit event triggers
├── UndoDialog.html             # Rollback / Undo interface for restoring past snapshots
├── Utils.js                    # Custom spreadsheet functions (SHEETNAME_DYNAMIC, GET_EXERCISE_DATE, etc.)
└── README.md                   # Project documentation
```

---

## 🛠️ Installation & Setup

1. **Prerequisites**: Install [Google Clasp](https://github.com/google/clasp) if deploying from the command line:
   ```bash
   npm install -g @google/clasp
   clasp login
   ```
2. **Clone & Push**:
   ```bash
   git clone https://github.com/tinihoc/resistance-script.git
   cd resistance-script
   clasp push
   ```
3. **Open Google Sheets**:
   * Reload the spreadsheet.
   * Access the custom menu: **🏋️ Workout > Select Exercises...**.
