# Weekly report automation (Excel macros)

Paste a raw data export, click once, and get a clean, formatted report. Click again to turn it into pivot tables.

| File | What it is |
|---|---|
| [`BuildReport.bas`](BuildReport.bas) | **Macro 1.** Copies the columns you pick, adds calculated columns and formats everything |
| [`CreatePivotReports.bas`](CreatePivotReports.bas) | **Macro 2.** Builds 3 pivot tables side by side from your report |
| [`report-headers.csv`](report-headers.csv) | The 22 headers for row 1 of your Report sheet (practice dataset) |

The full step-by-step guide (PDF) is free: comment **AUTOMATE** on my post and I'll send it to you.

## Quick start

1. **Get the practice data:** [Super Store - Retail Orders on Kaggle](https://www.kaggle.com/datasets/nocturnvoid/super-store-retail-orders) (free account needed to download).
2. **Make 2 sheets** in a new workbook: `Input` and `Report`.
3. **Paste the data** into cell A1 of `Input`, headers included.
4. **Add the Report headers:** open `report-headers.csv`, copy row 1, and paste it into cell A1 of `Report`.
5. **Add the macros:** open the VBA editor (Windows `Alt + F11` · Mac `fn + Option + F11`), click **Insert → Module**, and paste `BuildReport.bas`. Do the same in a **second** module for `CreatePivotReports.bas`.
6. **Run them:** press `Alt + F8` (Mac `fn + Option + F8`), run **BuildReport**, then **CreatePivotReports**.
7. **Save as `.xlsm`** (Excel Macro-Enabled Workbook), or Excel deletes the code.

## Use it with your own data

You only change the parts marked **EDIT** in the code:

- **BuildReport**
  - `EDIT 1`: which columns to copy (`Array("B", "A")` = Input column B → Report column A)
  - `EDIT 2`: calculated columns (`{ROW}` is filled in for every row)
  - `EDIT 3`: number formats
- **CreatePivotReports**
  - `PICK YOUR FIELDS`: what to group by and what to add up
  - Optional filter lines: delete the `'` in front to switch them on

## Good to know

- Works in the Excel desktop app (Windows and Mac). Excel on the web can't run macros.
- There's no undo after a macro runs, so save a copy first.
- `BuildReport` **adds** rows below what's already in Report each time you run it. Paste a fresh batch into `Input` before each run.
- `CreatePivotReports` deletes and rebuilds the `Pivot Reports` sheet every run.

The practice dataset belongs to its Kaggle author and isn't included here. Please download it from the link above.
