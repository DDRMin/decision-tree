# Phase 0: host checks (go/no-go)

These checks run the skeleton add-in in real Excel (Office Home 2024, 64-bit) before the full editor is built.

## Automated results (`tools/RunPhase0Checks.ps1`, hidden Excel instance)

| # | Check | Result |
|---|---|---|
| 3 | Custom XML part survives save as `.xlsx`, close and reopen | PASS |
| 6 | Render 500 nodes (1,500 shapes) | PASS, 0.75 s |
| 6 | Render 1,000 nodes (3,000 shapes) | PASS, about 5 s |
| 6 | Clear deletes only add-in shapes; the user's own shape survives | PASS (4.4 s for 3,000 shapes) |
| 8 | Excel exits cleanly with the add-in loaded | PASS, after the fix below |

Findings:
- **The exit bug was found and fixed.** The add-in held references to Excel (event sinks, panes, timer), and Excel-DNA's `AutoClose` doesn't run when Excel shuts down. So `EXCEL.EXE` stayed alive after closing. `ShutdownWatcher` now releases everything on `OnBeginShutdown`. A control run without the add-in confirmed the cause.
- **Shape time grows faster than node count:** 2× the nodes takes about 6.5× the time. The likely causes are naming each shape (Excel checks name uniqueness) and looking up shapes one at a time when deleting. Phase 2 will tag shapes in bulk, group each tree's shapes so they can be deleted in one call, and redraw only what changed. 500-node trees are already fast.

## Manual checks (need you at the screen)

Load the add-in:
1. Close all Excel windows.
2. Go to **File → Options → Add-ins → Manage: Excel Add-ins → Go… → Browse…**. Choose `src\DecisionTree.Excel\bin\Debug\net48\publish\DecisionTreeStudio64-packed.xll` and click **OK**.
   - If Excel says the file is blocked: right-click the `.xll`, choose **Properties**, tick **Unblock**, then **OK**.
3. A **DecisionTree Studio** tab appears on the ribbon. Click **Show Panel**.

| # | Check | How | Pass when |
|---|---|---|---|
| 1 | Works offline | Turn Wi-Fi off, restart Excel and click **Show Panel**. | The panel opens and the log shows `ready`. |
| 2 | One pane per window | Open two workbooks and click **Show Panel** in each. Close one of them. | Each window has its own pane. The log shows `Removed pane` for the closed one. |
| 4 | Selection mapping | Click **Render 500 nodes**. Click some cells, then click individual shapes. | The panel's *Selection* line shows `cells …` or `shape DTS_spike_N…` each time. |
| 5 | Ctrl+Z | Select an empty cell, click **OnUndo test**, then press **Ctrl+Z**. | The cell goes back to empty and the log shows `Ctrl+Z ran DTS_SpikeUndo`. |
| 7 | Clicking while editing a cell | Double-click a cell so you're typing in it. Click **Deferred command** in the panel, then press **Enter**. | The log shows `clicked` first and `ran` only after Enter. Nothing crashes. |

Send me the panel log, or screenshots, for each manual check.

Log file: `%LOCALAPPDATA%\DecisionTreeStudio\logs\dts-<date>.log`
