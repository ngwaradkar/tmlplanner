# Plantest dashboard design and speed update

This is a replacement-file package for the existing `ngwaradkar/plantest` project, based on commit `407aa2a3accd7f8b67e35ff14bd66dca98257ece`. No GitHub changes were made.

## Install on your computer

1. Stop Streamlit and make a backup copy of your existing project folder.
2. Extract this ZIP into that project folder. Replace `app.py`, `data_loader.py`, `allocation_engine.py`, and `requirements.txt`. Add the new `dashboard_ui.py` and `tests` folder. **All five application/dependency files are required.**
3. Keep your existing Excel reports, `TEST` folder, `.streamlit` settings, and `clear_to_build.db`. They are not included in this update.
4. Open a terminal in your project folder and run:

   ```bat
   python -m pip install -r requirements.txt
   python -m streamlit run app.py
   ```

The package was validated with Python 3.12. Dependency versions are pinned to the tested environment. A fresh virtual environment is recommended if your existing environment has conflicting packages.

## What changed

- Sidebar navigation: Overview, Summary & Excel Reports, Cockpit & Wiring Shortages, TCF1, TCF2, Total Float & Search, Quality Holds, Control Panel, and Telegram.
- Compact header and KPI cards, system fonts, light/dark appearance, mobile wrapping, and a new Overview with production counts, line status, blocking reasons, and vehicle search.
- Independent page rendering. Opening a line screen no longer builds all other screens or their Excel files.
- A session report snapshot keyed to report inputs, BOM, stock quantities, and model settings. Navigation and filters reuse the calculation; changed inputs invalidate it.
- Excel files are built only after **Prepare** is clicked, then offered for download. Prepared files are reused until their inputs change; downloads do not trigger a full rerun.
- Master-workbook sheets are read once and normalized directly, removing the intermediate Excel write/read cycle.
- File parsing caches include file modification information or buffer content, and separate different parser functions.
- A periodic Streamlit fragment checks for file/sync changes without redrawing the dashboard on every clock tick. Unchanged OneDrive workbooks reuse the existing report.
- Engine and Nova stocks use compact batch-entry forms. Click **Save & Recalculate**, then open a report page. Unsaved form edits are not applied.
- Float search has a compact column view, optional detail columns, literal searches, and filters that persist across pages.
- FIFO processing uses a BOM lookup dictionary. Its ordering and stock deduction rules are retained.
- Missing entire stock inputs produce a clear message rather than a `None.copy()` error. Stock sheet names tolerate differences in letter case.
- Incomplete OneDrive workbooks do not replace the current report set. The five core sheets are required; hourly production is optional.

## Telegram configuration

The exposed hard-coded token has been removed from this package. Rotate that token before further use. Configure the replacement in **Telegram Dispatcher → Telegram Bot Settings**, through environment variables, or in your existing `.streamlit/secrets.toml`:

```toml
TELEGRAM_BOT_TOKEN = "YOUR_NEW_TOKEN"
TELEGRAM_CHAT_ID = "YOUR_CHAT_ID"
```

Existing database-saved Telegram settings take precedence, so replace an old token there using the settings screen. No Telegram messages were sent during validation.

Automatic sending remains opt-in and requires an active Streamlit dashboard session with a generated report. It is not a separate always-running background service. Per-report database claims reduce duplicate dispatch across sessions and allow failed reports to retry during the scheduled minute. A process crash during an in-flight send is left pending rather than blindly sending a duplicate.

## Validation and limits

- All nine main screens executed successfully with the repository sample reports.
- Master, hourly, production matrix, cockpit/wiring, line-ready, line-blocked, float/pivot, and quality-hold exports were exercised.
- Seven sample-workbook parser results matched the original parser results, including the vehicle-code paint summary.
- Navigation reused the report snapshot, and changing starting stock invalidated it. Stock form saves, resets, and line float searches were checked.
- A synthetic 2,000-vehicle FIFO comparison produced identical allocation results and final balances. One local run took about 0.10 seconds for the updated engine versus 0.37 seconds for the original. This is an engine benchmark, not a promise about whole-app performance on your computer.
- Included regression tests can be run with `python -m unittest discover -s tests -v`.
- Validation used Streamlit's application test runner; desktop/mobile pixel layout was not verified in a browser. Authenticated OneDrive downloads and live Telegram delivery were not exercised.

This update focuses on design and execution speed. Existing business rules, EV stock defaults, and part-level missing-stock behavior have not been comprehensively revised. The earlier accuracy/security review should still be addressed separately before relying on the app as the sole production clearance authority.

To revert, stop the app and restore your backup code and dependency environment. Keep any stock/BOM changes you made in the database only if you intend to retain them.


## QA revision — 23 September 2026

Four fixes were added after deeper QA: restoring missing buffers from an unchanged synced workbook, preventing stale automatic sends while reports are paused or invalidated, clearing recovered sync errors, and detecting uploads with unchanged names/sizes but different content. See `QA_REPORT.md` for results, reproducible findings, and remaining limits. There are now 12 included unit/regression tests.
