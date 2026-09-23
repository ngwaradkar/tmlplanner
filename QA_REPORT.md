# Plantest QA report — 23 September 2026

## Result

**Frontend/performance revision: passed the automated checks listed below after four targeted fixes.**

**Overall production readiness: conditional.** The pre-existing missing-part allocation behavior remains open. Browser visual acceptance, authenticated OneDrive access, and actual Telegram delivery were not tested.

The delivered ZIP was extracted into an isolated QA folder. Tests used disposable databases and the repository's sample workbooks. GitHub was not modified. All Telegram delivery tests used mocks; no messages were sent.

Original ZIP SHA-256: `015d760e1ee28e23d42707460e42d5c4d767a6b1eff8654344f77304440a296d`

## Defects found and corrected

| ID | Severity | Reproduction and observed result | Correction | Retest |
|---|---|---|---|---|
| QA-01 | High | Synchronize a workbook, clear report buffers as reset does, then synchronize the same bytes. The app reported success but restored 0 of 6 buffers. | An unchanged-file shortcut now also requires the expected buffers to exist. A sync revision counter triggers refresh when identical bytes restore missing buffers. | Passed |
| QA-02 | High | Set report generation to paused with an older snapshot still present, then trigger a quarter-hour timer. The timer dispatched the cached bundle in a mocked test. | Input edits/reset invalidate snapshots and prepared exports. Auto-send requires an active report and no current sync error. | Passed |
| QA-03 | Medium | Successful sync → temporary download error → successful download of unchanged bytes. The old sync error remained set. | The unchanged-success path clears the error; the monitor refreshes after recovery. | Passed |
| QA-04 | Medium | Upload report A, then report B with the same filename and byte count. The second upload was ignored and A remained in memory. | Upload identity uses content hashing instead of filename and size alone. | Passed using the actual upload-processing path |

QA-01 through QA-03 concern the new cache/timer behavior. QA-04 was inherited from the original upload logic and was also corrected here.

## Automated coverage

| Area | Checks | Result |
|---|---:|---|
| Main screens | 9 | Passed |
| Excel preparation and workbook readability | 13 | Passed |
| Other workflow groups: navigation caching, stock invalidation, two line subviews, search/export isolation, stock forms, reset, empty float, concurrent dispatch claims, retry | 10 | Passed |
| Included unit/regression tests | 12 | Passed |
| Same-name/same-size upload replacement | 1 | Passed after correction |
| Randomized FIFO and backflush comparisons against original engine | 25 scenarios | Identical results and balances |
| Sample master-workbook parser comparisons against original loaders | 7 | Matched |

The 32 workflow checks above were executed through Streamlit's application test runner, with mock transport for the two dispatch checks. Exports were opened with OpenPyXL and checked for readable worksheets. This does not constitute a cell-by-cell audit of every exported business total.

### Screens

Overview; Summary & Excel Reports; Cockpit & Wiring Shortages; TCF1; TCF2; Total Float & Search; Quality Holds; Control Panel; Telegram Dispatcher.

### Export paths

Production matrix; hourly tracker; master summary; all cockpit/wiring parts; critical cockpit/wiring shortages; TCF1 blocked; TCF2 blocked; total float with pivots; PBS quality holds; paint-shop quality holds; BOM; TCF1 ready; TCF2 ready.

### State and failure cases

- Unchanged navigation reuses the calculation timestamp.
- Changing stock changes the input fingerprint and recalculates the report.
- Searching for `[` is treated literally and does not crash.
- Changing a float filter invalidates its prepared export without recalculating the full report.
- Filters survive page navigation.
- Both stock forms save successfully and invalidate old report snapshots.
- Reset zeros engine stocks, pauses report generation, and removes the old snapshot.
- An empty vehicle float is handled without an application exception.
- An incomplete synchronized workbook leaves prior buffers intact.
- Two concurrent mocked dispatchers send each report once.
- A partially failed mocked dispatch retries the failed report without resending already successful reports.

## Performance observations

In this local application-test environment, the recorded nine page interactions took **0.266–0.497 seconds** each after initial setup. These figures are script execution timings, not browser paint timings or production-server guarantees. Excel generation is deferred until Prepare is clicked, and filter/navigation checks confirmed report-snapshot reuse.

## Open findings and limits

1. **High — inherited allocation issue, confirmed again:** with a complete BOM, positive engine/wiring stock, and an empty cockpit stock dictionary, the engine still returns `Ready for TCF`. Missing required part-level stock should become `Stock unknown` or blocked. The performance revision retains that business behavior; it is not fixed in this QA patch.
2. The existing fixed Harrier EV starting clearance and other previously reviewed business-rule defaults remain unchanged. Calculation equivalence to the original code does not establish that these defaults are operationally correct.
3. Browser layout, mobile appearance, accessibility, and visual contrast were not inspected in a real browser. Browser installation was unavailable in the earlier setup, so only Streamlit element/execution checks are claimed.
4. OneDrive synchronization was tested with controlled download responses. Authentication, network latency, and real SharePoint policies were not exercised.
5. Telegram automation still needs an active dashboard session; it is not an independent background service. Live API delivery was not tested.
6. Multi-user stock editing/conflict resolution was not certified. The concurrency test covers Telegram delivery claims only.

## Corrected package

Use the latest `plantest-design-speed-update.zip` supplied with this report. Runtime corrections are in `app.py` and `dashboard_ui.py`; `tests/test_sync_regressions.py` adds regression coverage. The package also contains this report and the workflow result log.

Extract all files into your existing project folder after making a backup. Keep your Excel reports and database. Follow `README_UPDATE.md` for installation, including the pinned dependencies.

Run the included tests with:

```bash
python -m unittest discover -s tests -v
```

## Workflow result record

| Check | Result |
|---|---|
| Page: Overview | PASS |
| Page: Summary & Excel Reports | PASS |
| Export: prepare_export_prod_matrix_btn | PASS |
| Export: prepare_export_hourly_prod_btn | PASS |
| Export: prepare_export_summary_report | PASS |
| Page: Cockpit & Wiring Shortages | PASS |
| Export: prepare_export_cockpit_wiring_tab_all_parts | PASS |
| Export: prepare_export_cockpit_wiring_tab_critical_only | PASS |
| Page: TCF1 Line | PASS |
| Export: prepare_dl_blocked_tcf1 | PASS |
| Page: TCF2 Line | PASS |
| Export: prepare_dl_blocked_tcf2 | PASS |
| Page: Total Float & Search | PASS |
| Export: prepare_export_float_details_All | PASS |
| Page: Quality Holds | PASS |
| Export: prepare_export_pbs_quality_holds | PASS |
| Export: prepare_export_paintshop_quality_holds | PASS |
| Page: Control Panel | PASS |
| Export: prepare_download_bom | PASS |
| Page: Telegram Dispatcher | PASS |
| Navigation reuses calculation | PASS |
| Stock change invalidates calculation | PASS |
| Export: prepare_dl_ready_tcf1 | PASS |
| Subview: TCF1 Line | PASS |
| Export: prepare_dl_ready_tcf2 | PASS |
| Subview: TCF2 Line | PASS |
| Search/filter/export cache isolation | PASS |
| Stock forms and invalidation | PASS |
| Reset prevents stale snapshot reuse | PASS |
| Empty vehicle float | PASS |
| Concurrent Telegram claims (mocked) | PASS |
| Partial Telegram failure retry (mocked) | PASS |
