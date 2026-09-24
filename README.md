# qPCR studio

A local desktop workspace for raw qPCR data, ΔΔCt analysis, and customizable figures. Based on the supplied `qpcr_analyzer.html`; the original is preserved in `baseline.html`.

## Run

The packaged Windows app includes its Python engine and needs no separate Python installation. To develop from source, install Node.js and Python 3.12, then run:

```powershell
py -3.12 -m venv .venv
.venv\Scripts\python.exe -m pip install -r python/requirements.txt
npm install
npm start
```

Create a Windows portable app with `npm run build:python` followed by `npm run dist`. The executable is written to `dist/v1.1.0/`. Run `.venv\Scripts\python.exe python/verify_frozen.py` to check PNG/SVG/PDF support in the standalone engine before packaging.

## Workflow

1. Import CSV, XLSX, or XLS files. Use Sample, Target/Gene, and Ct/Cq column headings. One file represents one independent biological trial; technical wells are averaged within a trial and condition. Blank and nonnumeric Ct/Cq cells are skipped.
2. Choose targets, a reference gene, and a calibrator. Map sample names to conditions and run analysis. Missing reference/calibrator measurements produce missing results rather than zeroes.
3. Inspect calculations and export results as CSV.
4. In Figure studio choose bars, grouped bars, points, boxes, or paired lines. Customize colors, opacity, point size, gridlines, error bars, metric, scale, title, and threshold. Matplotlib figures use Seaborn styling and fit the available width (up to 820 logical pixels), with compressed bars and 45° labels. Long labels are shortened with an ellipsis. Export 300 DPI PNG, editable-text SVG, or vector PDF.
5. Save a project (Ctrl+S) to preserve raw data, exclusions, mappings, graph overrides, and figure controls. Open the JSON project to continue. Save before closing; there is no automatic save.

All data and Excel parsing stay local. No CDN or network connection is required after installation. Example data is explicitly loaded with Explore an example.

## Analysis details

Target and reference technical replicates are averaged per condition within each imported trial. ΔCt = target mean − reference mean. ΔΔCt uses the calibrator from the same trial. Expression = base^(−ΔΔCt), with base 2 by default. A custom base applies equally to target and reference; it is not an efficiency-corrected model with separate efficiencies.

Pandas/NumPy perform normalization. SciPy statistical tests use the selected graph metric and independent trial-level observations. Paired tests match trial IDs. Plot overrides are optional and are excluded from statistical tests by default. Tests are exploratory and do not apply multiple-comparison corrections. ANOVA requires two observations per condition; undefined tests (e.g. identical constant groups) show an explanation. The app cannot determine whether imported files are truly independent biological replicates. Log axes omit nonpositive observations and report the omitted count. Results are recalculated on navigation after raw data or analysis settings change.

The Electron main process owns a persistent Python subprocess and exchanges structured JSON over standard input/output. The sandboxed renderer can call only the predefined analysis, statistics, render, and health operations. No server, network, arbitrary Python execution, or user-provided shell command is involved. Work runs asynchronously; stale responses are discarded when data or figure settings change. Python exits when the desktop app closes.

## Development

`npm run build:source` rebuilds the renderer from the preserved baseline plus `scripts/build-source.cjs`, `scripts/enhancements.txt`, and `scripts/python-ui.txt`. Styling is in `src/studio.css`. The scientific engine is `python/engine.py`; the original JavaScript scientific engine is removed from the generated app. `npm run test:python` checks calculations and figure formats; `npm test` runs desktop integration tests. Set `PACKAGED_TEST=1` to test `dist/v1.1.0/win-unpacked` with its bundled Python engine. Existing project JSON files remain supported.
