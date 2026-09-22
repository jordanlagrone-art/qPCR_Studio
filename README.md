# qPCR studio

A local desktop workspace for raw qPCR data, ΔΔCt analysis, and customizable figures. Based on the supplied `qpcr_analyzer.html`; the original is preserved in `baseline.html`.

## Run

Install Node.js, then run `npm install` and `npm start`. Create a Windows portable app with `npm run dist`. The executable is written to `dist/`.

## Workflow

1. Import CSV, XLSX, or XLS files. Use Sample, Target/Gene, and Ct/Cq column headings. One file represents one independent biological trial; technical wells are averaged within a trial and condition. Blank and nonnumeric Ct/Cq cells are skipped.
2. Choose targets, a reference gene, and a calibrator. Map sample names to conditions and run analysis. Missing reference/calibrator measurements produce missing results rather than zeroes.
3. Inspect calculations and export results as CSV.
4. In Figure studio choose bars, grouped bars, points, boxes, or paired lines. Customize colors, opacity, point size, gridlines, error bars, metric, scale, title, and threshold. Figures fit the available width (up to 820 logical pixels), with compressed bars and 45° labels instead of growing wider as groups are added. Long labels are shortened with an ellipsis. PNG export is at least 2× resolution.
5. Save a project (Ctrl+S) to preserve raw data, exclusions, mappings, graph overrides, and figure controls. Open the JSON project to continue. Save before closing; there is no automatic save.

All data and Excel parsing stay local. No CDN or network connection is required after installation. Example data is explicitly loaded with Explore an example.

## Analysis details

Target and reference technical replicates are averaged per condition within each imported trial. ΔCt = target mean − reference mean. ΔΔCt uses the calibrator from the same trial. Expression = base^(−ΔΔCt), with base 2 by default. A custom base applies equally to target and reference; it is not an efficiency-corrected model with separate efficiencies.

Statistical tests use the selected graph metric and independent trial-level observations. Paired tests match trial IDs. Plot overrides are optional and are excluded from statistical tests by default. Tests are exploratory and do not apply multiple-comparison corrections. The app cannot determine whether imported files are truly independent biological replicates. Log axes omit nonpositive observations. Results are recalculated on navigation after raw data or analysis settings change.

## Development

`npm run build:source` rebuilds the renderer from the preserved baseline plus `scripts/build-source.cjs` and `scripts/enhancements.txt`. Styling is in `src/studio.css`. `npm test` runs desktop integration tests.
