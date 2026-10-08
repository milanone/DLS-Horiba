# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.
General rules for all of Francesco's projects (stack, code conventions, repository rules, git/GitHub, the
Drive mirror, how he works) are in `..\CLAUDE.md`; this file has the project-specific details.

## Purpose

Tkinter viewer for the `.nsz` files of the **Horiba SZ-100** dynamic light scattering (DLS) instrument: correlation
function, size distribution, parameters table, multi-sample overlay, Excel export, figure save/edit. The format was
reverse-engineered, so the parser is checked against the CSV exported by the instrument software.

## Running / tests

```bash
pythonw nsz_viewer.pyw [file.nsz ...]
py -m unittest discover -s tests -v
```

Sample data: `example data/lipo_bup_MLV_0minson_scattering_0729.nsz` and the software CSV of the same measurement
`..._0731.csv` (gitignored: real data, never commit). Tk tests need a screen; tests that need the sample are skipped
without it. The two `screenshot*.png` are the real program with that sample (`ImageGrab` on the window, topmost).

## Architecture

Single file `nsz_viewer.pyw`: parser functions (`f32`, `parse_props`, `load_nsz` / `_read_nsz`, `dump_nsz` for
debugging) and the class `NSZViewer` (sidebar with sample list, a `ttk.Notebook` with Autocorrelogram / Distribution /
Parameters tabs, two matplotlib `Figure`s `fig_acf` and `fig_size`). `self.samples = {path: data dict}`, `self.selected`
= paths shown. `_plot_acf` / `_plot_size` redraw on every control change, `_update_params_table` fills the table (up to 5
samples), `_write_xlsx` writes one workbook per sample, `_to_number_dist` converts intensity to number distribution
(Rayleigh, N proportional to I/d^6).

## The .nsz format (OLE compound file, read with `olefile`)

| Stream | Content and how it is read |
|---|---|
| `Object3` | lag times: 16 header floats, then tau (µs); trailing repeated baseline channels are cut where tau stops increasing (43 -> 37 channels in the sample) |
| `Object2` | correlation: 3 header floats, then one value per channel. These are **g1 squared** (g2-1): the software's CSV g1 equals `sqrt(value)` to 5e-7, CSV row k = channel 7 + k |
| `Object10` | fit: header `[1][n_fit][fit_offset]`, then values; `fit_offset` (6) = initial channels excluded from the fit; also g1 squared, CSV fit row k = `sqrt(fit[k+1])` |
| `Object7` | residuals (same offset as the fit); correlated (0.83) with the CSV residual of g1, not equal |
| `Object5` | diameters: 17 header floats, then 84 bins |
| `Object4` | intensity distribution: 3 header floats, then 84 values |
| `Object31` | cumulative distribution: 4 header floats, then 84 values |
| `Base1`, `Base34`, `Base38` | calculated results, measurement parameters, sample name; `parse_props` reads `name\0` + float32 pairs |

Details that matter: diameter fields (`CalcMean`, `CalcPeakPos`, `CalcPerDiameter`, `Cum_fMean`, ...) are stored as
**radii** (x2); the PDI is `Cum_fPi` because `CalcTotalPI` does not match the software; `MeasID` holds date and time.

**Bug fixed while writing the tests (2026-10-08):** the three distribution streams were read with the wrong header
lengths (16, 2, 2 instead of 17, 3, 4), so the diameter and intensity arrays were shifted by one bin (a spurious 0.30 nm
first bin, the last bin 8510 nm dropped) and the cumulative curve was shifted by one bin against the diameters. The
header lengths above reproduce all three CSV columns exactly (to the CSV rounding).

## Verified / not verified

- Verified on **one** measurement against the software CSV: Mean, SD, Mode, Z-average, PDI; size table (diameter,
  frequency, undersize); correlation and fit (as g1 squared). See `tests/test_nsz_viewer.py::TestRealSample`.
- The correlation is plotted as stored (g2-1, label `g2(tau) - 1`); the software shows g1. Offer to switch to g1
  (`sqrt`) if Francesco wants the same curve as the software.
- Not verified: other SZ-100 software versions or instruments, two-peak distributions (`Peak2`/`Mean2` exist in the
  parser but were never seen non-empty), and the time of day: `MeasID` gave `12:22:00` where the software prints
  `12:22:39`.
- `Temp_C` is 24.95 where the software prints 25.0 (rounding of a displayed value).

## PlotStyleKit

`_load_origin_style()` loads `origin_style.py` by path (local copy, then `../PlotStyleKit`) and calls
`applica_rcparams()` at import; `sample_colors()` then uses the Origin colour cycle (black first), else the fallback
`COLORS`. `_style_figure(fig)` applies `applica_stile_origin(ax, None, set_size=False)` to every axes, enlarges the fonts
for the pane and runs one `tight_layout` (no grid with the style; the dashed grid only without PlotStyleKit). File menu:
`Save Figure Image`, `Save Figure (pickle)`, `Edit Figure...` act on the active plot tab (`_current_figure()`) through
`_copy_figure()` (pickle copy, restyled: `single` 4:3 for one panel, `double` 16:9 for two). Style changes go in
PlotStyleKit, never here. (This project exposed that `applica_stile_origin` set `AutoMinorLocator` on log axes, which
matplotlib warns about; PlotStyleKit now skips non-linear axes.)

## Editing conventions (project)
- Follow `..\CLAUDE.md`: everything in English, surgical edits. The UI was in Italian and was translated in one pass
  (menus, labels, plot titles, Excel sheet names: `Distribuzione intensità` is now `Intensity distribution`, etc.).
- Never commit anything from `example data/`.
