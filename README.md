# DLS-Horiba

Desktop viewer for `.nsz` files from the **Horiba SZ-100** dynamic light scattering (DLS)
instrument. It reads the binary files directly, without the instrument software, plots the
correlation function and the particle size distribution, compares several samples, and
exports everything to Excel.

> **Unofficial.** This project is not affiliated with or endorsed by Horiba. The `.nsz`
> format was worked out by inspecting files, and the extracted values were validated against
> the CSV exported by the instrument software on a single sample. Check the results against
> your own exports before relying on them, especially with other instrument models.

## Features

- Opens `.nsz` files from a dialog, a whole folder, the command line or drag-and-drop
  (needs `tkinterdnd2`).
- Loads many samples and overlays the selected ones.
- **Autocorrelogram** tab: measured correlation function, Horiba fit and residuals.
- **Distribution** tab: size distribution by intensity or by number, log or linear x axis,
  optional cumulative curve.
- **Parameters** tab: date, solvent, scattering angle, wavelength, temperature, viscosity,
  refractive index, peak and mean sizes, SD, D10/D50/D90, mode, median, Z-average, PDI and span.
- **Excel export**: one `.xlsx` per sample with sheets for the ACF (data, fit, residuals),
  the intensity and number distributions and the parameters.

The interface is in Italian.

## Run

```
pip install -r requirements.txt
python nsz_viewer.pyw [file.nsz ...]
```

On Windows, `pythonw nsz_viewer.pyw` starts it without a console window.

## Notes on the format

An `.nsz` file is an OLE compound file. The viewer reads these streams:

| Stream | Content |
|---|---|
| `Object2`, `Object3` | Correlation function and its lag times (trailing baseline channels are trimmed) |
| `Object10`, `Object7` | Fit and residuals; the header holds the number of initial channels excluded from the fit |
| `Object5`, `Object4`, `Object31` | Size bins (84), intensity distribution, cumulative distribution |
| `Base1`, `Base34`, `Base38` | Calculated results, measurement parameters, sample name |

Two details matter when interpreting the values: the diameter fields (`CalcMean`, `CalcPeakPos`,
`CalcPerDiameter`, `Cum_fMean`, ...) are stored as **radii**, so the viewer multiplies them by 2;
and the PDI is taken from `Cum_fPi`, because `CalcTotalPI` does not match the software's value.

## Data

This repository contains code only: `.gitignore` excludes `*.nsz`, `*.csv` and `*.xlsx`, so your
measurements and exports stay on your machine.

## Requirements

Python 3, numpy, matplotlib, olefile, openpyxl; `tkinterdnd2` is optional.
