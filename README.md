# H-DMAPlot

**GUI tool for Dynamic Mechanical Analysis (DMA) data visualization and processing**

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python)
![License](https://img.shields.io/badge/License-MIT-green)
![Version](https://img.shields.io/badge/Version-1.4-informational)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?logo=windows)

H-DMAPlot is a Python desktop application for visualizing and processing Dynamic Mechanical Analysis data. It is part of the **H-SciTools** scientific software suite, developed to support materials engineering research.

---

## Features

**Data loading**
- Import `.txt` files exported by DMA instruments (two-line header format: column names + units)
- Automatic decimal separator detection (comma or period)
- Automatic encoding detection (UTF-8 / Latin-1)
- Multiple samples loaded simultaneously

**Axis configuration**
- Free column selection for the X axis (typically temperature)
- Multiple column selection for the Y axis (E', E'', tan δ, η*, etc.)
- Curve overlay mode with independent Y axes (twin axes)
- Automatic mapping of DMA column names to human-readable descriptions

**Data processing**
- Curve smoothing: Savitzky-Golay and Moving Average, with window size and polynomial degree controls
- Interval trimming by real column value (e.g. temperature range)
- Numerical offset per column and per sample
- Individual reset of all applied treatments

**Style and appearance**
- Color, line width, and line style configurable per sample
- Independent font size control for title, axis labels, ticks, legend, and annotations
- Configurable grid (X axis, Y axis, fixed or automatic interval)
- Manual X-axis limits

**Graph interactivity**
- Dynamic crosshair with floating tooltip on cursor movement
- Click to pin markers with coordinates and interpolated values
- Draggable pinned annotations with arrows connected to data points
- Double-click directly on titles or axis labels to edit them inline
- Automatic peak detection with controls for prominence, minimum distance, and maximum count

**Export**
- PNG (300 DPI)
- PDF with basic report (via ReportLab)
- Excel `.xlsx` with per-sample data sheets and a statistical summary tab

---

## Interface

The layout is split into a scrollable control panel (left) and the plot area (right):

```
┌────────────────────┬──────────────────────────────────────────┐
│  H-DMAPlot         │                                          │
│  ─────────────     │                                          │
│  1. Files          │           Plot area                      │
│  2. Axes           │       (interactive matplotlib)           │
│  3. Style          │                                          │
│  4. Smoothing      │                                          │
│  5. Data treatment │                                          │
│  6. Grid / Fonts   │                                          │
│  7. Peak detection │                                          │
│  8. Export         │                                          │
└────────────────────┴──────────────────────────────────────────┘
```

---

## Requirements

Python 3.10 or higher.

```
numpy
pandas
matplotlib
scipy
reportlab
openpyxl
```

Install dependencies:

```bash
pip install numpy pandas matplotlib scipy reportlab openpyxl
```

---

## Usage

### Running directly

```bash
python H-DMAPlot_v1_4.py
```

### Packaging with PyInstaller

```bash
pyinstaller --onefile --windowed --icon=DMAPlot.ico H-DMAPlot_v1_4.py
```

The resulting executable in `dist/` can be distributed without a Python installation.

> The `DMAPlot.ico` icon file is optional. If not found, the application starts normally without an icon.

---

## Supported file format

H-DMAPlot reads `.txt` files with the following structure:

```
Ts    t     f      E'     E''    tan_delta
[°C]  [s]   [Hz]   [MPa]  [MPa]  []
25.1  0.0   1.0    1850   45.2   0.0244
25.3  0.5   1.0    1848   44.9   0.0243
...
```

**Line 1:** space-separated column names  
**Line 2:** units in brackets  
**Subsequent lines:** numeric data

Lines starting with `&` or `#` are ignored. Sequential index columns (`index [#]`) are automatically discarded.

---

## Recognized DMA columns

| Column name  | Physical quantity                         |
|--------------|-------------------------------------------|
| `Ts`         | Sample temperature                        |
| `t`          | Time                                      |
| `f`          | Excitation frequency                      |
| `E'`         | Storage Modulus (elastic)                 |
| `E''`        | Loss Modulus (viscous)                    |
| `E*`         | Complex Modulus                           |
| `tan_delta`  | Loss Factor / Damping Factor              |
| `eta'`       | Dynamic Viscosity (real component)        |
| `eta*`       | Complex Viscosity                         |
| `D'`, `D''`  | Storage and Loss Compliance               |

---

## Code structure

```
H-DMAPlot_v1_4.py
│
├── _parse_txt()            # TXT file parser for DMA data
├── Amostra                 # Per-sample data model
│   └── get_serie()         # Returns series with trimming, offset, and smoothing applied
├── PlotCanvas              # Matplotlib plot management
│   ├── redesenhar()        # Main rendering method
│   ├── _desenhar_picos()   # Peak detection and annotation
│   └── crosshair           # Interactive cursor and pinned markers
├── PainelEsquerdo          # Control panel UI (tkinter)
├── DialogCorte             # Interval trim dialog
├── DialogOffset            # Numerical offset dialog
├── DialogSelecionarAmostraColuna  # Sample/column selector dialog
└── App                     # Main application controller
```

---

## Part of the H-SciTools suite

| Tool         | Purpose                                         |
|--------------|-------------------------------------------------|
| H-DMAPlot    | Dynamic Mechanical Analysis (DMA)               |
| H-TGAPlot    | Thermogravimetric Analysis (TGA)                |
| H-DRXPlot    | X-Ray Diffraction (XRD)                         |
| H-AnodPlot   | Electrochemical anodization curves              |

---

## Author

**Carlos Henrique Amaro da Silva**  
M.Sc. in Materials Technology and Industrial Processes — Universidade Feevale (2025)  
B.Sc. in Chemical Engineering (2023)

Research focus: surface treatments, anodization, and electrodeposition with biomedical applications.

GitHub: https://github.com/Leindsher

---

## License

This project is licensed under the [MIT License](LICENSE).
