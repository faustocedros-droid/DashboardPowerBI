# DashboardPowerBI

Interactive telemetry dashboard built from an Excel data source,
delivered both as a Power BI report and as a standalone HTML preview.

## Repository Layout

```
data/
  sample_telemetry.xlsx   ← source Excel file (columns 1–17, data from row 3)
docs/
  powerbi_setup.md        ← step-by-step Power BI setup guide
  dashboard.html          ← self-contained HTML preview (generated)
scripts/
  generate_sample_data.py ← creates / regenerates sample_telemetry.xlsx
  create_dashboard.py     ← reads the Excel file and writes dashboard.html
```

## Eight Charts

| # | Title                  | X-axis | Y-axis                                     |
|---|------------------------|--------|--------------------------------------------|
| 1 | Lap Time               | LAP    | LAPTIME (col 4, converted to SSS.mmm)      |
| 2 | Fuel per Lap           | LAP    | FUEL/LAP (col 5)                           |
| 3 | Water Temperature      | LAP    | T wat, Max (col 6) · T wat, avg (col 7)    |
| 4 | Oil Temperature        | LAP    | T oil, Max (col 8) · T oil, avg (col 9)    |
| 5 | Oil Pressure           | LAP    | Poil, Min (col 10) · Poil, avg (col 11)    |
| 6 | Alternator Voltage     | LAP    | Alt V, Min (col 12) · Alt V, avg (col 13)  |
| 7 | Fuel Pressure          | LAP    | Pfuel, Min (col 14) · Pfuel, avg (col 15)  |
| 8 | Lift Pump Current      | LAP    | LiftPump1 current, avg (col 16) · LiftPump2 current, avg (col 17) |

## Quick Start (HTML preview)

```bash
pip install pandas openpyxl plotly
python scripts/generate_sample_data.py   # creates data/sample_telemetry.xlsx
python scripts/create_dashboard.py       # creates docs/dashboard.html
# Open docs/dashboard.html in a browser
```

## Power BI Setup

See **[docs/powerbi_setup.md](docs/powerbi_setup.md)** for the full step-by-step guide,
including the Power Query M formula for the LAPTIME conversion
(`MM.SS.mmm` → `SSS.mmm`).

## Contributing

1. Fork the repository.
2. Create a feature branch: `git checkout -b feature/my-dashboard`.
3. Commit your changes and open a pull request.
