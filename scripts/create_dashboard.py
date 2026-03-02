"""
create_dashboard.py
-------------------
Reads data/sample_telemetry.xlsx and produces an interactive HTML dashboard
(docs/dashboard.html) containing all 8 charts described in the requirement.

LAPTIME conversion: MM.SS.mmm → SSS.mmm
  e.g. "1.23.456" (1 min 23.456 s) → 83.456 seconds
"""

import os
import re
import json
import pandas as pd

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_FILE = os.path.join(REPO_ROOT, "data", "sample_telemetry.xlsx")
OUT_FILE = os.path.join(REPO_ROOT, "docs", "dashboard.html")

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def parse_laptime(raw: str) -> float:
    """Convert MM.SS.mmm string to total seconds SSS.mmm (float)."""
    raw = str(raw).strip()
    # Expected format: MM.SS.mmm  (e.g. "1.23.456" or "0.58.123")
    m = re.match(r'^(\d+)\.(\d{2})\.(\d{3})$', raw)
    if m:
        minutes = int(m.group(1))
        seconds = int(m.group(2))
        millis = int(m.group(3))
        return round(minutes * 60 + seconds + millis / 1000.0, 3)
    # Fallback: try to interpret as a plain float (already in seconds)
    try:
        return float(raw)
    except ValueError:
        return None


def load_data(path: str) -> pd.DataFrame:
    # header=None so we control column names; skiprows=0 keeps row 1 visible
    df = pd.read_excel(path, header=None, skiprows=2)  # skip rows 1-2 (headers/units)
    df.columns = range(1, len(df.columns) + 1)         # 1-based column index

    # Drop rows where column 1 (LAP) is empty / NaN
    df = df[df[1].notna()].copy()
    df[1] = df[1].astype(int)

    # Convert LAPTIME (col 4) from MM.SS.mmm → SSS.mmm
    df["LAPTIME"] = df[4].apply(parse_laptime)

    return df


# ---------------------------------------------------------------------------
# Chart builders (return Plotly figure dicts)
# ---------------------------------------------------------------------------

def make_fig(laps, y_traces, title, y_title):
    """Return a minimal Plotly figure dict."""
    data = []
    colors = ["#1f77b4", "#ff7f0e", "#2ca02c"]
    for idx, (label, y_vals) in enumerate(y_traces):
        data.append({
            "type": "scatter",
            "mode": "lines+markers",
            "name": label,
            "x": laps,
            "y": y_vals,
            "line": {"color": colors[idx % len(colors)]},
            "marker": {"size": 6},
        })
    layout = {
        "title": {"text": title, "font": {"size": 16}},
        "xaxis": {"title": "LAP", "dtick": 1},
        "yaxis": {"title": y_title},
        "legend": {"orientation": "h", "y": -0.2},
        "margin": {"l": 60, "r": 20, "t": 50, "b": 80},
        "paper_bgcolor": "#ffffff",
        "plot_bgcolor": "#f9f9f9",
    }
    return {"data": data, "layout": layout}


def build_charts(df: pd.DataFrame) -> list[dict]:
    laps = df[1].tolist()

    charts = [
        # Chart 1 – LAPTIME
        make_fig(laps,
                 [("LAPTIME", df["LAPTIME"].tolist())],
                 "Chart 1 – Lap Time", "LAPTIME (s)"),
        # Chart 2 – FUEL/LAP
        make_fig(laps,
                 [("FUEL/LAP", df[5].tolist())],
                 "Chart 2 – Fuel per Lap", "FUEL/LAP (L)"),
        # Chart 3 – Water Temperature
        make_fig(laps,
                 [("T wat, Max", df[6].tolist()),
                  ("T wat, avg", df[7].tolist())],
                 "Chart 3 – Water Temperature", "Temperature (°C)"),
        # Chart 4 – Oil Temperature
        make_fig(laps,
                 [("T oil, Max", df[8].tolist()),
                  ("T oil, avg", df[9].tolist())],
                 "Chart 4 – Oil Temperature", "Temperature (°C)"),
        # Chart 5 – Oil Pressure
        make_fig(laps,
                 [("Poil, Min", df[10].tolist()),
                  ("Poil, avg", df[11].tolist())],
                 "Chart 5 – Oil Pressure", "Pressure (bar)"),
        # Chart 6 – Alternator Voltage
        make_fig(laps,
                 [("Alt V, Min", df[12].tolist()),
                  ("Alt V, avg", df[13].tolist())],
                 "Chart 6 – Alternator Voltage", "Voltage (V)"),
        # Chart 7 – Fuel Pressure
        make_fig(laps,
                 [("Pfuel, Min", df[14].tolist()),
                  ("Pfuel, avg", df[15].tolist())],
                 "Chart 7 – Fuel Pressure", "Pressure (bar)"),
        # Chart 8 – Lift Pump Current
        make_fig(laps,
                 [("LiftPump1 current, avg", df[16].tolist()),
                  ("LiftPump2 current, avg", df[17].tolist())],
                 "Chart 8 – Lift Pump Current", "Current (A)"),
    ]
    return charts


# ---------------------------------------------------------------------------
# HTML generator
# ---------------------------------------------------------------------------

PLOTLY_CDN = "https://cdn.plot.ly/plotly-2.32.0.min.js"


def render_html(charts: list[dict]) -> str:
    divs = []
    scripts = []
    for i, fig in enumerate(charts):
        div_id = f"chart{i}"
        divs.append(f'<div id="{div_id}" class="chart-box"></div>')
        data_json = json.dumps(fig["data"])
        layout_json = json.dumps(fig["layout"])
        scripts.append(
            f"Plotly.newPlot('{div_id}', {data_json}, {layout_json}, "
            f"{{responsive: true, displayModeBar: false}});"
        )

    divs_html = "\n    ".join(divs)
    scripts_html = "\n    ".join(scripts)

    return f"""<!DOCTYPE html>
<html lang="it">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Telemetry Dashboard</title>
  <script src="{PLOTLY_CDN}"></script>
  <style>
    body {{
      font-family: "Segoe UI", Arial, sans-serif;
      background: #f0f2f5;
      margin: 0;
      padding: 16px;
    }}
    h1 {{
      text-align: center;
      color: #1F4E79;
      margin-bottom: 24px;
    }}
    .grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(480px, 1fr));
      gap: 20px;
    }}
    .chart-box {{
      background: #fff;
      border-radius: 8px;
      box-shadow: 0 2px 6px rgba(0,0,0,.12);
      padding: 8px;
      min-height: 320px;
    }}
  </style>
</head>
<body>
  <h1>Telemetry Dashboard</h1>
  <div class="grid">
    {divs_html}
  </div>
  <script>
    {scripts_html}
  </script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print(f"Reading {EXCEL_FILE} …")
    df = load_data(EXCEL_FILE)
    print(f"  Loaded {len(df)} laps.")

    charts = build_charts(df)

    os.makedirs(os.path.dirname(OUT_FILE), exist_ok=True)
    html = render_html(charts)
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Dashboard written to: {OUT_FILE}")
