"""
CASRA KPI — Power BI layout mockup (synthetic sample data).

Generates a single HTML file you can open in a browser to preview a suggested
dashboard structure. This is NOT Power BI; it only illustrates layout and visuals.

Requirements:
    pip install plotly pandas

Usage:
    python mockups/casra_dashboard_mockup.py
    → opens mockups/casra_dashboard_mockup.html
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd

try:
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots
except ImportError as exc:
    raise SystemExit("Install plotly: pip install plotly pandas") from exc


OUTPUT_HTML = Path(__file__).resolve().parent / "casra_dashboard_mockup.html"

METRIC_BUCKETS = [
    "Storage Location",
    "QM Insp Type",
    "Valuation Type",
    "Batch MNGMT",
    "Serialized Profile",
    "Class MOA",
    "Unit of Measure",
    "MRP Area",
]

CHECK_LABELS = {
    "Check_SNP": "SNP",
    "Check_SLoc Missing": "SLoc Missing",
    "Check_SLoc_MRPInd": "SLoc MRP Ind",
    "Check_QMAT Extra": "QMAT Extra",
    "Check_QMAT Missing": "QMAT Missing",
    "Check_VType Extra": "VType Extra",
    "Check_VType Missing": "VType Missing",
    "Check_VType Error": "VType Error",
    "Check_Batch": "Batch",
    "Check_MOA": "MOA",
    "Check_Missing_Model": "Missing Model",
    "Check_Missing_MOA_Class": "Missing MOA Class",
    "Check_UofM": "UofM",
    "Check_MRPArea": "MRP Area",
}


def build_sample_master() -> pd.DataFrame:
    """Six months of KPI Master-style rows (decimals = rates)."""
    rows = [
        ("2025-11-15", "20251001", "20251031", 412, 0.042, 0.018, 0.031, 0.009, 0.012, 0.024, 0.011, 0.007, 0.002, 0.006, 0.162),
        ("2025-12-10", "20251101", "20251130", 385, 0.039, 0.016, 0.028, 0.008, 0.015, 0.022, 0.010, 0.008, 0.002, 0.005, 0.153),
        ("2026-01-08", "20251201", "20251231", 401, 0.037, 0.015, 0.026, 0.007, 0.011, 0.019, 0.009, 0.006, 0.002, 0.004, 0.136),
        ("2026-02-05", "20260101", "20260131", 368, 0.035, 0.014, 0.024, 0.008, 0.010, 0.017, 0.008, 0.007, 0.002, 0.005, 0.128),
        ("2026-03-04", "20260201", "20260228", 342, 0.033, 0.013, 0.022, 0.006, 0.009, 0.015, 0.007, 0.006, 0.002, 0.004, 0.117),
        ("2026-04-02", "20260301", "20260331", 279, 0.031, 0.012, 0.020, 0.005, 0.008, 0.014, 0.006, 0.005, 0.002, 0.004, 0.107),
    ]
    cols = [
        "Report Date", "Date From", "Date To", "Parts Created",
        *METRIC_BUCKETS, "Hazmat", "Total %",
    ]
    df = pd.DataFrame(rows, columns=cols)
    df["Report Date"] = pd.to_datetime(df["Report Date"])
    return df


def build_sample_detail(n: int = 40) -> pd.DataFrame:
    import random

    random.seed(42)
    parts = [f"PN-{100000 + i}" for i in range(n)]
    errors = [random.randint(0, 5) for _ in range(n)]
    checks = list(CHECK_LABELS.keys())

    data = {
        "Material Number": parts,
        "Created on": [date(2026, 3, random.randint(1, 28)) for _ in range(n)],
        "MTyp": ["HALB"] * n,
        "Plnt": [random.choice(["3000", "3100", "3200"]) for _ in range(n)],
        "Errors": errors,
    }
    for col in checks:
        data[col] = [1 if random.random() < (e / 6) else 0 for e in errors]

    return pd.DataFrame(data).sort_values("Errors", ascending=False)


def pct_text(value: float) -> str:
    return f"{value * 100:.1f}%"


def add_kpi_card(fig, row, col, title: str, value: str, subtitle: str, row_h=0.12):
    """Simulated KPI card using a shape + annotations."""
    x0, x1 = col * 0.24 + 0.02, (col + 1) * 0.24
    y0, y1 = 1.0 - row * row_h - 0.02, 1.0 - (row - 1) * row_h - 0.04

    fig.add_shape(
        type="rect", x0=x0, y0=y0, x1=x1, y1=y1,
        xref="paper", yref="paper",
        fillcolor="#f4f7fb", line=dict(color="#c5d3e8", width=1),
    )
    fig.add_annotation(
        x=(x0 + x1) / 2, y=y1 - 0.015, xref="paper", yref="paper",
        text=f"<b>{title}</b>", showarrow=False, font=dict(size=11, color="#5a6b7d"),
    )
    fig.add_annotation(
        x=(x0 + x1) / 2, y=(y0 + y1) / 2, xref="paper", yref="paper",
        text=f"<b>{value}</b>", showarrow=False, font=dict(size=22, color="#1a3a5c"),
    )
    fig.add_annotation(
        x=(x0 + x1) / 2, y=y0 + 0.012, xref="paper", yref="paper",
        text=subtitle, showarrow=False, font=dict(size=10, color="#7a8a9a"),
    )


def build_dashboard(master: pd.DataFrame, detail: pd.DataFrame) -> go.Figure:
    latest = master.iloc[-1]
    period = f"{latest['Date From']} → {latest['Date To']}"

    fig = make_subplots(
        rows=3,
        cols=2,
        row_heights=[0.38, 0.32, 0.30],
        column_widths=[0.55, 0.45],
        specs=[
            [{"type": "xy", "colspan": 2}, None],
            [{"type": "xy"}, {"type": "xy"}],
            [{"type": "table"}, {"type": "xy"}],
        ],
        subplot_titles=(
            "",
            "Error rate trend (Total %)",
            "Latest period — breakdown by category",
            "Top parts by error count (drill-down preview)",
            "Check-level flags — latest period (sample)",
        ),
        vertical_spacing=0.14,
        horizontal_spacing=0.08,
    )

    # --- Row 0: trend (full width) ---
    fig.add_trace(
        go.Scatter(
            x=master["Report Date"],
            y=master["Total %"] * 100,
            mode="lines+markers",
            name="Total %",
            line=dict(color="#1a3a5c", width=3),
            marker=dict(size=8),
        ),
        row=1,
        col=1,
    )
    for bucket, color in zip(
        ["Serialized Profile", "Storage Location", "Valuation Type"],
        ["#e67e22", "#3498db", "#9b59b6"],
    ):
        fig.add_trace(
            go.Scatter(
                x=master["Report Date"],
                y=master[bucket] * 100,
                mode="lines",
                name=bucket,
                line=dict(width=2, dash="dot", color=color),
            ),
            row=1,
            col=1,
        )
    fig.update_yaxes(title_text="Error rate (%)", row=1, col=1)
    fig.update_xaxes(title_text="Report date (run date)", row=1, col=1)

    # --- Row 1 left: category breakdown (latest month) ---
    rates = [latest[b] * 100 for b in METRIC_BUCKETS]
    fig.add_trace(
        go.Bar(
            y=METRIC_BUCKETS,
            x=rates,
            orientation="h",
            marker_color=[
                "#3498db" if r < 1.5 else "#e67e22" if r < 2.5 else "#c0392b"
                for r in rates
            ],
            text=[f"{r:.1f}%" for r in rates],
            textposition="outside",
            name="Category %",
        ),
        row=2,
        col=1,
    )
    fig.update_xaxes(title_text="% of parts flagged", row=2, col=1)
    fig.update_layout(bargap=0.25)

    # --- Row 1 right: stacked trend by category ---
    for bucket in METRIC_BUCKETS[:6]:
        fig.add_trace(
            go.Bar(
                x=master["Report Date"].dt.strftime("%b %Y"),
                y=master[bucket] * 100,
                name=bucket,
            ),
            row=2,
            col=2,
        )
    fig.update_layout(barmode="stack")
    fig.update_yaxes(title_text="Stacked % (top categories)", row=2, col=2)

    # --- Row 2 left: table top parts ---
    top = detail.head(12)
    fig.add_trace(
        go.Table(
            header=dict(
                values=["Material", "Plant", "Errors", "SNP", "SLoc", "QMAT", "VType"],
                fill_color="#1a3a5c",
                font=dict(color="white", size=11),
                align="left",
            ),
            cells=dict(
                values=[
                    top["Material Number"],
                    top["Plnt"],
                    top["Errors"],
                    top["Check_SNP"],
                    top["Check_SLoc Missing"],
                    top["Check_QMAT Missing"],
                    top["Check_VType Missing"],
                ],
                fill_color=[["#ffffff", "#f8fafc"] * 6],
                align="left",
                font=dict(size=10),
            ),
        ),
        row=3,
        col=1,
    )

    # --- Row 2 right: check-level counts (latest detail sample) ---
    check_cols = [c for c in CHECK_LABELS if c in detail.columns]
    counts = detail[check_cols].sum().sort_values(ascending=True)
    labels = [CHECK_LABELS[c] for c in counts.index]
    fig.add_trace(
        go.Bar(
            y=labels,
            x=counts.values,
            orientation="h",
            marker_color="#5dade2",
            text=counts.values,
            textposition="outside",
            name="Flag count",
        ),
        row=3,
        col=2,
    )
    fig.update_xaxes(title_text="Number of parts flagged", row=3, col=2)

    # KPI cards (paper coordinates, top band)
    add_kpi_card(fig, 1, 0, "Parts Created", f"{int(latest['Parts Created']):,}", period)
    add_kpi_card(fig, 1, 1, "Total Error Rate", pct_text(latest["Total %"]), "All categories combined")
    add_kpi_card(fig, 1, 2, "Serialized Profile", pct_text(latest["Serialized Profile"]), "After SNP exceptions")
    add_kpi_card(fig, 1, 3, "Parts with Issues", f"{int((detail['Errors'] > 0).sum())}", f"of {len(detail)} in detail sample")

    fig.update_layout(
        title=dict(
            text="<b>CASRA Material Master KPI</b> — layout mockup (sample data)",
            x=0.02,
            font=dict(size=20),
        ),
        template="plotly_white",
        height=1100,
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
        margin=dict(t=120, b=40, l=60, r=40),
        paper_bgcolor="#ffffff",
    )

    fig.add_annotation(
        text=(
            "<i>Suggested Power BI pages: (1) Executive summary — this view "
            "(2) Trend & targets (3) Part detail from SNP_Final. "
            "Slicers: period, plant, material group.</i>"
        ),
        xref="paper",
        yref="paper",
        x=0,
        y=1.08,
        showarrow=False,
        font=dict(size=11, color="#7a8a9a"),
        align="left",
    )

    return fig


def main() -> None:
    static_html = Path(__file__).resolve().parent / "casra_dashboard_mockup.html"
    if static_html.exists():
        print(f"Static mockup (no dependencies): {static_html}")

    try:
        import plotly.graph_objects as go  # noqa: F401
    except ImportError:
        print("Plotly not installed — open the static HTML mockup above.")
        print("Optional: pip install plotly pandas  then re-run for an interactive chart version.")
        return

    master = build_sample_master()
    detail = build_sample_detail()
    fig = build_dashboard(master, detail)
    interactive_html = Path(__file__).resolve().parent / "casra_dashboard_mockup_interactive.html"
    fig.write_html(interactive_html, include_plotlyjs="cdn", full_html=True)
    print(f"Interactive mockup saved: {interactive_html}")


if __name__ == "__main__":
    main()
