# /// script
# requires-python = ">=3.10"
# dependencies = [
#   "pandas",
#   "numpy",
#   "openpyxl",
#   "plotly"
# ]
# ///

from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

import build_business_cycle_indices as business_cycle


DEFAULT_START_DATE = pd.Timestamp("1980-01-01")


@dataclass(frozen=True)
class ComponentSpec:
    key: str
    name: str
    group: str
    contribution_workbook: str
    contribution_label: str
    cycle_workbook: str
    cycle_label: str
    color: str


GDP_SPEC = {
    "key": "gdp_current",
    "name": "GDP current prices",
    "workbook": "5206001_Key_Aggregates.xlsx",
    "label": "Gross domestic product: Current prices ;",
    "series_type": "Seasonally Adjusted",
}

GDP_PER_CAPITA_SPEC = {
    "key": "gdp_per_capita",
    "name": "GDP per capita",
    "workbook": "5206001_Key_Aggregates.xlsx",
    "label": "GDP per capita: Chain volume measures ;",
    "series_type": "Seasonally Adjusted",
}

COMPONENTS = [
    ComponentSpec(
        "mining",
        "Mining GVA",
        "Resource Cycle",
        "5206045_Industry_GVA_Current_Price.xlsx",
        "Mining (B) ; Gross value added at basic prices ;",
        "5206006_Industry_GVA.xlsx",
        "Mining (B) ;",
        "#00796b",
    ),
    ComponentSpec(
        "machinery_equipment",
        "Machinery and equipment",
        "Domestic Capex Cycle",
        "5206003_Expenditure_Current_Price.xlsx",
        "Private ; Gross fixed capital formation - Machinery and equipment - New ;",
        "5206002_Expenditure_Volume_Measures.xlsx",
        "Private ; Gross fixed capital formation - Machinery and equipment - New ;",
        "#1b9e77",
    ),
    ComponentSpec(
        "non_dwelling_engineering",
        "Non-dwelling engineering",
        "Domestic Capex Cycle",
        "5206003_Expenditure_Current_Price.xlsx",
        "Private ; Gross fixed capital formation - Non-dwelling construction - New engineering construction ;",
        "5206002_Expenditure_Volume_Measures.xlsx",
        "Private ; Gross fixed capital formation - Non-dwelling construction - New engineering construction ;",
        "#d95f02",
    ),
    ComponentSpec(
        "non_dwelling_building",
        "Non-dwelling building",
        "Domestic Capex Cycle",
        "5206003_Expenditure_Current_Price.xlsx",
        "Private ; Gross fixed capital formation - Non-dwelling construction - New building ;",
        "5206002_Expenditure_Volume_Measures.xlsx",
        "Private ; Gross fixed capital formation - Non-dwelling construction - New building ;",
        "#7570b3",
    ),
    ComponentSpec(
        "dwellings_new_used",
        "Dwellings - new and used",
        "Housing Cycle",
        "5206003_Expenditure_Current_Price.xlsx",
        "Private ; Gross fixed capital formation - Dwellings - New and Used ;",
        "5206002_Expenditure_Volume_Measures.xlsx",
        "Private ; Gross fixed capital formation - Dwellings - New and Used ;",
        "#1976d2",
    ),
    ComponentSpec(
        "dwellings_alterations",
        "Dwellings - alterations",
        "Housing Cycle",
        "5206003_Expenditure_Current_Price.xlsx",
        "Private ; Gross fixed capital formation - Dwellings - Alterations and additions ;",
        "5206002_Expenditure_Volume_Measures.xlsx",
        "Private ; Gross fixed capital formation - Dwellings - Alterations and additions ;",
        "#f9a825",
    ),
    ComponentSpec(
        "ownership_transfer_costs",
        "Ownership transfer costs",
        "Housing Cycle",
        "5206003_Expenditure_Current_Price.xlsx",
        "Private ; Gross fixed capital formation - Ownership transfer costs ;",
        "5206002_Expenditure_Volume_Measures.xlsx",
        "Private ; Gross fixed capital formation - Ownership transfer costs ;",
        "#c2185b",
    ),
    ComponentSpec(
        "hh_vehicles",
        "Purchase of vehicles",
        "Household Discretionary Cycle",
        "5206008_Household_Final_Consumption_Expenditure.xlsx",
        "Purchase of vehicles: Current prices ;",
        "5206008_Household_Final_Consumption_Expenditure.xlsx",
        "Purchase of vehicles: Chain volume measures ;",
        "#e53935",
    ),
    ComponentSpec(
        "hh_furnishings",
        "Furnishings and household equipment",
        "Household Discretionary Cycle",
        "5206008_Household_Final_Consumption_Expenditure.xlsx",
        "Furnishings and household equipment: Current prices ;",
        "5206008_Household_Final_Consumption_Expenditure.xlsx",
        "Furnishings and household equipment: Chain volume measures ;",
        "#43a047",
    ),
    ComponentSpec(
        "hh_hotels",
        "Hotels, cafes and restaurants",
        "Household Discretionary Cycle",
        "5206008_Household_Final_Consumption_Expenditure.xlsx",
        "Hotels, cafes and restaurants: Current prices ;",
        "5206008_Household_Final_Consumption_Expenditure.xlsx",
        "Hotels, cafes and restaurants: Chain volume measures ;",
        "#fb8c00",
    ),
    ComponentSpec(
        "hh_clothing",
        "Clothing and footwear",
        "Household Discretionary Cycle",
        "5206008_Household_Final_Consumption_Expenditure.xlsx",
        "Clothing and footwear: Current prices ;",
        "5206008_Household_Final_Consumption_Expenditure.xlsx",
        "Clothing and footwear: Chain volume measures ;",
        "#8e24aa",
    ),
]

GROUP_ORDER = [
    "Resource Cycle",
    "Domestic Capex Cycle",
    "Housing Cycle",
    "Household Discretionary Cycle",
]

GROUP_COLORS = {
    "Resource Cycle": "#005f56",
    "Domestic Capex Cycle": "#424242",
    "Housing Cycle": "#0d47a1",
    "Household Discretionary Cycle": "#bf360c",
}

GROUP_DESCRIPTIONS = {
    "Resource Cycle": "Mining real GVA",
    "Domestic Capex Cycle": "Machinery and non-dwelling construction",
    "Housing Cycle": "Dwellings, alterations, and transfer costs",
    "Household Discretionary Cycle": "Vehicles, furnishings, hospitality, and clothing",
}


def component_to_abs_specs(component: ComponentSpec) -> list[dict[str, str]]:
    return [
        {
            "key": f"{component.key}_contribution",
            "name": f"{component.name} current prices",
            "workbook": component.contribution_workbook,
            "label": component.contribution_label,
            "series_type": "Seasonally Adjusted",
        },
        {
            "key": f"{component.key}_cycle",
            "name": f"{component.name} real activity",
            "workbook": component.cycle_workbook,
            "label": component.cycle_label,
            "series_type": "Seasonally Adjusted",
        },
    ]


def load_series(national_accounts_dir: Path) -> dict[str, pd.Series]:
    component_specs = [
        spec
        for component in COMPONENTS
        for spec in component_to_abs_specs(component)
    ]
    specs = [GDP_SPEC, GDP_PER_CAPITA_SPEC, *component_specs]
    specs_by_workbook: dict[str, list[dict[str, str]]] = {}
    for spec in specs:
        specs_by_workbook.setdefault(spec["workbook"], []).append(spec)

    series_by_key: dict[str, pd.Series] = {}
    for workbook_name, workbook_specs in specs_by_workbook.items():
        series_by_key.update(
            business_cycle.load_series_batch(national_accounts_dir / workbook_name, workbook_specs)
        )
    return series_by_key


def build_dashboard_data(
    series_by_key: dict[str, pd.Series],
    start_date: pd.Timestamp,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    gdp = series_by_key[GDP_SPEC["key"]].rename("gdp_current")
    gdp_lag = gdp.shift(4).rename("gdp_current_lag_4q")

    component_frames: list[pd.DataFrame] = []
    for component in COMPONENTS:
        level = series_by_key[f"{component.key}_contribution"].rename("level")
        cycle_score = business_cycle.compute_standardized_growth_gap(
            series_by_key[f"{component.key}_cycle"]
        ).rename("cycle_score")
        aligned = pd.concat([level, gdp, gdp_lag], axis=1, sort=True).dropna(how="any")
        aligned = aligned.loc[aligned.index >= start_date]
        if aligned.empty:
            raise ValueError(f"No aligned observations for {component.name}.")

        contribution = (aligned["level"] - aligned["level"].shift(4)) / aligned["gdp_current_lag_4q"] * 100.0
        component_frame = pd.concat(
            [
                contribution.rename("contribution_to_gdp_growth"),
                cycle_score,
            ],
            axis=1,
            sort=True,
        )
        component_frame = component_frame.loc[component_frame.index >= start_date]
        component_frame = component_frame.dropna(subset=["cycle_score"])
        if component_frame.empty:
            raise ValueError(f"No cycle observations for {component.name}.")

        component_frames.append(
            pd.DataFrame(
                {
                    "date": component_frame.index,
                    "key": component.key,
                    "name": component.name,
                    "group": component.group,
                    "cycle_score": component_frame["cycle_score"],
                    "contribution_to_gdp_growth": component_frame["contribution_to_gdp_growth"],
                }
            ).dropna(subset=["cycle_score"])
        )

    component_data = pd.concat(component_frames, ignore_index=True)
    if component_data.empty:
        raise ValueError("Dashboard component data is empty.")

    group_frames: list[pd.DataFrame] = []
    for group in GROUP_ORDER:
        subset = component_data.loc[component_data["group"] == group]
        cycle_panel = (
            subset.pivot(index="date", columns="key", values="cycle_score")
            .sort_index()
            .dropna(how="any")
        )
        contribution_panel = (
            subset.pivot(index="date", columns="key", values="contribution_to_gdp_growth")
            .sort_index()
            .reindex(cycle_panel.index)
        )
        if cycle_panel.empty:
            raise ValueError(f"No balanced cycle panel for {group}.")

        group_frames.append(
            pd.DataFrame(
                {
                    "date": cycle_panel.index,
                    "group": group,
                    "cycle_score": cycle_panel.mean(axis=1),
                    "contribution_to_gdp_growth": contribution_panel.sum(axis=1, min_count=1),
                    "components": GROUP_DESCRIPTIONS[group],
                }
            )
        )

    group_data = pd.concat(group_frames, ignore_index=True)

    return component_data, group_data


def add_recession_shading(
    figure: go.Figure,
    recession_flag: pd.Series,
    row_date_ranges: dict[int, tuple[pd.Timestamp, pd.Timestamp]],
) -> None:
    for start, end in business_cycle.build_recession_intervals(recession_flag):
        for row_number, (row_start, row_end) in row_date_ranges.items():
            clipped_start = max(start, row_start)
            clipped_end = min(end, row_end)
            if clipped_start > clipped_end:
                continue
            figure.add_vrect(
                x0=clipped_start,
                x1=clipped_end,
                fillcolor="rgba(97, 97, 97, 0.10)",
                line_width=0,
                layer="below",
                row=row_number,
                col=1,
            )


def create_figure(
    component_data: pd.DataFrame,
    group_data: pd.DataFrame,
    recession_flag: pd.Series,
) -> go.Figure:
    latest_date = group_data["date"].max()
    latest = group_data.loc[group_data["date"] == latest_date].set_index("group")

    figure = make_subplots(
        rows=4,
        cols=1,
        shared_xaxes=True,
        vertical_spacing=0.045,
        subplot_titles=[
            f"{group}: latest cycle {latest.loc[group, 'cycle_score']:+.2f}"
            for group in GROUP_ORDER
        ],
    )

    row_date_ranges: dict[int, tuple[pd.Timestamp, pd.Timestamp]] = {}
    for row_number, group in enumerate(GROUP_ORDER, start=1):
        group_subset = group_data.loc[group_data["group"] == group].sort_values("date")
        row_date_ranges[row_number] = (group_subset["date"].min(), group_subset["date"].max())

        figure.add_trace(
            go.Scatter(
                x=group_subset["date"],
                y=group_subset["cycle_score"],
                mode="lines",
                name="real activity cycle score",
                line={"color": GROUP_COLORS[group], "width": 3.2},
                showlegend=False,
                customdata=pd.concat(
                    [
                        pd.DataFrame(business_cycle.quarter_labels(group_subset["date"]), columns=["quarter"]),
                        group_subset[["components"]].reset_index(drop=True),
                    ],
                    axis=1,
                ),
                hovertemplate=(
                    "%{customdata[0]}<br>"
                    "cycle score %{y:+.2f}<br>"
                    "%{customdata[1]}"
                    "<extra>%{fullData.name}</extra>"
                ),
            ),
            row=row_number,
            col=1,
        )

        figure.update_yaxes(
            title_text="cycle score",
            zeroline=True,
            zerolinecolor="#616161",
            zerolinewidth=1,
            row=row_number,
            col=1,
        )

    add_recession_shading(figure, recession_flag, row_date_ranges)

    figure.update_layout(
        title={
            "text": (
                "Australia Cyclical GDP Dashboard<br>"
                "<sup>Real activity standardized YoY growth-gap cycle; shading = two consecutive negative GDP per-capita quarters.</sup>"
            ),
            "x": 0.5,
        },
        template="plotly_white",
        hovermode="x unified",
        height=1050,
        showlegend=False,
        margin={"l": 75, "r": 75, "t": 120, "b": 60},
    )
    figure.update_xaxes(title_text="Quarter", row=4, col=1)
    return figure


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build a blog-style Australian cyclical GDP dashboard from current-price ABS components."
    )
    parser.add_argument(
        "--national-accounts-dir",
        type=Path,
        default=Path("abs_workbooks/australian_national_accounts"),
        help="Directory containing the ABS national accounts XLSX files.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("cyclical_dashboard.html"),
        help="Output HTML path for the Plotly dashboard.",
    )
    parser.add_argument(
        "--start-date",
        type=pd.Timestamp,
        default=DEFAULT_START_DATE,
        help="Start date for the dashboard sample.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_arguments()
    national_accounts_dir = args.national_accounts_dir.resolve()
    if not national_accounts_dir.exists():
        raise FileNotFoundError(f"National accounts directory not found: {national_accounts_dir}")

    series_by_key = load_series(national_accounts_dir)
    component_data, group_data = build_dashboard_data(series_by_key, args.start_date)
    _, recession_flag = business_cycle.compute_recession_flag(series_by_key[GDP_PER_CAPITA_SPEC["key"]])
    figure = create_figure(component_data, group_data, recession_flag)

    output_path = args.output.resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    figure.write_html(output_path, include_plotlyjs=True, full_html=True)

    latest_date = group_data["date"].max()
    latest = group_data.loc[group_data["date"] == latest_date].set_index("group").reindex(GROUP_ORDER)

    print(f"Wrote {output_path}")
    print("Latest cyclical dashboard values:")
    for group, row in latest.iterrows():
        print(f"- {group}: cycle={row['cycle_score']:+.2f}")


if __name__ == "__main__":
    try:
        main()
    except (FileNotFoundError, ValueError) as error:
        raise SystemExit(str(error)) from None
