# /// script
# requires-python = ">=3.10"
# dependencies = [
#   "pandas",
#   "openpyxl",
#   "plotly"
# ]
# ///

import json
import os
from dataclasses import dataclass, asdict
from datetime import datetime, timezone
from typing import Dict, List, Optional, Sequence, Tuple

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots


class ABSWorkbookError(Exception):
    """Raised when an ABS workbook does not match the expected schema."""


class SeriesNotFoundError(Exception):
    """Raised when a requested ABS series cannot be found."""


@dataclass(frozen=True)
class SeriesSpec:
    workbook_path: str
    description: str
    series_type: str = "Seasonally Adjusted"


@dataclass
class ExtractionAudit:
    workbook_path: str
    description: str
    series_type: str
    found: bool
    matched_column: Optional[int]
    start_date: Optional[str]
    end_date: Optional[str]
    observations: int
    error: Optional[str] = None


@dataclass
class SignalDiagnostics:
    clipped_std_observations: int
    total_observations: int


@dataclass
class CompositeResult:
    name: str
    composite: pd.DataFrame
    components: pd.DataFrame
    audits: List[ExtractionAudit]
    diagnostics: Dict[str, SignalDiagnostics]
    common_start_date: Optional[str]
    common_end_date: Optional[str]
    series_count: int


EXPECTED_DATA_SHEET = "Data1"
DESCRIPTION_ROW = 0
SERIES_TYPE_ROW = 2
DATA_START_ROW = 10
DATE_COLUMN = 0
EPSILON = 1e-6


def normalize_text(value: object) -> str:
    return " ".join(str(value).strip().split())


def validate_workbook_schema(workbook_path: str, sheet_name: str = EXPECTED_DATA_SHEET) -> None:
    if not os.path.exists(workbook_path):
        raise ABSWorkbookError(f"Workbook does not exist: {workbook_path}")

    try:
        xls = pd.ExcelFile(workbook_path)
    except Exception as exc:
        raise ABSWorkbookError(f"Failed to open workbook {workbook_path}: {exc}") from exc

    if sheet_name not in xls.sheet_names:
        raise ABSWorkbookError(
            f"Workbook {workbook_path} does not contain required sheet '{sheet_name}'. "
            f"Available sheets: {xls.sheet_names}"
        )

    df = pd.read_excel(workbook_path, sheet_name=sheet_name, header=None)
    if df.shape[0] <= DATA_START_ROW:
        raise ABSWorkbookError(
            f"Workbook {workbook_path} sheet '{sheet_name}' has only {df.shape[0]} rows; "
            f"expected at least {DATA_START_ROW + 1}."
        )
    if df.shape[1] <= DATE_COLUMN:
        raise ABSWorkbookError(
            f"Workbook {workbook_path} sheet '{sheet_name}' has no date column at position {DATE_COLUMN}."
        )

    parsed_dates = pd.to_datetime(df.iloc[DATA_START_ROW:, DATE_COLUMN], errors="coerce")
    if parsed_dates.notna().sum() == 0:
        raise ABSWorkbookError(
            f"Workbook {workbook_path} sheet '{sheet_name}' has no parseable dates starting at row {DATA_START_ROW}."
        )


def extract_series(
    spec: SeriesSpec,
    *,
    strict: bool = True,
    sheet_name: str = EXPECTED_DATA_SHEET,
) -> Tuple[pd.Series, ExtractionAudit]:
    validate_workbook_schema(spec.workbook_path, sheet_name=sheet_name)
    df = pd.read_excel(spec.workbook_path, sheet_name=sheet_name, header=None)

    row_desc = df.iloc[DESCRIPTION_ROW]
    row_type = df.iloc[SERIES_TYPE_ROW]
    target_desc = normalize_text(spec.description)

    col_idx: Optional[int] = None
    for i, (desc, stype) in enumerate(zip(row_desc, row_type)):
        if pd.notna(desc) and normalize_text(desc) == target_desc and str(stype) == spec.series_type:
            col_idx = i
            break

    if col_idx is None:
        error = (
            f"Series not found in workbook '{spec.workbook_path}': "
            f"description='{spec.description}', series_type='{spec.series_type}'"
        )
        audit = ExtractionAudit(
            workbook_path=spec.workbook_path,
            description=spec.description,
            series_type=spec.series_type,
            found=False,
            matched_column=None,
            start_date=None,
            end_date=None,
            observations=0,
            error=error,
        )
        if strict:
            raise SeriesNotFoundError(error)
        return pd.Series(dtype=float), audit

    data = df.iloc[DATA_START_ROW:, [DATE_COLUMN, col_idx]].copy()
    data.columns = ["date", "value"]
    data["date"] = pd.to_datetime(data["date"], errors="coerce")
    data["value"] = pd.to_numeric(data["value"], errors="coerce")
    data = data.dropna(subset=["date", "value"]).sort_values("date")

    series = pd.Series(data["value"].values, index=pd.DatetimeIndex(data["date"]), name=spec.description)
    series = series[~series.index.duplicated(keep="last")]

    audit = ExtractionAudit(
        workbook_path=spec.workbook_path,
        description=spec.description,
        series_type=spec.series_type,
        found=True,
        matched_column=col_idx,
        start_date=series.index.min().date().isoformat() if not series.empty else None,
        end_date=series.index.max().date().isoformat() if not series.empty else None,
        observations=int(series.shape[0]),
        error=None,
    )
    return series, audit


def compute_standardized_growth_gap(
    series: pd.Series,
    *,
    window: int,
    min_periods: int = 4,
    epsilon: float = EPSILON,
) -> Tuple[pd.Series, SignalDiagnostics]:
    yoy = series.pct_change(4)
    rolling_mean = yoy.rolling(window=window, min_periods=min_periods).mean()
    rolling_std = yoy.rolling(window=window, min_periods=min_periods).std()

    clipped_mask = rolling_std.notna() & (rolling_std.abs() < epsilon)
    rolling_std = rolling_std.clip(lower=epsilon)

    signal = (yoy - rolling_mean) / rolling_std
    signal.name = series.name

    diagnostics = SignalDiagnostics(
        clipped_std_observations=int(clipped_mask.sum()),
        total_observations=int(signal.notna().sum()),
    )
    return signal, diagnostics


def build_balanced_panel(series_map: Dict[str, pd.Series]) -> Tuple[pd.DataFrame, Optional[pd.Timestamp], Optional[pd.Timestamp]]:
    if not series_map:
        return pd.DataFrame(), None, None

    start_dates = []
    end_dates = []
    for series in series_map.values():
        valid = series.dropna()
        if valid.empty:
            raise ValueError(f"Series '{series.name}' has no valid observations after transformation.")
        start_dates.append(valid.index.min())
        end_dates.append(valid.index.max())

    common_start = max(start_dates)
    common_end = min(end_dates)
    if common_start > common_end:
        raise ValueError("No overlapping balanced sample exists across the requested series.")

    aligned = pd.concat(series_map.values(), axis=1)
    balanced = aligned.loc[(aligned.index >= common_start) & (aligned.index <= common_end)].copy()
    balanced = balanced.dropna(how="any")

    if balanced.empty:
        raise ValueError("Balanced panel is empty after intersecting valid date ranges and dropping missing values.")

    return balanced, common_start, common_end


def compute_recession_flag(gdp_series: pd.Series) -> pd.Series:
    qoq_growth = gdp_series.pct_change(1)
    recession_flag = ((qoq_growth < 0) & (qoq_growth.shift(1) < 0)).astype(int)
    recession_flag.name = "recession_flag"
    return recession_flag


def build_recession_intervals(recession_flag: pd.Series) -> List[Tuple[pd.Timestamp, pd.Timestamp]]:
    intervals: List[Tuple[pd.Timestamp, pd.Timestamp]] = []
    in_recession = False
    start: Optional[pd.Timestamp] = None
    sorted_flag = recession_flag.fillna(0).astype(int).sort_index()

    for date, flag in sorted_flag.items():
        if flag == 1 and not in_recession:
            start = date
            in_recession = True
        elif flag == 0 and in_recession:
            intervals.append((start, date))
            in_recession = False
            start = None

    if in_recession and start is not None:
        intervals.append((start, sorted_flag.index.max()))
    return intervals


def compute_composite(
    name: str,
    specs: Sequence[SeriesSpec],
    *,
    window: int,
    smooth_window: int,
    allow_missing: bool,
    epsilon: float = EPSILON,
) -> CompositeResult:
    audits: List[ExtractionAudit] = []
    transformed: Dict[str, pd.Series] = {}
    diagnostics: Dict[str, SignalDiagnostics] = {}

    for spec in specs:
        series, audit = extract_series(spec, strict=not allow_missing)
        audits.append(audit)
        if series.empty:
            continue
        signal, diag = compute_standardized_growth_gap(series, window=window, epsilon=epsilon)
        transformed[spec.description] = signal.rename(spec.description)
        diagnostics[spec.description] = diag

    if not transformed:
        raise ValueError(f"Composite '{name}' has no usable series.")

    panel, common_start, common_end = build_balanced_panel(transformed)
    series_count = panel.shape[1]

    diffusion = (panel < 0).sum(axis=1) / series_count
    negative_only = panel.where(panel < 0)
    depth = negative_only.mean(axis=1).fillna(0.0)
    combined_raw = -(diffusion * depth)
    combined_smoothed = combined_raw.rolling(window=smooth_window, min_periods=1).mean()

    composite_df = pd.DataFrame(
        {
            f"{name}": combined_smoothed,
            f"{name}_diffusion": diffusion,
            f"{name}_depth": depth,
            f"{name}_n_active": series_count,
        }
    )

    components_df = panel.copy()
    components_df.columns = [f"{name}__{col}" for col in components_df.columns]

    return CompositeResult(
        name=name,
        composite=composite_df,
        components=components_df,
        audits=audits,
        diagnostics=diagnostics,
        common_start_date=common_start.date().isoformat() if common_start is not None else None,
        common_end_date=common_end.date().isoformat() if common_end is not None else None,
        series_count=series_count,
    )


def plot_indices_html(
    output_html: str,
    combined: pd.DataFrame,
    recession_intervals: Sequence[Tuple[pd.Timestamp, pd.Timestamp]],
    *,
    start_date: str,
) -> None:
    plot_data = combined.loc[combined.index >= pd.Timestamp(start_date)].copy()
    plot_data = plot_data.sort_index()

    fig = make_subplots(
        rows=5,
        cols=1,
        shared_xaxes=True,
        vertical_spacing=0.03,
        subplot_titles=[
            "Leading index",
            "Coincident index",
            "GDP (YoY)",
            "GDP per capita (YoY)",
            "Lagging index",
        ],
    )

    series_to_plot = [
        ("leading", "Std. diff-depth"),
        ("coincident", "Std. diff-depth"),
        ("gdp", "YoY"),
        ("gdp_per_capita", "YoY"),
        ("lagging", "Std. diff-depth"),
    ]

    for row_num, (col, yaxis_title) in enumerate(series_to_plot, start=1):
        fig.add_trace(
            go.Scatter(
                x=plot_data.index,
                y=plot_data[col],
                mode="lines",
                name=col,
                hovertemplate="Date: %{x|%Y-%m-%d}<br>Value: %{y:,.4f}<extra></extra>",
            ),
            row=row_num,
            col=1,
        )

        for rec_start, rec_end in recession_intervals:
            if rec_end < plot_data.index.min() or rec_start > plot_data.index.max():
                continue
            fig.add_vrect(
                x0=rec_start,
                x1=rec_end,
                fillcolor="grey",
                opacity=0.2,
                line_width=0,
                row=row_num,
                col=1,
            )

        if col in {"leading", "coincident", "lagging"}:
            fig.add_hline(y=0, line_dash="dash", line_width=1, line_color="black", row=row_num, col=1)

        fig.update_yaxes(title_text=yaxis_title, row=row_num, col=1)

    fig.update_xaxes(title_text="Date", row=5, col=1)
    fig.update_layout(
        height=1200,
        hovermode="x unified",
        showlegend=False,
        margin=dict(l=60, r=30, t=80, b=50),
    )

    fig.write_html(output_html, include_plotlyjs=True, full_html=True)


def write_metadata(
    output_path: str,
    *,
    window: int,
    smooth_window: int,
    allow_missing: bool,
    epsilon: float,
    primary_output: str,
    components_output: str,
    plot_output: str,
    composites: Sequence[CompositeResult],
    missing_series: Sequence[ExtractionAudit],
) -> None:
    payload = {
        "run_timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "parameters": {
            "window": window,
            "smooth_window": smooth_window,
            "allow_missing": allow_missing,
            "epsilon": epsilon,
        },
        "outputs": {
            "primary_csv": primary_output,
            "components_csv": components_output,
            "plot_html": plot_output,
        },
        "composites": {},
        "missing_series": [asdict(audit) for audit in missing_series],
    }

    for result in composites:
        payload["composites"][result.name] = {
            "series_count": result.series_count,
            "common_start_date": result.common_start_date,
            "common_end_date": result.common_end_date,
            "audits": [asdict(audit) for audit in result.audits],
            "diagnostics": {
                key: asdict(value) for key, value in result.diagnostics.items()
            },
        }

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)


def main() -> None:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(script_dir, "abs_workbooks")

    output_csv = os.path.join(script_dir, "composite_indices.csv")
    output_components_csv = os.path.join(script_dir, "composite_components.csv")
    output_metadata_json = os.path.join(script_dir, "composite_metadata.json")
    output_html = os.path.join(script_dir, "composite_indices.html")

    window = 8
    smooth_window = 4
    allow_missing = False
    epsilon = EPSILON
    start_date = "1970-01-01"

    exp_volume_file = os.path.join(data_dir, "5206002_Expenditure_Volume_Measures.xlsx")
    key_aggregates_file = os.path.join(data_dir, "5206001_Key_Aggregates.xlsx")
    hh_cons_file = os.path.join(data_dir, "5206008_Household_Final_Consumption_Expenditure.xlsx")
    industry_gva_file = os.path.join(data_dir, "5206006_Industry_GVA.xlsx")

    leading_specs = [
        SeriesSpec(exp_volume_file, "Private ;  Gross fixed capital formation - Dwellings - New and Used ;"),
        SeriesSpec(exp_volume_file, "Private ;  Gross fixed capital formation - Dwellings - Alterations and additions ;"),
        SeriesSpec(exp_volume_file, "Private ;  Gross fixed capital formation - Machinery and equipment - New ;"),
        SeriesSpec(exp_volume_file, "Private ;  Gross fixed capital formation - Machinery and equipment - Net purchase of second hand assets ;"),
        SeriesSpec(hh_cons_file, "Purchase of vehicles: Chain volume measures ;"),
        SeriesSpec(hh_cons_file, "Furnishings and household equipment: Chain volume measures ;"),
    ]

    coincident_specs = [
        SeriesSpec(exp_volume_file, "Households ;  Final consumption expenditure ;"),
        SeriesSpec(key_aggregates_file, "Gross domestic product: Chain volume measures ;"),
        SeriesSpec(key_aggregates_file, "Hours worked: Index ;"),
        SeriesSpec(key_aggregates_file, "GDP per capita: Chain volume measures ;"),
    ]

    lagging_specs = [
        SeriesSpec(exp_volume_file, "Private ;  Gross fixed capital formation - Non-dwelling construction - Total ;"),
        SeriesSpec(industry_gva_file, "Education and training (P) ;"),
        SeriesSpec(industry_gva_file, "Health care and social assistance (Q) ;"),
        SeriesSpec(industry_gva_file, "Public administration and safety (O) ;"),
    ]

    leading_result = compute_composite(
        "leading",
        leading_specs,
        window=window,
        smooth_window=smooth_window,
        allow_missing=allow_missing,
        epsilon=epsilon,
    )
    coincident_result = compute_composite(
        "coincident",
        coincident_specs,
        window=window,
        smooth_window=smooth_window,
        allow_missing=allow_missing,
        epsilon=epsilon,
    )
    lagging_result = compute_composite(
        "lagging",
        lagging_specs,
        window=window,
        smooth_window=smooth_window,
        allow_missing=allow_missing,
        epsilon=epsilon,
    )

    gdp_spec = SeriesSpec(key_aggregates_file, "Gross domestic product: Chain volume measures ;")
    gdp_per_capita_spec = SeriesSpec(key_aggregates_file, "GDP per capita: Chain volume measures ;")

    gdp_series, gdp_audit = extract_series(gdp_spec, strict=True)
    gdp_per_capita_series, gdp_per_capita_audit = extract_series(gdp_per_capita_spec, strict=True)

    gdp_yoy = gdp_series.pct_change(4).rename("gdp")
    gdp_per_capita_yoy = gdp_per_capita_series.pct_change(4).rename("gdp_per_capita")

    gdp_qoq = gdp_series.pct_change(1).rename("gdp_qoq")
    recession_flag = compute_recession_flag(gdp_series)
    recession_intervals = build_recession_intervals(recession_flag)

    combined = pd.concat(
        [
            leading_result.composite,
            coincident_result.composite,
            lagging_result.composite,
            gdp_yoy,
            gdp_per_capita_yoy,
            gdp_qoq,
            recession_flag,
        ],
        axis=1,
        sort=True,
    ).sort_index()

    primary_output = combined[
        [
            "leading",
            "leading_diffusion",
            "leading_depth",
            "leading_n_active",
            "coincident",
            "coincident_diffusion",
            "coincident_depth",
            "coincident_n_active",
            "gdp",
            "gdp_per_capita",
            "lagging",
            "lagging_diffusion",
            "lagging_depth",
            "lagging_n_active",
            "gdp_qoq",
            "recession_flag",
        ]
    ]
    primary_output.to_csv(output_csv, index=True)

    components_output = pd.concat(
        [
            leading_result.components,
            coincident_result.components,
            lagging_result.components,
        ],
        axis=1,
    ).sort_index()
    components_output.to_csv(output_components_csv, index=True)

    missing_series = [
        audit
        for result in [leading_result, coincident_result, lagging_result]
        for audit in result.audits
        if not audit.found
    ]
    if not gdp_audit.found:
        missing_series.append(gdp_audit)
    if not gdp_per_capita_audit.found:
        missing_series.append(gdp_per_capita_audit)

    write_metadata(
        output_metadata_json,
        window=window,
        smooth_window=smooth_window,
        allow_missing=allow_missing,
        epsilon=epsilon,
        primary_output=output_csv,
        components_output=output_components_csv,
        plot_output=output_html,
        composites=[leading_result, coincident_result, lagging_result],
        missing_series=missing_series,
    )

    plot_indices_html(output_html, primary_output, recession_intervals, start_date=start_date)

    print(f"Saved primary output to {output_csv}")
    print(f"Saved components output to {output_components_csv}")
    print(f"Saved metadata output to {output_metadata_json}")
    print(f"Saved plot to {output_html}")


if __name__ == "__main__":
    main()
