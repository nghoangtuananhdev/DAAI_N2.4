"""Dashboard Measure, Metric và KPI cho Problem statement 3.

Google Colab:
1. Tải file kpi.py lên Colab và chạy:  %run kpi.py
2. Khi được hỏi, tải các workbook CSDL_v5 hoặc một file ZIP chứa chúng.
3. Dashboard tương tác và các file kết quả được tạo trong KPI_PS3_Output.

Chỉ bốn bảng là bắt buộc: Employee.xlsx, Order.xlsx, Payment.xlsx, Return.xlsx.
Order Item.xlsx được nhận diện để kiểm tra mở rộng nhưng không cộng vào doanh số,
vì doanh số được tính ở cấp hóa đơn từ Payment.
"""

from __future__ import annotations

from pathlib import Path
import argparse
import html
import importlib
import json
import subprocess
import sys
import zipfile


def ensure_packages() -> None:
    packages = {
        "pandas": "pandas",
        "numpy": "numpy",
        "openpyxl": "openpyxl",
        "xlsxwriter": "XlsxWriter",
        "plotly": "plotly",
        "ipywidgets": "ipywidgets",
    }
    for module, package in packages.items():
        try:
            importlib.import_module(module)
        except ImportError:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", package])


ensure_packages()

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
from IPython.display import HTML, display, clear_output
import ipywidgets as widgets


IN_COLAB = "google.colab" in sys.modules
REQUIRED_FILES = ["Employee.xlsx", "Order.xlsx", "Payment.xlsx", "Return.xlsx"]
OPTIONAL_FILES = ["Order Item.xlsx"]
VALID_STATUSES = {"delivered", "paid", "shipped", "returned"}

MEASURE_DICTIONARY = pd.DataFrame([
    ["Employee", "sales_employee_id", "Khóa định danh nhân viên"],
    ["Employee", "sales_employee_name", "Tên nhân viên để hiển thị"],
    ["Employee", "years_experience", "Số năm kinh nghiệm"],
    ["Order", "order_id", "Khóa hóa đơn"],
    ["Order", "sales_employee_id", "Nhân viên phụ trách đơn"],
    ["Order", "order_date", "Ngày phát sinh đơn"],
    ["Order", "order_status", "Trạng thái dùng để ghi nhận doanh số"],
    ["Payment", "payment_value", "Giá trị thanh toán của đơn"],
    ["Return", "refund_amount", "Giá trị hoàn tiền"],
    ["Return", "return_id", "Khóa lượt trả hàng"],
], columns=["table", "measure_column", "meaning"])

METRIC_DICTIONARY = pd.DataFrame([
    ["gross_sales", "SUM(payment_value) của delivered/paid/shipped/returned"],
    ["refund_amount", "SUM(refund_amount) của các đơn được ghi nhận"],
    ["net_sales", "gross_sales - refund_amount"],
    ["valid_orders", "DISTINCTCOUNT(order_id) của các đơn được ghi nhận"],
    ["average_net_sales_per_order", "net_sales / valid_orders"],
    ["rolling_12m_net_sales", "Tổng net_sales trong 12 tháng gần nhất"],
    ["cumulative_net_sales", "Doanh thu thuần lũy kế theo thời gian"],
], columns=["metric", "formula"])

KPI_DICTIONARY = pd.DataFrame([
    ["sales_contribution_pct", "Doanh thu nhân viên / tổng doanh thu cùng kỳ", "Cao hơn = đóng góp lớn hơn"],
    ["mom_growth_pct", "(Tháng hiện tại - tháng trước) / tháng trước", "Theo dõi xu hướng ngắn hạn"],
    ["yoy_growth_pct", "(Kỳ hiện tại - cùng kỳ năm trước) / cùng kỳ năm trước", "Theo dõi tăng trưởng đã giảm ảnh hưởng mùa vụ"],
    ["average_mom_growth_pct", "Trung bình các MoM hợp lệ trong kỳ chọn", "Tăng trưởng tháng trung bình"],
    ["average_yoy_growth_pct", "Trung bình các YoY hợp lệ trong kỳ chọn", "Tăng trưởng năm trung bình"],
    ["positive_growth_month_pct", "Số tháng MoM > 0 / số tháng so sánh được", "Tần suất tăng trưởng"],
    ["sales_volatility_pct", "Độ lệch chuẩn doanh thu tháng / trung bình tháng", "Thấp hơn = ổn định hơn"],
    ["sales_stability_pct", "1 / (1 + sales_volatility_pct)", "Cao hơn = ổn định hơn"],
    ["rolling_12m_growth_pct", "12 tháng gần nhất so với 12 tháng liền trước", "Xu hướng dài hạn"],
], columns=["kpi", "formula", "interpretation"])


def normalize_id(series: pd.Series) -> pd.Series:
    result = series.astype("string").fillna("").str.strip()
    return result.str.replace(r"\.0$", "", regex=True)


def safe_extract_zip(payload: bytes, target: Path) -> None:
    from io import BytesIO
    with zipfile.ZipFile(BytesIO(payload)) as archive:
        for member in archive.infolist():
            if member.is_dir():
                continue
            name = Path(member.filename).name
            if name in REQUIRED_FILES + OPTIONAL_FILES:
                (target / name).write_bytes(archive.read(member))


def upload_in_colab(target: Path) -> None:
    from google.colab import files
    target.mkdir(parents=True, exist_ok=True)
    print("Chọn 4 workbook bắt buộc hoặc một file ZIP chứa CSDL_v5:")
    print(", ".join(REQUIRED_FILES))
    uploaded = files.upload()
    for filename, payload in uploaded.items():
        if filename.lower().endswith(".zip"):
            safe_extract_zip(payload, target)
        elif Path(filename).name in REQUIRED_FILES + OPTIONAL_FILES:
            (target / Path(filename).name).write_bytes(payload)


def locate_files(base: Path) -> dict[str, Path]:
    found: dict[str, Path] = {}
    if base.exists():
        for filename in REQUIRED_FILES + OPTIONAL_FILES:
            candidates = sorted(base.rglob(filename), key=lambda p: (len(p.parts), str(p)))
            if candidates:
                found[filename] = candidates[0]
    missing = [f for f in REQUIRED_FILES if f not in found]
    if missing and IN_COLAB:
        upload_dir = Path("/content/kpi_ps3_upload")
        upload_in_colab(upload_dir)
        return locate_files(upload_dir)
    if missing:
        raise FileNotFoundError(
            "Thiếu các file: " + ", ".join(missing)
            + f". Hãy đặt chúng trong {base.resolve()} hoặc dùng --data-dir."
        )
    return found


def require_columns(frame: pd.DataFrame, filename: str, columns: list[str]) -> None:
    missing = [c for c in columns if c not in frame.columns]
    if missing:
        raise ValueError(f"{filename} thiếu cột: {missing}")


def numeric_column(frame: pd.DataFrame, column: str, filename: str) -> pd.Series:
    value = pd.to_numeric(frame[column], errors="coerce")
    bad = value.isna() & frame[column].notna()
    if bad.any():
        examples = frame.loc[bad, column].astype(str).head(5).tolist()
        raise ValueError(f"{filename}.{column} có {int(bad.sum())} giá trị không phải số: {examples}")
    return value.fillna(0.0)


def load_and_validate(paths: dict[str, Path]) -> tuple[dict[str, pd.DataFrame], pd.DataFrame]:
    employee = pd.read_excel(paths["Employee.xlsx"], engine="openpyxl", dtype=object)
    order = pd.read_excel(paths["Order.xlsx"], engine="openpyxl", dtype=object)
    payment = pd.read_excel(paths["Payment.xlsx"], engine="openpyxl", dtype=object)
    returns = pd.read_excel(paths["Return.xlsx"], engine="openpyxl", dtype=object)

    require_columns(employee, "Employee.xlsx", ["sales_employee_id", "sales_employee_name", "years_experience"])
    require_columns(order, "Order.xlsx", ["order_id", "sales_employee_id", "order_date", "order_status"])
    require_columns(payment, "Payment.xlsx", ["order_id", "payment_value"])
    require_columns(returns, "Return.xlsx", ["return_id", "order_id", "refund_amount"])

    employee["sales_employee_id"] = normalize_id(employee["sales_employee_id"])
    order["order_id"] = normalize_id(order["order_id"])
    order["sales_employee_id"] = normalize_id(order["sales_employee_id"])
    payment["order_id"] = normalize_id(payment["order_id"])
    returns["order_id"] = normalize_id(returns["order_id"])
    returns["return_id"] = normalize_id(returns["return_id"])
    order["order_status"] = order["order_status"].astype("string").str.strip().str.lower()
    order["order_date"] = pd.to_datetime(order["order_date"], errors="coerce")
    payment["payment_value"] = numeric_column(payment, "payment_value", "Payment.xlsx")
    returns["refund_amount"] = numeric_column(returns, "refund_amount", "Return.xlsx")

    problems: list[dict[str, object]] = []
    checks = [
        ("Employee PK trùng", int(employee.duplicated("sales_employee_id").sum())),
        ("Order PK trùng", int(order.duplicated("order_id").sum())),
        ("Payment PK trùng", int(payment.duplicated("order_id").sum())),
        ("Return PK trùng", int(returns.duplicated("return_id").sum())),
        ("Order ngày không hợp lệ", int(order["order_date"].isna().sum())),
        ("Order thiếu Employee", int((~order["sales_employee_id"].isin(employee["sales_employee_id"])).sum())),
        ("Payment thiếu Order", int((~payment["order_id"].isin(order["order_id"])).sum())),
        ("Return thiếu Order", int((~returns["order_id"].isin(order["order_id"])).sum())),
        ("Order thiếu Payment", int((~order["order_id"].isin(payment["order_id"])).sum())),
    ]
    for name, count in checks:
        problems.append({"check": name, "error_rows": count, "passed": count == 0})
    quality = pd.DataFrame(problems)
    failed = quality.loc[~quality["passed"]]
    if not failed.empty:
        raise ValueError("Kiểm tra dữ liệu không đạt:\n" + failed.to_string(index=False))

    tables = {"Employee": employee, "Order": order, "Payment": payment, "Return": returns}
    return tables, quality


def build_kpi_model(tables: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    employee, order = tables["Employee"], tables["Order"]
    payment, returns = tables["Payment"], tables["Return"]

    refund_by_order = returns.groupby("order_id", as_index=False)["refund_amount"].sum()
    fact = (
        order[["order_id", "sales_employee_id", "order_date", "order_status"]]
        .merge(payment[["order_id", "payment_value"]], on="order_id", validate="one_to_one")
        .merge(refund_by_order, on="order_id", how="left", validate="one_to_one")
    )
    fact["refund_amount"] = fact["refund_amount"].fillna(0.0)
    fact["is_valid_order"] = fact["order_status"].isin(VALID_STATUSES)
    fact["gross_sales"] = np.where(fact["is_valid_order"], fact["payment_value"], 0.0)
    fact["recognized_refund"] = np.where(fact["is_valid_order"], fact["refund_amount"], 0.0)
    fact["net_sales"] = fact["gross_sales"] - fact["recognized_refund"]
    fact["month"] = fact["order_date"].dt.to_period("M")
    fact["year"] = fact["order_date"].dt.year

    start_month, end_month = fact["month"].min(), fact["month"].max()
    all_months = pd.period_range(start_month, end_month, freq="M")
    employee_ids = employee["sales_employee_id"].drop_duplicates().sort_values()
    full_index = pd.MultiIndex.from_product(
        [employee_ids, all_months], names=["sales_employee_id", "month"]
    )

    monthly = fact.groupby(["sales_employee_id", "month"], observed=True).agg(
        gross_sales=("gross_sales", "sum"),
        refund_amount=("recognized_refund", "sum"),
        net_sales=("net_sales", "sum"),
        valid_orders=("is_valid_order", "sum"),
    ).reindex(full_index, fill_value=0).reset_index()
    monthly["month_date"] = monthly["month"].dt.to_timestamp()
    monthly["year"] = monthly["month"].dt.year
    monthly["year_month"] = monthly["month"].astype(str)
    monthly["average_net_sales_per_order"] = np.divide(
        monthly["net_sales"], monthly["valid_orders"],
        out=np.full(len(monthly), np.nan), where=monthly["valid_orders"].ne(0)
    )
    system_month = monthly.groupby("month")["net_sales"].transform("sum")
    monthly["sales_contribution_pct"] = np.divide(
        monthly["net_sales"], system_month,
        out=np.zeros(len(monthly)), where=system_month.ne(0)
    )
    group = monthly.groupby("sales_employee_id", sort=False)
    monthly["previous_month_net_sales"] = group["net_sales"].shift(1)
    monthly["mom_growth_pct"] = (
        (monthly["net_sales"] - monthly["previous_month_net_sales"])
        / monthly["previous_month_net_sales"].replace(0, np.nan)
    )
    monthly["previous_year_net_sales"] = group["net_sales"].shift(12)
    monthly["yoy_growth_pct"] = (
        (monthly["net_sales"] - monthly["previous_year_net_sales"])
        / monthly["previous_year_net_sales"].replace(0, np.nan)
    )
    monthly["cumulative_net_sales"] = group["net_sales"].cumsum()
    monthly["rolling_12m_net_sales"] = group["net_sales"].transform(
        lambda s: s.rolling(12, min_periods=1).sum()
    )
    monthly["previous_rolling_12m"] = group["rolling_12m_net_sales"].shift(12)
    monthly["rolling_12m_growth_pct"] = (
        (monthly["rolling_12m_net_sales"] - monthly["previous_rolling_12m"])
        / monthly["previous_rolling_12m"].replace(0, np.nan)
    )
    monthly = monthly.merge(
        employee[["sales_employee_id", "sales_employee_name", "years_experience"]],
        on="sales_employee_id", how="left", validate="many_to_one"
    )

    annual = monthly.groupby(
        ["sales_employee_id", "sales_employee_name", "year"], as_index=False
    ).agg(
        gross_sales=("gross_sales", "sum"),
        refund_amount=("refund_amount", "sum"),
        net_sales=("net_sales", "sum"),
        valid_orders=("valid_orders", "sum"),
    )
    annual["average_net_sales_per_order"] = annual["net_sales"] / annual["valid_orders"].replace(0, np.nan)
    annual_total = annual.groupby("year")["net_sales"].transform("sum")
    annual["sales_contribution_pct"] = annual["net_sales"] / annual_total.replace(0, np.nan)
    annual["previous_year_net_sales"] = annual.groupby("sales_employee_id")["net_sales"].shift(1)
    annual["yoy_growth_pct"] = (
        (annual["net_sales"] - annual["previous_year_net_sales"])
        / annual["previous_year_net_sales"].replace(0, np.nan)
    )
    return {"fact": fact, "monthly": monthly, "annual": annual}


def period_employee_summary(monthly: pd.DataFrame, start: pd.Timestamp, end: pd.Timestamp) -> pd.DataFrame:
    selected = monthly.loc[monthly["month_date"].between(start, end)].copy()
    grouped = selected.groupby(["sales_employee_id", "sales_employee_name"], as_index=False).agg(
        gross_sales=("gross_sales", "sum"),
        refund_amount=("refund_amount", "sum"),
        net_sales=("net_sales", "sum"),
        valid_orders=("valid_orders", "sum"),
        average_mom_growth_pct=("mom_growth_pct", "mean"),
        average_yoy_growth_pct=("yoy_growth_pct", "mean"),
        mean_monthly_sales=("net_sales", "mean"),
        monthly_sales_std=("net_sales", lambda s: s.std(ddof=0)),
        positive_growth_month_pct=("mom_growth_pct", lambda s: float((s > 0).sum() / s.notna().sum()) if s.notna().any() else np.nan),
    )
    total = grouped["net_sales"].sum()
    grouped["sales_contribution_pct"] = grouped["net_sales"] / total if total else np.nan
    grouped["average_net_sales_per_order"] = grouped["net_sales"] / grouped["valid_orders"].replace(0, np.nan)
    grouped["sales_volatility_pct"] = grouped["monthly_sales_std"] / grouped["mean_monthly_sales"].replace(0, np.nan)
    grouped["sales_stability_pct"] = 1 / (1 + grouped["sales_volatility_pct"])
    return grouped.sort_values("net_sales", ascending=False).reset_index(drop=True)


def dashboard_payload(model: dict[str, pd.DataFrame], employee_id: str, start, end):
    monthly = model["monthly"]
    start_ts = pd.Timestamp(start).to_period("M").to_timestamp()
    end_ts = pd.Timestamp(end).to_period("M").to_timestamp()
    ranking = period_employee_summary(monthly, start_ts, end_ts)
    selected_all = monthly.loc[monthly["month_date"].between(start_ts, end_ts)]

    if employee_id == "__ALL__":
        timeline = selected_all.groupby(["month", "month_date", "year_month"], as_index=False).agg(
            gross_sales=("gross_sales", "sum"), refund_amount=("refund_amount", "sum"),
            net_sales=("net_sales", "sum"), valid_orders=("valid_orders", "sum")
        )
        timeline["sales_contribution_pct"] = 1.0
        title = "Toàn bộ nhân viên"
    else:
        timeline = selected_all.loc[selected_all["sales_employee_id"].eq(employee_id)].copy()
        row = ranking.loc[ranking["sales_employee_id"].eq(employee_id)]
        title = employee_id if row.empty else f"{employee_id} – {row.iloc[0]['sales_employee_name']}"

    timeline = timeline.sort_values("month_date")
    timeline["previous_month_net_sales"] = timeline["net_sales"].shift(1)
    timeline["mom_growth_pct"] = (
        (timeline["net_sales"] - timeline["previous_month_net_sales"])
        / timeline["previous_month_net_sales"].replace(0, np.nan)
    )
    timeline["previous_year_net_sales"] = timeline["net_sales"].shift(12)
    timeline["yoy_growth_pct"] = (
        (timeline["net_sales"] - timeline["previous_year_net_sales"])
        / timeline["previous_year_net_sales"].replace(0, np.nan)
    )
    timeline["rolling_12m_net_sales"] = timeline["net_sales"].rolling(12, min_periods=1).sum()

    gross = timeline["gross_sales"].sum()
    refund = timeline["refund_amount"].sum()
    net = timeline["net_sales"].sum()
    orders = timeline["valid_orders"].sum()
    mean = timeline["net_sales"].mean()
    volatility = timeline["net_sales"].std(ddof=0) / mean if mean else np.nan
    system_total = selected_all["net_sales"].sum()
    if employee_id == "__ALL__":
        contribution = 1.0
    else:
        contribution = net / system_total if system_total else np.nan
    summary = {
        "gross_sales": gross, "refund_amount": refund, "net_sales": net,
        "valid_orders": int(orders), "average_net_sales_per_order": net / orders if orders else np.nan,
        "sales_contribution_pct": contribution,
        "latest_mom_growth_pct": timeline["mom_growth_pct"].dropna().iloc[-1] if timeline["mom_growth_pct"].notna().any() else np.nan,
        "latest_yoy_growth_pct": timeline["yoy_growth_pct"].dropna().iloc[-1] if timeline["yoy_growth_pct"].notna().any() else np.nan,
        "average_mom_growth_pct": timeline["mom_growth_pct"].mean(),
        "average_yoy_growth_pct": timeline["yoy_growth_pct"].mean(),
        "positive_growth_month_pct": (timeline["mom_growth_pct"] > 0).sum() / timeline["mom_growth_pct"].notna().sum() if timeline["mom_growth_pct"].notna().any() else np.nan,
        "sales_volatility_pct": volatility,
        "sales_stability_pct": 1 / (1 + volatility) if pd.notna(volatility) else np.nan,
    }
    return title, timeline, ranking, summary


def format_number(value) -> str:
    return "—" if pd.isna(value) else f"{value:,.2f}"


def format_pct(value) -> str:
    return "—" if pd.isna(value) else f"{value:.2%}"


def make_figures(title: str, timeline: pd.DataFrame, ranking: pd.DataFrame):
    sales = go.Figure()
    sales.add_trace(go.Scatter(x=timeline["month_date"], y=timeline["net_sales"], name="Doanh thu thuần", mode="lines+markers"))
    sales.add_trace(go.Scatter(x=timeline["month_date"], y=timeline["rolling_12m_net_sales"], name="Lũy kế 12 tháng", mode="lines"))
    sales.update_layout(title=f"Doanh thu theo thời gian — {title}", hovermode="x unified", template="plotly_white", yaxis_title="Doanh thu")

    growth = go.Figure()
    growth.add_trace(go.Bar(x=timeline["month_date"], y=timeline["mom_growth_pct"] * 100, name="MoM %"))
    growth.add_trace(go.Scatter(x=timeline["month_date"], y=timeline["yoy_growth_pct"] * 100, name="YoY %", mode="lines+markers"))
    growth.add_hline(y=0, line_color="#666")
    growth.update_layout(title="Tăng trưởng MoM và YoY", hovermode="x unified", template="plotly_white", yaxis_title="%")

    contribution = px.line(timeline, x="month_date", y=timeline["sales_contribution_pct"] * 100,
                           title=f"Tỷ trọng đóng góp theo tháng — {title}", labels={"y": "% đóng góp", "month_date": "Tháng"})
    contribution.update_layout(template="plotly_white")

    scatter_data = ranking.replace([np.inf, -np.inf], np.nan).dropna(subset=["average_yoy_growth_pct", "sales_stability_pct"])
    scatter = px.scatter(
        scatter_data, x="average_yoy_growth_pct", y="sales_contribution_pct",
        size="net_sales", color="sales_stability_pct", hover_name="sales_employee_name",
        hover_data=["sales_employee_id", "net_sales", "positive_growth_month_pct"],
        title="Ma trận đóng góp – tăng trưởng – ổn định của nhân viên",
        labels={"average_yoy_growth_pct": "YoY trung bình", "sales_contribution_pct": "Tỷ trọng đóng góp", "sales_stability_pct": "Điểm ổn định"},
        color_continuous_scale="Viridis",
    )
    scatter.update_layout(template="plotly_white")
    scatter.update_xaxes(tickformat=".1%")
    scatter.update_yaxes(tickformat=".2%")
    return sales, growth, contribution, scatter


def summary_html(title: str, summary: dict[str, object]) -> str:
    cards = [
        ("Doanh thu thuần", format_number(summary["net_sales"])),
        ("Tỷ trọng đóng góp", format_pct(summary["sales_contribution_pct"])),
        ("Số đơn hợp lệ", f"{summary['valid_orders']:,}"),
        ("Doanh thu/đơn", format_number(summary["average_net_sales_per_order"])),
        ("MoM mới nhất", format_pct(summary["latest_mom_growth_pct"])),
        ("YoY mới nhất", format_pct(summary["latest_yoy_growth_pct"])),
        ("MoM trung bình", format_pct(summary["average_mom_growth_pct"])),
        ("YoY trung bình", format_pct(summary["average_yoy_growth_pct"])),
        ("Tháng tăng trưởng", format_pct(summary["positive_growth_month_pct"])),
        ("Biến động", format_pct(summary["sales_volatility_pct"])),
        ("Điểm ổn định", format_pct(summary["sales_stability_pct"])),
    ]
    body = "".join(
        f'<div style="padding:12px;border:1px solid #dce3ec;border-radius:8px;background:#fff"><div style="color:#607086;font-size:12px">{html.escape(label)}</div><div style="font-size:20px;font-weight:700;color:#17365d">{html.escape(value)}</div></div>'
        for label, value in cards
    )
    return f'<h2 style="color:#17365d">KPI Problem statement 3 — {html.escape(title)}</h2><div style="display:grid;grid-template-columns:repeat(4,minmax(150px,1fr));gap:8px">{body}</div>'


def export_outputs(model: dict[str, pd.DataFrame], quality: pd.DataFrame, output_dir: Path) -> tuple[Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    monthly = model["monthly"].copy()
    annual = model["annual"].copy()
    overall = period_employee_summary(monthly, monthly["month_date"].min(), monthly["month_date"].max())
    monthly["month"] = monthly["month"].astype(str)
    for frame in [monthly, annual, overall]:
        numeric = frame.select_dtypes(include=[np.number]).columns
        frame[numeric] = frame[numeric].replace([np.inf, -np.inf], np.nan)

    excel_path = output_dir / "KPI_Problem_Statement_3.xlsx"
    with pd.ExcelWriter(excel_path, engine="xlsxwriter") as writer:
        MEASURE_DICTIONARY.to_excel(writer, sheet_name="Measure_Dictionary", index=False)
        METRIC_DICTIONARY.to_excel(writer, sheet_name="Metric_Dictionary", index=False)
        KPI_DICTIONARY.to_excel(writer, sheet_name="KPI_Dictionary", index=False)
        monthly.to_excel(writer, sheet_name="Monthly_Employee_KPI", index=False)
        annual.to_excel(writer, sheet_name="Annual_Employee_KPI", index=False)
        overall.to_excel(writer, sheet_name="Employee_Summary", index=False)
        quality.to_excel(writer, sheet_name="Data_Quality", index=False)
        workbook = writer.book
        header = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#274B70", "border": 0})
        pct = workbook.add_format({"num_format": "0.00%"})
        money = workbook.add_format({"num_format": "#,##0.00"})
        for sheet_name, frame in {
            "Measure_Dictionary": MEASURE_DICTIONARY, "Metric_Dictionary": METRIC_DICTIONARY,
            "KPI_Dictionary": KPI_DICTIONARY, "Monthly_Employee_KPI": monthly,
            "Annual_Employee_KPI": annual, "Employee_Summary": overall, "Data_Quality": quality,
        }.items():
            ws = writer.sheets[sheet_name]
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, len(frame), max(len(frame.columns) - 1, 0))
            ws.hide_gridlines(2)
            for i, col in enumerate(frame.columns):
                ws.write(0, i, col, header)
                width = min(max(len(str(col)) + 2, 12), 30)
                fmt = pct if str(col).endswith("_pct") else money if any(x in str(col) for x in ["sales", "refund"]) else None
                ws.set_column(i, i, width, fmt)

    title, timeline, ranking, summary = dashboard_payload(
        model, "__ALL__", monthly["month_date"].min(), monthly["month_date"].max()
    )
    figures = make_figures(title, timeline, ranking)
    fragments = [pio.to_html(fig, full_html=False, include_plotlyjs="cdn" if i == 0 else False) for i, fig in enumerate(figures)]
    html_path = output_dir / "KPI_Problem_Statement_3.html"
    html_path.write_text(
        "<!doctype html><html><head><meta charset='utf-8'><title>KPI PS3</title></head><body style='font-family:Arial;background:#f5f7fa;padding:20px'>"
        + summary_html(title, summary) + "".join(fragments) + "</body></html>", encoding="utf-8"
    )
    return excel_path, html_path


def show_dashboard(model: dict[str, pd.DataFrame]) -> None:
    if IN_COLAB:
        from google.colab import output
        output.enable_custom_widget_manager()

    monthly = model["monthly"]
    employees = (
        monthly[["sales_employee_id", "sales_employee_name"]].drop_duplicates()
        .sort_values(["sales_employee_name", "sales_employee_id"])
    )
    options = [("Toàn bộ nhân viên", "__ALL__")] + [
        (f"{row.sales_employee_id} – {row.sales_employee_name}", row.sales_employee_id)
        for row in employees.itertuples(index=False)
    ]
    employee_widget = widgets.Dropdown(options=options, value="__ALL__", description="Nhân viên:", layout=widgets.Layout(width="520px"))
    start_widget = widgets.DatePicker(description="Từ tháng:", value=monthly["month_date"].min().date())
    end_widget = widgets.DatePicker(description="Đến tháng:", value=monthly["month_date"].max().date())
    refresh = widgets.Button(description="Cập nhật dashboard", button_style="primary", icon="refresh")
    output_area = widgets.Output()

    def render(_=None):
        with output_area:
            clear_output(wait=True)
            if start_widget.value is None or end_widget.value is None or start_widget.value > end_widget.value:
                print("Khoảng thời gian không hợp lệ.")
                return
            title, timeline, ranking, summary = dashboard_payload(
                model, employee_widget.value, start_widget.value, end_widget.value
            )
            display(HTML(summary_html(title, summary)))
            for fig in make_figures(title, timeline, ranking):
                fig.show()
            table = ranking.head(20).copy()
            pct_cols = [c for c in table if c.endswith("_pct")]
            display(HTML("<h3>Top 20 nhân viên trong khoảng thời gian chọn</h3>"))
            display(table.style.format({
                "net_sales": "{:,.2f}", "gross_sales": "{:,.2f}", "refund_amount": "{:,.2f}",
                **{c: "{:.2%}" for c in pct_cols},
            }).hide(axis="index"))

    refresh.on_click(render)
    display(widgets.VBox([widgets.HBox([employee_widget, start_widget, end_widget]), refresh, output_area]))
    render()


def main() -> None:
    parser = argparse.ArgumentParser(description="KPI Problem statement 3", add_help=True)
    parser.add_argument("--data-dir", default="/content" if IN_COLAB else ".", help="Thư mục chứa các workbook")
    parser.add_argument("--output-dir", default="/content/KPI_PS3_Output" if IN_COLAB else "KPI_PS3_Output")
    parser.add_argument("--no-dashboard", action="store_true", help="Chỉ kiểm tra và xuất file, không hiện widget")
    args, _unknown = parser.parse_known_args()

    paths = locate_files(Path(args.data_dir))
    print("Đã nhận diện dữ liệu:")
    for name in REQUIRED_FILES:
        print(f"- {name}: {paths[name]}")
    tables, quality = load_and_validate(paths)
    model = build_kpi_model(tables)
    excel_path, html_path = export_outputs(model, quality, Path(args.output_dir))
    print("\nKiểm tra dữ liệu: ĐẠT")
    print(f"Khoảng thời gian: {model['fact']['order_date'].min().date()} đến {model['fact']['order_date'].max().date()}")
    print(f"Số nhân viên: {model['monthly']['sales_employee_id'].nunique():,}")
    print(f"Doanh thu thuần: {model['fact']['net_sales'].sum():,.2f}")
    print(f"Đã xuất: {excel_path.resolve()}")
    print(f"Đã xuất: {html_path.resolve()}")
    if not args.no_dashboard:
        show_dashboard(model)


if __name__ == "__main__":
    main()
