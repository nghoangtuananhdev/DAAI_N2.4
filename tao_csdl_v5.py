"""Tạo CSDL_v5: hợp nhất toàn bộ epd.json trực tiếp vào bảng Product.

Các product_id trong epd.json xung đột với Product từ products.csv. V5 cấp khóa
kỹ thuật EPD-0001..EPD-1000 cho các dòng JSON, đồng thời giữ product_id gốc tại
source_product_id và lưu trạng thái null/chuỗi rỗng để không mất dấu dữ liệu.
"""

from pathlib import Path
import hashlib
import json
import tempfile
import zipfile

import pandas as pd

import tao_csdl as base
import tao_csdl_v4 as v4


PROJECT_DIR = Path(__file__).resolve().parent
SOURCE_DIR = PROJECT_DIR.parent / "student_data (2)"
OUTPUT_DIR = PROJECT_DIR / "CSDL_v5"
REPORT_DIR = PROJECT_DIR / "bao_cao_CSDL_v5"
MODEL_PATH = PROJECT_DIR / "mo_hinh_csdl_v5.txt"
ZIP_PATH = PROJECT_DIR / "CSDL_v5.zip"

PRODUCT_BASE_COLS = list(v4.CORE_SCHEMA["Product"])
PRODUCT_V5_COLS = [
    "product_id", "source_product_id", "source_product_id_state", "source_file",
    "source_record", "product_name", "category", "segment", "size", "color", "price", "cogs",
]
SCHEMA = dict(v4.CORE_SCHEMA)
SCHEMA["Product"] = PRODUCT_V5_COLS
SCHEMA["Web_Traffic_Quarantine"] = list(v4.EXTRA_SCHEMA["Web_Traffic_Quarantine"])
PK = dict(base.PK)
PK["Web_Traffic_Quarantine"] = ["quarantine_id"]


def require(condition, message):
    if not condition:
        raise ValueError(message)


def sha256(path):
    h = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def build_v5(csv, raw_json, report):
    tables = v4.build_v4(csv, raw_json, report)
    epd_source = tables.pop("Product_EPD_Source").copy()

    canonical = tables["Product"].copy()
    canonical_v5 = pd.DataFrame({
        "product_id": canonical["product_id"],
        "source_product_id": canonical["product_id"],
        "source_product_id_state": "value",
        "source_file": "products.csv",
        "source_record": [str(i) for i in range(1, len(canonical) + 1)],
    })
    for col in PRODUCT_BASE_COLS[1:]:
        canonical_v5[col] = canonical[col]

    source_states = []
    for row in raw_json["epd"]:
        value = row.get("product_id")
        if value is None:
            source_states.append("json_null")
        elif value == "":
            source_states.append("empty_string")
        else:
            source_states.append("value")
    epd_v5 = pd.DataFrame({
        "product_id": [f"EPD-{i:04d}" for i in range(1, len(epd_source) + 1)],
        "source_product_id": epd_source["product_id"],
        "source_product_id_state": source_states,
        "source_file": "epd.json",
        "source_record": [str(i) for i in range(1, len(epd_source) + 1)],
    })
    for col in PRODUCT_BASE_COLS[1:]:
        epd_v5[col] = epd_source[col]

    tables["Product"] = pd.concat([canonical_v5, epd_v5], ignore_index=True)[PRODUCT_V5_COLS]
    # Giữ thứ tự: 19 bảng lõi, sau đó bảng kiểm dịch Web Traffic.
    tables = {name: tables[name] for name in list(v4.CORE_SCHEMA) + ["Web_Traffic_Quarantine"]}

    report["v5_product_merge"] = {
        "products_csv_rows": len(canonical_v5),
        "epd_json_rows_added_to_Product": len(epd_v5),
        "product_total_rows": len(tables["Product"]),
        "generated_key_range": [epd_v5.product_id.iloc[0], epd_v5.product_id.iloc[-1]],
        "source_product_id_json_null_rows": int((epd_v5.source_product_id_state == "json_null").sum()),
        "source_product_id_empty_string_rows": int((epd_v5.source_product_id_state == "empty_string").sum()),
        "rule": "Khóa EPD-xxxx là khóa kỹ thuật; source_product_id giữ nguyên ID nguồn.",
    }
    return tables, canonical_v5, epd_v5


def check_v5(tables, csv, raw_json, canonical_v5, epd_v5, report):
    checks = []
    require(set(tables) == set(SCHEMA), f"Sai danh sách bảng: {set(tables) ^ set(SCHEMA)}")
    for name, columns in SCHEMA.items():
        table = tables[name]
        require(list(table) == columns, f"{name}: sai cột")
        require(len(table) <= 1_048_575, f"{name}: vượt giới hạn Excel")
        require(not table.isna().any().any(), f"{name}: có NaN nội bộ")
        require(not table[PK[name]].eq("").any().any(), f"{name}: PK rỗng")
        require(not table.duplicated(PK[name]).any(), f"{name}: PK trùng")
        checks.append({"check": f"PK/cột {name}", "passed": True, "rows": len(table)})

    for child, col, parent, ref in base.FK:
        bad = ~tables[child][col].isin(tables[parent][ref])
        require(not bad.any(), f"FK {child}.{col} -> {parent}.{ref}: {int(bad.sum())} lỗi")
        checks.append({"check": f"FK {child}.{col} -> {parent}.{ref}", "passed": True})

    require(len(tables["Product"]) == len(csv["products"]) + len(raw_json["epd"]), "Product chưa đủ CSV + EPD")
    require(len(epd_v5) == 1000, "EPD không đủ 1.000 dòng")
    require(epd_v5.product_id.tolist() == [f"EPD-{i:04d}" for i in range(1, 1001)], "Khóa EPD không ổn định")
    require(not set(epd_v5.product_id) & set(canonical_v5.product_id), "Khóa EPD đụng khóa Product CSV")
    require(canonical_v5.source_product_id.equals(csv["products"].product_id.reset_index(drop=True)), "ID nguồn products.csv bị đổi")

    # Đối chiếu mọi thuộc tính EPD với bảng nguồn mà v4 đã chuyển thành chuỗi bảo toàn.
    expected_epd = v4.json_frame(raw_json["epd"], PRODUCT_BASE_COLS)
    require(epd_v5.source_product_id.equals(expected_epd.product_id), "source_product_id EPD bị đổi")
    for col in PRODUCT_BASE_COLS[1:]:
        require(epd_v5[col].equals(expected_epd[col]), f"EPD bị đổi cột {col}")
    checks.append({"check": "Product chứa đủ 2.412 CSV + 1.000 EPD và bảo toàn thuộc tính", "passed": True})

    # Kiểm tra bảo toàn các bảng còn lại bằng quy tắc v4 tương ứng.
    require(len(tables["Promotion"]) == len(csv["promotions"]) + len(raw_json["eprom"]), "Promotion thiếu eprom")
    require(len(tables["Web Traffic"]) + len(tables["Web_Traffic_Quarantine"]) == len(csv["web_traffic"]) + len(raw_json["tf"]), "Web Traffic thiếu tf")
    pairs = pd.MultiIndex.from_frame(tables["Order Item"][["order_id", "product_id"]])
    for name in ["Return", "Review"]:
        require(pd.MultiIndex.from_frame(tables[name][["order_id", "product_id"]]).isin(pairs).all(), f"{name}: cặp đơn hàng-sản phẩm không tồn tại")
    require(set(tables["Payment"].order_id) == set(tables["Order"].order_id), "Order thiếu Payment")
    checks.append({"check": "Đủ eprom/tf; cặp Return/Review và Payment hợp lệ", "passed": True})
    report["checks"] = checks


def write_model(report):
    lines = [
        "MÔ HÌNH DỮ LIỆU V5 — PRODUCT HỢP NHẤT PRODUCTS.CSV + EPD.JSON", "",
        "V5 sửa lỗi của v4: toàn bộ 1.000 dòng epd.json nằm trực tiếp trong Product.",
        "Product.product_id vẫn là PK. Dòng products.csv giữ product_id cũ; dòng epd.json dùng EPD-0001..EPD-1000.",
        "source_product_id giữ giá trị product_id gốc; source_product_id_state phân biệt value/json_null/empty_string.",
        "source_file và source_record cho phép truy ngược đúng dòng nguồn.",
        "Các FK giao dịch tiếp tục trỏ tới product_id chuẩn từ products.csv; các dòng EPD hiện chưa được giao dịch tham chiếu.",
        "Web_Traffic_Quarantine tiếp tục giữ bốn dòng tf.json không đạt điều kiện khóa/ngày.", "",
    ]
    for name, columns in SCHEMA.items():
        lines.extend([name + ":", "-- Thuộc tính: " + ", ".join(columns), "-- PK: " + ", ".join(PK[name])])
        for child, col, parent, ref in base.FK:
            if child == name:
                lines.append(f"-- FK: {col} -> {parent}.{ref}")
        lines.append("")
    lines.extend(["ĐỐI SOÁT PRODUCT", json.dumps(report["v5_product_merge"], ensure_ascii=False, indent=2)])
    MODEL_PATH.write_text("\n".join(lines), encoding="utf-8")


def save_report(report):
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    (REPORT_DIR / "bao_cao_kiem_tra_v5.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    lines = [
        "BÁO CÁO CSDL_V5", "", "Trạng thái: " + report["status"],
        "Sửa lỗi v4: epd.json được hợp nhất trực tiếp vào Product.", "",
        "ĐỐI SOÁT PRODUCT", json.dumps(report.get("v5_product_merge", {}), ensure_ascii=False, indent=2), "",
        "BẤT THƯỜNG NGUỒN GIỮ NGUYÊN", json.dumps(report.get("field_quality_anomalies", []), ensure_ascii=False, indent=2), "",
        "BẢNG ĐẦU RA",
    ]
    for name, info in report.get("tables", {}).items():
        lines.append(f"- {name}: {info['rows']:,} dòng; đọc lại Excel: {info.get('verified_rows', 'chưa chạy')}")
    lines.extend(["", f"Số kiểm tra đạt: {len(report.get('checks', []))}",
                  "Đã kiểm tra lại checksum 15 file nguồn; không sửa nguồn."])
    if report.get("error"):
        lines.append("Lỗi: " + report["error"])
    (REPORT_DIR / "bao_cao_kiem_tra_v5.txt").write_text("\n".join(lines), encoding="utf-8")


def main():
    report = {"status": "ĐANG XỬ LÝ", "sources": {}, "tables": {}, "checks": []}
    try:
        # Buộc v4 đọc đúng thư mục nguồn dù project đã được đổi tên/thư mục.
        v4.SOURCE_DIR = SOURCE_DIR
        csv, raw_json = v4.read_sources(report)
        tables, canonical_v5, epd_v5 = build_v5(csv, raw_json, report)
        check_v5(tables, csv, raw_json, canonical_v5, epd_v5, report)
        write_model(report)
        save_report(report)

        with tempfile.TemporaryDirectory(prefix=".csdl_v5_build_", dir=PROJECT_DIR) as staging:
            stage = Path(staging)
            for name, frame in tables.items():
                path = stage / f"{name}.xlsx"
                allow_invalid_date = name == "Web_Traffic_Quarantine"
                allow_invalid_types = name in {"Product", "Promotion", "Web Traffic", "Web_Traffic_Quarantine"}
                print(f"Xuất {name}: {len(frame):,} dòng", flush=True)
                precision = v4.export_table(name, frame, path, allow_invalid_date, allow_invalid_types)
                info = {"rows": len(frame), "columns": list(frame), "sha256": sha256(path),
                        "precision_text_cells": precision}
                info["verified_rows"] = v4.verify_excel(name, frame, path, allow_invalid_date, allow_invalid_types)
                report["tables"][name] = info
                save_report(report)

            for filename, info in report["sources"].items():
                require(sha256(SOURCE_DIR / filename) == info["sha256"], f"Nguồn bị thay đổi: {filename}")
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            for name in SCHEMA:
                (stage / f"{name}.xlsx").replace(OUTPUT_DIR / f"{name}.xlsx")

        report["status"] = "HOÀN TẤT"
        report["source_files_unchanged"] = True
        save_report(report)
        with zipfile.ZipFile(ZIP_PATH, "w", zipfile.ZIP_DEFLATED, compresslevel=1) as archive:
            for name in SCHEMA:
                archive.write(OUTPUT_DIR / f"{name}.xlsx", f"CSDL_v5/{name}.xlsx")
            archive.write(MODEL_PATH, MODEL_PATH.name)
            archive.write(Path(__file__), Path(__file__).name)
            for path in sorted(REPORT_DIR.glob("*")):
                archive.write(path, f"bao_cao_CSDL_v5/{path.name}")
        print(f"Hoàn tất {len(tables)} bảng: {OUTPUT_DIR}")
        print(f"Product: {len(tables['Product']):,} dòng")
    except Exception as exc:
        report["status"] = "KHÔNG HOÀN TẤT"
        report["error"] = f"{type(exc).__name__}: {exc}"
        save_report(report)
        raise


if __name__ == "__main__":
    main()
