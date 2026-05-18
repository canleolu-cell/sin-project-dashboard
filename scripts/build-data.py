import argparse
import json
import re
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import openpyxl


HEADER_ALIASES = {
    "import_lot": "进口批次 / 货品",
    "import_amount": "进口发票金额 (USD)",
    "payment_status": "货款已付",
    "export_lot": "出口批次(Export Lot)",
    "import_containers": "进口柜数",
    "export_containers": "出口柜数",
    "load_plan": "单柜装载分布",
    "pallets": "托数",
    "boxes": "箱数",
    "import_pcs": "进口件数(Import PCS)",
    "export_pcs": "出口件数(Export PCS)",
    "net_weight_kg": "发沪净重(KG)",
    "gross_weight_kg": "发沪毛重(KG)",
    "shanghai_unit_price": "到沪电芯单价(USD/EA)",
    "shanghai_sales_amount": "到沪销售总价(USD)",
    "europe_etd": "欧洲ETD",
    "singapore_eta": "新加坡ETA",
    "second_leg_vessel": "二程发沪船名",
    "singapore_etd": "新加坡ETD",
    "shanghai_eta": "上海ETA",
    "remark": "最新物流与节点状态",
}


def normalize_header(value):
    text = "" if value is None else str(value)
    text = re.sub(r"<br\s*/?>", "", text, flags=re.IGNORECASE)
    return re.sub(r"\s+", "", text)


def clean_text(value):
    if value is None:
        return ""
    text = str(value)
    text = re.sub(r"<br\s*/?>", " ", text, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", text).strip()


def serializable(value):
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, Decimal):
        return float(value)
    return value


def normalize_date(value):
    value = serializable(value)
    return clean_text(value) if value not in (None, "") else "待定"


def parse_date(value, default_year):
    text = normalize_date(value)
    if not text or text in {"待定", "待确认", "--"}:
        return None

    iso_match = re.search(r"(20\d{2})[-/](\d{1,2})[-/](\d{1,2})", text)
    if iso_match:
        year, month, day = map(int, iso_match.groups())
        return date(year, month, day)

    us_match = re.search(r"(\d{1,2})/(\d{1,2})/(20\d{2})", text)
    if us_match:
        month, day, year = map(int, us_match.groups())
        return date(year, month, day)

    chinese_match = re.search(r"(\d{1,2})月(\d{1,2})日", text)
    if chinese_match:
        month, day = map(int, chinese_match.groups())
        return date(default_year, month, day)

    return None


def to_float(value):
    if value in (None, ""):
        return 0.0
    if isinstance(value, str):
        text = value.replace(",", "").replace("$", "").strip()
        match = re.search(r"-?\d+(?:\.\d+)?", text)
        return float(match.group()) if match else 0.0
    return float(value)


def to_number(value):
    number = to_float(value)
    return int(number) if number.is_integer() else number


def parse_lot(raw_value):
    text = clean_text(raw_value)
    match = re.search(r"Lot\s*([0-9]+(?:\.[0-9]+)?)", text, re.IGNORECASE)
    if not match:
        return None

    lot_id = match.group(1)
    product = text[match.end():].strip(" -*()")
    if not product:
        product = "电池电芯"

    product_type = "Module" if "模组" in product else "Cell"
    return {
        "id": lot_id,
        "idNumber": float(lot_id),
        "product": product,
        "type": product_type,
    }


def infer_payment_status(lot_number):
    return "已付" if lot_number < 11 else "待付"


def normalize_payment_status(value, lot_number):
    text = clean_text(value)
    if not text:
        return infer_payment_status(lot_number)

    lowered = text.lower()
    if lowered in {"yes", "y", "paid", "true", "已付", "已付款"}:
        return "已付"
    if lowered in {"no", "n", "unpaid", "false", "待付", "未付", "未付款"}:
        return "待付"
    if "estimate" in lowered or "预计" in text or "暂定" in text:
        due = text.replace("estimate", "").strip()
        return f"预计付款 {due}" if due else text
    return text


def infer_status(remark, europe_etd, singapore_eta, singapore_etd, shanghai_eta, as_of_date):
    combined = f"{remark} {singapore_eta} {singapore_etd} {shanghai_eta}"
    if "已抵上海" in combined or "已抵沪" in combined:
        return "抵沪"
    if singapore_etd and singapore_etd != "待定" and shanghai_eta and shanghai_eta != "待定":
        return "在途"

    singapore_eta_date = parse_date(singapore_eta, as_of_date.year)
    if singapore_eta_date and singapore_eta_date <= as_of_date:
        return "在新"
    if europe_etd and europe_etd not in {"待定", "待确认"}:
        return "在途"
    return "待排"


def infer_stage(status, singapore_etd, shanghai_eta):
    if status == "抵沪":
        return "arrived_shanghai"
    if singapore_etd != "待定" and shanghai_eta != "待定":
        return "in_second_leg"
    if status == "在新":
        return "in_singapore"
    return "first_leg"


def build_column_map(header_row):
    normalized_to_index = {
        normalize_header(value): index
        for index, value in enumerate(header_row)
        if value is not None
    }
    column_map = {}
    missing = []

    for field, header in HEADER_ALIASES.items():
        normalized = normalize_header(header)
        if normalized not in normalized_to_index:
            missing.append(header)
        else:
            column_map[field] = normalized_to_index[normalized]

    if missing:
        raise ValueError(f"Excel 缺少必要表头: {', '.join(missing)}")
    return column_map


def cell(row, column_map, field):
    index = column_map[field]
    return row[index] if index < len(row) else None


def build_lots(workbook_path, as_of_date):
    workbook = openpyxl.load_workbook(workbook_path, data_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        raise ValueError("Excel 是空文件")

    column_map = build_column_map(rows[0])
    lots = []

    for row in rows[1:]:
        raw_lot = cell(row, column_map, "import_lot")
        lot = parse_lot(raw_lot)
        if lot is None:
            continue

        if "全局合计" in clean_text(raw_lot):
            continue

        europe_etd = normalize_date(cell(row, column_map, "europe_etd"))
        singapore_eta = normalize_date(cell(row, column_map, "singapore_eta"))
        singapore_etd = normalize_date(cell(row, column_map, "singapore_etd"))
        shanghai_eta = normalize_date(cell(row, column_map, "shanghai_eta"))
        remark = clean_text(cell(row, column_map, "remark"))
        status = infer_status(remark, europe_etd, singapore_eta, singapore_etd, shanghai_eta, as_of_date)

        import_containers = to_number(cell(row, column_map, "import_containers"))
        export_containers = to_number(cell(row, column_map, "export_containers"))

        lot.update(
            {
                "importAmount": round(to_float(cell(row, column_map, "import_amount")), 2),
                "exportLot": clean_text(cell(row, column_map, "export_lot")),
                "importContainers": import_containers,
                "exportContainers": export_containers,
                "savedContainers": to_number(import_containers - export_containers),
                "loadPlan": clean_text(cell(row, column_map, "load_plan")),
                "pallets": to_number(cell(row, column_map, "pallets")),
                "boxes": to_number(cell(row, column_map, "boxes")),
                "importPcs": to_number(cell(row, column_map, "import_pcs")),
                "exportPcs": to_number(cell(row, column_map, "export_pcs")),
                "netWeightKg": to_number(cell(row, column_map, "net_weight_kg")),
                "grossWeightKg": to_number(cell(row, column_map, "gross_weight_kg")),
                "shanghaiUnitPrice": to_float(cell(row, column_map, "shanghai_unit_price")),
                "shanghaiSalesAmount": round(to_float(cell(row, column_map, "shanghai_sales_amount")), 2),
                "europeEtd": europe_etd,
                "singaporeEta": singapore_eta,
                "secondLegVessel": clean_text(cell(row, column_map, "second_leg_vessel")),
                "singaporeEtd": singapore_etd,
                "shanghaiEta": shanghai_eta,
                "status": status,
                "stage": infer_stage(status, singapore_etd, shanghai_eta),
                "pay": normalize_payment_status(cell(row, column_map, "payment_status"), lot["idNumber"]),
                "remark": remark,
            }
        )
        lots.append(lot)

    return sorted(lots, key=lambda item: item["idNumber"])


def summarize_statuses(lots):
    counts = {}
    for lot in lots:
        counts[lot["status"]] = counts.get(lot["status"], 0) + 1
    return " | ".join(f"{count} 票{status}" for status, count in counts.items()) or "暂无数据"


def build_summary_groups(lots):
    definitions = [
        ("1-5 批次进口金额", "1-5 批次", 1, 5, "success"),
        ("6-10 批次进口金额", "6-10 批次", 6, 10.999, "carbon"),
        ("11+ 批次进口金额", "11+ 批次", 11, float("inf"), "danger"),
    ]
    groups = []

    for label, short_label, minimum, maximum, card_class in definitions:
        group_lots = [lot for lot in lots if minimum <= lot["idNumber"] <= maximum]
        groups.append(
            {
                "label": label,
                "shortLabel": short_label,
                "cardClass": card_class,
                "lotCount": len(group_lots),
                "importAmount": round(sum(lot["importAmount"] for lot in group_lots), 2),
                "importContainers": sum(lot["importContainers"] for lot in group_lots),
                "exportContainers": sum(lot["exportContainers"] for lot in group_lots),
                "savedContainers": sum(lot["savedContainers"] for lot in group_lots),
                "description": summarize_statuses(group_lots),
            }
        )

    return groups


def build_payload(workbook_path):
    as_of = date.today()
    lots = build_lots(workbook_path, as_of)
    if not lots:
        raise ValueError("没有找到 Lot 明细行")

    return {
        "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "sourceFile": Path(workbook_path).name,
        "lotRange": f"Lot {lots[0]['id']}-{lots[-1]['id']}",
        "totals": {
            "lotCount": len(lots),
            "importAmount": round(sum(lot["importAmount"] for lot in lots), 2),
            "importContainers": sum(lot["importContainers"] for lot in lots),
            "exportContainers": sum(lot["exportContainers"] for lot in lots),
            "savedContainers": sum(lot["savedContainers"] for lot in lots),
        },
        "summaryGroups": build_summary_groups(lots),
        "lots": lots,
    }


def main():
    parser = argparse.ArgumentParser(description="Build dashboard JSON from Excel lot details.")
    parser.add_argument("--input", default="data/lot-details.xlsx", help="Source Excel file")
    parser.add_argument("--output", default="data/lots.json", help="Output JSON file")
    args = parser.parse_args()

    payload = build_payload(args.input)
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {output_path} with {len(payload['lots'])} lot rows")


if __name__ == "__main__":
    main()

