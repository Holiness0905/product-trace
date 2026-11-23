#!/usr/bin/env python3
"""把表格或文本形式的产品数据转换为追溯页数据 (data/products.json) 与二维码 (qrcode/*.png)。

支持的输入:
- .xls / .xlsx: 第一个工作表。
- .txt / .tsv / .csv: 制表符或逗号分隔的表格, 版式与 Excel 相同
  (字段名列为空的一行, 其产品列为产品名; 字段名列中为字段, 其余列为各产品的值)。
- .txt 纯文本: 每个产品之间用空行分隔, 每行形如 "字段名：值"。
  不含 "字段名：" 的行, 以及字段名不在配置中的行, 会作为上一字段的补充行。

字段规则见 config.json。

用法:
    python tools/build.py [数据文件] [--base-url URL] [--stdout]

不指定数据文件时, 使用 input/ 中最新修改的数据文件。
--stdout 只把生成的 JSON 打印到终端, 不写入文件。
"""
import argparse
import csv
import datetime
import io
import json
import os
import re
from urllib.parse import quote

import qrcode
import xlrd
from xlrd import xldate

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_FILE = os.path.join(ROOT, "config.json")
INPUT_DIR = os.path.join(ROOT, "input")
DATA_FILE = os.path.join(ROOT, "data", "products.json")
QR_DIR = os.path.join(ROOT, "qrcode")
SUPPORTED = (".xls", ".xlsx", ".txt", ".tsv", ".csv")


def num_text(value):
    if float(value).is_integer():
        return str(int(value))
    return str(value)


def read_text(path):
    with open(path, "rb") as f:
        raw = f.read()
    for enc in ("utf-8-sig", "gbk"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    raise SystemExit(f"无法识别文件编码: {path}")


def grid_from_xls(path):
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)
    rows = []
    for r in range(sheet.nrows):
        row = []
        for c in range(sheet.ncols):
            cell = sheet.cell(r, c)
            if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                text = ""
            elif cell.ctype == xlrd.XL_CELL_DATE:
                text = xldate.xldate_as_datetime(cell.value, book.datemode).strftime("%Y-%m-%d")
            elif cell.ctype == xlrd.XL_CELL_NUMBER:
                text = num_text(cell.value)
            else:
                text = str(cell.value).strip()
            row.append(text)
        rows.append(row)
    return rows


def grid_from_xlsx(path):
    import openpyxl

    sheet = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    rows = []
    for values in sheet.iter_rows(values_only=True):
        row = []
        for v in values:
            if v is None:
                text = ""
            elif isinstance(v, datetime.datetime):
                text = v.strftime("%Y-%m-%d")
            elif isinstance(v, (int, float)):
                text = num_text(v)
            else:
                text = str(v).strip()
            row.append(text)
        rows.append(row)
    return rows


def products_from_grid(rows, id_label):
    """返回每个产品的 (字段名, 值) 列表; 找不到编号字段时返回 None。"""
    id_row = id_col = None
    for r, row in enumerate(rows):
        for c, v in enumerate(row):
            if v.strip() == id_label:
                id_row, id_col = r, c
                break
        if id_row is not None:
            break
    if id_row is None:
        return None

    header_row = None
    for r in range(id_row):
        row = rows[r]
        if id_col < len(row) and row[id_col].strip():
            continue
        if any(v.strip() for v in row[id_col + 1:]):
            header_row = r
            break
    if header_row is None:
        raise SystemExit("找不到产品名称行 (字段名列为空、产品列有名称的一行)")

    header = rows[header_row]
    product_cols = [c for c in range(id_col + 1, len(header)) if header[c].strip()]
    products = []
    for c in product_cols:
        pairs = []
        for r in range(header_row + 1, len(rows)):
            row = rows[r]
            label = row[id_col].strip() if id_col < len(row) else ""
            value = row[c].strip() if c < len(row) else ""
            if label or value:
                pairs.append((label, value))
        products.append(pairs)
    return products


def products_from_blocks(text):
    products = []
    for block in re.split(r"\n\s*\n", text.strip()):
        pairs = []
        for line in block.splitlines():
            line = line.strip()
            if not line:
                continue
            m = re.match(r"^([^：:]+)[：:](.*)$", line)
            if m:
                pairs.append((m.group(1).strip(), m.group(2).strip()))
            else:
                pairs.append(("", line))
        if pairs:
            products.append(pairs)
    return products


def load_products(path, id_label):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        products = products_from_grid(grid_from_xls(path), id_label)
    elif ext == ".xlsx":
        products = products_from_grid(grid_from_xlsx(path), id_label)
    else:
        text = read_text(path)
        if ext == ".csv":
            rows = list(csv.reader(io.StringIO(text)))
        else:
            rows = [line.split("\t") for line in text.splitlines()]
        products = products_from_grid(rows, id_label)
        if products is None:
            products = products_from_blocks(text)
    if not products:
        raise SystemExit(f"{path} 中没有可用的产品数据")
    return products


def known_labels(cfg):
    labels = {cfg["id_label"], cfg["site"]["warranty_label"]} | set(cfg["card"])
    for sec in cfg["sections"]:
        for field in sec["fields"]:
            labels.add(field if isinstance(field, str) else field["label"])
    return labels


def group_values(pairs, known):
    """按字段名归类。字段名为空的行, 或字段名未配置但有值的行, 归入上一个字段。"""
    groups = {}
    current = None
    for label, value in pairs:
        if label in known:
            current = label
            groups.setdefault(label, [])
            if value:
                groups[label].append(value)
        elif label:
            if value and current:
                groups[current].append(f"{label}：{value}")
        elif value and current:
            groups[current].append(value)
    return groups


def normalize_date(raw):
    m = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", raw)
    if not m:
        return raw
    year, month, day = m.groups()
    return f"{int(year)}年{int(month)}月{int(day)}日"


def apply_rewrites(text, rules):
    for pattern, repl in rules:
        text = re.sub(pattern, repl, text)
    return text


def fallback(label, cfg):
    if label not in cfg["fallbacks"]:
        raise SystemExit(f"字段 '{label}' 为沿用值, 但 config.json 的 fallbacks 中没有默认值")
    return cfg["fallbacks"][label]


def line_text(label, raw, cfg):
    if raw in cfg["carry_markers"]:
        text = fallback(label, cfg)
    elif "改为" in raw:
        head, new = (part.strip() for part in raw.split("改为", 1))
        fmt = cfg["change_formats"].get(head, cfg["change_format"])
        text = fmt.format(head=head, new=new)
    else:
        text = raw
    return apply_rewrites(text, cfg["rewrites"])


def test_text(label, raw, unit, cfg):
    if raw in cfg["carry_markers"]:
        return fallback(label, cfg)
    if unit and re.fullmatch(r"-?\d+(\.\d+)?", raw):
        return raw + unit
    return raw


def build_product(pairs, cfg, index):
    groups = group_values(pairs, known_labels(cfg))

    def values_of(label):
        vals = groups.get(label)
        if not vals:
            raise SystemExit(f"第 {index} 个产品缺少字段 '{label}'")
        return vals

    product_id = values_of(cfg["id_label"])[0]

    card = []
    for label in cfg["card"]:
        value = values_of(label)[0]
        if label in cfg["date_fields"]:
            value = normalize_date(value)
        card.append({"label": label, "value": value})

    sections = []
    for sec in cfg["sections"]:
        items = []
        if sec["kind"] == "lines":
            for label in sec["fields"]:
                lines = [line_text(label, v, cfg) for v in values_of(label)]
                items.append({"name": label, "lines": lines})
        elif sec["kind"] == "tests":
            for field in sec["fields"]:
                raw = values_of(field["label"])[0]
                items.append({
                    "name": field.get("name", field["label"]),
                    "value": test_text(field["label"], raw, field.get("unit", ""), cfg),
                    "result": cfg["result_default"],
                })
        else:
            raise SystemExit(f"未知的 section kind: {sec['kind']}")
        sections.append({"title": sec["title"], "icon": sec.get("icon", ""),
                         "kind": sec["kind"], "items": items})

    return {
        "id": product_id,
        "card": card,
        "sections": sections,
        "warranty": values_of(cfg["site"]["warranty_label"])[0],
    }


def latest_input():
    if not os.path.isdir(INPUT_DIR):
        raise SystemExit("未指定数据文件, 且 input/ 目录不存在")
    files = [os.path.join(INPUT_DIR, f) for f in os.listdir(INPUT_DIR)
             if f.lower().endswith(SUPPORTED)]
    if not files:
        raise SystemExit("input/ 中没有 .xls/.xlsx/.txt/.csv 数据文件")
    return max(files, key=os.path.getmtime)


def qr_filename(product_id):
    return re.sub(r'[\\/:*?"<>|]', "_", product_id) + ".png"


def main():
    parser = argparse.ArgumentParser(description="生成产品追溯数据与二维码")
    parser.add_argument("source", nargs="?", help="数据文件, 默认为 input/ 中最新的文件")
    parser.add_argument("--base-url", help="追溯页基础地址, 默认读取 config.json 的 base_url")
    parser.add_argument("--stdout", action="store_true", help="仅打印 JSON, 不写入文件")
    args = parser.parse_args()

    with open(CONFIG_FILE, encoding="utf-8") as f:
        cfg = json.load(f)

    source = args.source or latest_input()
    base_url = args.base_url or cfg["base_url"]
    if not base_url.endswith("/"):
        base_url += "/"

    raw_products = load_products(source, cfg["id_label"])
    products = [build_product(p, cfg, i) for i, p in enumerate(raw_products, start=1)]

    ids = [p["id"] for p in products]
    duplicates = sorted({i for i in ids if ids.count(i) > 1})
    if duplicates:
        raise SystemExit(f"出厂编号重复: {', '.join(duplicates)}")

    site = dict(cfg["site"])
    site["generated_at"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    data = {"site": site, "products": products}
    text = json.dumps(data, ensure_ascii=False, indent=2)

    if args.stdout:
        print(text)
        return

    os.makedirs(os.path.dirname(DATA_FILE), exist_ok=True)
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        f.write(text)

    os.makedirs(QR_DIR, exist_ok=True)
    keep = set()
    for p in products:
        name = qr_filename(p["id"])
        keep.add(name)
        url = f"{base_url}trace.html?id={quote(p['id'])}"
        img = qrcode.make(
            url,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            box_size=10,
            border=4,
        )
        img.save(os.path.join(QR_DIR, name))
        print(f"{p['id']}  {url}")

    for name in os.listdir(QR_DIR):
        if name.endswith(".png") and name not in keep:
            os.remove(os.path.join(QR_DIR, name))

    print(f"已生成 {len(products)} 个产品, 数据来源: {os.path.basename(source)}")


if __name__ == "__main__":
    main()
