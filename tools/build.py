"""从 二维码.xls 生成 data/products.json 与各产品二维码 PNG。

用法:
    python tools/build.py https://<用户名>.github.io/product-trace/
"""
import json
import os
import sys

import qrcode
import xlrd
from xlrd import xldate

REPO_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_XLS = os.path.join(os.path.dirname(REPO_DIR), "二维码.xls")
DATA_FILE = os.path.join(REPO_DIR, "data", "products.json")
QR_DIR = os.path.join(REPO_DIR, "qrcode")

HEADER_ROW = 2   # 第 3 行: 案例名
FIELD_START = 3  # 第 4 行起为字段
LABEL_COL = 3    # D 列: 字段名
PRODUCT_START_COL = 4  # E 列起为各产品

ID_LABEL = "出厂编号"
MODEL_LABEL = "产品型号"


def format_cell(cell, label, datemode):
    if cell.ctype == xlrd.XL_CELL_EMPTY:
        return ""
    if cell.ctype == xlrd.XL_CELL_NUMBER:
        value = cell.value
        if "日期" in label:
            return xldate.xldate_as_datetime(value, datemode).strftime("%Y-%m-%d")
        return str(int(value)) if float(value).is_integer() else str(value)
    return str(cell.value).strip()


def read_products(book):
    sheet = book.sheet_by_index(0)
    products = []
    for col in range(PRODUCT_START_COL, sheet.ncols):
        name = str(sheet.cell_value(HEADER_ROW, col)).strip()
        if not name:
            continue
        items = []
        for row in range(FIELD_START, sheet.nrows):
            label = str(sheet.cell_value(row, LABEL_COL)).strip()
            value = format_cell(sheet.cell(row, col), label, book.datemode)
            if not label and not value:
                continue
            items.append({"label": label, "value": value})

        fields = {item["label"]: item["value"] for item in items if item["label"]}
        product_id = fields.get(ID_LABEL, "")
        if not product_id:
            raise SystemExit(f"{name} 缺少 '{ID_LABEL}'，无法生成二维码")
        products.append({
            "id": product_id,
            "model": fields.get(MODEL_LABEL, ""),
            "items": items,
        })
    return products


def main():
    base_url = sys.argv[1] if len(sys.argv) > 1 else "https://YOUR_USER.github.io/product-trace/"
    if not base_url.endswith("/"):
        base_url += "/"

    book = xlrd.open_workbook(SRC_XLS)
    products = read_products(book)

    os.makedirs(os.path.dirname(DATA_FILE), exist_ok=True)
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump({"products": products}, f, ensure_ascii=False, indent=2)

    os.makedirs(QR_DIR, exist_ok=True)
    for p in products:
        url = f"{base_url}trace.html?id={p['id']}"
        img = qrcode.make(
            url,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            box_size=10,
            border=4,
        )
        img.save(os.path.join(QR_DIR, f"{p['id']}.png"))
        print(f"{p['id']}  {url}")


if __name__ == "__main__":
    main()
