"""从 二维码.xls 生成 data/products.json 与各产品二维码 PNG。

Excel 中的修改指令按以下规则处理:
- "按原来的"、"不变": 使用参考页面中的默认值 (DEFAULTS / TEST_DEFAULTS)
- "XX改为YY": 生成 "XX：YY" 形式的文本，充注量单独处理

用法:
    python tools/build.py https://<用户名>.github.io/product-trace/
"""
import datetime
import json
import os
import re
import sys

import qrcode
import xlrd
from xlrd import xldate

REPO_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_XLS = os.path.join(os.path.dirname(REPO_DIR), "二维码.xls")
DATA_FILE = os.path.join(REPO_DIR, "data", "products.json")
QR_DIR = os.path.join(REPO_DIR, "qrcode")

HEADER_ROW = 2         # 第 3 行: 案例名
FIELD_START = 3        # 第 4 行起为字段
LABEL_COL = 3          # D 列: 字段名
PRODUCT_START_COL = 4  # E 列起为各产品

ID_LABEL = "出厂编号"

COMPANY = "江苏奥利维尔环境设备有限公司"
SYSTEM_NAME = "全生命周期追溯管理系统 V2.0"
VALIDITY_YEARS = 15
CERTS = [
    {"cert": "质量管理", "code": "ISO 9001:2015"},
    {"cert": "环境管理", "code": "ISO 14001:2015"},
    {"cert": "职业健康", "code": "ISO 45001:2018"},
]

MATERIAL_LABELS = ["压缩机", "换热器铜管", "风机", "制冷剂"]
PROCESS_LABELS = ["生产线别", "制冷剂充注", "过程检验"]
TEST_LABELS = ["制冷量", "制热量", "能效比", "绝缘电阻", "泄露电流",
               "电气强度", "制冷剂泄露", "噪音等级"]
TEST_DISPLAY = {"泄露电流": "泄漏电流", "制冷剂泄露": "制冷剂泄漏", "噪音等级": "噪声等级"}
TEST_UNITS = {"制冷量": "kW", "制热量": "kW", "泄露电流": "mA", "噪音等级": "dB(A)"}

# 参考页面 (case1-E15.html) 中的原始值, Excel 填写"按原来的"时使用
DEFAULTS = {
    "换热器铜管": ["金龙铜业 · 入厂复检合格"],
    "制冷剂": ["R410A环保制冷剂"],
}
# Excel 填写"不变"时使用
TEST_DEFAULTS = {
    "绝缘电阻": ">500MΩ",
    "电气强度": "AC1800V/5s",
    "制冷剂泄露": "≤2g/年",
}


def format_cell(cell, label, datemode):
    if cell.ctype == xlrd.XL_CELL_EMPTY:
        return ""
    # 带日期格式的数字单元格在 xls 中为 XL_CELL_DATE, 需一并处理
    if cell.ctype in (xlrd.XL_CELL_NUMBER, xlrd.XL_CELL_DATE):
        value = cell.value
        if "日期" in label:
            return xldate.xldate_as_datetime(value, datemode).strftime("%Y-%m-%d")
        return str(int(value)) if float(value).is_integer() else str(value)
    return str(cell.value).strip()


def read_groups(sheet, col, datemode):
    """按字段名分组。D 列为空的行归入上一个字段名。"""
    groups = {}
    current = None
    for row in range(FIELD_START, sheet.nrows):
        label = str(sheet.cell_value(row, LABEL_COL)).strip()
        if label:
            current = label
        if current is None:
            continue
        value = format_cell(sheet.cell(row, col), current, datemode)
        if value:
            groups.setdefault(current, []).append(value)
    return groups


def cn_date(iso):
    year, month, day = iso.split("-")
    return f"{int(year)}年{int(month)}月{int(day)}日"


def material_lines(label, values):
    if values == ["按原来的"]:
        return DEFAULTS[label]
    # 在"批次"、"型式试验"前加分隔符: 谷轮批次X -> 谷轮 · 批次X
    return [re.sub(r"(?<=\S)(批次|型式试验)", r" · \1", v) for v in values]


def process_lines(values):
    lines = []
    for v in values:
        if "改为" in v:
            head, new = v.split("改为", 1)
            if head == "充注量":
                lines.append(f"R410A · 充注量{new}kg")
            else:
                lines.append(f"{head}：{new}")
        else:
            lines.append(v)
    return lines


def test_value(label, v):
    if v == "不变":
        return TEST_DEFAULTS[label]
    return v + TEST_UNITS.get(label, "")


def build_product(name, groups):
    def require(label):
        values = groups.get(label)
        if not values:
            raise SystemExit(f"{name} 缺少 '{label}'")
        return values

    product_id = require(ID_LABEL)[0]
    materials = []
    for label in MATERIAL_LABELS:
        materials.append({"name": label, "lines": material_lines(label, require(label))})

    process = []
    for label in PROCESS_LABELS:
        process.append({"name": label, "lines": process_lines(require(label))})

    tests = []
    for label in TEST_LABELS:
        tests.append({
            "name": TEST_DISPLAY.get(label, label),
            "value": test_value(label, require(label)[0]),
            "result": "合格",
        })

    return {
        "id": product_id,
        "model": require("产品型号")[0],
        "batch": require("生产批次")[0],
        "date": cn_date(require("生产日期")[0]),
        "materials": materials,
        "process": process,
        "tests": tests,
        "warranty": require("保修期限")[0],
    }


def main():
    base_url = sys.argv[1] if len(sys.argv) > 1 else "https://YOUR_USER.github.io/product-trace/"
    if not base_url.endswith("/"):
        base_url += "/"

    book = xlrd.open_workbook(SRC_XLS)
    sheet = book.sheet_by_index(0)
    products = []
    for col in range(PRODUCT_START_COL, sheet.ncols):
        name = str(sheet.cell_value(HEADER_ROW, col)).strip()
        if not name:
            continue
        groups = read_groups(sheet, col, book.datemode)
        products.append(build_product(name, groups))

    data = {
        "company": COMPANY,
        "system": SYSTEM_NAME,
        "validityYears": VALIDITY_YEARS,
        "generatedAt": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "certs": CERTS,
        "products": products,
    }
    os.makedirs(os.path.dirname(DATA_FILE), exist_ok=True)
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

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
