# 产品追溯系统

产品追溯页与二维码生成工具。支持从 Excel 表格、文本文件或网页录入产品数据，自动生成追溯页数据和二维码，部署到 GitHub Pages 后扫码即可查看。

## 快速开始

### 方式一：网页录入（推荐）

直接用浏览器打开 `entry.html`，无需安装任何依赖：

1. 填写产品基本信息（型号、批次、编号、日期）
2. 填写原材料、生产过程、工序人员、出厂检验数据
3. 点击「加入产品列表」，页面自动生成二维码
4. 点击「导出 products.json」下载数据文件
5. 将下载的文件替换 `data/products.json`，提交到 GitHub

### 方式二：脚本批量生成

适合批量处理 Excel 或文本数据：

```bash
# 安装依赖
pip install "qrcode[pil]" xlrd openpyxl

# 生成数据和二维码
python tools/build.py input/二维码.xls
python tools/build.py input/数据.txt --base-url https://你的用户名.github.io/product-trace/
```

不指定文件时，脚本自动使用 `input/` 目录中最新修改的文件。

## 目录结构

```
product-trace/
├── config.json         配置文件（字段映射、页面文案、证书信息）
├── trace.html          追溯页模板（部署后通过 ?id=编号 访问）
├── entry.html          网页录入页面（本地使用）
├── data/
│   └── products.json   产品数据（由脚本或录入页生成）
├── qrcode/             二维码图片目录
├── input/              数据源文件（不提交到仓库）
└── tools/
    └── build.py        数据转换与二维码生成脚本
```

## 支持的输入格式

### Excel 格式（.xls / .xlsx）

第一行为表头，每行一个产品。表头字段名与 `config.json` 中的 `sections` 字段对应。

示例表头：
```
产品型号 | 生产批次 | 出厂编号 | 生产日期 | 压缩机 | 换热器铜管 | ...
```

### 分隔符文本（.txt / .tsv / .csv）

制表符或逗号分隔，第一行为表头，格式与 Excel 相同。

### 纯文本块格式（.txt）

每个产品之间用空行分隔，每行写作 `字段名：值`：

```
产品型号：E15
生产批次：2026040253
出厂编号：OLY-2026040253
生产日期：2026年4月2日
压缩机：美芝压缩机 · 批次20260325

产品型号：E25
生产批次：2026050078
出厂编号：OLY-2026050078
...
```

## 特殊值处理

在表格或文本中，以下写法会被特殊处理：

| 写法 | 效果 |
|------|------|
| `按原来的` 或 `不变` | 使用 config.json 中 fallbacks 的默认值 |
| `XX改为YY` | 显示为 `XX：YY` |
| `充注量改为3.5` | 显示为 `R410A · 充注量3.5kg`（特殊格式） |

## 配置说明

编辑 `config.json` 可自定义：

### 基本设置

```json
{
  "base_url": "https://用户名.github.io/product-trace/",
  "id_label": "出厂编号",
  "card": ["产品型号", "生产批次", "出厂编号", "生产日期"],
  "date_fields": ["生产日期"]
}
```

- `base_url`：GitHub Pages 地址，用于生成二维码链接
- `id_label`：作为产品唯一标识的字段名
- `card`：追溯页顶部卡片显示的字段
- `date_fields`：需要格式化为日期的字段

### 默认值

```json
{
  "fallbacks": {
    "换热器铜管": "金龙铜业 · 入厂复检合格",
    "制冷剂": "R410A环保制冷剂"
  }
}
```

当字段值为「按原来的」或「不变」时，使用对应的默认值。

### 页面分区

```json
{
  "sections": [
    {
      "title": "原材料追溯",
      "icon": "📦",
      "kind": "lines",
      "fields": ["压缩机", "换热器铜管", "风机", "制冷剂"]
    },
    {
      "title": "出厂检验报告",
      "icon": "✅",
      "kind": "tests",
      "fields": [
        {"label": "制冷量", "name": "制冷量", "unit": "kW"}
      ]
    }
  ]
}
```

- `kind: "lines"`：每个字段显示为一行文本
- `kind: "tests"`：显示为检测项表格（项目 / 结果 / 结论）

### 页面文案

```json
{
  "site": {
    "title": "奥利维尔空调设备",
    "subtitle": "✓ 全生命周期质量追溯",
    "logo": "❄️",
    "company": "江苏奥利维尔环境设备有限公司",
    "certs": [
      {"cert": "质量管理", "code": "ISO 9001:2015"}
    ]
  }
}
```

## 发布到 GitHub Pages

### 首次设置

1. 在 GitHub 创建仓库（如 `product-trace`）
2. 进入仓库 Settings → Pages
3. Source 选择 `Deploy from a branch`
4. Branch 选择 `main`，目录选择 `/ (root)`
5. 点击 Save

### 更新数据

```bash
cd product-trace

# 生成数据（如果使用脚本方式）
python tools/build.py input/数据.xls

# 提交并推送
git add -A
git commit -m "更新产品数据"
git push
```

推送后等待 1-2 分钟，GitHub Pages 自动部署。

### 访问追溯页

```
https://用户名.github.io/product-trace/trace.html?id=OLY-2026040253
```

扫描生成的二维码即可跳转到对应产品的追溯页。

## 本地预览

```bash
cd product-trace
python -m http.server 8080
```

然后访问 `http://localhost:8080/trace.html?id=OLY-2026040253`

## 常见问题

### 二维码不显示

`entry.html` 的二维码功能依赖 CDN 加载 qrcodejs 库。如果网络不通，录入和导出功能仍可正常使用，只是不显示预览二维码。

### 日期显示为数字

确保 `config.json` 中的 `date_fields` 包含日期字段名。脚本会自动将 Excel 日期序列号转换为 `YYYY年M月D日` 格式。

### 追溯页显示「未找到产品」

检查 URL 参数 `?id=` 的值是否与 `data/products.json` 中的 `id` 字段匹配。
