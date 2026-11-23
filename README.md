# product-trace

产品追溯页与二维码生成。把表格或文本数据放入 `input/`，运行脚本后自动生成追溯页数据和二维码。

## 目录

```
config.json        字段映射、页面文案、证书与保修信息
input/             放置待处理的数据文件 (不提交到仓库)
data/products.json 由脚本生成的追溯数据
qrcode/            由脚本生成的二维码 PNG
trace.html         追溯页 (读取 data/products.json)
tools/build.py     数据转换与二维码生成脚本
```

## 使用

依赖: `pip install "qrcode[pil]" xlrd openpyxl`

```
python tools/build.py input/二维码.xls
python tools/build.py input/样例.txt --base-url https://<用户名>.github.io/product-trace/
```

不指定数据文件时，脚本使用 `input/` 中最新修改的数据文件。

支持的输入:
- `.xls` / `.xlsx`: 第一个工作表，版式与 `二维码.xls` 相同。
- `.txt` / `.tsv` / `.csv`: 制表符或逗号分隔，版式与 Excel 相同。
- `.txt` 纯文本: 每个产品之间用空行分隔，每行写作 `字段名：值`。

## 网页录入

`entry.html` 是纯前端录入页，可直接用浏览器打开，无需运行脚本：

- 填写产品基本信息、原材料、生产过程、工序人员、出厂检验项。
- 点击「加入产品列表」后，页面生成该产品的追溯地址和二维码。
- 「导出 products.json」下载数据文件，替换 `data/products.json` 后提交即可。
- 「导入 products.json」可载入已有数据继续编辑。

二维码库从 CDN 加载，离线时仍可录入和导出，但不显示二维码。

## 修改规则

在表格中，同一字段填写以下内容时，脚本按规则处理:
- `按原来的` / `不变`: 使用 `config.json` 中 `fallbacks` 的默认值。
- `XX改为YY`: 显示为 `XX：YY`。充注量按 `change_formats` 显示为 `R410A · 充注量YYkg`。

## 修改配置

- 新增字段: 在 `config.json` 的 `sections` 中添加字段名。
- 修改页面文案、证书、保修: 编辑 `site` 段。
- 修改默认值: 编辑 `fallbacks`。

## 发布

1. 运行 `tools/build.py`，更新 `data/products.json` 和 `qrcode/`。
2. `git add -A`、`git commit`、`git push`。
3. GitHub Pages 从 `main` 分支根目录发布。
