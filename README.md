<div align="center">

# 公文格式排版 · gongwen-format

把任意 Word 文档套成党政机关公文格式 —— **只改格式，不改一个字**

[![Python](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)]()
[![Skill](https://img.shields.io/badge/agent%20skill-compatible-purple.svg)](SKILL.md)

</div>

---

## 这是什么

一份写好的 `.docx` 丢进去，出来一份符合公文排版的 `.docx`：

- 页面 A4，上 3.7 / 下 3.5 / 左 2.8 / 右 2.6 cm
- 主标题方正小标宋简体二号居中，一级标题黑体三号，二级标题楷体三号
- 正文仿宋_GB2312 三号，数字字母 Times New Roman，两端对齐，首行缩进 2 字符，行距固定值 30 磅
- 页脚 `— 1 —` 页码，单页居右空一字、双页居左空一字

**正文文字、表格、图片、自动编号、脚注、域，一个都不动。** 校验脚本会逐字比对，确认改动前后正文完全一致。

## 效果

| 元素 | 改之前（常见乱格式） | 改之后 |
|---|---|---|
| 页面 | Word 默认 2.54 / 3.17 cm | 3.7 / 3.5 / 2.8 / 2.6 cm |
| 标题 | 黑体加粗、左对齐 | 方正小标宋简体、二号、不加粗、居中 |
| 一级标题 `一、` | 宋体、五号 | 黑体、三号、缩进 2 字符 |
| 正文 | 微软雅黑 12pt、单倍行距、无缩进 | 仿宋_GB2312 16pt、固定 30 磅、缩进 2 字符、两端对齐 |
| 数字 | 中文字体 | Times New Roman |
| 页码 | 无 / 手打的数字 | `— 1 —` 真页码域，奇偶分居左右 |

## 为什么不"新建文档再粘贴"

网上多数方案是读原文 → 新建 Document → 按样式重写。这条路会丢东西：表格样式、图片锚定、自动编号、脚注、交叉引用、页码域，重建一遍基本全没。

本项目走另一条路：**在源文件的 OOXML 上做外科手术**——只改 `w:pPr`（段落属性）和 `w:rPr`（字符属性），`w:tbl` / `w:drawing` / `w:numPr` / `w:fldSimple` 原样保留。

## 安装

### 方式一：作为 agent skill 安装（推荐）

```bash
npx skills add lanxii8888/gongwen-format
```

装好后对 AI 说「把这份文档按公文格式改一下」即可。

### 方式二：手动放进 skill 目录

```bash
git clone https://github.com/lanxii8888/gongwen-format.git \
  ~/.dsh/skills/gongwen-format
```

DeepSeek Harness / Claude Code 等兼容 `SKILL.md` 的运行时都会自动发现它。

### 方式三：只当命令行工具用

```bash
pip install python-docx lxml
python scripts/gongwen_format.py 输入.docx 输出.docx
```

## 快速开始

```bash
# 1. 先演练：只看每段被判成什么，不生成文件
python scripts/gongwen_format.py 工作报告.docx --dry-run

# 2. 正式排版（不覆盖源文件）
python scripts/gongwen_format.py 工作报告.docx 工作报告_公文格式.docx

# 3. 校验：结构合法 + 文字零丢失
python scripts/validate_docx.py 工作报告_公文格式.docx 工作报告.docx
```

输出两个文件：

- `工作报告_公文格式.docx` —— 成品
- `工作报告_公文格式_排版报告.txt` —— 逐段列出「第几段被判成什么」，方便核对

### 演示

```text
$ python scripts/gongwen_format.py 周报.docx --dry-run
#0   主标题   工作周报
#1   正文     正文内容……
#2   空段
#6   一级标题  一、标题设置
#7   正文     主标题使用方正小标宋简体字，不加粗，二号字，居中排布。如有副标题……
#8   一级标题  二、正文设置
```

## 参数

| 参数 | 说明 |
|---|---|
| `--dry-run` | 只输出段落识别结果，不写文件 |
| `--title "标题原文"` | 手动指定主标题（按文字完全匹配，忽略空格） |
| `--title-index N` | 手动指定第 N 段（从 0 起）为主标题 |
| `--title-bold` | 主标题加粗（默认不加粗） |
| `--no-page-number` | 不生成页码 |
| `--keep-header` | 保留原页眉（默认清空） |
| `--table-size 14` | 表格内文字字号，默认三号；宽表建议降到 14 或 12 |
| `--config spec.json` | 加载自定义规格 |
| `--dump-config spec.json` | 导出当前规格为 JSON，改完再 `--config` 加载 |

## 自定义规格

不同单位、不同模板的细则常不一样（比如有的要求 28 磅行距、二级标题楷体加粗）。把规格抽出来，不用改代码：

```bash
python scripts/gongwen_format.py 输入.docx --dump-config my_spec.json
```

导出后按需修改：

```json
{
  "body_line_pt": 28,
  "indent_chars": 200,
  "style": {
    "h2": { "east_asia": "楷体_GB2312", "size_pt": 16, "bold": true }
  },
  "page": { "top_cm": 3.7, "bottom_cm": 3.5, "left_cm": 2.8, "right_cm": 2.6 }
}
```

只写要改的字段即可，其余沿用默认值（深合并）。

```bash
python scripts/gongwen_format.py 输入.docx 输出.docx --config my_spec.json
```

## 段落识别规则

按顺序判定，命中即停：

| 类型 | 判定条件 |
|---|---|
| `title` 主标题 | 首个非空、非纯图片段落；≤40 字；不以 `。；，` 结尾 |
| `subtitle` 副标题 | 紧跟主标题，且以 `——` 开头 |
| `h1` 一级标题 | 有大纲级别 0；或 ≤45 字、不以 `。` 结尾、匹配 `^一、` |
| `h2` 二级标题 | 同上，匹配 `^（一）` |
| `h3` 三级标题 | 同上，匹配 `^1.` |
| `h4` 四级标题 | 同上，匹配 `^（1）` |
| `sign` 落款 | 最后一个长段落（>60 字）之后的短段落（≤30 字） |
| `image` 图片段 | 段内只有图片、无文字 |
| `empty` / `body` | 其余 |

判错时优先用 `--title` / `--title-index` 覆盖；规则本身在 `scripts/gongwen_format.py` 的 `classify()` 里，改起来是几行正则。

## 校验

```bash
python scripts/validate_docx.py 输出.docx 输入.docx
```

检查四项，任一失败返回码为 1：

1. 包内所有 XML 可解析、zip 结构完好
2. `w:pPr` / `w:rPr` 子元素顺序合法（顺序错 Word 会报「内容有问题」）
3. 页脚 PAGE 域存在、`evenAndOddHeaders` 已开启、单双页页脚齐备
4. **正文文字与源文件逐字一致**

## 实现要点

这四件事不做，格式看着对、实际会跑偏：

- **清主题字体** 源文档常带 `asciiTheme` / `eastAsiaTheme`，不清掉会盖过显式字体，改了也没用。
- **清自动段间距** `beforeAutospacing` / `afterAutospacing` 不清，段距不受 `w:spacing` 控制。
- **关文档网格** 段落必须写 `snapToGrid=0`，否则「固定值 30 磅」会被网格吸成别的值。
- **用字符缩进** 写 `firstLineChars="200"`（同时写 `firstLine` 兼容旧版本），换字号不跑偏。

页码用 `PAGE` 域 + `w:dirty="true"` 强制重算，不是写死的数字。

中文思路细节见 [references/spec.md](references/spec.md)。

## 已知限制

- **不生成红头、版记、印章位置** —— 这些是机构特定的，需要人工排
- **`.doc` 老格式不支持**，请先用 Word 另存为 `.docx`
- **文本框内的文字不处理**（`w:txbxContent`），会被跳过
- **表格只改字体**，不动列宽、框线、单元格合并；三号字在窄表里会换行变高
- **多节文档统一页边距**，如需保留分节差异要改 `setup_sections()`
- 字体缺失时（如本机没装「方正小标宋简体」）由 Word 自动替换，属正常现象

## 目录结构

```
gongwen-format/
├── SKILL.md                     agent 运行时读取的技能说明
├── README.md
├── LICENSE
├── requirements.txt
├── scripts/
│   ├── gongwen_format.py        排版引擎
│   ├── validate_docx.py         校验工具
│   └── spec_default.json        默认规格
├── references/
│   ├── spec.md                  规格与 OOXML 对照
│   └── troubleshooting.md       故障排查
└── tests/
    └── test_smoke.py            冒烟测试
```

## 开发

```bash
pip install -r requirements.txt
python -m unittest discover -s tests -v
```

测试会在临时目录里造一份带标题、四级标题、落款、表格的文档，跑完排版后断言页边距、字体、字号、行距、缩进、页码域是否正确，并校验文字未变。不需要安装 Word 或 LibreOffice。

## 贡献

欢迎开 issue 提你的单位模板细则，或者直接 PR 改 `classify()` 的识别规则。改规则时请顺手补一条测试用例。

## License

[MIT](LICENSE)

---

## English

**gongwen-format** reformats any `.docx` into the Chinese government-document (公文) layout — fonts, sizes, 30pt fixed line spacing, 2-character first-line indent, GB/T 9704 page margins, and `— 1 —` page numbers.

It edits the source document's OOXML in place instead of rebuilding it, so tables, images, numbering, footnotes and fields survive untouched. A bundled validator proves the body text is byte-for-byte identical before and after.

```bash
python scripts/gongwen_format.py input.docx output.docx
python scripts/validate_docx.py output.docx input.docx
```

Works as a standalone CLI or as an agent skill (`SKILL.md`).
