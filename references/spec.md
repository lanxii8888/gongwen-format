# 规格与 OOXML 对照

改规格前先读这份，知道每个数字落在哪个 XML 元素上。

## 字号对照（Word 中文习惯叫法）

| 叫法 | 磅值 | `w:sz` 值（半磅） |
|---|---|---|
| 初号 | 42 | 84 |
| 小初 | 36 | 72 |
| 一号 | 26 | 52 |
| 小一 | 24 | 48 |
| 二号 | 22 | 44 |
| 小二 | 18 | 36 |
| 三号 | 16 | 32 |
| 小三 | 15 | 30 |
| 四号 | 14 | 28 |
| 小四 | 12 | 24 |
| 五号 | 10.5 | 21 |

## 页面

```xml
<w:sectPr>
  <w:pgSz w:w="11906" w:h="16838"/>          <!-- A4，单位 twip：1cm = 567 twip -->
  <w:pgMar w:top="2098" w:right="1474"
           w:bottom="1984" w:left="1587"
           w:header="850" w:footer="992" w:gutter="0"/>
</w:sectPr>
```

3.7cm = 2098，3.5cm = 1984，2.8cm = 1587，2.6cm = 1474。

## 正文段落

```xml
<w:pPr>
  <w:widowControl w:val="0"/>
  <w:snapToGrid w:val="0"/>                  <!-- 关键：否则固定行距被网格改写 -->
  <w:spacing w:before="0" w:after="0"
             w:beforeAutospacing="0" w:afterAutospacing="0"
             w:line="600" w:lineRule="exact"/>  <!-- 30 磅 = 600 twip -->
  <w:ind w:firstLineChars="200" w:firstLine="640"/>  <!-- 2 字符 / 2×16pt×20 -->
  <w:jc w:val="both"/>
</w:pPr>
<w:rPr>
  <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"
            w:eastAsia="仿宋_GB2312" w:cs="Times New Roman"/>
  <w:sz w:val="32"/><w:szCs w:val="32"/>
  <w:b w:val="0"/><w:bCs w:val="0"/>
</w:rPr>
```

要点：

- **`firstLineChars` 优先于 `firstLine`**。按字符缩进，换字号不会跑偏；同时写 `firstLine` 是为兼容不认 chars 的旧版本。
- **中西文不用拆 run**。Word 按字符选字体：汉字取 `eastAsia`，数字字母取 `ascii`。一个 run 里写 `2026年` 会自动混排，不需要切成两个 run。
- **必须删主题属性**（`asciiTheme`/`eastAsiaTheme`/`cstheme`/`hint`），否则主题字体盖过显式字体。
- **`w:b` 要显式写 `0`**，只删元素会继承样式里的加粗。

## 标题类

与正文同一套 pPr，区别：

- 主/副标题：`align="center"`，不写 `w:line`（单倍行距），不设首行缩进。
- 一至四级标题：与正文相同（两端对齐 + 首行缩进 2 字符 + 固定行距），只换字体/加粗。

## 落款

```xml
<w:pPr>
  <w:spacing ... w:line="600" w:lineRule="exact"/>
  <w:ind w:rightChars="200" w:right="560"/>   <!-- 署名右空 2 字 -->
  <w:jc w:val="right"/>
</w:pPr>
```

日期行右空 4 字（`rightChars="400"`）。

## 页码

两套页脚：默认（奇数页）右对齐、偶数页左对齐，靠 `w:evenAndOddHeaders` 生效。

```xml
<!-- settings.xml -->
<w:evenAndOddHeaders/>

<!-- sectPr -->
<w:footerReference w:type="default" r:id="rId7"/>
<w:footerReference w:type="even" r:id="rId6"/>

<!-- footer1.xml：— 1 — -->
<w:p>
  <w:pPr><w:ind w:rightChars="100" w:right="280"/><w:jc w:val="right"/></w:pPr>
  <w:r><w:rPr>...宋体 14pt...</w:rPr><w:t>—</w:t></w:r>
  <w:r><w:rPr>...</w:rPr><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>
  <w:r><w:rPr>...</w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
  <w:r><w:rPr>...</w:rPr><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r><w:rPr>...</w:rPr><w:t>1</w:t></w:r>
  <w:r><w:rPr>...</w:rPr><w:fldChar w:fldCharType="end"/></w:r>
  ...
</w:p>
```

`w:dirty="true"` 让 Word 打开时重算页码，避免显示占位数字 1。

注意：`w:fldChar` 和 `w:t` 必须在 `w:rPr` **之后**，顺序错了 Word 会报文件损坏。

## 表格

只改单元格内段落的字体字号、清掉首行缩进、行距改单倍，**不动** `tblGrid` / `tblBorders` / `tcW`。

```xml
<w:pPr><w:widowControl w:val="0"/><w:spacing .../><w:jc w:val="left"/></w:pPr>
```

单元格内不能套固定行距 30 磅——行高会跟着涨，表格被撑得很高。

## 与 GB/T 9704 的差异

本 skill 的默认值来自用户提供的模板，不是国标原文：

| 项 | GB/T 9704-2012 | 本 skill 默认 |
|---|---|---|
| 行距 | 28 磅左右 | **30 磅** |
| 页边距 | 上 37 下 35 左 28 右 26 mm | 一致 |
| 页码 | 4 号半角宋体，一字线，底端外侧 | 一致 |

要回到国标值，改 `--config` 里的 `body_line_pt` 为 28。
