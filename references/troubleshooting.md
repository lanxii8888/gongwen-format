# 故障排查

## 校验不通过

### `[错误] w:pPr: 子元素顺序错误`

OOXML 对 `pPr`/`rPr` 子元素顺序有硬性要求，顺序错了 Word 可能直接报"内容有问题"。

修法：不要 `append` 到末尾，用脚本里的 `_ordered_put()` / `set_para_format()`。自己加元素时按 `PPR_ORDER` / `RPR_ORDER` 里的顺序插入。

### `[错误] 正文文字与源文件不一致`

说明有代码路径在**重建**段落（新建 `w:p` 或给 `paragraph.text` 赋值）。正确做法是只改已有 `w:r` 的 `rPr`。检查：

- 有没有 `p.text = ...`
- 有没有 `doc.add_paragraph()` 处理正文
- 有没有把 `w:t` 重新拼接过

### `[错误] settings 缺少 evenAndOddHeaders`

单双页页码分居左右依赖这个开关。用 `doc.settings.odd_and_even_pages_header_footer = True`，别手工 `append` 到 settings 末尾（顺序同样有要求）。

### 页脚里没有 PAGE 域 / 只有一个页脚

- 必须 `is_linked_to_previous = False` 才会生成独立页脚部件
- 奇数页用 `section.footer`，偶数页用 `section.even_page_footer`

## 识别不对

### 标题认错

```bash
python gongwen_format.py 输入.docx --dry-run          # 先看是第几段
python gongwen_format.py 输入.docx 输出.docx --title-index 2
```

标题文字里带书名号、破折号时，`--title` 要做完全匹配（空格会被忽略）。

### 落款被当成正文

`sign` 规则 = "最后一个 >60 字段落之后的 ≤30 字段落"。如果文末还有附件清单、说明等长段，落款会被漏掉。两个办法：

1. 把附件清单写成 >60 字的长段（不现实）
2. 改 `classify()`：加关键词匹配，如 `^(承诺人|签字|盖章|年.*月.*日)`

### `（一）` 开头的正文被认成二级标题

规则限定了 ≤45 字且不以 `。` 结尾，但短句仍可能误判。把该段改长一点，或调整长度阈值。

### 编号列表缩进没变

带 `numPr` 的段落会跳过缩进设置（否则编号和文字会错位）。这是故意的。要统一缩进得改 `numbering.xml` 的 `w:ind`。

## 打开文件时报错

### Word 提示"内容有问题，是否修复"

按概率排查：

1. `pPr`/`rPr` 子元素顺序错（最常见）
2. `settings.xml` 里手工插入了元素，位置不对
3. 页脚里 `w:fldChar` 写在了 `w:rPr` 前面
4. zip 里有重复的部件名（重建包时 `write` 了同名成员）

先用 `validate_docx.py` 跑一遍，它会报出顺序问题。

### Word 打不开 / 一直转圈

多半是文件级问题（OneDrive 占位、受保护视图、宏），跟排版无关。把文件复制到本地磁盘再试。

## 其他

### `.doc` 老格式

python-docx 读不了。让用户用 Word"另存为 .docx"，或转档后再跑。

### 本机没有 `方正小标宋简体` / `仿宋_GB2312`

Word 会按字体替换表自动替换，显示效果与装了字体的大致一致。不要为此改配置——改了反而在用户机器上不对。

### 多节文档

页码只装在第一节，其余节 `linked_to_previous = True`，会自动跟随。若各节页边距不同，`setup_sections()` 会统一覆盖——如果用户要保留分节差异，注释掉那一步。

### Word COM 起不来

`New-Object -ComObject Word.Application` 报"Word 出现问题"是环境问题（注册表、受保护视图、激活状态），换机器也一样。不要为了看效果反复重试——会不停弹窗。改用 `office-render`，或直接交付文件让用户打开。
