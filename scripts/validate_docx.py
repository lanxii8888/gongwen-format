# -*- coding: utf-8 -*-
"""校验生成的 docx：包完整性、XML 元素顺序、页码域、文字零丢失。

用法：
    python validate_docx.py 输出.docx [原文件.docx]

第二参数给出时会比对正文文字是否完全一致。发现问题时返回码为 1。
"""
import sys
import zipfile
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
TWIP = 1440.0  # 每英寸 1440 twips

PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
             'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens',
             'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE',
             'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
             'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
             'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
             'sectPr', 'pPrChange']

RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike',
             'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid',
             'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz',
             'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign',
             'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath']

errors, warns = [], []


def check_order(el, order, label):
    prev, seen = -1, set()
    for child in el:
        tag = etree.QName(child).localname
        if tag not in order:
            errors.append('%s: 未知子元素 w:%s' % (label, tag))
            continue
        idx = order.index(tag)
        if idx < prev:
            errors.append('%s: 子元素顺序错误 w:%s' % (label, tag))
        prev = idx
        if tag in seen:
            errors.append('%s: 重复子元素 w:%s' % (label, tag))
        seen.add(tag)


def body_text(xml_bytes):
    root = etree.fromstring(xml_bytes)
    return ''.join(t.text or '' for t in root.find(W + 'body').iter(W + 't'))


def main(path, source=None):
    z = zipfile.ZipFile(path)
    if z.testzip():
        errors.append('zip 结构损坏')

    for n in z.namelist():
        if n.endswith(('.xml', '.rels')):
            try:
                etree.fromstring(z.read(n))
            except Exception as e:
                errors.append('%s: XML 解析失败 %s' % (n, e))

    doc = etree.fromstring(z.read('word/document.xml'))
    body = doc.find(W + 'body')

    for ppr in doc.iter(W + 'pPr'):
        check_order(ppr, PPR_ORDER, 'w:pPr')
    for rpr in doc.iter(W + 'rPr'):
        check_order(rpr, RPR_ORDER, 'w:rPr')

    if etree.QName(body[-1]).localname != 'sectPr':
        errors.append('body 最后一个元素不是 sectPr')

    sect = body.find(W + 'sectPr')
    pg_sz = sect.find(W + 'pgSz')
    pg_mar = sect.find(W + 'pgMar')
    size = (int(pg_sz.get(W + 'w')) / TWIP * 2.54, int(pg_sz.get(W + 'h')) / TWIP * 2.54)
    mar = {k.split('}')[-1]: int(v) / TWIP * 2.54 for k, v in pg_mar.items()}
    print('纸张：%.2f×%.2f cm' % size)
    print('页边距：上 %.2f 下 %.2f 左 %.2f 右 %.2f cm' % (
        mar['top'], mar['bottom'], mar['left'], mar['right']))

    refs = {}
    for r in sect.findall(W + 'footerReference'):
        rid = r.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        refs[r.get(W + 'type')] = rid
    print('页脚引用：%s' % refs)

    footers = [n for n in z.namelist() if n.startswith('word/footer') and n.endswith('.xml')]
    if refs and len(footers) < 2:
        errors.append('页脚不足两个（单双页页码需要 default + even）')
    for n in footers:
        raw = z.read(n).decode('utf-8')
        ok = ('PAGE' in raw) and ('fldChar' in raw) and raw.count('—') >= 2
        print('%s：PAGE域=%s 一字线×%d' % (n, 'PAGE' in raw, raw.count('—')))
        if not ok:
            warns.append('%s 页码域内容可能不完整' % n)
        for r in etree.fromstring(z.read(n)).iter(W + 'rPr'):
            check_order(r, RPR_ORDER, n + '/w:rPr')

    text = body_text(z.read('word/document.xml'))
    n_p = len([c for c in body.iterchildren() if etree.QName(c).localname == 'p'])
    n_t = len([c for c in body.iterchildren() if etree.QName(c).localname == 'tbl'])
    n_img = len(list(body.iter(W + 'drawing')))
    print('正文 %d 字，顶层段落 %d，表格 %d，图片 %d' % (len(text), n_p, n_t, n_img))

    if source:
        src = body_text(zipfile.ZipFile(source).read('word/document.xml'))
        same = (src == text)
        print('文字与原文一致：%s（原 %d 字 / 新 %d 字）' % ('是' if same else '否！', len(src), len(text)))
        if not same:
            errors.append('正文文字与源文件不一致')

    print('\n=== 校验结果 ===')
    for e in errors:
        print('  [错误] %s' % e)
    for w in warns:
        print('  [提示] %s' % w)
    if not errors:
        print('  通过')
        return 0
    return 1


if __name__ == '__main__':
    sys.exit(main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None))


