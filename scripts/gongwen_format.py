#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按公文格式规格重排 Word 文档（在原文件上改格式，不重建文档）。

规格默认值 = 党政机关公文常用排版；可用 --config 覆盖，或 --dump-config 导出后改。

用法：
    python gongwen_format.py 输入.docx [输出.docx] [--config spec.json] [选项]

选项：
    --title TEXT          直接指定主标题文字
    --title-index N       指定第 N 段（从 0 起）为主标题
    --title-bold          主标题加粗（默认不加粗）
    --no-page-number      不生成页码
    --keep-header         保留原页眉
    --table-size PT       表格内文字字号（默认取配置）
    --dump-config PATH    导出当前配置为 JSON 模板
    --dry-run             只输出识别结果，不写文件
"""
import argparse
import copy
import json
import os
import re
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm

XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

# ------------------------------------------------------------------ 默认规格

DEFAULTS = {
    "page": {
        "width_cm": 21.0, "height_cm": 29.7,
        "top_cm": 3.7, "bottom_cm": 3.5,
        "left_cm": 2.8, "right_cm": 2.6,
        "header_cm": 1.5, "footer_cm": 1.75
    },
    "style": {
        "title":    {"east_asia": "方正小标宋简体", "size_pt": 22, "bold": False},
        "subtitle": {"east_asia": "楷体_GB2312",    "size_pt": 16, "bold": False},
        "h1":       {"east_asia": "黑体",           "size_pt": 16, "bold": False},
        "h2":       {"east_asia": "楷体_GB2312",    "size_pt": 16, "bold": False},
        "h3":       {"east_asia": "仿宋_GB2312",    "size_pt": 16, "bold": True},
        "h4":       {"east_asia": "仿宋_GB2312",    "size_pt": 16, "bold": True},
        "body":     {"east_asia": "仿宋_GB2312",    "size_pt": 16, "bold": False},
        "sign":     {"east_asia": "仿宋_GB2312",    "size_pt": 16, "bold": False}
    },
    "latin_font": "Times New Roman",
    "body_line_pt": 30,
    "indent_chars": 200,
    "table": {"size_pt": 16},
    "pagenum": {"enabled": True, "east_asia": "宋体", "size_pt": 14,
                "space_chars": 100, "dash": "—"},
    "header": {"clear": True}
}

# ------------------------------------------------------------------ 识别规则

RE_H1 = re.compile(r'^\s*[一二三四五六七八九十百零〇]+\s*、')
RE_H2 = re.compile(r'^\s*[（(]\s*[一二三四五六七八九十百零〇]+\s*[）)]')
RE_H3 = re.compile(r'^\s*\d{1,2}\s*[.．、]\s*(?!\d)')
RE_H4 = re.compile(r'^\s*[（(]\s*\d{1,2}\s*[）)]')
RE_DATE = re.compile(r'^[\s〇零一二三四五六七八九十百千0-9]{0,6}年\s*[\s〇零一二三四五六七八九十百千0-9]{1,3}月'
                     r'\s*[\s〇零一二三四五六七八九十百千0-9]{1,3}日$')

KIND_CN = {
    'title': '主标题', 'subtitle': '副标题', 'h1': '一级标题', 'h2': '二级标题',
    'h3': '三级标题', 'h4': '四级标题', 'body': '正文', 'sign': '落款',
    'empty': '空段', 'image': '图片段',
}

RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike',
             'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid',
             'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz',
             'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign',
             'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath']

THEME_ATTRS = ['asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme', 'hint']
P_KEEP = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr']


# ------------------------------------------------------------------ XML 工具

def _ordered_put(rpr, tag, attrs):
    """在 rPr 内按 schema 顺序写入子元素（顺序错了 Word 会报文件损坏）。"""
    old = rpr.find(qn('w:' + tag))
    if old is not None:
        for k, v in attrs.items():
            old.set(qn('w:' + k), str(v))
        return old
    el = OxmlElement('w:' + tag)
    for k, v in attrs.items():
        el.set(qn('w:' + k), str(v))
    idx = RPR_ORDER.index(tag)
    pos = len(rpr)
    for i, child in enumerate(rpr):
        name = child.tag.split('}')[-1]
        if name in RPR_ORDER and RPR_ORDER.index(name) > idx:
            pos = i
            break
    rpr.insert(pos, el)
    return el


def set_run_format(run_el, ea_font, sz_pt, bold=None, latin_font='Times New Roman'):
    rpr = run_el.find(qn('w:rPr'))
    if rpr is None:
        rpr = OxmlElement('w:rPr')
        run_el.insert(0, rpr)
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts')
        pos = 1 if rpr.find(qn('w:rStyle')) is not None else 0
        rpr.insert(pos, rf)
    # 主题字体引用会盖掉显式字体，必须删掉
    for a in THEME_ATTRS:
        rf.attrib.pop(qn('w:' + a), None)
    rf.set(qn('w:ascii'), latin_font)
    rf.set(qn('w:hAnsi'), latin_font)
    rf.set(qn('w:eastAsia'), ea_font)
    rf.set(qn('w:cs'), latin_font)
    if sz_pt:
        half = str(int(round(sz_pt * 2)))
        _ordered_put(rpr, 'sz', {'val': half})
        _ordered_put(rpr, 'szCs', {'val': half})
    if bold is not None:
        _ordered_put(rpr, 'b', {'val': '1' if bold else '0'})
        _ordered_put(rpr, 'bCs', {'val': '1' if bold else '0'})
    return rpr


def set_para_format(p, align=None, line_pt=None, before_pt=0, after_pt=0,
                    first_line_chars=0, left_chars=0, right_chars=0, base_sz=16):
    """重建 pPr：清掉源文档的缩进/行距/自动间距，再按规范重写。"""
    el = p._p
    old = el.find(qn('w:pPr'))
    keep = {}
    if old is not None:
        for tag in P_KEEP + ['numPr', 'rPr', 'sectPr']:
            sub = old.find(qn('w:' + tag))
            if sub is not None:
                keep[tag] = copy.deepcopy(sub)
        el.remove(old)

    new = OxmlElement('w:pPr')
    for tag in P_KEEP:
        if tag in keep:
            new.append(keep[tag])
    wc = OxmlElement('w:widowControl')
    wc.set(qn('w:val'), '0')
    new.append(wc)
    has_num = 'numPr' in keep
    if has_num:
        new.append(keep['numPr'])
    sg = OxmlElement('w:snapToGrid')   # 不对齐文档网格，否则固定行距会被改写
    sg.set(qn('w:val'), '0')
    new.append(sg)

    sp = OxmlElement('w:spacing')
    sp.set(qn('w:before'), str(int(round(before_pt * 20))))
    sp.set(qn('w:after'), str(int(round(after_pt * 20))))
    sp.set(qn('w:beforeAutospacing'), '0')
    sp.set(qn('w:afterAutospacing'), '0')
    if line_pt:
        sp.set(qn('w:line'), str(int(round(line_pt * 20))))
        sp.set(qn('w:lineRule'), 'exact')
    new.append(sp)

    if not has_num and (first_line_chars or left_chars or right_chars):
        ind = OxmlElement('w:ind')
        tw = lambda chars: str(int(round(chars / 100.0 * base_sz * 20)))
        if first_line_chars:
            ind.set(qn('w:firstLineChars'), str(int(first_line_chars)))
            ind.set(qn('w:firstLine'), tw(first_line_chars))
        if left_chars:
            ind.set(qn('w:leftChars'), str(int(left_chars)))
            ind.set(qn('w:left'), tw(left_chars))
        if right_chars:
            ind.set(qn('w:rightChars'), str(int(right_chars)))
            ind.set(qn('w:right'), tw(right_chars))
        new.append(ind)
    if align:
        jc = OxmlElement('w:jc')
        jc.set(qn('w:val'), align)
        new.append(jc)
    if 'rPr' in keep:
        new.append(keep['rPr'])
    if 'sectPr' in keep:
        new.append(keep['sectPr'])
    el.insert(0, new)


# ------------------------------------------------------------------ 段落识别

def para_text(p):
    return ''.join(t.text or '' for t in p._p.iter(qn('w:t')))


def has_drawing(p):
    return p._p.find(qn('w:drawing')) is not None or p._p.find(qn('w:pict')) is not None


def outline_level(p):
    pPr = p._p.find(qn('w:pPr'))
    if pPr is None:
        return None
    ol = pPr.find(qn('w:outlineLvl'))
    if ol is None:
        return None
    try:
        return int(ol.get(qn('w:val')))
    except (TypeError, ValueError):
        return None


def current_jc(p):
    pPr = p._p.find(qn('w:pPr'))
    if pPr is None:
        return None
    jc = pPr.find(qn('w:jc'))
    return jc.get(qn('w:val')) if jc is not None else None


def classify(paras, args):
    texts = [para_text(p).strip() for p in paras]
    kinds, notes = {}, {}

    title_idx = None
    if args.title_index is not None:
        title_idx = args.title_index
        notes[title_idx] = '按 --title-index 指定'
    elif args.title:
        target = args.title.replace(' ', '')
        for i, t in enumerate(texts):
            if t.replace(' ', '') == target:
                title_idx = i
                notes[i] = '按 --title 文本匹配'
                break
    else:
        for i, t in enumerate(texts):
            if not t or has_drawing(paras[i]):
                continue
            if len(t) <= 40 and not t.endswith(('。', '；', '，')):
                title_idx = i
                notes[i] = '首段短句自动识别为主标题'
            break

    last_long = -1
    for i, t in enumerate(texts):
        if i != title_idx and len(t) > 60:
            last_long = i
    sign_start = last_long + 1

    for i, p in enumerate(paras):
        t = texts[i]
        if i == title_idx:
            kinds[i] = 'title'
            continue
        if not t and not has_drawing(p):
            kinds[i] = 'empty'
            continue
        if has_drawing(p) and not t:
            kinds[i] = 'image'
            continue
        if title_idx is not None and i == title_idx + 1 and t.startswith(('——', '—')):
            kinds[i] = 'subtitle'
            notes[i] = '紧跟标题且以破折号开头，识别为副标题'
            continue
        lvl = outline_level(p)
        if lvl is not None and 0 <= lvl <= 3:
            kinds[i] = 'h%d' % (lvl + 1)
            notes[i] = '按大纲级别 %d' % lvl
            continue
        if len(t) <= 45 and not t.endswith('。'):
            if RE_H1.match(t):
                kinds[i] = 'h1'
                continue
            if RE_H2.match(t):
                kinds[i] = 'h2'
                continue
            if RE_H4.match(t):
                kinds[i] = 'h4'
                continue
            if RE_H3.match(t):
                kinds[i] = 'h3'
                continue
        if i >= sign_start and len(t) <= 30 and not RE_H1.match(t):
            kinds[i] = 'sign'
            notes[i] = '文末短段落，识别为落款'
            continue
        kinds[i] = 'body'
    return kinds, notes, title_idx


# ------------------------------------------------------------------ 套用格式

def apply_para(p, kind, cfg):
    style = cfg['style']
    bsz = style['body']['size_pt']
    line = cfg['body_line_pt']
    ind = cfg['indent_chars']

    if kind == 'title':
        set_para_format(p, align='center', line_pt=None, base_sz=style['title']['size_pt'])
        spec = style['title']
    elif kind == 'subtitle':
        set_para_format(p, align='center', line_pt=line, base_sz=style['subtitle']['size_pt'])
        spec = style['subtitle']
    elif kind in ('h1', 'h2', 'h3', 'h4'):
        set_para_format(p, align='both', line_pt=line, first_line_chars=ind, base_sz=bsz)
        spec = style[kind]
    elif kind == 'sign':
        right = 400 if RE_DATE.match(para_text(p).strip()) else 200
        set_para_format(p, align='right', line_pt=line, right_chars=right, base_sz=bsz)
        spec = style['sign']
    elif kind == 'image':
        set_para_format(p, align='center', line_pt=None)
        spec = None
    elif kind == 'empty':
        set_para_format(p, align='left', line_pt=line)
        spec = None
    else:
        set_para_format(p, align='both', line_pt=line, first_line_chars=ind, base_sz=bsz)
        spec = style['body']

    if spec is None:
        return
    for run_el in p._p.iter(qn('w:r')):
        set_run_format(run_el, spec['east_asia'], spec['size_pt'], spec['bold'], cfg['latin_font'])


def format_tables(doc, cfg, size_pt):
    ea = cfg['style']['body']['east_asia']
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    set_para_format(p, align=current_jc(p) or 'left', line_pt=None,
                                    first_line_chars=0, base_sz=size_pt)
                    for run_el in p._p.iter(qn('w:r')):
                        set_run_format(run_el, ea, size_pt, None, cfg['latin_font'])


def setup_sections(doc, cfg):
    pg = cfg['page']
    for s in doc.sections:
        s.page_width = Cm(pg['width_cm'])
        s.page_height = Cm(pg['height_cm'])
        s.top_margin = Cm(pg['top_cm'])
        s.bottom_margin = Cm(pg['bottom_cm'])
        s.left_margin = Cm(pg['left_cm'])
        s.right_margin = Cm(pg['right_cm'])
        s.header_distance = Cm(pg['header_cm'])
        s.footer_distance = Cm(pg['footer_cm'])
        s.gutter = Cm(0)
        try:
            s.different_first_page_header_footer = False
        except Exception:
            pass


def clear_container(container):
    el = container._element
    for p in list(el.findall(qn('w:p'))):
        el.remove(p)


def build_footer(footer, align, chars, cfg):
    clear_container(footer)
    p = footer.add_paragraph()
    pn = cfg['pagenum']
    spec = dict(align=align, line_pt=None, before_pt=0, after_pt=0, base_sz=pn['size_pt'])
    spec['right_chars' if align == 'right' else 'left_chars'] = chars
    set_para_format(p, **spec)

    dash = pn.get('dash', '—')
    p.add_run(dash)
    p.add_run(' ')
    r_begin = p.add_run()
    fc = OxmlElement('w:fldChar')
    fc.set(qn('w:fldCharType'), 'begin')
    fc.set(qn('w:dirty'), 'true')
    r_begin._element.append(fc)
    r_instr = p.add_run()
    it = OxmlElement('w:instrText')
    it.set(XML_SPACE, 'preserve')
    it.text = ' PAGE '
    r_instr._element.append(it)
    r_sep = p.add_run()
    fc2 = OxmlElement('w:fldChar')
    fc2.set(qn('w:fldCharType'), 'separate')
    r_sep._element.append(fc2)
    p.add_run('1')
    r_end = p.add_run()
    fc3 = OxmlElement('w:fldChar')
    fc3.set(qn('w:fldCharType'), 'end')
    r_end._element.append(fc3)
    p.add_run(' ')
    p.add_run(dash)

    for run_el in p._p.iter(qn('w:r')):
        set_run_format(run_el, pn['east_asia'], pn['size_pt'], False, cfg['latin_font'])


def setup_page_numbers(doc, cfg, keep_header):
    sec = doc.sections[0]
    for s in doc.sections[1:]:
        for attr in ('footer', 'even_page_footer', 'header'):
            try:
                getattr(s, attr).is_linked_to_previous = True
            except Exception:
                pass
    try:
        doc.settings.odd_and_even_pages_header_footer = True
    except Exception:
        settings = doc.settings.element
        if settings.find(qn('w:evenAndOddHeaders')) is None:
            settings.append(OxmlElement('w:evenAndOddHeaders'))

    even = sec.even_page_footer
    even.is_linked_to_previous = False
    build_footer(even, 'left', cfg['pagenum']['space_chars'], cfg)
    odd = sec.footer
    odd.is_linked_to_previous = False
    build_footer(odd, 'right', cfg['pagenum']['space_chars'], cfg)

    if cfg['header'].get('clear', True) and not keep_header:
        for s in doc.sections:
            try:
                h = s.header
                clear_container(h)
                h.add_paragraph()
            except Exception:
                pass


def fix_default_styles(doc, cfg):
    """把文档默认字体/字号也改成规范值，方便用户后续手动编辑。"""
    body = cfg['style']['body']
    styles_el = doc.styles.element
    dd = styles_el.find(qn('w:docDefaults'))
    if dd is not None:
        rpr_default = dd.find(qn('w:rPrDefault'))
        rpr = None
        if rpr_default is not None:
            rpr = rpr_default.find(qn('w:rPr'))
        if rpr is None:
            rpr = OxmlElement('w:rPr')
            (rpr_default if rpr_default is not None else dd).append(rpr)
        set_run_format(rpr, body['east_asia'], body['size_pt'], None, cfg['latin_font'])
    try:
        normal = doc.styles['Normal']
        set_run_format(normal.element, body['east_asia'], body['size_pt'], None, cfg['latin_font'])
    except Exception:
        pass


# ------------------------------------------------------------------ 报告

def write_report(path, src, out, paras, kinds, notes, cfg):
    lines = ['公文格式排版报告', '=' * 60,
             '源文件：%s' % src, '输出　：%s' % out, '']
    pg = cfg['page']
    lines.append('页面：%.1f×%.1fcm，上 %.1f / 下 %.1f / 左 %.1f / 右 %.1f cm' % (
        pg['width_cm'], pg['height_cm'], pg['top_cm'], pg['bottom_cm'],
        pg['left_cm'], pg['right_cm']))
    b = cfg['style']['body']
    lines.append('正文：%s %.1fpt，数字字母 %s，两端对齐，首行缩进 %.1f 字符，行距固定值 %d 磅' % (
        b['east_asia'], b['size_pt'], cfg['latin_font'],
        cfg['indent_chars'] / 100.0, cfg['body_line_pt']))
    if cfg['pagenum'].get('enabled', True):
        pn = cfg['pagenum']
        lines.append('页码：%.1fpt %s，%s 1 %s，单页居右空一字、双页居左空一字' % (
            pn['size_pt'], pn['east_asia'], pn.get('dash', '—'), pn.get('dash', '—')))
    lines.append('')
    counter = {}
    for i in sorted(kinds):
        counter[kinds[i]] = counter.get(kinds[i], 0) + 1
    lines.append('段落识别：' + '，'.join(
        '%s %d 段' % (KIND_CN.get(k, k), v) for k, v in counter.items()))
    lines.append('')
    lines.append('逐段明细：')
    for i in sorted(kinds):
        txt = para_text(paras[i]).strip().replace('\n', ' ')
        if len(txt) > 34:
            txt = txt[:34] + '…'
        note = '  （%s）' % notes[i] if i in notes else ''
        lines.append('  #%-3d %-6s %s%s' % (i, KIND_CN.get(kinds[i], kinds[i]), txt, note))
    lines.append('')
    lines.append('识别有误时：用 --title 或 --title-index 指定标题，或调整 classify() 规则。')
    with open(path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))


# ------------------------------------------------------------------ 配置

def merge(base, override):
    out = copy.deepcopy(base)
    for k, v in (override or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = merge(out[k], v)
        else:
            out[k] = v
    return out


def load_config(path):
    if not path:
        return copy.deepcopy(DEFAULTS)
    with open(path, 'r', encoding='utf-8') as f:
        user = json.load(f)
    return merge(DEFAULTS, user)


# ------------------------------------------------------------------ 主流程

def main():
    ap = argparse.ArgumentParser(description='按公文格式规格重排 Word 文档')
    ap.add_argument('input')
    ap.add_argument('output', nargs='?')
    ap.add_argument('--config', help='JSON 规格文件（覆盖默认值）')
    ap.add_argument('--dump-config', metavar='PATH', help='导出默认配置模板')
    ap.add_argument('--title')
    ap.add_argument('--title-index', type=int)
    ap.add_argument('--title-bold', action='store_true')
    ap.add_argument('--no-page-number', action='store_true')
    ap.add_argument('--keep-header', action='store_true')
    ap.add_argument('--table-size', type=float)
    ap.add_argument('--dry-run', action='store_true')
    args = ap.parse_args()

    cfg = load_config(args.config)
    if args.title_bold:
        cfg['style']['title']['bold'] = True
    if args.no_page_number:
        cfg['pagenum']['enabled'] = False

    if args.dump_config:
        with open(args.dump_config, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        print('配置模板已写入：%s' % args.dump_config)
        return

    if args.input.lower().endswith('.doc'):
        sys.exit('这是旧版 .doc。请用 Word 打开后“另存为 .docx”，或先转档再运行本脚本。')
    if not os.path.exists(args.input):
        sys.exit('找不到文件：%s' % args.input)

    out = args.output or os.path.splitext(args.input)[0] + '_公文格式.docx'
    if os.path.abspath(out) == os.path.abspath(args.input):
        sys.exit('输出文件不能覆盖源文件')

    doc = Document(args.input)
    paras = list(doc.paragraphs)
    kinds, notes, title_idx = classify(paras, args)
    if args.title and title_idx is None:
        print('警告：--title 在正文中找不到完全一致的段落，已退回自动识别。')

    if args.dry_run:
        for i in sorted(kinds):
            t = para_text(paras[i]).strip()
            print('#%-3d %-6s %s' % (i, KIND_CN.get(kinds[i], kinds[i]), t[:40]))
        return

    setup_sections(doc, cfg)
    fix_default_styles(doc, cfg)
    for i, p in enumerate(paras):
        apply_para(p, kinds.get(i, 'body'), cfg)
    format_tables(doc, cfg, args.table_size or cfg['table']['size_pt'])
    if cfg['pagenum'].get('enabled', True):
        setup_page_numbers(doc, cfg, args.keep_header)

    doc.save(out)
    rep = os.path.splitext(out)[0] + '_排版报告.txt'
    write_report(rep, args.input, out, paras, kinds, notes, cfg)

    print('已输出：%s' % out)
    print('报告　：%s' % rep)
    if title_idx is None:
        print('警告：未识别到主标题，首段按正文处理。可用 --title 或 --title-index 指定。')


if __name__ == '__main__':
    main()
