# -*- coding: utf-8 -*-
"""冒烟测试：造一份乱格式文档，排版后断言各项属性。

    python -m unittest discover -s tests -v

不需要安装 Word / LibreOffice。
"""
import contextlib
import os
import shutil
import subprocess
import sys
import uuid
import unittest
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt

ROOT = Path(__file__).resolve().parents[1]
FORMAT = ROOT / 'scripts' / 'gongwen_format.py'
VALIDATE = ROOT / 'scripts' / 'validate_docx.py'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# 临时目录放在仓库内：某些沙箱环境不允许写系统 %TEMP%
TMP_ROOT = ROOT / '.tmp'


def newtmp(prefix='t-'):
    """自己建目录，不用 tempfile.mkdtemp——它创建的 0700 目录在沙箱下不可写。"""
    TMP_ROOT.mkdir(parents=True, exist_ok=True)
    d = TMP_ROOT / (prefix + uuid.uuid4().hex[:8])
    d.mkdir()
    return d


@contextlib.contextmanager
def tmpdir():
    d = newtmp()
    try:
        yield d
    finally:
        shutil.rmtree(d, ignore_errors=True)

BODY_TEXT = [
    '为贯彻落实上级部署，我单位自2026年9月起组织开展专项工作，现将有关情况报告如下，请予审阅。',
    '成立工作专班，明确责任分工，建立周调度机制，累计召开推进会十二次，覆盖全部业务条线。',
    '主要负责同志牵头抓总，分管领导具体负责，各处室按职责分工推进，确保任务到岗、责任到人。',
    '制定任务清单和时间表，逐项明确完成标准和验收方式，实行台账管理、挂图作业、逐月通报。',
]


def build_source(path):
    """造一份典型乱格式文档：黑体加粗标题、微软雅黑正文、无缩进、无页码。"""
    doc = Document()

    def add(text, size=12, bold=False, font='微软雅黑'):
        p = doc.add_paragraph()
        r = p.add_run(text)
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.name = font
        r._element.rPr.rFonts.set(qn('w:eastAsia'), font)
        return p

    add('关于开展专项工作的报告', 20, True, '黑体')
    add('各单位：', 12)
    add('一、总体情况', 12, True)
    add(BODY_TEXT[0])
    add('（一）主要做法', 12, True)
    add(BODY_TEXT[1])
    add('1.加强组织领导', 12, True)
    add(BODY_TEXT[2])
    add('（1）细化措施', 12, True)
    add(BODY_TEXT[3])
    add('二〇二六年十一月九日', 12)
    doc.save(path)


def run(script, *args):
    env = dict(os.environ, PYTHONIOENCODING='utf-8')
    return subprocess.run(
        [sys.executable, str(script)] + [str(a) for a in args],
        cwd=str(ROOT), capture_output=True, text=True,
        encoding='utf-8', errors='replace', env=env
    )


def body_text(path):
    with zipfile.ZipFile(path) as z:
        import lxml.etree as etree
        root = etree.fromstring(z.read('word/document.xml'))
        return ''.join(t.text or '' for t in root.find(W + 'body').iter(W + 't'))


class FormatTest(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        TMP_ROOT.mkdir(parents=True, exist_ok=True)
        cls.tmp = newtmp('fmt-')
        cls.src = cls.tmp / 'in.docx'
        cls.out = cls.tmp / 'out.docx'
        build_source(cls.src)
        cls.proc = run(FORMAT, cls.src, cls.out)
        if cls.proc.returncode != 0:
            raise AssertionError('排版脚本失败：%s\n%s' % (cls.proc.stdout, cls.proc.stderr))
        cls.doc = Document(str(cls.out))
        cls.paras = cls.doc.paragraphs

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    # --- 页面 ---
    def test_page_size_and_margins(self):
        s = self.doc.sections[0]
        self.assertAlmostEqual(s.page_width.cm, 21.0, places=1)
        self.assertAlmostEqual(s.page_height.cm, 29.7, places=1)
        self.assertAlmostEqual(s.top_margin.cm, 3.7, places=1)
        self.assertAlmostEqual(s.bottom_margin.cm, 3.5, places=1)
        self.assertAlmostEqual(s.left_margin.cm, 2.8, places=1)
        self.assertAlmostEqual(s.right_margin.cm, 2.6, places=1)

    # --- 文字 ---
    def test_text_unchanged(self):
        self.assertEqual(body_text(self.src), body_text(self.out))

    # --- 标题 ---
    def test_title(self):
        p = self.paras[0]
        self.assertEqual(p.alignment, WD_ALIGN_PARAGRAPH.CENTER)
        run0 = p.runs[0]
        self.assertEqual(run0._element.rPr.rFonts.get(qn('w:eastAsia')), '方正小标宋简体')
        self.assertEqual(run0.font.size.pt, 22)
        self.assertFalse(run0.font.bold)

    # --- 正文 ---
    def test_body_paragraph(self):
        p = self.paras[3]
        self.assertEqual(p.alignment, WD_ALIGN_PARAGRAPH.JUSTIFY)
        ind = p._p.pPr.find(qn('w:ind'))
        self.assertEqual(ind.get(qn('w:firstLineChars')), '200')
        sp = p._p.pPr.find(qn('w:spacing'))
        self.assertEqual(sp.get(qn('w:line')), '600')
        self.assertEqual(sp.get(qn('w:lineRule')), 'exact')
        run0 = p.runs[0]
        self.assertEqual(run0._element.rPr.rFonts.get(qn('w:eastAsia')), '仿宋_GB2312')
        self.assertEqual(run0._element.rPr.rFonts.get(qn('w:ascii')), 'Times New Roman')
        self.assertEqual(run0.font.size.pt, 16)
        self.assertFalse(run0.font.bold)

    def test_no_theme_font_leak(self):
        for p in self.paras:
            for r in p.runs:
                rf = r._element.rPr.rFonts
                self.assertIsNone(rf.get(qn('w:eastAsiaTheme')))

    def test_snap_to_grid_disabled(self):
        for p in self.paras:
            pPr = p._p.pPr
            if pPr is None:
                continue
            sg = pPr.find(qn('w:snapToGrid'))
            if sg is not None:
                self.assertEqual(sg.get(qn('w:val')), '0')

    # --- 各级标题 ---
    def test_headings(self):
        expect = {2: ('黑体', False), 4: ('楷体_GB2312', False),
                  6: ('仿宋_GB2312', True), 8: ('仿宋_GB2312', True)}
        for idx, (font, bold) in expect.items():
            run0 = self.paras[idx].runs[0]
            self.assertEqual(run0._element.rPr.rFonts.get(qn('w:eastAsia')), font,
                             '第 %d 段字体不对' % idx)
            self.assertEqual(bool(run0.font.bold), bold, '第 %d 段加粗不对' % idx)
            self.assertEqual(run0.font.size.pt, 16)

    # --- 落款 ---
    def test_signature_right_aligned(self):
        p = self.paras[-1]
        self.assertEqual(p.alignment, WD_ALIGN_PARAGRAPH.RIGHT)
        ind = p._p.pPr.find(qn('w:ind'))
        self.assertEqual(ind.get(qn('w:rightChars')), '400')   # 日期右空 4 字

    # --- 页码 ---
    def test_page_number_field(self):
        with zipfile.ZipFile(self.out) as z:
            footers = [n for n in z.namelist() if n.startswith('word/footer')]
            self.assertGreaterEqual(len(footers), 2, '应同时有奇偶页页脚')
            for name in footers:
                raw = z.read(name).decode('utf-8')
                self.assertIn('PAGE', raw)
                self.assertIn('fldChar', raw)
                self.assertIn('—', raw)

    def test_even_and_odd_enabled(self):
        with zipfile.ZipFile(self.out) as z:
            self.assertIn('evenAndOddHeaders', z.read('word/settings.xml').decode('utf-8'))


class CliTest(unittest.TestCase):

    def test_dry_run_writes_nothing(self):
        with tmpdir() as tmp:
            src = Path(tmp) / 'in.docx'
            build_source(src)
            proc = run(FORMAT, src, '--dry-run')
            self.assertEqual(proc.returncode, 0)
            self.assertIn('主标题', proc.stdout)
            self.assertEqual(sorted(p.name for p in Path(tmp).iterdir()), ['in.docx'])

    def test_refuses_to_overwrite_source(self):
        with tmpdir() as tmp:
            src = Path(tmp) / 'in.docx'
            build_source(src)
            proc = run(FORMAT, src, src)
            self.assertNotEqual(proc.returncode, 0)

    def test_rejects_legacy_doc(self):
        with tmpdir() as tmp:
            fake = Path(tmp) / 'old.doc'
            fake.write_bytes(b'not a real doc')
            proc = run(FORMAT, fake, Path(tmp) / 'out.docx')
            self.assertNotEqual(proc.returncode, 0)
            self.assertIn('.doc', proc.stdout + proc.stderr)

    def test_validator_passes(self):
        with tmpdir() as tmp:
            src = Path(tmp) / 'in.docx'
            out = Path(tmp) / 'out.docx'
            build_source(src)
            run(FORMAT, src, out)
            proc = run(VALIDATE, out, src)
            self.assertEqual(proc.returncode, 0, proc.stdout)
            self.assertIn('通过', proc.stdout)


if __name__ == '__main__':
    unittest.main()
