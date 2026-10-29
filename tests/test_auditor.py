"""执行: python -m unittest discover -s tests -p 'test_*.py' -v"""
from pathlib import Path
import sys
import tempfile
import unittest
from shutil import copy2
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from word_audit import Auditor, audit_dir, text_of, contains_han

class AuditTests(unittest.TestCase):
    def setUp(self):
        self.good=ROOT/'examples'/'01_符合规范的测试稿.docx'
        self.bad=ROOT/'examples'/'02_故意设置错误的测试稿.docx'

    def test_good_no_false_positive(self):
        rows=Auditor(self.good.name).check_document(self.good)
        self.assertEqual(rows, [])

    def test_bad_margin(self):
        rows=Auditor(self.bad.name).check_document(self.bad)
        self.assertTrue(any(r[2]=='页边距' and '2.00 cm' in r[4] for r in rows))

    def test_bad_line_spacing(self):
        rows=Auditor(self.bad.name).check_document(self.bad)
        self.assertTrue(any(r[2]=='行距' and '360 / auto' in r[4] for r in rows))

    def test_bad_font_size(self):
        rows=Auditor(self.bad.name).check_document(self.bad)
        self.assertTrue(any(r[2]=='字号' and '10.5 pt' in r[4] for r in rows))

    def test_bad_exactly_three_errors(self):
        rows=Auditor(self.bad.name).check_document(self.bad)
        self.assertEqual(len([r for r in rows if r[1]=='错误']),3)

    def test_batches_and_local_output(self):
        with tempfile.TemporaryDirectory() as t:
            d=Path(t); inp=d/'drafts'; inp.mkdir()
            copy2(self.good, inp/self.good.name); copy2(self.bad, inp/self.bad.name)
            out,summary=audit_dir(inp,d)
            self.assertEqual(len(summary),2)
            self.assertTrue((out/'审校报告.html').exists())
            self.assertTrue((out/'稿件汇总.csv').exists())
            self.assertTrue((out/'问题明细.csv').exists())
            html=(out/'审校报告.html').read_text('utf-8')
            self.assertIn('02_故意设置错误',html)

    def test_unparsable_does_not_kill_batch(self):
        with tempfile.TemporaryDirectory() as t:
            f=Path(t)/'broken.docx'; f.write_bytes(b'notdocx')
            rows=Auditor(f.name).check_document(f)
            self.assertTrue(any(r[2]=='无法解析' and r[1]=='错误' for r in rows))

    def test_extra_typographic_and_reference_checks(self):
        f=ROOT/'examples'/'03_多种异常测试稿.docx'
        rows=Auditor(f.name).check_document(f)
        cats={r[2] for r in rows}
        for required in ('页面大小','首行缩进','中文字体','章节语言','姓名语言','引用编号','参考文献编号'):
            self.assertIn(required,cats)

    def test_figure_reference_in_body_not_caption(self):
        import xml.etree.ElementTree as ET
        from word_audit import category_of
        class MinimalStyles:
            def p_props(self, p): return ([], 'Normal')
            styles = {'Normal': {'name': 'Normal'}}
        ns='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        p=ET.fromstring(f'<w:p xmlns:w="{ns}"><w:r><w:t>图 2 的距离数据支持这一解释。</w:t></w:r></w:p>')
        self.assertEqual(category_of(p, 4, False, MinimalStyles()), '正文')

    def test_reference_left_alignment_allowed(self):
        import xml.etree.ElementTree as ET
        from word_audit import Auditor, Styles
        p=ET.fromstring('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>[1] Example article.</w:t></w:r></w:p>')
        a=Auditor('sample.docx')
        styles=Styles(None); styles.doc_p=None; styles.doc_r=None
        a._paragraph_check(p, '[1] Example article.', '参考文献条目', '参考文献测试', styles, 999)
        self.assertFalse(any(r[2]=='段落对齐' for r in a.rows))

    def test_text_language_detection(self):
        self.assertTrue(contains_han('清华大学'))
        self.assertFalse(contains_han('Tsinghua'))

if __name__=='__main__': unittest.main()