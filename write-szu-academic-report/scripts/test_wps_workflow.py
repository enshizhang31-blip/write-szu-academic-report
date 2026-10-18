"""Isolated regression tests. Mock renderer tests never certify real office software."""
import unittest
from unittest.mock import patch
from copy import deepcopy
import test_report_tools
from report_core import *
from report_reflow import reflow, candidate_reason
from report_verify import verify, CHECKS
from report_visual import line_risks
from PIL import Image, ImageDraw


class WorkflowTests(unittest.TestCase):
    setUp=test_report_tools.ReportTests.setUp
    tearDown=test_report_tools.ReportTests.tearDown

    def test_office_nested_heading_metadata(self):
        from report_editors import normalize_headings
        h={'level':1,'label':'一、','text':'实验方法'}
        self.assertEqual(normalize_headings([[h],None,[[h]]]),[h,h])
        self.assertEqual(normalize_headings(h),[h])
        with self.assertRaises(ValueError):normalize_headings('invalid')

    def test_wps_collapsed_grid_and_short_row_pagination(self):
        d=Document();t=d.add_table(rows=3,cols=2);t.autofit=False
        node(t._tbl.tblPr,'tblCellSpacing',w=0,type='dxa')
        node(t.rows[1]._tr.get_or_add_trPr(),'tblCellSpacing',w=0,type='dxa')
        t.cell(1,0).text='短行';t.cell(2,0).text='长'*1000
        node(node(t.cell(1,0)._tc.get_or_add_tcPr(),'tcBorders'),'bottom',val='double',sz=8)
        format_table(t,'data')
        self.assertFalse(t._tbl.xpath('.//w:tblCellSpacing'))
        self.assertFalse(t._tbl.xpath('.//w:tcBorders'))
        self.assertEqual(int(value(t._tbl.tblPr,'tblW','w')),sum(int(c.get(qn('w:w'))) for c in t._tbl.tblGrid))
        self.assertTrue(t.rows[1]._tr.xpath('./w:trPr/w:cantSplit'))
        self.assertFalse(t.rows[2]._tr.xpath('./w:trPr/w:cantSplit'))
        self.assertTrue(t.rows[0]._tr.xpath('./w:trPr/w:tblHeader'))

    def fixture(self, merged=False):
        d=Document();d.add_paragraph('封面保留')
        t=d.add_table(rows=5,cols=1)
        for i,c in enumerate([r.cells[0] for r in t.rows]):c.text='正文'+str(i)
        nested=t.cell(1,0).add_table(rows=2,cols=2);nested.cell(0,0).text='真实数据';format_table(nested,'data')
        from code_box import add_code_box
        add_code_box(t.cell(2,0),'print("test")')
        t.cell(4,0).text='指导教师签字：'
        if merged:t.cell(0,0).merge(t.cell(1,0))
        src=self.root/'source.docx';d.save(src)
        prepared=self.root/'prepared.docx';mf=self.root/'m.json'
        prepare(src,prepared,mf,layout_tables=[0])
        return prepared,load_manifest(mf)

    def convert(self):
        src,m=self.fixture();out=self.root/'flow.docx';mf=self.root/'flow.json'
        reflow(src,out,m,mf,[0]);return src,m,out,load_manifest(mf)

    def test_reflow_preserves_order_nested_tables_code_and_teacher(self):
        src,m,out,updated=self.convert();before=Document(src);after=Document(out)
        self.assertEqual([p.text for p in paragraphs(before) if p.text],[p.text for p in paragraphs(after) if p.text])
        self.assertEqual(audit_protection(after,updated,out),[])
        self.assertEqual(len(updated['paragraphs']),len(m['paragraphs']))
        self.assertTrue(updated['tables'][-1]['retired_by_reflow'])
        self.assertEqual(len(tables(before)),len(tables(after)))
        self.assertEqual([p.text for p in after.tables[-1].cell(0,0).paragraphs],['指导教师签字：'])
        self.assertFalse(any(candidate_reason(t) for t in after.tables))

    def test_reflow_rejects_merge_and_unapproved_table(self):
        src,m=self.fixture(merged=True)
        with self.assertRaises(ValueError):reflow(src,self.root/'bad.docx',m,self.root/'bad.json',[0])
        self.assertFalse((self.root/'bad.docx').exists())
        with self.assertRaises(ValueError):reflow(src,self.root/'bad.docx',m,self.root/'bad.json',[1])

    def test_reflow_bounds_and_original_format_remain_protected(self):
        src,old,out,m=self.convert();d=Document(out)
        start=d.element.body.xpath('.//w:bookmarkStart')[0]
        p=d.add_paragraph('合法新正文');start.getparent().addnext(p._p)
        self.assertEqual(audit_protection(d,m),[])
        d.paragraphs[0].runs[0].bold=True
        self.assertIn('protected_paragraph_changed',{i['rule'] for i in audit_protection(d,m)})

    def test_anchor_cannot_expand_to_cover(self):
        src,old,out,m=self.convert();d=Document(out)
        start=d.element.body.xpath('.//w:bookmarkStart')[0].getparent()
        d.element.body.remove(start);d.element.body.insert(0,start)
        self.assertIn('flow_region_protected_members_changed',{i['rule'] for i in audit_protection(d,m)})

    def test_format_new_flow_content_and_protect_teacher(self):
        src,old,out,m=self.convert();d=Document(out)
        start=d.element.body.xpath('.//w:bookmarkStart')[0].getparent()
        p=d.add_paragraph('新增正常正文');start.addnext(p._p)
        draft=self.root/'draft.docx';d.save(draft);final=self.root/'formatted.docx'
        format_document(draft,final,m)
        result=Document(final)
        p=next(p for p in result.paragraphs if p.text=='新增正常正文')
        self.assertEqual(effective_p(p,'ind','firstLine'),'420')
        self.assertEqual(audit_protection(result,m,final),[])
        result.tables[-1].cell(0,0).add_paragraph('不应代签')
        self.assertTrue(audit_protection(result,m))

    def test_body_wrapper_and_orphan_heading_risks(self):
        src,m=self.fixture();d=Document(src)
        self.assertIn('body_layout_table',{i['rule'] for i in audit(d,m)['issues']})
        from report_tools import render_risks
        im=Image.new('RGB',(400,600),'white')
        risks=render_risks([im],['正文末尾孤立标题'],[{'page':1,'text':'孤立标题'}])
        self.assertIn('possible_orphan_heading',{x['rule'] for x in risks})

    def test_legacy_mode_explicit_and_settings_protected(self):
        d=Document();d.add_paragraph('原文')
        for x in d.settings.element.xpath('.//w:compatSetting'):
            if x.get(qn('w:name'))=='compatibilityMode':x.set(qn('w:val'),'12')
        src=self.root/'legacy.docx';d.save(src);out=self.root/'new.docx';mf=self.root/'m.json'
        with self.assertRaises(ValueError):prepare(src,out,mf)
        prepare(src,out,mf,upgrade_compatibility=True);m=load_manifest(mf)
        self.assertTrue(m['compatibility_transition']['explicit'])
        d=Document(out);node(d.settings.element,'zoom',percent=150)
        self.assertIn('protected_settings_changed',{i['rule'] for i in audit_protection(d,m)})

    def test_thin_borders_and_conflicting_overrides(self):
        d=Document();ensure_styles(d);t=d.add_table(rows=4,cols=2);t.style='Table Grid'
        node(t.cell(0,0)._tc.get_or_add_tcPr(),'tcBorders').append(OxmlElement('w:bottom'))
        format_table(t,'data')
        self.assertIsNone(t._tbl.tblPr.find(qn('w:tblStyle')))
        for edge in t._tbl.tblPr.find(qn('w:tblBorders')):
            self.assertEqual((edge.get(qn('w:val')),edge.get(qn('w:sz')),edge.get(qn('w:color'))),('single','4','000000'))
        self.assertEqual([value(r.cells[0]._tc.tcPr,'shd','fill') for r in t.rows],['FFFFFF','FFFFFF','F2F2F2','FFFFFF'])
        for val,size in [('dash',4),('double',4),('single',8)]:
            node(node(t._tbl.tblPr,'tblBorders'),'bottom',val=val,sz=size,color='000000')
            self.assertGreater(audit(d)['errors'],0)
        format_table(t,'data');node(node(t.cell(1,0)._tc.tcPr,'tcBorders'),'bottom',val='nil')
        self.assertIn('cell_border_override',{i['rule'] for i in audit(d)['issues']})

    def test_visual_gap_double_and_clipping_hints(self):
        im=Image.new('RGB',(600,800),'white');dr=ImageDraw.Draw(im)
        dr.rectangle((50,100,550,700),outline='black',width=2)
        dr.line((50,300,550,300),fill='black',width=2);dr.rectangle((240,298,260,302),fill='white')
        dr.line((50,500,550,500),fill='black',width=2);dr.line((50,507,550,507),fill='black',width=2)
        dr.rectangle((0,0,4,4),fill='black')
        rules={x['rule'] for x in line_risks(im)}
        self.assertTrue({'possible_border_gap','possible_duplicate_border','possible_page_edge_clipping'}.issubset(rules))

    def fake_render(self,src,out,editor,*args):
        out.mkdir(parents=True,exist_ok=True);(out/'report.pdf').write_bytes(b'test fixture only')
        (out/'page-1.png').write_bytes(b'test fixture only')
        r={'backend':editor,'version':'TEST-MOCK','source_sha256':digest(src),'pdf_sha256':digest(out/'report.pdf'),
           'pages':1,'headings':[],'risks':[],'page_image_sha256':{'page-1.png':digest(out/'page-1.png')}}
        write_json(out/'render.json',r);return r

    def verification_fixture(self):
        d=Document();ensure_styles(d);p=d.add_paragraph('测试正文',style='SZUBody');set_font(p.runs[0],10.5,latin='Times New Roman')
        src=self.root/'v.docx';d.save(src);return src,self.root/'renders'

    def test_missing_editor_blocks_dual_acceptance(self):
        src,out=self.verification_fixture()
        with patch('report_verify.render',side_effect=RuntimeError('missing editor')):
            r,code=verify(src,None,out)
        self.assertEqual(code,2);self.assertFalse(r['dual_editor_verified'])

    def test_wrong_renderer_cannot_impersonate_required_editor(self):
        src,out=self.verification_fixture()
        with patch('report_verify.render',side_effect=self.fake_render):verify(src,None,out)
        p=out/'wps/render.json';r=load_manifest(p);r['backend']='word';write_json(p,r)
        r,code=verify(src,None,out,reuse=True)
        self.assertEqual(code,2);self.assertFalse(r['dual_editor_verified'])

    def test_review_pending_then_stale_source_pdf_image_rejected(self):
        src,out=self.verification_fixture()
        with patch('report_verify.render',side_effect=self.fake_render):r,code=verify(src,None,out)
        self.assertEqual(code,3);self.assertFalse(r['delivery_ready'])
        review=load_manifest(out/'review-template.json');review.update(reviewer='TEST ONLY',semantic_pass=True,semantic_notes='Mock workflow test')
        for x in review['editors'].values():x.update(pages_reviewed=[1],visual_pass=True,visual_notes='Mock workflow test',checks_passed=CHECKS,warning_rules_reviewed=[i['rule'] for i in r['issues'] if i['severity']=='warning'])
        rp=self.root/'review.json';write_json(rp,review)
        r,code=verify(src,None,out,reuse=True,review_path=rp);self.assertEqual(code,0)
        for filename in ['report.pdf','page-1.png']:
            p=out/'wps'/filename;original=p.read_bytes();p.write_bytes(b'changed')
            r,code=verify(src,None,out,reuse=True,review_path=rp)
            self.assertEqual(code,2);self.assertFalse(r['dual_editor_verified']);p.write_bytes(original)
        review['source_sha256']='stale';write_json(rp,review)
        r,code=verify(src,None,out,reuse=True,review_path=rp);self.assertEqual(code,3)


if __name__=='__main__':unittest.main()
