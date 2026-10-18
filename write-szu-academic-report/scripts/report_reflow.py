"""Explicit layout-table migration with retained protection records and bounded flow regions."""
from copy import deepcopy
from pathlib import Path
from uuid import uuid4
import re
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table
from report_core import (audit_protection, paragraphs, tables, table_id, value, node,
                         canon, pid, ancestors, digest, write_json, save_new)

TEACHER = re.compile(r'^\s*(指导教师批阅意见|成绩评定|指导教师签字|备注)')

def flow_members(doc, manifest):
    """Return paragraph elements in intact named bookmark intervals; fail closed."""
    result=set(); problems=[]
    flat=list(doc.element.body.iter()); positions={el:i for i,el in enumerate(flat)}
    for r in manifest.get('flow_regions',[]):
        starts=doc.element.body.xpath('.//w:bookmarkStart[@w:name="'+r['name']+'"]')
        ends=doc.element.body.xpath('.//w:bookmarkEnd[@w:id="'+r['bookmark_id']+'"]')
        if len(starts)!=1 or len(ends)!=1 or starts[0].get(qn('w:id'))!=r['bookmark_id'] or positions[starts[0]]>=positions[ends[0]]:
            problems.append({'severity':'error','rule':'flow_region_anchor_missing_or_invalid','region':r['name']});continue
        lo,hi=positions[starts[0]],positions[ends[0]]
        result.update(p._p for p in paragraphs(doc) if lo<positions[p._p]<hi)
        # Protected endpoints prevent the region from being moved to the cover/teacher area.
        ids=[pid(p._p) for p in paragraphs(doc) if lo<positions[p._p]<hi and pid(p._p)]
        expected=r['protected_ids']
        protected={x['id'] for x in manifest['paragraphs']}
        if [x for x in ids if x in protected]!=[x for x in expected if x in protected]:
            problems.append({'severity':'error','rule':'flow_region_protected_members_changed','region':r['name']})
    return result,problems

def candidate_reason(t):
    from report_core import table_role
    if table_role(t) in ('code','form'):return None
    rows=t._tbl.findall(qn('w:tr'))
    if not rows or any(len(r.findall(qn('w:tc')))!=1 for r in rows):return None
    ps=[p for r in rows for p in r.findall('./'+qn('w:tc')+'/'+qn('w:p'))]
    if any(TEACHER.match(''.join(p.itertext())) for p in ps) and len(rows)<=3:return None
    styles=[value(p.find(qn('w:pPr')),'pStyle') or '' for p in ps]
    if len(rows)>=4 or sum(s=='SZUBody' for s in styles)>=4 or any(s.startswith('SZUHeading') for s in styles):
        return 'Single-column wrapper contains report paragraphs/headings; explicitly authorize reflow or review as a genuine form.'
    return None

def _anchor(name,bid,end=False):
    p=OxmlElement('w:p');pr=node(p,'pPr');node(pr,'spacing',before=0,after=0,line=20,lineRule='exact')
    b=OxmlElement('w:bookmarkEnd' if end else 'w:bookmarkStart');b.set(qn('w:id'),bid)
    if not end:b.set(qn('w:name'),name)
    p.append(b);return p

def reflow(source, output, manifest, manifest_out, indices):
    if not indices:raise ValueError('Explicit --layout-table index is required; reflow never guesses authorization.')
    if Path(output).resolve()==Path(source).resolve():raise ValueError('Reflow requires a new output file.')
    doc=Document(source);m=deepcopy(manifest)
    problems=audit_protection(doc,m,source)
    if problems:raise ValueError('Protection failed before reflow: '+str(problems[:3]))
    ts=tables(doc);records={r['id']:r for r in m['tables']}
    if len(set(indices))!=len(indices) or any(i<0 or i>=len(ts) for i in indices):raise ValueError('Invalid or duplicate layout table index.')
    selected=[ts[i] for i in indices]
    for t in selected:
        rec=records.get(table_id(t))
        if not rec or not rec['layout']:raise ValueError('Selected table is not authorized for layout changes in manifest.')
        if ancestors(t._tbl,'tbl'):raise ValueError('Nested layout wrappers require manual review; no automatic unwrapping.')
        if len(t.columns)!=1 or t._tbl.xpath('./w:tr/w:tc/w:tcPr/w:gridSpan | ./w:tr/w:tc/w:tcPr/w:vMerge | ./w:tr/w:tc/w:tcPr/w:hMerge'):
            raise ValueError('Merged or multi-column layout requires manual review.')
        if t._tbl.xpath('./w:tblPr/w:tblpPr | .//wp:anchor | ./w:tr/w:tc/w:tcPr/w:textDirection'):
            raise ValueError('Floating/positioned layout requires manual review.')
        for row in t._tbl.findall(qn('w:tr')):
            cells=row.findall(qn('w:tc'))
            if len(cells)!=1 or any(x.tag not in {qn('w:tcPr'),qn('w:p'),qn('w:tbl')} for x in cells[0]):
                raise ValueError('Unsupported cell structure; manual review required.')
    max_bid=max([int(x.get(qn('w:id'))) for x in doc.element.body.xpath('.//w:bookmarkStart') if x.get(qn('w:id'),'').isdigit()]+[0])
    migrations=[];new_form_ids=set()
    for t in selected:
        ident=table_id(t);rec=records[ident]
        old_tables={pid(p._p):[value(a.find(qn('w:tblPr')),'tblCaption') for a in ancestors(p._p,'tbl')] for p in paragraphs(doc) if t._tbl in ancestors(p._p,'tbl')}
        groups=[]
        for tr in t._tbl.findall(qn('w:tr')):
            tc=tr.find(qn('w:tc'))
            teacher=any(TEACHER.match(''.join(p.xpath('.//w:t/text()'))) for p in tc.findall(qn('w:p')))
            if not groups or groups[-1][0]!=teacher:groups.append((teacher,[]))
            groups[-1][1].append(tr)
        for teacher,rows in groups:
            if teacher:
                form=OxmlElement('w:tbl');form.append(deepcopy(t._tbl.tblPr));form.append(deepcopy(t._tbl.tblGrid))
                fid='SZU_FORM_'+uuid4().hex;node(form.find(qn('w:tblPr')),'tblCaption',val=fid)
                # Preserve original visual properties; classify in manifest, not by changing style.
                for tr in rows:form.append(tr)
                t._tbl.addprevious(form);new_form_ids.add(fid)
            else:
                max_bid+=1;bid=str(max_bid);name='SZU_FLOW_'+uuid4().hex[:20]
                t._tbl.addprevious(_anchor(name,bid));protected_ids=[]
                for tr in rows:
                    for el in list(tr.find(qn('w:tc'))):
                        if el.tag==qn('w:tcPr'):continue
                        t._tbl.addprevious(el)
                        protected_ids.extend(pid(p) for p in el.iter(qn('w:p')) if pid(p))
                        if el.tag==qn('w:p') and pid(el) and pid(el) not in protected_ids:protected_ids.append(pid(el))
                t._tbl.addprevious(_anchor(name,bid,True))
                m.setdefault('flow_regions',[]).append({'name':name,'bookmark_id':bid,'source_table':ident,'protected_ids':protected_ids})
        t._tbl.getparent().remove(t._tbl)
        rec['retired_by_reflow']=True
        migrations.append({'source_table':ident,'source_table_record':deepcopy(rec),'paragraph_locations_before':old_tables})
    byid={pid(p._p):p for p in paragraphs(doc)}
    for r in m['paragraphs']:
        p=byid[r['id']]
        new=[value(a.find(qn('w:tblPr')),'tblCaption') for a in ancestors(p._p,'tbl')]
        if new!=r['tables']:
            r.setdefault('original_tables',list(r['tables']));r['tables']=new
    current=[]
    for t in tables(doc):
        ident=table_id(t);newanc=[value(a.find(qn('w:tblPr')),'tblCaption') for a in ancestors(t._tbl,'tbl')]
        if ident in records:
            r=records[ident]
            if r['ancestor_tables']!=newanc:r.setdefault('original_ancestor_tables',r['ancestor_tables']);r['ancestor_tables']=newanc
        else:
            if ident not in new_form_ids:raise ValueError('Unexpected table after reflow.')
            r={'id':ident,'layout':False,'role':'form','properties':canon(t._tbl.tblPr,('tblCaption',)),
               'grid':canon(t._tbl.tblGrid),'rows':[[canon(row._tr.trPr),[canon(c._tc.tcPr) for c in row.cells]] for row in t.rows],
               'ancestor_tables':newanc}
        current.append(r)
    # Retired records are retained as transformation evidence and must stay absent.
    m['tables']=current+[r for r in m['tables'] if r.get('retired_by_reflow')]
    m['version']=2;m.setdefault('transformations',[]).append({'kind':'reflow','source_sha256':digest(source),'tables':migrations})
    save_new(doc,source,output)
    after=audit_protection(Document(output),m,output)
    if after:raise ValueError('Protection failed after reflow: '+str(after[:3]))
    write_json(manifest_out,m)
    return {'output':str(output),'manifest':str(manifest_out),'unwrapped_tables':len(selected),'protected_paragraphs':len(m['paragraphs']),'review_required':True}
