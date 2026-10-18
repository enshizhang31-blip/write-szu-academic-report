"""Per-editor evidence-bound verification; no renderer substitution."""
from pathlib import Path
import hashlib
import json
from docx import Document
from report_core import audit, digest, load_manifest, write_json, issue, issues_result, paragraphs
from report_editors import render

CHECKS=['closed_frames','continuous_table_grid','readable_figures','no_orphan_headings','figure_caption_pairs','whitespace_reviewed','thin_single_borders','no_body_layout_wrapper','clipping_overlap_reviewed']

def render_fingerprint(r):
    data={k:r.get(k) for k in ('source_sha256','backend','version','pdf_sha256','page_image_sha256','dpi','headings','risks')}
    return hashlib.sha256(json.dumps(data,sort_keys=True,ensure_ascii=False).encode('utf-8')).hexdigest()

def verify(docx,manifest,out,primary='wps',compat='word',reuse=False,review_path=None,dpi=200,soffice=None,evidence=None,source_dir=None):
    from report_tools import evidence_audit,source_check,rendered_numbering_issues,validate_review
    docx=Path(docx);out=Path(out);out.mkdir(parents=True,exist_ok=True)
    result=audit(Document(docx),manifest,docx);issues=result['issues']+evidence_audit(evidence)+source_check(docx,source_dir)
    editors=list(dict.fromkeys([primary]+([] if compat=='none' else [compat])))
    review=load_manifest(review_path) if review_path else {}
    review_map=review.get('editors',{})
    report_hash=digest(docx);states={};pending=[];template={'source_sha256':report_hash,'reviewer':'','semantic_pass':False,'semantic_notes':'','editors':{}}
    for editor in editors:
        rd=out/editor
        try:
            if reuse:
                rendered=load_manifest(rd/'render.json')
                if rendered.get('backend')!=editor:raise ValueError('Backend identity mismatch')
                if rendered.get('source_sha256')!=report_hash or rendered.get('pdf_sha256')!=digest(rd/'report.pdf'):raise ValueError('Stale source/PDF hash')
                if set(rendered.get('page_image_sha256',{}))!={f'page-{i}.png' for i in range(1,rendered['pages']+1)}:raise ValueError('Incomplete image inventory')
                for name,sha in rendered['page_image_sha256'].items():
                    if digest(rd/name)!=sha:raise ValueError('Stale page image: '+name)
                for risk in rendered.get('risks',[]):
                    if risk.get('evidence_image') and digest(rd/risk['evidence_image'])!=risk.get('evidence_sha256'):raise ValueError('Stale risk evidence image')
            else:rendered=render(docx,rd,editor,soffice,dpi)
            local_issues=rendered_numbering_issues(rendered)
            if editor in ('wps','word'):
                expected=sum(p.style.name.startswith('SZUHeading') for p in paragraphs(Document(docx)))
                if expected!=len(rendered.get('headings',[])):issue(local_issues,'rendered_heading_count_mismatch',expected=expected,actual=len(rendered.get('headings',[])))
            if rendered.get('roundtrip') and not rendered['roundtrip'].get('numbering_stable'):issue(local_issues,'roundtrip_numbering_changed')
            issues.extend({**i,'editor':editor} for i in local_issues)
            blank={'backend':editor,'version':rendered.get('version'),'render_pdf_sha256':rendered['pdf_sha256'],'render_sha256':render_fingerprint(rendered),'pages_reviewed':[],
                   'visual_pass':False,'visual_notes':'','checks_passed':[],'warning_rules_reviewed':[],'render_risk_rules_reviewed':[]}
            template['editors'][editor]=blank
            r=review_map.get(editor,{})
            # Validate in memory without generating review files which imply human approval.
            checks=[]
            if review.get('source_sha256')!=report_hash:checks.append('review_stale_source')
            if not review.get('reviewer') or not review.get('semantic_pass') or not review.get('semantic_notes'):checks.append('semantic_review_pending')
            if r.get('backend')!=editor or r.get('version')!=rendered.get('version'):checks.append('review_editor_identity_mismatch')
            if r.get('render_pdf_sha256')!=rendered['pdf_sha256']:checks.append('review_stale_render')
            if r.get('render_sha256')!=render_fingerprint(rendered):checks.append('review_stale_render_bundle')
            if sorted(r.get('pages_reviewed',[]))!=list(range(1,rendered['pages']+1)):checks.append('visual_pages_incomplete')
            if not r.get('visual_pass') or not r.get('visual_notes'):checks.append('visual_review_pending')
            if not set(CHECKS).issubset(r.get('checks_passed',[])):checks.append('visual_checklist_incomplete')
            warn={i['rule'] for i in issues if i['severity']=='warning'}
            if not warn.issubset(r.get('warning_rules_reviewed',[])):checks.append('warnings_not_reviewed')
            risks={i['rule'] for i in rendered.get('risks',[])}
            if not risks.issubset(r.get('render_risk_rules_reviewed',[])):checks.append('render_risks_not_reviewed')
            states[editor]={'status':'passed' if not checks and not local_issues else 'review_pending','version':rendered.get('version'),'pages':rendered['pages'],'pending':checks,'risks':rendered.get('risks',[])}
            pending.extend(editor+':'+x for x in checks)
        except Exception as exc:
            states[editor]={'status':'blocked','error':str(exc)};pending.append(editor+':render_unavailable')
    result.update(issues_result(issues));ready=not result['errors'] and all(s['status']=='passed' for s in states.values())
    result.update(primary_editor=primary,compatibility_editor=compat,editors=states,review_pending=pending,delivery_ready=ready,
                  dual_editor_verified=ready and 'wps' in states and 'word' in states,
                  semantic_review='passed' if ready else 'pending',visual_review='passed' if ready else 'pending')
    if not review_path:write_json(out/'review-template.json',template)
    code=1 if result['errors'] else (2 if any(s['status']=='blocked' for s in states.values()) else (0 if ready else 3))
    return result,code
