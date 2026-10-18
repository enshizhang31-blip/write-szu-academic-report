#!/usr/bin/env python3
"""Compare against prepare's per-region baseline, allowing legitimate filling."""
import argparse,json
from docx import Document
from report_core import audit_protection, digest, issues_result, load_manifest, write_json
def audit(template_path, report_path, manifest_path=None):
    if not manifest_path:
        raise ValueError('A prepare-generated --manifest is required; whole-table equality is not a valid preservation test.')
    m=load_manifest(manifest_path)
    if digest(template_path)!=m['template_sha256']:
        raise ValueError('Template does not match manifest baseline.')
    return issues_result(audit_protection(Document(report_path),m,report_path))
if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('template');p.add_argument('report');p.add_argument('--manifest',required=True);p.add_argument('--json')
    a=p.parse_args()
    try:
        r=audit(a.template,a.report,a.manifest)
        if a.json:write_json(a.json,r)
        print(json.dumps(r,ensure_ascii=False,indent=2));raise SystemExit(1 if r['errors'] else 0)
    except ValueError as e:p.exit(2,str(e)+'\n')
