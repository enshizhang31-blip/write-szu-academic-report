#!/usr/bin/env python3
"""Scoped formatter; never globally restyle an unknown template."""
import argparse
from report_core import format_document, load_manifest
def normalize(path, output, manifest=None, generated_all=False):
    return format_document(path, output, manifest, generated_all=generated_all)
if __name__ == '__main__':
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('input');p.add_argument('output');p.add_argument('--manifest');p.add_argument('--generated-all',action='store_true')
    a=p.parse_args()
    try: print(normalize(a.input,a.output,load_manifest(a.manifest),a.generated_all))
    except Exception as e: p.exit(2,str(e)+'\n')
