#!/usr/bin/env python3
"""Compatibility entrypoint; prefer report_tools.py audit."""
import sys
from pathlib import Path
from docx import Document
from report_core import audit as audit_document
from report_tools import main
def audit(path: Path):
    return audit_document(Document(path), path=path)
if __name__ == '__main__':
    sys.exit(main(['audit', *sys.argv[1:]]))
