"""One tool-using agent over evidence-grounded sources.

session.py      the workbook as read once: contracts per account, source version, drafts
tools.py        the capabilities that already existed, wrapped so a failure is never a success
summary.py      the numerical summary, every figure traced to a fact
export.py       a deck from validated candidates, with the touched slides checked
controller.py   one loop: the model chooses a tool, the controller enforces what code can

The legacy pipeline (fdd_utils/ai/pipeline.py: Generator -> Auditor -> Validator) is
untouched and remains the default; see controller.py for the flag.
"""
