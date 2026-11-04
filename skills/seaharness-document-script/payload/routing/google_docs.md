# Google Docs routing

1. Produce and validate a local `.docx` with this skill first.
2. Before any Google Docs import, run `"$DSH_OFFICE_PYTHON" scripts/google_docs_title_sanitize.py input.docx --out sanitized.docx` and `"$DSH_OFFICE_PYTHON" scripts/google_docs_title_sanitize.py sanitized.docx --check`.
3. If the current session has a Drive/Docs import tool, use it with native Docs upload semantics. If it does not, deliver the sanitized DOCX and tell the user: native Google Docs import is unavailable; here is an importable DOCX.
4. Do not type into a blank Google Doc with Browser Use as the main path unless the user explicitly asks.
