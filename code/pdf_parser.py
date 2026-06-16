"""
Pluggable PDF text parser.

Used only by the text-input sensitivity conditions (`text_md`). The benchmark's
main `pdf` condition attaches the PDF natively to the model and does NOT use this
parser, so swapping the parser here affects only those sensitivity conditions.

The published study used pdfplumber. To experiment with a different parser, replace
`extract_text` below with another implementation that takes a PDF path and returns
the extracted text as a single string — the rest of the pipeline is unchanged.
"""
from __future__ import annotations

from pathlib import Path


def extract_text(pdf_path: str | Path) -> str:
    """Extract text from a PDF using pdfplumber (study default)."""
    import pdfplumber

    texts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                texts.append(text.strip())
    return "\n\n".join(texts)
