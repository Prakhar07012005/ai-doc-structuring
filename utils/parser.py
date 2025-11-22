import pdfplumber

def parse_pdf(filepath: str) -> str:
    """
    Extracts text from a PDF file, concatenating pages with newlines.
    """
    text = []
    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            text.append(page_text)
    return "\n".join(text)
