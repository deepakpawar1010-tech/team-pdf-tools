import re
import fitz


def extract_latex_text_from_pdf(pdf_bytes: bytes) -> str:
    """
    Extracts text and embedded LaTeX formulas from a PDF document.
    Cleans up PDF extraction artifacts such as BOM, zero-width spaces,
    and trailing empty attachment markers.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages_text: list[str] = []

    for page in doc:
        raw_text = page.get_text()
        if not raw_text:
            continue

        # Clean invisible/problematic unicode characters
        cleaned = (
            raw_text.replace("\ufeff", "")
            .replace("\u200b", "")
            .replace("\xa0", " ")
        )

        # Remove standalone 'No attachment' artifact lines
        cleaned = re.sub(r"^\s*No attachment\s*$", "", cleaned, flags=re.MULTILINE)
        # Remove empty QP Associated Images banner if standalone
        cleaned = re.sub(r"^\s*QP Associated Images\*?\s*$", "", cleaned, flags=re.MULTILINE)

        # Strip trailing whitespace on each line
        lines = [line.rstrip() for line in cleaned.splitlines()]
        page_clean = "\n".join(lines).strip()

        if page_clean:
            pages_text.append(page_clean)

    # Join pages with double newlines
    full_text = "\n\n".join(pages_text)

    # Normalize 3+ consecutive newlines into 2
    full_text = re.sub(r"\n{3,}", "\n\n", full_text)

    return full_text.strip()
