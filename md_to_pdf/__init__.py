from pathlib import Path
from .converter import convert_markdown_to_html
from .renderer import render_html_to_pdf
from .extractor import extract_latex_text_from_pdf


def render_markdown_to_pdf(
    markdown_text: str,
    output_pdf_path: Path | None = None,
    preset: str = "exam",
    paper_size: str = "A4",
    render_math: bool = True,
    font_family: str = "sans",
) -> bytes:
    html_content = convert_markdown_to_html(
        markdown_text=markdown_text,
        preset=preset,
        paper_size=paper_size,
        render_math=render_math,
        font_family=font_family,
    )
    return render_html_to_pdf(html_content, output_pdf_path=output_pdf_path)


__all__ = [
    "convert_markdown_to_html",
    "render_html_to_pdf",
    "render_markdown_to_pdf",
    "extract_latex_text_from_pdf",
]
