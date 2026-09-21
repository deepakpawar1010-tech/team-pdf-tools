import html
import re
from pathlib import Path
import markdown

KATEX_CSS_CDN = "https://cdn.jsdelivr.net/npm/katex@0.16.10/dist/katex.min.css"
KATEX_JS_CDN = "https://cdn.jsdelivr.net/npm/katex@0.16.10/dist/katex.min.js"
KATEX_AUTO_RENDER_CDN = "https://cdn.jsdelivr.net/npm/katex@0.16.10/dist/contrib/auto-render.min.js"


def _protect_latex_math(text: str) -> tuple[str, dict[str, str]]:
    """
    Extracts LaTeX math formulas ($$, $, \\[, \\() and replaces them with placeholders
    so Python Markdown does not mangle underscores, asterisks, or backslashes.
    """
    placeholders: dict[str, str] = {}
    counter = 0

    # 1. Protect block math: $$...$$ and \[...\]
    def replace_block(match: re.Match) -> str:
        nonlocal counter
        token = f"MATHBLOCKTOKEN{counter}ENDTOKEN"
        counter += 1
        placeholders[token] = match.group(0)
        return token

    text = re.sub(r"\$\$.*?\$\$", replace_block, text, flags=re.DOTALL)
    text = re.sub(r"\\\[.*?\\\]", replace_block, text, flags=re.DOTALL)

    # 2. Protect inline math: \(...\) and $...$
    def replace_inline(match: re.Match) -> str:
        nonlocal counter
        token = f"MATHINLINETOKEN{counter}ENDTOKEN"
        counter += 1
        placeholders[token] = match.group(0)
        return token

    text = re.sub(r"\\\(.*?\\\)", replace_inline, text, flags=re.DOTALL)
    # Match single $, but avoid matching currency like $100 or double $$
    text = re.sub(r"(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)", replace_inline, text)

    return text, placeholders


def _restore_latex_math(html_text: str, placeholders: dict[str, str]) -> str:
    """Restores protected LaTeX math formulas into the rendered HTML."""
    for token, math_expr in placeholders.items():
        html_text = html_text.replace(token, math_expr)
    return html_text


def _format_exam_markdown(text: str) -> str:
    """
    Enhances question paper elements (Q1, Q2, Question:, Key:, Explanation:, etc.)
    into structured HTML cards while preserving standard Markdown.
    """
    lines = text.splitlines()
    formatted_lines: list[str] = []
    in_question_card = False

    # Regex for question headers: Q1, Q2, **Q11**, Q16), etc.
    q_start_pattern = re.compile(r"^\s*(?:\*\*)?(?:Q\d+|Question\s+\d+)[\):.]*(?:\*\*)?\s*$", re.IGNORECASE)
    # Subject banners: "Maths", "Physics", "Chemistry", "Biology", "Maths-A", etc.
    subject_pattern = re.compile(r"^\s*(?:#\s*)?(Maths|Physics|Chemistry|Biology|Mathematics|Science)(?:-[A-Z])?\s*$", re.IGNORECASE)
    # Top test paper title (e.g. HS/AP/Mock/JEE MAIN_OLY_8TH STD_E2_46304)
    exam_title_pattern = re.compile(r"^\s*(?:HS/|AP/|Mock/|JEE|NEET|CBSE|SSC|OLY).{5,80}$", re.IGNORECASE)
    # Option patterns: A) ..., B) ... or 1) ..., 2) ...
    option_pattern = re.compile(r"^\s*([A-D1-4])[\).]\s+(.*)$")

    in_options = False
    in_meta = False

    for line in lines:
        stripped = line.strip()

        # Check for top exam title banner
        if exam_title_pattern.match(stripped) and not stripped.startswith("#"):
            if in_options:
                formatted_lines.append("</div>")
                in_options = False
            if in_meta:
                formatted_lines.append("</div>")
                in_meta = False
            if in_question_card:
                formatted_lines.append("</div>")
                in_question_card = False
            formatted_lines.append(f"<div class='exam-main-title'>{html.escape(stripped)}</div>")
            continue

        # Check for major Subject Banner
        sub_m = subject_pattern.match(stripped)
        if sub_m:
            if in_options:
                formatted_lines.append("</div>")
                in_options = False
            if in_meta:
                formatted_lines.append("</div>")
                in_meta = False
            if in_question_card:
                formatted_lines.append("</div>")
                in_question_card = False
            subj_name = stripped.lstrip("#").strip()
            formatted_lines.append(f"<div class='subject-banner'><h2>{html.escape(subj_name)}</h2></div>")
            continue

        # Check for Question start
        if q_start_pattern.match(stripped):
            if in_options:
                formatted_lines.append("</div>")
                in_options = False
            if in_meta:
                formatted_lines.append("</div>")
                in_meta = False
            if in_question_card:
                formatted_lines.append("</div>")
            clean_q = stripped.replace("*", "").strip()
            formatted_lines.append("<div class='question-card'>")
            formatted_lines.append(f"<div class='question-header'>{html.escape(clean_q)}</div>")
            in_question_card = True
            continue

        # Check for Options (A) / 1) / etc.)
        opt_m = option_pattern.match(stripped)
        if opt_m:
            if in_meta:
                formatted_lines.append("</div>")
                in_meta = False
            if not in_options:
                formatted_lines.append("<div class='options-grid'>")
                in_options = True
            opt_letter, opt_text = opt_m.group(1), opt_m.group(2)
            formatted_lines.append(f"<div class='option-item'><span class='opt-label'>{opt_letter})</span> <span class='opt-text'>{opt_text}</span></div>")
            continue
        else:
            if in_options:
                formatted_lines.append("</div>")
                in_options = False

        # Format Key: lines
        if re.match(r"^\s*Key\s*:\s*", stripped, re.IGNORECASE):
            if in_meta:
                formatted_lines.append("</div>")
                in_meta = False
            key_val = re.sub(r"^\s*Key\s*:\s*", "", stripped, flags=re.IGNORECASE)
            formatted_lines.append(f"<div class='key-row'><span class='key-badge'>Key:</span> <span class='key-value'>{key_val}</span></div>")
            continue

        # Format Explanation: lines
        if re.match(r"^\s*Explanation\s*:\s*", stripped, re.IGNORECASE):
            if in_meta:
                formatted_lines.append("</div>")
                in_meta = False
            exp_text = re.sub(r"^\s*Explanation\s*:\s*", "", stripped, flags=re.IGNORECASE)
            formatted_lines.append(f"<div class='explanation-box'><strong>Explanation:</strong> {exp_text}</div>")
            continue

        # Format Metadata fields (Topic, Subtopic, Difficulty Level, etc.)
        if re.match(r"^\s*(?:Topic|Subtopic|Difficulty Level|Question Type|Question Mode|Level|Worksheet|Source)\s*:\s*", stripped, re.IGNORECASE):
            if not in_meta:
                formatted_lines.append("<div class='meta-container'>")
                in_meta = True
            parts = stripped.split(":", 1)
            label = parts[0].strip()
            val = parts[1].strip() if len(parts) > 1 else ""
            formatted_lines.append(f"<div class='meta-row'><span class='meta-label'>{html.escape(label)}:</span> <span class='meta-value'>{html.escape(val)}</span></div>")
            continue
        else:
            if in_meta:
                formatted_lines.append("</div>")
                in_meta = False

        formatted_lines.append(line)

    if in_options:
        formatted_lines.append("</div>")
    if in_meta:
        formatted_lines.append("</div>")
    if in_question_card:
        formatted_lines.append("</div>")

    return "\n".join(formatted_lines)


def get_css_for_preset(preset: str = "exam", paper_size: str = "A4", font_family: str = "sans") -> str:
    page_size = "A4" if paper_size.upper() == "A4" else "letter"
    font_stack = (
        "'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif"
        if font_family == "sans"
        else "'Computer Modern', 'Times New Roman', Times, serif"
    )

    base_css = f"""
    @page {{
        size: {page_size};
        margin: 12mm 14mm;
        @bottom-right {{
            content: counter(page);
            font-size: 9pt;
            color: #64748b;
        }}
    }}
    * {{
        box-sizing: border-box;
    }}
    body {{
        font-family: {font_stack};
        font-size: 10pt;
        line-height: 1.5;
        color: #1e293b;
        margin: 0;
        padding: 0;
        background: #ffffff;
    }}
    h1, h2, h3, h4 {{
        color: #0f172a;
        margin-top: 1em;
        margin-bottom: 0.4em;
        font-weight: 700;
        break-after: avoid;
        page-break-after: avoid;
    }}
    table {{
        width: 100%;
        border-collapse: collapse;
        margin: 10px 0;
        font-size: 9.5pt;
    }}
    th, td {{
        border: 1px solid #cbd5e1;
        padding: 6px 10px;
        text-align: left;
    }}
    th {{
        background-color: #f1f5f9;
        font-weight: 700;
    }}
    p {{
        margin: 0.3em 0;
    }}
    ul, ol {{
        margin: 0.3em 0;
        padding-left: 20px;
    }}
    li {{
        margin-bottom: 0.2em;
    }}
    code {{
        background: #f1f5f9;
        padding: 2px 5px;
        border-radius: 4px;
        font-family: Consolas, Monaco, monospace;
        font-size: 9pt;
    }}
    pre {{
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 6px;
        padding: 10px;
        overflow-x: auto;
    }}
    .katex {{
        font-size: 1.05em;
    }}
    @media print {{
        body {{
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
        }}
        .question-card, .subject-banner, .exam-main-title {{
            break-inside: avoid;
            page-break-inside: avoid;
        }}
        .subject-banner, .exam-main-title {{
            break-after: avoid;
            page-break-after: avoid;
        }}
    }}
    """

    if preset == "exam":
        exam_css = """
        .exam-main-title {
            text-align: center;
            font-size: 12pt;
            font-weight: 800;
            color: #d97706;
            border: 1.5px solid #f59e0b;
            background: #fffbeb;
            padding: 6px 12px;
            border-radius: 6px;
            margin-bottom: 10px;
            letter-spacing: 0.02em;
            break-after: avoid;
            page-break-after: avoid;
        }
        .subject-banner {
            background: #f1f5f9;
            border-top: 2px solid #0f172a;
            border-bottom: 2px solid #0f172a;
            text-align: center;
            padding: 4px 10px;
            margin: 12px 0 8px;
            break-after: avoid;
            page-break-after: avoid;
        }
        .subject-banner h2 {
            margin: 0;
            font-size: 13pt;
            font-weight: 800;
            color: #0f172a;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }
        .question-card {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 6px;
            padding: 8px 12px;
            margin-bottom: 10px;
            page-break-inside: avoid;
            break-inside: avoid;
            box-shadow: 0 1px 2px rgba(0,0,0,0.02);
        }
        .question-header {
            font-weight: 800;
            color: #2563eb;
            font-size: 10.5pt;
            margin-bottom: 4px;
            border-bottom: 1px solid #f1f5f9;
            padding-bottom: 2px;
        }
        .options-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(130px, 1fr));
            gap: 6px 14px;
            margin: 6px 0 8px;
            padding-left: 4px;
        }
        .option-item {
            font-size: 10pt;
            display: flex;
            align-items: baseline;
            gap: 5px;
        }
        .opt-label {
            font-weight: 700;
            color: #1e293b;
        }
        .key-row {
            margin: 6px 0;
            font-size: 10pt;
        }
        .key-badge {
            display: inline-block;
            background: #16a34a;
            color: white;
            font-weight: 700;
            padding: 1px 7px;
            border-radius: 4px;
            font-size: 9pt;
        }
        .key-value {
            font-weight: 800;
            color: #16a34a;
            margin-left: 4px;
        }
        .explanation-box {
            background: #f8fafc;
            border-left: 3px solid #3b82f6;
            padding: 6px 10px;
            margin: 6px 0;
            font-size: 9pt;
            color: #334155;
            border-radius: 0 4px 4px 0;
        }
        .meta-container {
            margin-top: 6px;
            padding-top: 4px;
            border-top: 1px dashed #e2e8f0;
        }
        .meta-row {
            font-size: 8.5pt;
            color: #64748b;
            line-height: 1.35;
        }
        .meta-label {
            font-weight: 700;
            color: #475569;
        }
        """
        return base_css + exam_css

    elif preset == "academic":
        academic_css = """
        body {
            line-height: 1.65;
        }
        h1 {
            text-align: center;
            font-size: 16pt;
            border-bottom: 1px solid #334155;
            padding-bottom: 8px;
        }
        h2 {
            font-size: 13pt;
            border-bottom: 1px solid #cbd5e1;
            padding-bottom: 4px;
        }
        .question-card {
            margin-bottom: 16px;
            page-break-inside: avoid;
            break-inside: avoid;
        }
        """
        return base_css + academic_css

    else:
        # Clean Modern
        clean_css = """
        h1 {
            font-size: 18pt;
            color: #0f172a;
        }
        .question-card {
            border-left: 3px solid #6366f1;
            padding-left: 12px;
            margin-bottom: 14px;
            page-break-inside: avoid;
            break-inside: avoid;
        }
        """
        return base_css + clean_css


def convert_markdown_to_html(
    markdown_text: str,
    preset: str = "exam",
    paper_size: str = "A4",
    render_math: bool = True,
    font_family: str = "sans",
) -> str:
    """
    Converts Markdown text with LaTeX math formulas into standalone HTML
    ready for headless browser rendering to PDF.
    """
    if render_math:
        protected_md, math_placeholders = _protect_latex_math(markdown_text)
    else:
        protected_md = markdown_text
        math_placeholders = {}

    if preset == "exam":
        protected_md = _format_exam_markdown(protected_md)

    html_body = markdown.markdown(
        protected_md,
        extensions=["tables", "fenced_code", "def_list", "nl2br"],
    )

    if render_math:
        html_body = _restore_latex_math(html_body, math_placeholders)

    css = get_css_for_preset(preset=preset, paper_size=paper_size, font_family=font_family)

    katex_scripts = ""
    if render_math:
        katex_scripts = f"""
        <link rel="stylesheet" href="{KATEX_CSS_CDN}">
        <script defer src="{KATEX_JS_CDN}"></script>
        <script defer src="{KATEX_AUTO_RENDER_CDN}" onload="
            renderMathInElement(document.body, {{
                delimiters: [
                    {{left: '$$', right: '$$', display: true}},
                    {{left: '$', right: '$', display: false}},
                    {{left: '\\\\(', right: '\\\\)', display: false}},
                    {{left: '\\\\[', right: '\\\\]', display: true}}
                ],
                throwOnError: false
            }});
        "></script>
        """

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Markdown to PDF</title>
    {katex_scripts}
    <style>
    {css}
    </style>
</head>
<body>
    {html_body}
</body>
</html>"""
