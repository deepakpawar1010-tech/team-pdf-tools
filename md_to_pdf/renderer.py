import logging
import os
import platform
import shutil
import subprocess
import tempfile
from pathlib import Path

logger = logging.getLogger(__name__)

WINDOWS_BROWSER_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
    os.path.expandvars(r"%LocalAppData%\Microsoft\Edge\Application\msedge.exe"),
]

LINUX_BROWSER_COMMANDS = [
    "chromium",
    "chromium-browser",
    "google-chrome",
    "google-chrome-stable",
]


def find_browser_executable() -> str | None:
    system = platform.system()
    if system == "Windows":
        for path in WINDOWS_BROWSER_PATHS:
            if os.path.exists(path):
                return path
        for cmd in ["chrome", "msedge"]:
            p = shutil.which(cmd)
            if p:
                return p
    else:
        for cmd in LINUX_BROWSER_COMMANDS:
            p = shutil.which(cmd)
            if p:
                return p
    return None


def render_html_to_pdf(html_content: str, output_pdf_path: Path | None = None) -> bytes:
    browser = find_browser_executable()
    if not browser:
        raise RuntimeError(
            "No headless browser (Chrome, Edge, or Chromium) found on the server. "
            "Please ensure Google Chrome or Chromium is installed."
        )

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        html_file = tmp_path / "document.html"
        pdf_file = tmp_path / "document.pdf"

        html_file.write_text(html_content, encoding="utf-8")

        cmd = [
            browser,
            "--headless",
            "--disable-gpu",
            "--no-pdf-header-footer",
            "--run-all-compositor-stages-before-draw",
        ]

        if platform.system() != "Windows":
            cmd.extend(["--no-sandbox", "--disable-dev-shm-usage"])

        cmd.extend([
            f"--print-to-pdf={pdf_file.resolve()}",
            html_file.resolve().as_uri(),
        ])

        logger.info("Executing headless browser command: %s", " ".join(cmd))
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=45,
            )
            if result.returncode != 0:
                logger.error("Browser failed with code %s: %s", result.returncode, result.stderr)
                raise RuntimeError(f"Browser PDF generation failed: {result.stderr or result.stdout}")
        except subprocess.TimeoutExpired:
            raise TimeoutError("PDF rendering timed out after 45 seconds.")

        if not pdf_file.exists() or pdf_file.stat().st_size == 0:
            raise RuntimeError("Browser finished but no PDF was generated.")

        pdf_bytes = pdf_file.read_bytes()

        if output_pdf_path:
            output_pdf_path.parent.mkdir(parents=True, exist_ok=True)
            output_pdf_path.write_bytes(pdf_bytes)

        return pdf_bytes
