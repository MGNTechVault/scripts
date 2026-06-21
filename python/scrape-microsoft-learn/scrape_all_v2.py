#!/usr/bin/env python3
"""Microsoft Learn Scraper v2 — downloads learning paths as Markdown, DOCX, TXT, or PDF.

Author:  Matthew Gonzalez Nieves
Version: 1.0.0
Date:    31-03-2025
License: MIT
Website: https://mgntechvault.com

Purpose:
    This tool is intended for personal offline study using AI tools such as
    Microsoft Copilot Notebook and Google NotebookLM. Export your Microsoft Learn
    learning paths and feed them directly to your AI notebook of choice.

    This script is NOT intended to stress or flood Microsoft Learn (DDoS).
    Use it responsibly and only for your own study materials.

Usage:
    python scrape_all_v2.py --format markdown       # required flag
    python scrape_all_v2.py --format docx
    python scrape_all_v2.py --format txt
    python scrape_all_v2.py --format pdf
    python scrape_all_v2.py                         # interactive choice if --format is omitted
    python scrape_all_v2.py --pick                  # interactive learning-path picker
    python scrape_all_v2.py --locale nl-nl          # picker catalog locale (default en-us)
    python scrape_all_v2.py --config custom.yaml    # use a different config file
    python scrape_all_v2.py --urls URL1 URL2        # override URLs via CLI
    python scrape_all_v2.py --output filename       # override output name (extension auto-determined)
    python scrape_all_v2.py --headed                # show browser window
    python scrape_all_v2.py -v                      # verbose logging

    When neither --urls nor config paths are given, an interactive picker lists
    every current Microsoft Learn learning path so you can select one or more.

Requirements:
    pip install playwright pyyaml python-docx
    playwright install chromium
"""

import argparse
import json
import logging
import re
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

import yaml
from playwright.sync_api import sync_playwright, Page, TimeoutError as PwTimeout

logger = logging.getLogger(__name__)

VALID_FORMATS = ("markdown", "docx", "txt", "pdf")
FORMAT_EXTENSIONS = {"markdown": ".md", "docx": ".docx", "txt": ".txt", "pdf": ".pdf"}

# Browser user-agent reused for scraping and the catalog fetch.
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

# Microsoft Learn public catalog API — lists every learning path with its URL.
CATALOG_API = "https://learn.microsoft.com/api/catalog/?type=learningPaths&locale={locale}"
DEFAULT_LOCALE = "en-us"
PICKER_RESULT_LIMIT = 40

# --- Selectors (centralised so changes only need to happen in one place) ---
CONTENT_SELECTOR = "#unit-inner-section"
FALLBACK_CONTENT_SELECTOR = "div[role='main']"
UNIT_LIST_SELECTOR = "#unit-list"
NOISE_SELECTORS = (
    "button", "footer", "nav",
    ".display-flex",
    "#next-section", "#previous-section",
    ".assessment-table",
    ".unit-page-learning-objectives",
)
MIN_CONTENT_LENGTH = 50


# ── Data models ───────────────────────────────────────────────


@dataclass
class Unit:
    title: str
    url: str
    content: str


@dataclass
class Module:
    title: str
    url: str
    units: list[Unit] = field(default_factory=list)


@dataclass
class LearningPath:
    title: str
    url: str
    modules: list[Module] = field(default_factory=list)


@dataclass
class ScrapeStats:
    paths: int = 0
    modules: int = 0
    units_scraped: int = 0
    units_skipped: int = 0
    nav_failures: int = 0
    failed_urls: list[str] = field(default_factory=list)

    def summary(self) -> str:
        lines = [
            f"Paths processed:   {self.paths}",
            f"Modules processed: {self.modules}",
            f"Units scraped:     {self.units_scraped}",
            f"Units skipped:     {self.units_skipped}",
            f"Navigation errors: {self.nav_failures}",
        ]
        if self.failed_urls:
            lines.append("Failed URLs:")
            for url in self.failed_urls:
                lines.append(f"  - {url}")
        return "\n".join(lines)


# ── Utility functions ─────────────────────────────────────────


def clean_text(text: str) -> str:
    """Remove redundant blank lines."""
    if not text:
        return ""
    return re.sub(r"\n\s*\n", "\n\n", text.strip())


# ── Browser helpers ───────────────────────────────────────────


def safe_goto(page: Page, url: str, *, timeout: int = 20_000, retries: int = 3) -> bool:
    """Navigate to a URL with exponential backoff on failure."""
    for attempt in range(retries):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=timeout)
            return True
        except Exception as e:
            wait = 2 ** attempt
            logger.warning(
                "Navigation failed (%d/%d) %s — %s, retrying in %ds",
                attempt + 1, retries, url, e, wait,
            )
            time.sleep(wait)
    logger.error("Permanently failed: %s", url)
    return False


def wait_for_content(page: Page, timeout: int = 10_000) -> bool:
    """Wait until the content selector is visible on the page."""
    try:
        page.wait_for_selector(CONTENT_SELECTOR, state="attached", timeout=timeout)
        return True
    except PwTimeout:
        logger.debug("Content selector not found within timeout on %s", page.url)
        return False


def dismiss_cookie_banner(page: Page) -> None:
    """Click away the cookie banner if present."""
    for selector in ("button#onetrust-accept-btn-handler", "button.accept-cookies"):
        try:
            page.click(selector, timeout=2000)
            logger.debug("Cookie banner dismissed via %s", selector)
            return
        except PwTimeout:
            continue


def extract_content(page: Page) -> str | None:
    """Extract content via a DOM clone — does not mutate the original."""
    noise_css = ", ".join(NOISE_SELECTORS)
    result = page.evaluate(
        """([contentSel, fallbackSel, noiseSel]) => {
            const el = document.querySelector(contentSel)
                    || document.querySelector(fallbackSel);
            if (!el) return null;
            const clone = el.cloneNode(true);
            clone.querySelectorAll(noiseSel).forEach(e => e.remove());
            return clone.innerText.trim();
        }""",
        [CONTENT_SELECTOR, FALLBACK_CONTENT_SELECTOR, noise_css],
    )
    if result and len(result) > MIN_CONTENT_LENGTH:
        return clean_text(result)
    return None


# ── URL extraction ────────────────────────────────────────────


def normalize_url(base: str, href: str) -> str:
    """Build a full URL without query params, with trailing slash stripped."""
    return urljoin(base, href).split("?")[0].rstrip("/")


def get_module_urls(page: Page) -> list[str]:
    """Retrieve unique module URLs from a learning path page."""
    links = page.query_selector_all("a[href*='/modules/']")
    seen: set[str] = set()
    urls: list[str] = []
    for link in links:
        href = link.get_attribute("href")
        if not href:
            continue
        full = normalize_url(page.url, href)
        if full not in seen:
            seen.add(full)
            urls.append(full)
    return urls


def get_unit_urls(page: Page, module_url: str) -> list[str]:
    """Retrieve unique unit URLs from a module page."""
    container = page.query_selector(UNIT_LIST_SELECTOR)
    if not container:
        logger.debug("Selector '%s' not found, falling back to <main>", UNIT_LIST_SELECTOR)
        container = page.query_selector("main")
    if not container:
        return []

    normalized_module = module_url.rstrip("/")
    seen: set[str] = set()
    urls: list[str] = []
    for link in container.query_selector_all("a"):
        href = link.get_attribute("href")
        if not href:
            continue
        full = normalize_url(page.url, href)
        # Skip the module link itself and external links
        if full == normalized_module:
            continue
        if full not in seen:
            seen.add(full)
            urls.append(full)
    return urls


# ── Core scraping ─────────────────────────────────────────────


def scrape_unit(page: Page, url: str) -> Unit | None:
    """Scrape a single unit (lesson) and return a Unit, or None on failure."""
    if not safe_goto(page, url):
        return None
    wait_for_content(page)

    h1 = page.query_selector("h1")
    title = h1.inner_text().strip() if h1 else "Unknown lesson"
    content = extract_content(page)

    if not content:
        logger.info("  No usable content: %s", url)
        return None
    return Unit(title=title, url=url, content=content)


def scrape_module(page: Page, url: str, stats: ScrapeStats, seen_urls: set[str]) -> Module | None:
    """Scrape a module including all its units."""
    if not safe_goto(page, url):
        stats.nav_failures += 1
        stats.failed_urls.append(url)
        return None

    h1 = page.query_selector("h1")
    title = h1.inner_text().strip() if h1 else "Unknown module"
    module = Module(title=title, url=url)

    unit_urls = get_unit_urls(page, url)
    logger.info("    %d units found in '%s'", len(unit_urls), title)

    for idx, u_url in enumerate(unit_urls, 1):
        if u_url in seen_urls:
            logger.debug("    Unit already seen, skipping: %s", u_url)
            continue
        seen_urls.add(u_url)

        logger.info("      [%d/%d] %s", idx, len(unit_urls), u_url.split("/")[-1])
        unit = scrape_unit(page, u_url)
        if unit:
            module.units.append(unit)
            stats.units_scraped += 1
        else:
            stats.units_skipped += 1

    stats.modules += 1
    return module


def scrape_path(page: Page, url: str, stats: ScrapeStats, seen_urls: set[str]) -> LearningPath | None:
    """Scrape a learning path including all its modules and units."""
    logger.info("Learning path: %s", url.split("/")[-2])
    if not safe_goto(page, url):
        stats.nav_failures += 1
        stats.failed_urls.append(url)
        return None

    h1 = page.query_selector("h1")
    title = h1.inner_text().strip() if h1 else "Unknown path"
    path = LearningPath(title=title, url=url)

    module_urls = get_module_urls(page)
    logger.info("  %d modules found in '%s'", len(module_urls), title)

    for m_idx, m_url in enumerate(module_urls, 1):
        logger.info("  [%d/%d] Processing module...", m_idx, len(module_urls))
        module = scrape_module(page, m_url, stats, seen_urls)
        if module and module.units:
            path.modules.append(module)

    stats.paths += 1
    return path


# ── Writers ───────────────────────────────────────────────────


def _atomic_write_text(tmp: Path, output: Path, content_fn) -> None:
    """Write via a temporary file and atomically rename on success."""
    with open(tmp, "w", encoding="utf-8") as f:
        content_fn(f)
    tmp.replace(output)


def write_markdown(
    paths: list[LearningPath],
    output: Path,
    title: str = "AZ-305 Course Material",
) -> None:
    """Write to Markdown (.md). Atomic write via .tmp file."""
    tmp = output.with_suffix(".tmp")

    def _write(f):
        f.write(f"# {title}\n\n")
        f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n")

        f.write("## Table of Contents\n\n")
        for lp in paths:
            f.write(f"- **{lp.title}**\n")
            for mod in lp.modules:
                f.write(f"  - {mod.title} ({len(mod.units)} lessons)\n")
        f.write("\n---\n\n")

        for lp in paths:
            f.write(f"## {lp.title}\n\n")
            f.write(f"Source: {lp.url}\n\n")
            for mod in lp.modules:
                f.write(f"### {mod.title}\n\n")
                for unit in mod.units:
                    f.write(f"#### {unit.title}\n\n")
                    f.write(f"{unit.content}\n\n")
                    f.write("---\n\n")
            f.flush()

    _atomic_write_text(tmp, output, _write)
    logger.info("Markdown written to %s", output)


def write_txt(
    paths: list[LearningPath],
    output: Path,
    title: str = "AZ-305 Course Material",
) -> None:
    """Write to plain text (.txt). Atomic write via .tmp file."""
    tmp = output.with_suffix(".tmp")
    separator = "=" * 72
    thin_sep = "-" * 72

    def _write(f):
        f.write(f"{title}\n")
        f.write(f"{separator}\n")
        f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n")

        f.write("TABLE OF CONTENTS\n")
        f.write(f"{thin_sep}\n")
        for lp in paths:
            f.write(f"  {lp.title}\n")
            for mod in lp.modules:
                f.write(f"    - {mod.title} ({len(mod.units)} lessons)\n")
        f.write(f"\n{separator}\n\n")

        for lp in paths:
            f.write(f"{lp.title.upper()}\n")
            f.write(f"{separator}\n")
            f.write(f"Source: {lp.url}\n\n")
            for mod in lp.modules:
                f.write(f"{mod.title}\n")
                f.write(f"{thin_sep}\n\n")
                for unit in mod.units:
                    f.write(f"  {unit.title}\n")
                    f.write(f"  {'·' * (len(unit.title) + 2)}\n\n")
                    # Indent each content line
                    for line in unit.content.splitlines():
                        f.write(f"  {line}\n")
                    f.write(f"\n{thin_sep}\n\n")
            f.flush()

    _atomic_write_text(tmp, output, _write)
    logger.info("TXT written to %s", output)


def write_docx(
    paths: list[LearningPath],
    output: Path,
    title: str = "AZ-305 Course Material",
) -> None:
    """Write to a Word document (.docx)."""
    try:
        from docx import Document
        from docx.shared import Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
    except ImportError:
        logger.error(
            "python-docx is not installed. Run: pip install python-docx"
        )
        raise SystemExit(1)

    doc = Document()

    # Document title
    heading = doc.add_heading(title, level=0)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Generation date
    date_para = doc.add_paragraph(
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    )
    date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    # Table of contents (text-based)
    doc.add_heading("Table of Contents", level=1)
    for lp in paths:
        toc_path = doc.add_paragraph(style="List Bullet")
        toc_path.add_run(lp.title).bold = True
        for mod in lp.modules:
            toc_mod = doc.add_paragraph(style="List Bullet 2")
            toc_mod.add_run(f"{mod.title} ({len(mod.units)} lessons)")

    doc.add_page_break()

    # Content
    for lp in paths:
        doc.add_heading(lp.title, level=1)
        source_para = doc.add_paragraph()
        source_para.add_run("Source: ").bold = True
        source_para.add_run(lp.url)

        for mod in lp.modules:
            doc.add_heading(mod.title, level=2)

            for unit in mod.units:
                doc.add_heading(unit.title, level=3)
                # Split content on blank lines — each block becomes its own paragraph
                for block in unit.content.split("\n\n"):
                    block = block.strip()
                    if block:
                        doc.add_paragraph(block)

                doc.add_paragraph("─" * 40)

    tmp = output.with_suffix(".tmp")
    doc.save(tmp)
    tmp.replace(output)
    logger.info("DOCX written to %s", output)


PDF_STYLE = """
    body { font-family: -apple-system, Segoe UI, Helvetica, Arial, sans-serif;
           font-size: 11pt; line-height: 1.5; color: #1b1b1b; margin: 2em; }
    h1 { font-size: 22pt; border-bottom: 2px solid #0067b8; padding-bottom: .2em; }
    h2 { font-size: 17pt; color: #0067b8; margin-top: 1.4em;
         page-break-before: always; }
    h2:first-of-type { page-break-before: avoid; }
    h3 { font-size: 14pt; margin-top: 1em; }
    h4 { font-size: 12pt; margin-top: .8em; color: #333; }
    .meta { color: #666; font-size: 9pt; }
    .source { color: #666; font-size: 9pt; word-break: break-all; }
    .toc a, .content { color: inherit; }
    pre { white-space: pre-wrap; font-family: inherit; }
    hr { border: none; border-top: 1px solid #ddd; margin: 1em 0; }
"""


def _html_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def _render_html(paths: list[LearningPath], title: str) -> str:
    """Build a single self-contained HTML document for PDF rendering."""
    out: list[str] = [
        "<!DOCTYPE html><html><head><meta charset='utf-8'>",
        f"<title>{_html_escape(title)}</title>",
        f"<style>{PDF_STYLE}</style></head><body>",
        f"<h1>{_html_escape(title)}</h1>",
        f"<p class='meta'>Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>",
        "<h2 style='page-break-before:avoid'>Table of Contents</h2><ul class='toc'>",
    ]
    for lp in paths:
        out.append(f"<li><strong>{_html_escape(lp.title)}</strong><ul>")
        for mod in lp.modules:
            out.append(
                f"<li>{_html_escape(mod.title)} ({len(mod.units)} lessons)</li>"
            )
        out.append("</ul></li>")
    out.append("</ul>")

    for lp in paths:
        out.append(f"<h2>{_html_escape(lp.title)}</h2>")
        out.append(f"<p class='source'>Source: {_html_escape(lp.url)}</p>")
        for mod in lp.modules:
            out.append(f"<h3>{_html_escape(mod.title)}</h3>")
            for unit in mod.units:
                out.append(f"<h4>{_html_escape(unit.title)}</h4>")
                out.append(f"<pre class='content'>{_html_escape(unit.content)}</pre>")
                out.append("<hr>")
    out.append("</body></html>")
    return "".join(out)


def write_pdf(
    paths: list[LearningPath],
    output: Path,
    title: str = "AZ-305 Course Material",
) -> None:
    """Write to PDF (.pdf) by rendering HTML in headless Chromium."""
    html = _render_html(paths, title)
    tmp = output.with_suffix(".tmp")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_content(html, wait_until="load")
        page.pdf(
            path=str(tmp),
            format="A4",
            print_background=True,
            margin={"top": "1.5cm", "bottom": "1.5cm", "left": "1.2cm", "right": "1.2cm"},
        )
        browser.close()
    tmp.replace(output)
    logger.info("PDF written to %s", output)


def write_output(
    paths: list[LearningPath],
    output: Path,
    fmt: str,
    title: str,
) -> None:
    """Dispatcher: select the correct writer based on the specified format."""
    writers = {
        "markdown": write_markdown,
        "txt": write_txt,
        "docx": write_docx,
        "pdf": write_pdf,
    }
    writers[fmt](paths, output, title)


# ── Config & CLI ──────────────────────────────────────────────


def load_config(path: str) -> dict:
    """Load YAML configuration. Returns empty dict if file does not exist."""
    try:
        with open(path) as f:
            return yaml.safe_load(f) or {}
    except FileNotFoundError:
        logger.debug("Config file '%s' not found, using defaults", path)
        return {}


# ── Learning path picker ──────────────────────────────────────


def fetch_learning_paths(locale: str = DEFAULT_LOCALE) -> list[dict]:
    """Fetch every learning path from the Microsoft Learn catalog API.

    Uses a real headless Chromium page because a plain HTTP request is blocked
    (HTTP 403) by Microsoft Learn's bot protection.
    """
    url = CATALOG_API.format(locale=locale)
    logger.info("Fetching learning path catalog (%s)…", locale)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(user_agent=USER_AGENT)
        page = context.new_page()
        try:
            if not safe_goto(page, url):
                return []
            raw = page.evaluate("() => document.body.innerText")
        finally:
            browser.close()

    try:
        data = json.loads(raw)
    except (json.JSONDecodeError, TypeError):
        logger.error("Could not parse catalog response as JSON.")
        return []

    paths = data.get("learningPaths", [])
    paths.sort(key=lambda lp: (lp.get("title") or "").lower())
    logger.info("Catalog loaded: %d learning paths available.", len(paths))
    return paths


def pick_learning_paths(paths: list[dict]) -> tuple[list[str], str | None]:
    """Interactively let the user search and select one or more learning paths.

    Returns the selected URLs and the title of the first selection (used as a
    default document title). Returns ([], None) if the user cancels.
    """
    if not paths:
        print("\nNo learning paths could be retrieved from the catalog.")
        return [], None

    print()
    print("Learning path picker")
    print("=" * 42)
    print(f"{len(paths)} learning paths available.")

    def _ask(prompt: str) -> str:
        try:
            return input(prompt).strip()
        except (EOFError, KeyboardInterrupt):
            print("\nCancelled.")
            raise SystemExit(0)

    while True:
        term = _ask(
            "\nSearch by keyword (e.g. 'azure fundamentals'), or Enter to list all: "
        ).lower()

        if term:
            matches = [
                lp for lp in paths
                if term in (lp.get("title") or "").lower()
                or term in (lp.get("summary") or "").lower()
            ]
        else:
            matches = paths

        if not matches:
            print(f"  No learning paths match '{term}'. Try another keyword.")
            continue

        if len(matches) > PICKER_RESULT_LIMIT:
            print(
                f"  {len(matches)} matches — showing first {PICKER_RESULT_LIMIT}. "
                "Refine your keyword to narrow the list."
            )
            shown = matches[:PICKER_RESULT_LIMIT]
        else:
            shown = matches

        print()
        for idx, lp in enumerate(shown, 1):
            print(f"  {idx:>3}  {lp.get('title', '(untitled)')}")
        print()

        choice = _ask(
            "Enter number(s) to select (comma-separated), "
            "or 's' to search again, 'q' to quit: "
        ).lower()

        if choice in ("q", "quit"):
            print("Cancelled.")
            return [], None
        if choice in ("", "s", "search"):
            continue

        try:
            indices = [int(c) for c in choice.replace(" ", "").split(",") if c]
        except ValueError:
            print("  Invalid input — enter numbers like '1' or '1,3,5'.")
            continue

        selected = [shown[i - 1] for i in indices if 1 <= i <= len(shown)]
        if not selected:
            print("  No valid numbers in range — try again.")
            continue

        urls = [lp["url"] for lp in selected if lp.get("url")]
        if not urls:
            print("  Selected paths have no usable URL — try again.")
            continue

        print("\n-> Selected:")
        for lp in selected:
            print(f"     • {lp.get('title')}")
        print()
        return urls, selected[0].get("title")


def ask_format_interactively() -> str:
    """Ask the user interactively for an output format when --format is not provided."""
    options = {"1": "markdown", "2": "docx", "3": "txt", "4": "pdf"}
    print()
    print("Choose an output format")
    print("=" * 42)
    print("  1  ->  Markdown  (.md)")
    print("  2  ->  Word      (.docx)")
    print("  3  ->  Plain text (.txt)")
    print("  4  ->  PDF       (.pdf)")
    print()

    while True:
        try:
            choice = input(
                "Enter 1, 2, 3 or 4 (or type 'markdown', 'docx', 'txt', 'pdf'): "
            ).strip().lower()
        except (EOFError, KeyboardInterrupt):
            print("\nCancelled.")
            raise SystemExit(0)

        # Accept both number and name
        if choice in options:
            selected = options[choice]
            print(f"-> Selected format: {selected}\n")
            return selected
        if choice in VALID_FORMATS:
            print(f"-> Selected format: {choice}\n")
            return choice

        print(f"  Invalid choice '{choice}'. Choose 1, 2, 3, 4 or type the format name.")


def resolve_format(raw: str | None) -> str:
    """Validate the format. If None: ask interactively. If invalid: error + exit."""
    if raw is None:
        return ask_format_interactively()
    fmt = raw.strip().lower()
    if fmt not in VALID_FORMATS:
        print(
            f"\nError: '{raw}' is not a valid format.\n"
            f"Choose one of: {', '.join(VALID_FORMATS)}\n\n"
            f"Usage:  python scrape_all_v2.py --format markdown\n"
            f"        python scrape_all_v2.py --format docx\n"
            f"        python scrape_all_v2.py --format txt\n"
            f"        python scrape_all_v2.py --format pdf\n",
            file=sys.stderr,
        )
        raise SystemExit(1)
    return fmt


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description=(
            "Microsoft Learn Scraper — downloads learning paths as Markdown, DOCX, TXT, or PDF.\n\n"
            "  --format is required. If omitted you will be prompted interactively.\n"
            "  When no URLs are given, an interactive learning-path picker is shown."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python scrape_all_v2.py --format markdown            # pick path(s) interactively\n"
            "  python scrape_all_v2.py --format pdf --pick           # force the picker\n"
            "  python scrape_all_v2.py --format docx --output MyCourse --urls URL\n"
            "  python scrape_all_v2.py --format txt --config other.yaml -v\n"
        ),
    )
    p.add_argument(
        "--format",
        dest="fmt",
        choices=VALID_FORMATS,
        metavar="FORMAT",
        help=f"Output format — required. Choose from: {', '.join(VALID_FORMATS)}",
    )
    p.add_argument("--config", default="config.yaml", help="YAML config file (default: config.yaml)")
    p.add_argument("--urls", nargs="+", help="Learning path URLs (overrides config)")
    p.add_argument(
        "--pick",
        action="store_true",
        help="Pick learning paths interactively from the Microsoft Learn catalog "
             "(runs automatically when no URLs are provided)",
    )
    p.add_argument(
        "--locale",
        default=DEFAULT_LOCALE,
        help=f"Catalog locale for the picker (default: {DEFAULT_LOCALE}, e.g. nl-nl)",
    )
    p.add_argument("--output", help="Output filename without extension, or full path (extension is auto-determined)")
    p.add_argument("--title", help="Title of the output document")
    p.add_argument("--headed", action="store_true", help="Show the browser window (for debugging)")
    p.add_argument("-v", "--verbose", action="store_true", help="Verbose logging (debug level)")
    return p.parse_args()


# ── Main ──────────────────────────────────────────────────────


def main() -> None:
    args = parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler("scraper.log", encoding="utf-8"),
        ],
    )

    # ── Determine format (required — interactive if not provided) ─────────
    fmt = resolve_format(args.fmt)

    # ── Load config — CLI args take priority over YAML ────────────────────
    cfg = load_config(args.config)
    urls = args.urls or cfg.get("paths", [])
    title = args.title or cfg.get("title", "AZ-305 Course Material")

    # ── No URLs (or --pick): let the user choose from the live catalog ────
    if args.pick or not urls:
        catalog = fetch_learning_paths(args.locale)
        urls, picked_title = pick_learning_paths(catalog)
        if not urls:
            logger.info("No learning paths selected — nothing to do.")
            raise SystemExit(0)
        # Use the picked path's title as the document title unless one was set.
        if not args.title and not cfg.get("title") and picked_title:
            title = picked_title

    # Determine output filename — extension always follows the format
    raw_output = args.output or cfg.get("output", "output")
    output_path = Path(raw_output).with_suffix(FORMAT_EXTENSIONS[fmt])

    logger.info(
        "Starting scraper — format: %s | %d learning paths | output: %s",
        fmt, len(urls), output_path,
    )

    stats = ScrapeStats()
    seen_urls: set[str] = set()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not args.headed)
        context = browser.new_context(user_agent=USER_AGENT)
        page = context.new_page()

        # Handle cookie banner on first navigation
        if urls:
            safe_goto(page, urls[0])
            dismiss_cookie_banner(page)

        results: list[LearningPath] = []
        for url in urls:
            lp = scrape_path(page, url, stats, seen_urls)
            if lp and lp.modules:
                results.append(lp)

        browser.close()

    if results:
        write_output(results, output_path, fmt=fmt, title=title)
    else:
        logger.warning("No results to write.")

    # Summary
    logger.info("Done!\n%s", stats.summary())


if __name__ == "__main__":
    main()
