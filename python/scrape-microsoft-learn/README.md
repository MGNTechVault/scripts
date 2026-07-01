# Microsoft Learn Scraper

**Author:** Matthew Gonzalez
**Version:** 1.0.0
**Date:** 31-03-2025
**License:** MIT
**Website:** [mgntechvault.com](https://mgntechvault.com)

Download Microsoft Learn training paths and convert them to Markdown, DOCX, TXT, or PDF files.

Don't know the URL? Skip it — a **source menu** lets you choose between the live Microsoft Learn catalog (search & pick any learning path, or a whole **course** like MD-102 that bundles several paths) or the paths saved in `config.yaml`.

Intended for personal offline study using AI tools such as **Microsoft Copilot Notebook** and **Google NotebookLM** — export your learning paths and feed them directly to your AI notebook for smarter studying.

> **Note:** This tool is not intended to stress or flood Microsoft Learn (DDoS). Use it responsibly and only for your own study materials.

## Requirements

```bash
pip install playwright pyyaml python-docx
playwright install chromium
```

## Usage

Basic usage — choose format interactively:
```bash
python scrape_all_v2.py
```

Specify format directly:
```bash
python scrape_all_v2.py --format markdown
python scrape_all_v2.py --format docx
python scrape_all_v2.py --format txt
python scrape_all_v2.py --format pdf
```

### Source menu

When you run the scraper **without** `--urls`, a TUI menu asks where the
content should come from:

```
Where should the content come from?
==========================================
  1  ->  Live learning paths  (search & pick individual paths)
  2  ->  Live courses         (exam-style, bundles paths, e.g. MD-102)
  3  ->  Config file          (config.yaml, 6 path(s))
```

```bash
python scrape_all_v2.py --format pdf                # no URLs -> source menu opens
python scrape_all_v2.py --format markdown --pick    # skip menu, force learning-path picker
python scrape_all_v2.py --format pdf --courses      # skip menu, force course picker
```

- **1 — Live learning paths**: fetches the catalog and opens the picker (below).
- **2 — Live courses**: lists all 150+ Microsoft Learn courses. A course (e.g.
  MD-102) has no lessons of its own — it bundles several learning paths/modules,
  which are all scraped into one document. Each entry shows its **course number**
  and how many paths it expands to, e.g. `MD-102T00 — Microsoft 365 Endpoint
  Administrator  (8 paths/modules)`. The course number becomes the document
  title and the **output filename** (e.g. `MD-102T00.pdf`) unless you pass
  `--output`. Courses without a number fall back to a slug of the title.
- **3 — Config file**: uses the `paths` from `config.yaml` (only offered when the file has paths).

`--pick` / `--courses` skip the menu. In a non-interactive session (cron/pipe)
the menu is skipped and the config paths are used.

### Picker

When you pick from the live catalog (paths or courses) you get an interactive picker. In it you can:
- Type a keyword (e.g. `azure fundamentals`) to filter by title/summary
- Select one **or more** paths with comma-separated numbers (e.g. `1,3,5`)
- Use a different catalog language with `--locale` (default `en-us`, e.g. `--locale nl-nl`)

The title of the first selected path is used as the default document title.

> The catalog is fetched through the same headless browser the scraper uses, so
> it works even though Microsoft Learn's API blocks plain HTTP requests.

## Configuration

Edit `config.yaml` to set:
- `paths`: List of Microsoft Learn URLs to scrape
- `output`: Output filename (extension added automatically)
- `title`: Document title

Override via command line:
```bash
python scrape_all_v2.py --format markdown --output MyFile --config custom.yaml
python scrape_all_v2.py --format docx --urls URL1 URL2 URL3
```

## Options

- `--format` — Output format: markdown, docx, txt, or pdf (required)
- `--config` — Path to config file (default: config.yaml)
- `--urls` — Override URLs from command line
- `--pick` — Skip the source menu and go straight to the live learning-path picker
- `--courses` — Skip the source menu and go straight to the live course picker (bundles paths, e.g. MD-102)
- `--locale` — Catalog locale for the picker (default: en-us, e.g. nl-nl)
- `--output` — Output filename without extension
- `--title` — Document title
- `--headed` — Show browser window (for debugging)
- `-v, --verbose` — Verbose logging

## Output

The script generates:
- A single document with table of contents
- All modules and units organized hierarchically
- Timestamps and source URLs
- Log file: `scraper.log`

Supported formats preserve structure and content from Microsoft Learn.

## Notes

- Respects page load times and cookie banners
- Skips navigation elements and assessments
- Retries failed URLs with exponential backoff
- Writes atomically to prevent incomplete files
