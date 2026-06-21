# Microsoft Learn Scraper

**Author:** Matthew Gonzalez
**Version:** 1.0.0
**Date:** 31-03-2025
**License:** MIT
**Website:** [mgntechvault.com](https://mgntechvault.com)

Download Microsoft Learn training paths and convert them to Markdown, DOCX, TXT, or PDF files.

Don't know the URL? Skip it — an interactive **learning path picker** lists every current Microsoft Learn learning path so you can search and select one or more.

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

### Learning path picker

When you run the scraper **without** providing URLs (no `--urls` and no `paths`
in `config.yaml`), it fetches the full Microsoft Learn catalog and shows an
interactive picker:

```bash
python scrape_all_v2.py --format pdf          # no URLs -> picker opens
python scrape_all_v2.py --format markdown --pick   # force the picker
```

In the picker you can:
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
- `--pick` — Pick learning paths interactively from the catalog (auto when no URLs)
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
