# Internal Link Builder

A Python tool for SEOs and content teams: give it a keyword → URL map and a
Word document (or a whole folder of them), and it hyperlinks the right
phrases automatically, directly inside the `.docx` — formatting preserved,
no manual find-and-replace.

## What it does

- Reads a keyword map from **`.xlsx` or `.csv`**: column A = phrase, column B
  = URL, column C = optional per-phrase link limit. A header row is detected
  automatically.
- Links each phrase up to its limit (**default: once per document**) —
  never more.
- Whole-word, case-insensitive matching by default (`Shoes` matches `shoes`,
  not `shoescraper`).
- **Finds matches that span multiple text runs** — e.g. a phrase that's half
  bold, half plain, or split up inside the document's underlying XML.
- Never double-links text that's already a hyperlink.
- Preserves the original formatting (bold, italic, size, colour, etc.) on
  both sides of a new link.
- When two phrases could match the same spot, the **longer one wins**
  (`data analysis` is linked instead of just `data`).
- Walks body paragraphs, tables (including nested tables), and — optionally
  — headers and footers.
- Can process a **single file or a whole folder** of `.docx` files in one
  run.
- Can write a CSV/JSON **report** of what got linked, how many times, and
  what was never matched.

## Install

```bash
pip install -r requirements.txt
```

## Usage

```bash
# Single file
python internal_linker.py --map examples/map.csv --docx article.docx

# A whole folder, recursively, with a custom output directory
python internal_linker.py -m map.xlsx -d docs/ --recursive --out-dir linked/

# Preview without writing anything
python internal_linker.py -m map.xlsx -d article.docx --dry-run

# Also link inside headers/footers, and save a report
python internal_linker.py -m map.xlsx -d article.docx --include-headers-footers --report report.csv
```

This writes `article_linked.docx` next to the original by default.

### Options

| Flag | Description |
|---|---|
| `-m, --map` | Path to the `.xlsx` or `.csv` keyword map (required). |
| `-d, --docx` | Path to a `.docx` file **or a folder** to batch-process (required). |
| `-o, --output` | Custom output path (single-file mode only). Default: `<name>_linked.docx`. |
| `--out-dir` | Directory to write linked files into (batch mode). Default: next to each source file. |
| `--recursive` | When `--docx` is a folder, search it recursively. |
| `--match-case` | Case-sensitive matching. |
| `--max-occurrences N` | Default max times a phrase may be linked per document (default `1`). Overridable per row via column C of the map. |
| `--per-file-limits` | In batch mode, reset every phrase's counter for each file (default: limits apply once, across the whole batch). |
| `--include-headers-footers` | Also link inside headers and footers. |
| `--color HEX` | Link colour, e.g. `0563C1` (default), or `none` to keep the original colour. |
| `--no-underline` | Don't underline the new links. |
| `--report PATH` | Write a `.csv` or `.json` summary (file, phrase, URL, times linked). |
| `--dry-run` | Report what would be linked without saving anything. |
| `-q, --quiet` | Only print warnings/errors and the final summary. |

### Example map (`examples/map.csv`)

| Keyword | URL | Max |
|---|---|---|
| technical seo | https://example.com/technical-seo | |
| knowledge graph | https://example.com/knowledge-graph | |
| internal linking | /guides/internal-linking | 2 |
| data analysis | https://example.com/data-analysis | |

URLs can be absolute (`https://…`) or relative (`/guides/…`). Leave `Max`
blank to use the `--max-occurrences` default (1).

## How matching works

Every eligible phrase is compiled into a single whole-word regex, longest
phrase first, so overlapping phrases resolve sensibly at the same spot. A
phrase's match can straddle more than one text run inside the document's
XML (for example, half of it bold and half not) — the paragraph's runs are
flattened into plain text with a position map back to each run, so a match
like that is still found, and the surrounding runs are rebuilt with each
side's own original formatting intact. Once a phrase hits its occurrence
cap it drops out of the matcher for the rest of the run.

## Good to know / limitations

- "First occurrence" is resolved body paragraphs first, then tables, then
  (if enabled) headers/footers.
- Always review the output. Automation gets you most of the way there; a
  human should still sanity-check anchor relevance.

## Tests

```bash
pip install -r requirements-dev.txt
pytest tests/ -v
```
