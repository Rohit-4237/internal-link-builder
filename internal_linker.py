#!/usr/bin/env python3
"""
internal_linker.py
-------------------
Turn a keyword -> URL map into real hyperlinks inside Word (.docx) documents.

Point it at a spreadsheet (.xlsx or .csv) where one column is a phrase and
another is the URL that phrase should point to, then at a .docx file (or a
whole folder of them), and it inserts a live hyperlink at each phrase, up to
a configurable number of times per phrase, leaving everything else - text,
formatting, existing links - untouched.

Usage:
    python internal_linker.py --map keywords.csv --docx article.docx
    python internal_linker.py -m keywords.xlsx -d docs/ --recursive --out-dir linked/
    python internal_linker.py -m keywords.xlsx -d article.docx --dry-run

Requirements:
    pip install python-docx openpyxl
"""

import argparse
import csv
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

try:
    import openpyxl
except ImportError:
    sys.exit("Missing dependency: pip install openpyxl")

try:
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.opc.constants import RELATIONSHIP_TYPE
    from docx.shared import RGBColor
    from docx.text.run import Run
except ImportError:
    sys.exit("Missing dependency: pip install python-docx")


# --------------------------------------------------------------------------- #
# Keyword rules: one phrase, one destination, one usage cap
# --------------------------------------------------------------------------- #
@dataclass
class LinkRule:
    phrase: str
    url: str
    occurrence_cap: int = 1
    times_applied: int = 0

    @property
    def has_capacity(self):
        return self.times_applied < self.occurrence_cap


def _looks_like_header(row):
    """A header row's second cell won't look like an actual URL or path."""
    if not row or len(row) < 2 or row[1] is None:
        return False
    candidate = str(row[1]).strip().lower()
    if not candidate:
        return False
    return not (candidate.startswith("http") or candidate.startswith("/") or candidate.startswith("#"))


def _read_rows(path: Path):
    suffix = path.suffix.lower()
    if suffix == ".csv":
        with open(path, newline="", encoding="utf-8-sig") as fh:
            return [row for row in csv.reader(fh) if any(cell.strip() for cell in row if cell)]
    if suffix in (".xlsx", ".xlsm"):
        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
        sheet = workbook.active
        return [list(row) for row in sheet.iter_rows(values_only=True) if row and any(row)]
    raise ValueError(f"Unsupported map format: '{suffix}' (use .xlsx or .csv)")


def load_rules(path, case_sensitive=False, default_cap=1):
    """Read (phrase, url[, cap]) rows into LinkRule objects, skipping a header row if present."""
    path = Path(path)
    rows = _read_rows(path)
    if not rows:
        sys.exit(f"No rows found in {path}")
    if _looks_like_header(rows[0]):
        rows = rows[1:]

    rules = []
    seen_keys = set()
    for line_no, row in enumerate(rows, start=1):
        phrase = str(row[0]).strip() if row and row[0] not in (None, "") else ""
        url = str(row[1]).strip() if len(row) > 1 and row[1] not in (None, "") else ""
        if not phrase or not url:
            print(f"  (skipping row {line_no}: needs both a phrase and a URL - got {phrase!r} -> {url!r})")
            continue

        cap = default_cap
        if len(row) > 2 and row[2] not in (None, ""):
            try:
                cap = max(1, int(str(row[2]).strip()))
            except ValueError:
                print(f"  (row {line_no}: '{row[2]}' isn't a whole number, using the default cap of {default_cap})")

        key = phrase if case_sensitive else phrase.lower()
        if key in seen_keys:
            print(f"  (duplicate phrase '{phrase}' ignored - keeping the first one)")
            continue
        seen_keys.add(key)
        rules.append(LinkRule(phrase=phrase, url=url, occurrence_cap=cap))

    if not rules:
        sys.exit("No usable phrase/URL rows were found in the map.")
    return rules


# --------------------------------------------------------------------------- #
# Matching engine: one alternation regex over every phrase that still has
# capacity, longest phrase first so overlapping phrases resolve sensibly,
# rebuilt only when a phrase gets used up.
# --------------------------------------------------------------------------- #
class LinkEngine:
    def __init__(self, rules, case_sensitive=False):
        self.rules = rules
        self.case_sensitive = case_sensitive
        self._by_key = {(r.phrase if case_sensitive else r.phrase.lower()): r for r in rules}
        self._pattern = None

    def _compiled_pattern(self):
        if self._pattern is None:
            active = sorted((r for r in self.rules if r.has_capacity), key=lambda r: len(r.phrase), reverse=True)
            if not active:
                return None
            alternation = "|".join(re.escape(r.phrase) for r in active)
            flags = re.UNICODE if self.case_sensitive else re.UNICODE | re.IGNORECASE
            self._pattern = re.compile(rf"\b(?:{alternation})\b", flags)
        return self._pattern

    def next_match(self, text):
        """The single earliest still-eligible match in `text`, or None."""
        pattern = self._compiled_pattern()
        if pattern is None:
            return None
        for m in pattern.finditer(text):
            rule = self._by_key.get(m.group(0) if self.case_sensitive else m.group(0).lower())
            if rule is not None:
                return m.start(), m.end(), rule
        return None

    def consume(self, rule):
        rule.times_applied += 1
        if not rule.has_capacity:
            self._pattern = None  # next compile will drop this phrase

    def reset(self):
        for rule in self.rules:
            rule.times_applied = 0
        self._pattern = None


# --------------------------------------------------------------------------- #
# Formatting snapshots: lift a run's visible style into plain data so it can
# be re-applied to brand-new runs via python-docx's own Font API, instead of
# hand-copying raw XML tags.
# --------------------------------------------------------------------------- #
@dataclass
class StyleSnapshot:
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    strike: Optional[bool] = None
    subscript: Optional[bool] = None
    superscript: Optional[bool] = None
    all_caps: Optional[bool] = None
    small_caps: Optional[bool] = None
    name: Optional[str] = None
    size: Optional[object] = None
    color_rgb: Optional[object] = None
    highlight_color: Optional[object] = None


def capture_style(run) -> StyleSnapshot:
    font = run.font
    snapshot = StyleSnapshot(
        bold=run.bold, italic=run.italic, underline=run.underline,
        strike=font.strike, subscript=font.subscript, superscript=font.superscript,
        all_caps=font.all_caps, small_caps=font.small_caps, name=font.name,
    )
    snapshot.size = font.size
    snapshot.highlight_color = font.highlight_color
    if font.color is not None and font.color.type is not None:
        snapshot.color_rgb = font.color.rgb
    return snapshot


def apply_style(run, style: StyleSnapshot):
    if style is None:
        return
    for attr in ("bold", "italic", "underline"):
        value = getattr(style, attr)
        if value is not None:
            setattr(run, attr, value)
    font = run.font
    for attr in ("strike", "subscript", "superscript", "all_caps", "small_caps", "name", "size", "highlight_color"):
        value = getattr(style, attr)
        if value is not None:
            setattr(font, attr, value)
    if style.color_rgb is not None:
        font.color.rgb = style.color_rgb


# --------------------------------------------------------------------------- #
# Building the replacement XML: a hyperlink run, flanked by whatever plain
# text survives on either side of the match.
# --------------------------------------------------------------------------- #
def _new_styled_run(paragraph, text, style):
    element = OxmlElement("w:r")
    run = Run(element, paragraph)
    if text:
        run.text = text
    apply_style(run, style)
    return run


def _make_hyperlink(paragraph, url, text, style, link_color, underline_links):
    rel_id = paragraph.part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rel_id)
    inner = _new_styled_run(paragraph, text, style)
    if link_color:
        inner.font.color.rgb = RGBColor.from_string(link_color)
    if underline_links:
        inner.underline = True
    hyperlink.append(inner._r)
    return hyperlink


# --------------------------------------------------------------------------- #
# Cross-run aware scanning: a phrase may straddle two or more <w:r> runs
# (half bold, half not, etc.) - map the paragraph's flattened text back to
# the runs it came from so a match like that can still be found and linked.
# --------------------------------------------------------------------------- #
def _flatten_runs(paragraph):
    pieces, spans, cursor = [], [], 0
    for index, run in enumerate(paragraph.runs):
        chunk = run.text or ""
        if chunk:
            spans.append((cursor, cursor + len(chunk), index))
            pieces.append(chunk)
            cursor += len(chunk)
    return "".join(pieces), spans


def _runs_under(spans, start, end):
    hits = []
    for span_start, span_end, run_index in spans:
        if span_end <= start or span_start >= end:
            continue
        hits.append((run_index, max(start, span_start) - span_start, min(end, span_end) - span_start))
    return hits


def _rewrite_span(paragraph, matched_text, overlap, url, link_color, underline_links):
    runs = paragraph.runs
    first_index, first_offset, _ = overlap[0]
    last_index, _, last_offset = overlap[-1]
    first_run, last_run = runs[first_index], runs[last_index]

    lead_style = capture_style(first_run)
    lead_text = first_run.text[:first_offset]
    tail_text = last_run.text[last_offset:]

    # Runs entirely swallowed by the match (strictly between the first and
    # last touched run) contribute nothing and are simply dropped.
    for run_index, _, _ in overlap[1:-1]:
        element = runs[run_index]._r
        element.getparent().remove(element)

    first_element = first_run._r
    first_run.text = lead_text
    hyperlink = _make_hyperlink(paragraph, url, matched_text, lead_style, link_color, underline_links)
    first_element.addnext(hyperlink)
    if not lead_text:
        first_element.getparent().remove(first_element)

    if first_index == last_index:
        if tail_text:
            hyperlink.addnext(_new_styled_run(paragraph, tail_text, lead_style)._r)
        return

    last_element = last_run._r
    if tail_text:
        last_run.text = tail_text  # keep the tail's own original formatting
        hyperlink.addnext(last_element)  # relocates the element
    else:
        last_element.getparent().remove(last_element)


def link_paragraph(paragraph, engine, link_color, underline_links):
    """Apply every eligible match in this paragraph, one at a time (the run
    layout changes after each rewrite, so the text is re-flattened each pass)."""
    applied = 0
    while True:
        text, spans = _flatten_runs(paragraph)
        if not text:
            return applied
        found = engine.next_match(text)
        if found is None:
            return applied
        start, end, rule = found
        overlap = _runs_under(spans, start, end)
        if not overlap:
            return applied
        _rewrite_span(paragraph, text[start:end], overlap, rule.url, link_color, underline_links)
        engine.consume(rule)
        applied += 1


# --------------------------------------------------------------------------- #
# Walking a whole document: body paragraphs, table cells (nested tables
# included), and - opt-in - headers/footers.
# --------------------------------------------------------------------------- #
def _table_paragraphs(table):
    for row in table.rows:
        for cell in row.cells:
            yield from cell.paragraphs
            for nested in cell.tables:
                yield from _table_paragraphs(nested)


def iter_document_paragraphs(document, with_headers_footers=False):
    yield from document.paragraphs
    for table in document.tables:
        yield from _table_paragraphs(table)

    if with_headers_footers:
        for section in document.sections:
            for part in (section.header, section.footer,
                         section.first_page_header, section.first_page_footer,
                         section.even_page_header, section.even_page_footer):
                if part is None:
                    continue
                try:
                    yield from part.paragraphs
                    for table in part.tables:
                        yield from _table_paragraphs(table)
                except Exception:
                    continue


def link_document(document, engine, link_color, underline_links, with_headers_footers=False):
    total = 0
    for paragraph in iter_document_paragraphs(document, with_headers_footers):
        if not any(rule.has_capacity for rule in engine.rules):
            break
        total += link_paragraph(paragraph, engine, link_color, underline_links)
    return total


def discover_docx_files(target, recursive):
    target = Path(target)
    if target.is_file():
        return [target]
    pattern = "**/*.docx" if recursive else "*.docx"
    return sorted(p for p in target.glob(pattern) if not p.name.startswith("~$"))


# --------------------------------------------------------------------------- #
# Reporting
# --------------------------------------------------------------------------- #
def export_summary(path, run_log):
    path = Path(path)
    rows = [
        {
            "file": str(docx_path),
            "phrase": rule.phrase,
            "url": rule.url,
            "times_linked": rule.times_applied,
            "occurrence_cap": rule.occurrence_cap,
        }
        for docx_path, rules_snapshot, _ in run_log
        for rule in rules_snapshot
    ]
    if path.suffix.lower() == ".json":
        path.write_text(json.dumps(rows, indent=2), encoding="utf-8")
        return
    with open(path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=["file", "phrase", "url", "times_linked", "occurrence_cap"])
        writer.writeheader()
        writer.writerows(rows)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def _snapshot(rules):
    return [LinkRule(r.phrase, r.url, r.occurrence_cap, r.times_applied) for r in rules]


def main(argv=None):
    parser = argparse.ArgumentParser(description="Insert hyperlinks into .docx file(s) from a keyword/URL map.")
    parser.add_argument("-m", "--map", required=True, help="Path to the .xlsx or .csv map (phrase, URL, optional cap).")
    parser.add_argument("-d", "--docx", required=True, help="A .docx file, or a folder to batch-process.")
    parser.add_argument("-o", "--output", help="Output path (single-file mode only). Default: <name>_linked.docx.")
    parser.add_argument("--out-dir", help="Directory to write linked files into (batch mode). Default: next to each source file.")
    parser.add_argument("--recursive", action="store_true", help="When --docx is a folder, search it recursively.")
    parser.add_argument("--match-case", action="store_true", help="Case-sensitive matching (default: case-insensitive).")
    parser.add_argument("--max-occurrences", type=int, default=1,
                         help="Default cap on links per phrase per document (default: 1). A map row's 3rd column overrides this.")
    parser.add_argument("--per-file-limits", action="store_true",
                         help="In batch mode, reset every phrase's cap for each file (default: caps apply once, across the whole batch).")
    parser.add_argument("--include-headers-footers", action="store_true", help="Also link inside headers and footers.")
    parser.add_argument("--color", default="0563C1", help="Link colour as hex, or 'none' to leave colour untouched. Default: 0563C1.")
    parser.add_argument("--no-underline", action="store_true", help="Don't underline the inserted links.")
    parser.add_argument("--report", help="Write a .csv or .json summary of what was linked.")
    parser.add_argument("--dry-run", action="store_true", help="Show what would be linked without saving anything.")
    parser.add_argument("-q", "--quiet", action="store_true", help="Only print warnings and the final summary.")
    args = parser.parse_args(argv)

    map_path = Path(args.map)
    docx_target = Path(args.docx)
    if not map_path.exists():
        sys.exit(f"Map file not found: {map_path}")
    if not docx_target.exists():
        sys.exit(f"Docx path not found: {docx_target}")

    link_color = None if args.color.lower() == "none" else args.color.lstrip("#")
    underline_links = not args.no_underline

    def log(message):
        if not args.quiet:
            print(message)

    log(f"Loading map from {map_path} ...")
    rules = load_rules(map_path, case_sensitive=args.match_case, default_cap=args.max_occurrences)
    log(f"  {len(rules)} phrase/URL rule(s) loaded.\n")
    engine = LinkEngine(rules, case_sensitive=args.match_case)

    files = discover_docx_files(docx_target, args.recursive)
    if not files:
        sys.exit(f"No .docx files found at {docx_target}")
    if args.output and len(files) > 1:
        sys.exit("--output only works with a single input file; use --out-dir for batch mode.")

    run_log = []
    grand_total = 0
    for docx_path in files:
        if args.per_file_limits:
            engine.reset()

        log(f"Processing {docx_path} ...")
        document = Document(str(docx_path))
        total = link_document(document, engine, link_color, underline_links, args.include_headers_footers)
        grand_total += total

        linked_now = [r.phrase for r in rules if r.times_applied > 0]
        log(f"  Linked {total} occurrence(s) this file: {', '.join(linked_now) if linked_now else '(none)'}")

        if not args.dry_run:
            if args.output:
                out_path = Path(args.output)
            elif args.out_dir:
                out_dir = Path(args.out_dir)
                out_dir.mkdir(parents=True, exist_ok=True)
                out_path = out_dir / (docx_path.stem + "_linked.docx")
            else:
                out_path = docx_path.with_name(docx_path.stem + "_linked.docx")
            document.save(str(out_path))
            log(f"  Saved: {out_path}\n")

        run_log.append((docx_path, _snapshot(rules), total))

    never_matched = [r.phrase for r in rules if r.times_applied == 0]
    print(f"\nDone. {grand_total} link(s) inserted across {len(files)} file(s).")
    if never_matched:
        print(f"Never matched ({len(never_matched)}): {', '.join(never_matched)}")

    if args.report:
        export_summary(args.report, run_log)
        print(f"Report written: {args.report}")

    if args.dry_run:
        print("\nDry run - no files were written.")


if __name__ == "__main__":
    main()
