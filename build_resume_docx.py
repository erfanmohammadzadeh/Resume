# -*- coding: utf-8 -*-
"""Build a resume .docx from the Markdown file next to this script.

Edit the Markdown, then run:

    python build_resume_docx.py

The .docx is rewritten from the Markdown. Nothing in the resume is hardcoded.

    python build_resume_docx.py --watch

rebuilds whenever you save the Markdown.

Markdown shape this script understands
--------------------------------------
# Your Name
**Headline under the name**

Contact line with optional [label](https://...) links

## Any section title

Plain paragraphs, including **bold** and [links](https://...).

Skills — a Markdown table (header row is skipped):

| Area | Tools |
|---|---|
| Languages | C++, Python |

Jobs — title line, italic date line, then bullets:

**Job title** — Company · City
*Month Year – Month Year*
- Bullet with **highlights**

Projects — title line, description, optional italic tools line:

**Project name** — Place · Date
What you did.
*Tools, libraries*

Education — bold title, then a details line, then an optional paragraph:

**Degree**
School · City · dates · GPA
Optional notes.

Certificates — title line, then a URL or [label](url):

**Certificate name** — Issuer · Date
https://...

Label lists:

- **Language:** level
"""

from __future__ import annotations

import argparse
import re
import sys
import time
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

NAVY = "1B365D"
ACCENT = "2C5F8A"
MUTED = "555555"
BODY = "222222"

ENTRY_HEADER = re.compile(r"^\*\*(.+?)\*\*(.*)$")
LEADING_SEP = re.compile(r"^(?:—|–|-)\s*")
INLINE = re.compile(
    r"\*\*(.+?)\*\*"
    r"|\*(.+?)\*"
    r"|\[([^\]]+)\]\(([^)]+)\)"
    r"|(https?://[^\s)]+)"
)
LABEL_BULLET = re.compile(r"^\*\*([^*]+?:\s*)\*\*\s*(.+)$")
TABLE_RULE = re.compile(r"^:?-{3,}:?$")
SOLE_LINK = re.compile(r"^\[([^\]]+)\]\(([^)]+)\)$")
SOLE_URL = re.compile(r"^https?://\S+$")


class Span:
    def __init__(self, text, bold=False, italic=False, url=None):
        self.text = text
        self.bold = bold
        self.italic = italic
        self.url = url


class Builder:
    def __init__(self):
        self.links = []  # (rid, url) in first-seen order

    def rid_for(self, url):
        for rid, existing in self.links:
            if existing == url:
                return rid
        rid = "rIdLink%d" % (len(self.links) + 1)
        self.links.append((rid, url))
        return rid


def xml_text(text):
    preserve = bool(text[:1].isspace() or text[-1:].isspace())
    attrs = ' xml:space="preserve"' if preserve else ""
    return "<w:t%s>%s</w:t>" % (attrs, escape(text))


def r(text, *, bold=False, italic=False, size=21, color=BODY, font="Calibri"):
    if not text:
        return ""
    rpr = [
        '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>' % (font, font, font),
        '<w:sz w:val="%d"/>' % size,
        '<w:szCs w:val="%d"/>' % size,
        '<w:color w:val="%s"/>' % color,
    ]
    if bold:
        rpr.append("<w:b/><w:bCs/>")
    if italic:
        rpr.append("<w:i/><w:iCs/>")
    return "<w:r><w:rPr>%s</w:rPr>%s</w:r>" % ("".join(rpr), xml_text(text))


def hyperlink(builder, url, text, *, size=18, color=ACCENT):
    rid = builder.rid_for(url)
    rpr = (
        '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
        '<w:sz w:val="%d"/><w:szCs w:val="%d"/>'
        '<w:color w:val="%s"/><w:u w:val="single"/>' % (size, size, color)
    )
    return (
        '<w:hyperlink r:id="%s" w:history="1">'
        "<w:r><w:rPr>%s</w:rPr>%s</w:r></w:hyperlink>"
        % (rid, rpr, xml_text(text))
    )


def p(runs, *, align="left", before=0, after=60, line=240, border=False):
    chunks = [piece for piece in runs if piece]
    if not chunks:
        return ""
    jc = '<w:jc w:val="%s"/>' % align
    sp = (
        '<w:spacing w:before="%d" w:after="%d" w:line="%d" w:lineRule="auto"/>'
        % (before, after, line)
    )
    bdr = ""
    if border:
        bdr = (
            '<w:pBdr><w:bottom w:val="single" w:sz="12" w:space="4" '
            'w:color="%s"/></w:pBdr>' % NAVY
        )
    return "<w:p><w:pPr>%s%s%s</w:pPr>%s</w:p>" % (jc, sp, bdr, "".join(chunks))


def heading(text):
    return p(
        [r(text.upper(), bold=True, size=22, color=NAVY)],
        before=200,
        after=80,
        border=True,
    )


def parse_inline(text):
    spans = []
    pos = 0
    for match in INLINE.finditer(text):
        if match.start() > pos:
            spans.append(Span(text[pos : match.start()]))
        if match.group(1) is not None:
            spans.append(Span(match.group(1), bold=True))
        elif match.group(2) is not None:
            spans.append(Span(match.group(2), italic=True))
        elif match.group(3) is not None:
            spans.append(Span(match.group(3), url=match.group(4).strip()))
        else:
            url = match.group(5)
            spans.append(Span(url, url=url))
        pos = match.end()
    if pos < len(text):
        spans.append(Span(text[pos:]))
    return [span for span in spans if span.text]


def render_spans(builder, spans, *, size=20, color=BODY, bold=False, italic=False):
    pieces = []
    for span in spans:
        if span.url:
            pieces.append(hyperlink(builder, span.url, span.text, size=size, color=ACCENT))
        else:
            pieces.append(
                r(
                    span.text,
                    bold=bold or span.bold,
                    italic=italic or span.italic,
                    size=size,
                    color=color,
                )
            )
    return pieces


def is_entry_header(line):
    return line.startswith("**") and "**" in line[2:]


def is_wrapped_italic(line):
    return len(line) >= 2 and line.startswith("*") and line.endswith("*") and not line.startswith("**")


def unwrap_italic(line):
    if is_wrapped_italic(line):
        return line[1:-1].strip()
    return line


def parse_header(line):
    match = ENTRY_HEADER.match(line.strip())
    if not match:
        return line.strip(), None
    title = match.group(1).strip()
    rest = LEADING_SEP.sub("", match.group(2).strip()).strip()
    return title, rest or None


def sole_link(line):
    match = SOLE_LINK.match(line.strip())
    if match:
        return match.group(1).strip(), match.group(2).strip()
    if SOLE_URL.match(line.strip()):
        return line.strip(), line.strip()
    return None


def strip_comments(text):
    return re.sub(r"<!--.*?-->", "", text, flags=re.S)


def parse_resume(text):
    lines = strip_comments(text).splitlines()
    name = ""
    headline = ""
    contact = []
    sections = []
    i = 0
    n = len(lines)

    def blank(idx):
        return not lines[idx].strip() or lines[idx].strip() == "---"

    while i < n and blank(i):
        i += 1
    if i < n and lines[i].startswith("# "):
        name = lines[i][2:].strip()
        i += 1
    while i < n and blank(i):
        i += 1
    if i < n and lines[i].strip().startswith("**") and not lines[i].startswith("##"):
        headline = unwrap_italic(lines[i].strip())
        if headline.startswith("**") and headline.endswith("**"):
            headline = headline[2:-2].strip()
        i += 1
    while i < n and not lines[i].startswith("##"):
        stripped = lines[i].strip()
        if stripped and stripped != "---":
            contact.append(stripped)
        i += 1

    while i < n:
        if not lines[i].startswith("## "):
            i += 1
            continue
        title = lines[i][3:].strip()
        i += 1
        body = []
        while i < n and not lines[i].startswith("## "):
            body.append(lines[i].rstrip())
            i += 1
        sections.append((title, body))

    if not name:
        raise SystemExit("Markdown needs a top heading: # Your Name")
    return {"name": name, "headline": headline, "contact": contact, "sections": sections}


def is_table_rule(cells):
    return bool(cells) and all(TABLE_RULE.match(cell.replace(" ", "")) for cell in cells)


def render_table(builder, rows_lines):
    rows = []
    for line in rows_lines:
        cells = [cell.strip() for cell in line.strip().strip("|").split("|")]
        if is_table_rule(cells):
            if rows:
                rows.pop()
            continue
        rows.append(cells)
    parts = []
    for cells in rows:
        if not cells or not any(cells):
            continue
        label = cells[0]
        value = " · ".join(cell for cell in cells[1:] if cell)
        runs = [r(label + ":  ", bold=True, size=20)]
        runs.extend(render_spans(builder, parse_inline(value), size=20))
        parts.append(p(runs, before=20, after=20))
    return parts


def render_bullet_list(builder, bullet_lines):
    texts = [line.strip()[2:].strip() for line in bullet_lines]
    labels = [LABEL_BULLET.match(text) for text in texts]
    parts = []
    if all(labels):
        for match in labels:
            parts.append(
                p(
                    [
                        r(match.group(1) + "  ", bold=True, size=20),
                        *render_spans(builder, parse_inline(match.group(2)), size=20),
                    ],
                    before=20,
                    after=20,
                )
            )
        return parts
    for text in texts:
        parts.append(bullet(builder, text))
    return parts


def bullet(builder, text):
    ppr = (
        "<w:pPr>"
        '<w:spacing w:before="20" w:after="40" w:line="240" w:lineRule="auto"/>'
        '<w:ind w:left="360" w:hanging="200"/>'
        "</w:pPr>"
    )
    mark = r("•  ", size=20, color=ACCENT)
    runs = "".join(render_spans(builder, parse_inline(text), size=20))
    return "<w:p>%s%s%s</w:p>" % (ppr, mark, runs)


def render_entry(builder, lines, first):
    raw = [line.strip() for line in lines if line.strip() and line.strip() != "---"]
    if not raw:
        return []
    title, meta = parse_header(raw[0])
    rest = raw[1:]
    bullets = [line[2:].strip() for line in rest if line.startswith("- ")]
    other = [line for line in rest if not line.startswith("- ")]

    date = None
    if bullets and other and is_wrapped_italic(other[0]):
        date = unwrap_italic(other[0])
        other = other[1:]

    link = None
    if not bullets and len(other) == 1:
        link = sole_link(other[0])
        if link:
            other = []

    tech = None
    if not bullets and other and is_wrapped_italic(other[-1]):
        tech = unwrap_italic(other[-1])
        other = other[:-1]

    before = 40 if first else 150
    parts = []

    if bullets:
        runs = [r(title, bold=True, size=22, color=BODY)]
        if meta:
            runs.append(r("  |  ", size=20, color=MUTED))
            runs.append(r(meta, bold=True, size=21, color=ACCENT))
        parts.append(p(runs, before=before, after=0))
        if date:
            parts.append(
                p(
                    render_spans(builder, parse_inline(date), size=18, color=MUTED, italic=True),
                    before=0,
                    after=40,
                )
            )
        for paragraph in other:
            parts.append(p(render_spans(builder, parse_inline(paragraph), size=20), after=40))
        for item in bullets:
            parts.append(bullet(builder, item))
        return parts

    if link:
        runs = [r(title, bold=True, size=21)]
        if meta:
            runs.append(r("  —  " + meta, size=20, color=MUTED))
        parts.append(p(runs, before=before, after=20))
        label, url = link
        parts.append(p([hyperlink(builder, url, label, size=18)], after=60))
        return parts

    if meta:
        runs = [r(title, bold=True, size=21)]
        runs.append(r("  —  " + meta, size=19, color=MUTED))
        parts.append(p(runs, before=before, after=20))
        for index, paragraph in enumerate(other):
            last = index == len(other) - 1 and not tech
            parts.append(
                p(render_spans(builder, parse_inline(paragraph), size=20), after=40 if last else 20)
            )
        if tech:
            parts.append(
                p(
                    render_spans(builder, parse_inline(tech), size=18, color=MUTED, italic=True),
                    before=0,
                    after=40,
                )
            )
        return parts

    parts.append(p([r(title, bold=True, size=21)], before=before, after=0))
    if other:
        parts.append(
            p(
                render_spans(builder, parse_inline(other[0]), size=18, color=MUTED, italic=True),
                before=0,
                after=40,
            )
        )
        for paragraph in other[1:]:
            parts.append(p(render_spans(builder, parse_inline(paragraph), size=20), after=60))
    return parts


def render_section_body(builder, lines):
    parts = []
    i = 0
    n = len(lines)
    first_entry = True
    while i < n:
        stripped = lines[i].strip()
        if not stripped or stripped == "---":
            i += 1
            continue
        if stripped.startswith("|"):
            j = i
            while j < n and lines[j].strip().startswith("|"):
                j += 1
            parts.extend(render_table(builder, lines[i:j]))
            i = j
            continue
        if stripped.startswith("- "):
            j = i
            while j < n and lines[j].strip().startswith("- "):
                j += 1
            parts.extend(render_bullet_list(builder, lines[i:j]))
            i = j
            continue
        if is_entry_header(stripped):
            j = i + 1
            while j < n and not is_entry_header(lines[j].strip()):
                j += 1
            parts.extend(render_entry(builder, lines[i:j], first_entry))
            first_entry = False
            i = j
            continue
        paragraph = [stripped]
        i += 1
        while i < n:
            nxt = lines[i].strip()
            if (
                not nxt
                or nxt == "---"
                or nxt.startswith("|")
                or nxt.startswith("- ")
                or is_entry_header(nxt)
            ):
                break
            paragraph.append(nxt)
            i += 1
        parts.append(p(render_spans(builder, parse_inline(" ".join(paragraph)), size=20), after=80))
    return parts


def build_document(resume, builder):
    body = []
    body.append(
        p(
            [r(resume["name"].upper(), bold=True, size=44, color=NAVY)],
            align="center",
            after=40,
        )
    )
    if resume["headline"]:
        body.append(
            p(
                render_spans(builder, parse_inline(resume["headline"]), size=18, color=ACCENT),
                align="center",
                after=80,
            )
        )
    contact = resume["contact"]
    for index, line in enumerate(contact):
        last = index == len(contact) - 1
        body.append(
            p(
                render_spans(builder, parse_inline(line), size=18, color=MUTED),
                align="center",
                before=0,
                after=120 if last else 20,
                border=last,
            )
        )
    for title, lines in resume["sections"]:
        body.append(heading(title))
        body.extend(render_section_body(builder, lines))

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"\n'
        '            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
        "  <w:body>\n    %s\n"
        "    <w:sectPr>\n"
        '      <w:pgSz w:w="12240" w:h="15840"/>\n'
        '      <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" '
        'w:header="360" w:footer="360"/>\n'
        "    </w:sectPr>\n  </w:body>\n</w:document>\n"
        % "".join(body)
    )
    return document_xml


def package_xml(resume, builder):
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""
    link_xml = []
    for rid, url in builder.links:
        safe = escape(url, {'"': "&quot;"})
        link_xml.append(
            '  <Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="%s" TargetMode="External"/>'
            % (rid, safe)
        )
    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n'
        "%s\n</Relationships>\n" % "\n".join(link_xml)
    )
    styles = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
        <w:sz w:val="21"/><w:szCs w:val="21"/>
        <w:color w:val="222222"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="60" w:line="240" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
</w:styles>
"""
    safe_name = escape(resume["name"])
    core = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"\n'
        '                   xmlns:dc="http://purl.org/dc/elements/1.1/"\n'
        '                   xmlns:dcterms="http://purl.org/dc/terms/"\n'
        '                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
        "  <dc:title>%s — Resume</dc:title>\n"
        "  <dc:creator>%s</dc:creator>\n"
        "  <cp:lastModifiedBy>%s</cp:lastModifiedBy>\n"
        "</cp:coreProperties>\n" % (safe_name, safe_name, safe_name)
    )
    app = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Python</Application>
</Properties>
"""
    return content_types, rels, doc_rels, styles, core, app


def write_docx(path, resume, builder, document_xml):
    content_types, rels, doc_rels, styles, core, app = package_xml(resume, builder)
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", rels)
        archive.writestr("word/document.xml", document_xml.encode("utf-8"))
        archive.writestr("word/styles.xml", styles)
        archive.writestr("word/_rels/document.xml.rels", doc_rels)
        archive.writestr("docProps/core.xml", core.encode("utf-8"))
        archive.writestr("docProps/app.xml", app)


def output_path(resume, folder):
    safe = re.sub(r'[<>:"/\\|?*]', "", resume["name"]).strip()
    safe = re.sub(r"\s+", "_", safe) or "Resume"
    return folder / ("%s_Resume.docx" % safe)


def find_markdown(explicit):
    if explicit:
        path = Path(explicit)
        if not path.exists():
            raise SystemExit("Markdown file not found: %s" % path)
        return path
    here = Path(__file__).resolve().parent
    preferred = sorted(here.glob("Resume_*.md"))
    if len(preferred) == 1:
        return preferred[0]
    found = sorted(here.glob("*.md"))
    if len(found) == 1:
        return found[0]
    if not found:
        raise SystemExit("Put a .md resume next to build_resume_docx.py")
    names = ", ".join(path.name for path in found)
    raise SystemExit("More than one .md file. Pass the path. Found: %s" % names)


def generate(md_path):
    text = md_path.read_text(encoding="utf-8")
    resume = parse_resume(text)
    builder = Builder()
    document_xml = build_document(resume, builder)
    # Fail before writing if the Word XML is not well formed.
    from xml.etree import ElementTree

    ElementTree.fromstring(document_xml.encode("utf-8"))
    folder = md_path.resolve().parent
    destination = output_path(resume, folder)
    try:
        write_docx(destination, resume, builder, document_xml)
    except PermissionError:
        raise SystemExit("Close %s in Word, then run this script again." % destination.name)
    desktop = Path.home() / "Desktop" / destination.name
    if desktop.resolve() != destination.resolve():
        try:
            write_docx(desktop, resume, builder, document_xml)
        except OSError:
            pass
    print("Read  %s" % md_path.name)
    print("Wrote %s (%d bytes)" % (destination, destination.stat().st_size))
    return destination


def main(argv=None):
    parser = argparse.ArgumentParser(description="Build a .docx resume from a Markdown file.")
    parser.add_argument("markdown", nargs="?", help="Resume Markdown. Defaults to the .md file next to this script.")
    parser.add_argument("--watch", action="store_true", help="Rebuild each time the Markdown file is saved.")
    args = parser.parse_args(argv)
    md_path = find_markdown(args.markdown)
    if not args.watch:
        generate(md_path)
        return
    print("Watching %s — save the file to rebuild. Ctrl+C to stop." % md_path.name)
    last = None
    while True:
        try:
            mtime = md_path.stat().st_mtime_ns
        except OSError as exc:
            print(exc)
            time.sleep(1)
            continue
        if mtime != last:
            last = mtime
            try:
                generate(md_path)
            except SystemExit as exc:
                print(exc)
        time.sleep(0.8)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nStopped.")
        sys.exit(0)
