#!/usr/bin/env python3
"""Convert index.html to a .docx that imports cleanly into Google Docs.

Stdlib only. Understands the markup used in index.html: h1, .contact, h2, p,
li, dl.skills (dt/dd), .row (left span + right-aligned span), .loc, .role,
.company. Right-aligned dates use a right tab stop, which Google Docs keeps.

usage: html2docx.py index.html out.docx
"""
import sys
import zipfile
from html import escape
from html.parser import HTMLParser

FONT = "Times New Roman"
SIZE = 21            # half-points (10.5pt; Docs renders slightly smaller than Chrome)
MARGIN_X = 720       # twips (0.5in)
MARGIN_Y = 576       # twips (0.4in)
TEXT_W = 12240 - 2 * MARGIN_X
LABEL_W = 2448       # skills label column (1.7in)

BLOCK = {"h1", "h2", "p", "li", "dt", "dd"}
BLOCK_CLASSES = {"row", "loc", "role"}


class Para:
    def __init__(self, kind):
        self.kind = kind     # name, contact, h2, p, li, skill, row, loc, role
        self.runs = []       # (text, bold, italic) or ("\t",)

    def add(self, text, bold=False, italic=False):
        if text:
            self.runs.append((text, bold, italic))

    def tab(self):
        self.runs.append(("\t",))


class Parser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.paras = []
        self.stack = []       # (tag, classes)
        self.cur = None
        self.cur_depth = None
        self.in_body = False
        self.skip = 0
        self.row_spans = 0
        self.contact = None   # list of [lines] per column

    def classes(self):
        out = set()
        for _, c in self.stack:
            out |= c
        return out

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        cls = set((a.get("class") or "").split())
        if tag in ("style", "script", "title"):
            self.skip += 1
        if tag == "body":
            self.in_body = True
        if tag in ("br", "meta", "link"):
            if tag == "br" and self.contact is not None:
                self.contact[-1].append("")
            elif tag == "br" and self.cur:
                self.cur.add(" ")
            return
        self.stack.append((tag, cls))
        if not self.in_body or self.skip:
            return
        if "contact" in cls:
            self.contact = []
            return
        if self.contact is not None and tag == "div" and len(self.stack) >= 2 and "contact" in self.stack[-2][1]:
            self.contact.append([""])
            return
        if self.cur is None and (tag in BLOCK or cls & BLOCK_CLASSES):
            kind = {"h1": "name", "h2": "h2", "p": "p", "li": "li", "dt": "skill", "dd": "skill"}.get(tag)
            if "row" in cls:
                kind = "row"
            elif "loc" in cls:
                kind = "loc"
            elif "role" in cls:
                kind = "role"
            if tag == "dd" and self.paras and self.paras[-1].kind == "skill":
                self.cur = self.paras.pop()   # continue the dt paragraph
                self.cur.tab()
            else:
                self.cur = Para(kind)
            self.cur_depth = len(self.stack)
            self.row_spans = 0
            return
        if self.cur and self.cur.kind == "row" and tag == "span" and len(self.stack) == self.cur_depth + 1:
            self.row_spans += 1
            if self.row_spans == 2:
                self.cur.tab()

    def handle_endtag(self, tag):
        if tag in ("style", "script", "title"):
            self.skip -= 1
        if not self.stack:
            return
        depth = len(self.stack)
        cls = self.stack[-1][1]
        self.stack.pop()
        if "contact" in cls and self.contact is not None:
            p = Para("contact")
            cols = [[l.strip() for l in c if l.strip()] for c in self.contact]
            for i in range(max(len(c) for c in cols)):
                if i:
                    p.runs.append(("\n",))
                p.add(cols[0][i] if i < len(cols[0]) else "")
                p.tab()
                p.add(cols[1][i] if len(cols) > 1 and i < len(cols[1]) else "")
            self.paras.append(p)
            self.contact = None
            return
        if self.cur is not None and depth == self.cur_depth:
            self.paras.append(self.cur)
            self.cur = None

    def handle_data(self, data):
        if self.skip or not self.in_body:
            return
        text = " ".join(data.split())
        if data[:1].isspace() and text:
            text = " " + text
        if data[-1:].isspace() and text:
            text = text + " "
        if self.contact is not None and self.contact:
            self.contact[-1][-1] += text
            return
        if self.cur is None or not text.strip() and not self.cur.runs:
            return
        c = self.classes()
        tags = {t for t, _ in self.stack}
        bold = bool({"company"} & c) or bool({"b", "strong", "dt", "h1", "h2"} & tags)
        italic = "role" in c
        if self.cur.kind == "role":
            italic = True
        self.cur.add(text, bold, italic)


def run_xml(r, size=SIZE, caps=False, spacing=None):
    if r == ("\t",):
        return "<w:r><w:tab/></w:r>"
    if r == ("\n",):
        return "<w:r><w:br/></w:r>"
    text, bold, italic = r
    props = f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/>'
    props += "<w:b/>" if bold else ""
    props += "<w:i/>" if italic else ""
    props += "<w:caps/>" if caps else ""
    props += f'<w:spacing w:val="{spacing}"/>' if spacing else ""
    props += f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
    return f'<w:r><w:rPr>{props}</w:rPr><w:t xml:space="preserve">{escape(text, quote=False)}</w:t></w:r>'


def tidy(runs):
    """Trim edge whitespace around tabs/ends and merge adjacent same-format runs."""
    out = []
    for r in runs:
        if len(r) == 3 and out and len(out[-1]) == 3 and out[-1][1:] == r[1:]:
            out[-1] = (out[-1][0] + r[0],) + r[1:]
        else:
            out.append(r)
    for i, r in enumerate(out):
        if len(r) != 3:
            continue
        t = r[0]
        if i == 0 or len(out[i - 1]) == 1:
            t = t.lstrip()
        if i == len(out) - 1 or len(out[i + 1]) == 1:
            t = t.rstrip()
        out[i] = (t,) + r[1:]
    return [r for r in out if len(r) == 1 or r[0]]


def para_xml(p):
    runs = tidy(p.runs)
    right_tab = f'<w:tabs><w:tab w:val="right" w:pos="{TEXT_W}"/></w:tabs>'
    if p.kind == "name":
        ppr = '<w:jc w:val="center"/><w:spacing w:before="0" w:after="40"/>'
        body = "".join(run_xml((t[0], True, False), size=36) for t in runs if len(t) == 3)
    elif p.kind == "contact":
        ppr = right_tab + '<w:spacing w:before="0" w:after="60"/>'
        body = "".join(run_xml(r, size=20) for r in runs)
    elif p.kind == "h2":
        ppr = ('<w:keepNext/><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/></w:pBdr>'
               '<w:jc w:val="center"/><w:spacing w:before="160" w:after="80"/>')
        body = "".join(run_xml(r, size=22, caps=True, spacing=20) for r in runs)
    elif p.kind == "li":
        ppr = ('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
               '<w:spacing w:before="0" w:after="0"/><w:ind w:left="320" w:hanging="200"/>')
        body = "".join(run_xml(r) for r in runs)
    elif p.kind == "skill":
        ppr = (f'<w:tabs><w:tab w:val="left" w:pos="{LABEL_W}"/></w:tabs>'
               f'<w:spacing w:before="0" w:after="0"/><w:ind w:left="{LABEL_W}" w:hanging="{LABEL_W}"/>')
        body = "".join(run_xml(r) for r in runs)
    elif p.kind == "row":
        ppr = right_tab + '<w:keepNext/><w:spacing w:before="0" w:after="0"/>'
        body = "".join(run_xml(r) for r in runs)
    elif p.kind in ("loc", "role"):
        ppr = '<w:keepNext/><w:spacing w:before="0" w:after="0"/>'
        body = "".join(run_xml(r, size=20 if p.kind == "loc" else SIZE) for r in runs)
    else:  # p
        ppr = '<w:spacing w:before="0" w:after="60"/>'
        body = "".join(run_xml(r) for r in runs)
    return f"<w:p><w:pPr>{ppr}</w:pPr>{body}</w:p>"


def build(paras):
    out = []
    for i, p in enumerate(paras):
        # a little air before each company row that follows a bullet list
        is_company = p.kind == "row" and p.runs and len(p.runs[0]) == 3 and p.runs[0][1]
        if is_company and i and paras[i - 1].kind == "li":
                out.append('<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>'
                           f'<w:r><w:rPr><w:sz w:val="8"/></w:rPr></w:r></w:p>')
        out.append(para_xml(p))
    sect = (f'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
            f'<w:pgMar w:top="{MARGIN_Y}" w:right="{MARGIN_X}" w:bottom="{MARGIN_Y}" w:left="{MARGIN_X}" '
            f'w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{"".join(out)}{sect}</w:body></w:document>')


CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>'''

STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}" w:eastAsia="{FONT}"/><w:sz w:val="{SIZE}"/><w:szCs w:val="{SIZE}"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>'''

NUMBERING = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#9679;"/><w:lvlJc w:val="left"/>
<w:pPr><w:ind w:left="320" w:hanging="200"/></w:pPr><w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}"/><w:sz w:val="14"/></w:rPr></w:lvl>
</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>'''


def main(src, dst):
    p = Parser()
    p.feed(open(src, encoding="utf-8").read())
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", build(p.paras))
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/numbering.xml", NUMBERING)


if __name__ == "__main__":
    main(*sys.argv[1:3])
