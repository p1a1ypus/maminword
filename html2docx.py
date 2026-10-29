"""HTML редактора -> DOCX.

Документ собирается поверх исходного файла: тело заменяется, а стили, колонтитулы,
сноски, нумерация, картинки и прочие части пакета остаются как были.
"""
import base64
import copy
import io
import os
import re

import lxml.html
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Emu
from lxml import etree

from docx_common import (CSS_LINE_FACTOR, HIGHLIGHT_BY_RGB, NS, PPR_ORDER, R_NS, RPR_ORDER, SECTPR_ORDER,
                         TBLPR_ORDER, TCPR_ORDER, BORDER_ORDER, Styles, W, WV, check_zip, dec_id, insert_ordered,
                         local, mk, to_int)

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "templates", "blank.docx")

BLOCK_TAGS = {"p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "ul", "ol", "li", "table", "blockquote", "pre",
              "hr", "address", "section", "article", "header", "footer", "main", "nav", "aside", "figure",
              "figcaption", "dl", "dt", "dd", "center", "form", "fieldset", "details", "summary", "caption",
              "thead", "tbody", "tfoot", "tr", "td", "th"}
PARA_TAGS = {"p", "h1", "h2", "h3", "h4", "h5", "h6", "div", "blockquote", "pre", "address", "dt", "dd",
             "figcaption", "center", "summary", "caption", "li"}
SKIP_TAGS = {"script", "style", "head", "title", "meta", "link", "noscript", "template", "colgroup", "col"}
BAD_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f﻿￾￿\ud800-\udfff]")
NAMED_COLORS = {"black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "008000", "blue": "0000FF",
                "yellow": "FFFF00", "gray": "808080", "grey": "808080", "silver": "C0C0C0", "maroon": "800000",
                "navy": "000080", "purple": "800080", "teal": "008080", "olive": "808000", "lime": "00FF00",
                "aqua": "00FFFF", "cyan": "00FFFF", "fuchsia": "FF00FF", "magenta": "FF00FF",
                "orange": "FFA500"}
FONT_SIZES = {"1": 8, "2": 10, "3": 12, "4": 14, "5": 18, "6": 24, "7": 36}
CSS_TO_NUMFMT = {"decimal": "decimal", "lower-alpha": "lowerLetter", "lower-latin": "lowerLetter",
                 "upper-alpha": "upperLetter", "upper-latin": "upperLetter", "lower-roman": "lowerRoman",
                 "upper-roman": "upperRoman", "decimal-leading-zero": "decimalZero",
                 "lower-greek": "lowerLetter", "cyrillic-lower": "russianLower"}


# ------------------------------------------------------------------ разбор CSS
def parse_style(s):
    out = {}
    for part in re.split(r";(?![^(]*\))", s or ""):
        if ":" in part:
            k, v = part.split(":", 1)
            out[k.strip().lower()] = v.strip()
    return out


def color_hex(v):
    if not v:
        return None
    v = v.strip().lower().replace("!important", "").strip()
    if v in ("transparent", "inherit", "initial", "none", "currentcolor", "unset"):
        return None
    if v in NAMED_COLORS:
        return NAMED_COLORS[v]
    m = re.match(r"#([0-9a-f]{3}|[0-9a-f]{6})\b", v)
    if m:
        h = m.group(1)
        if len(h) == 3:
            h = "".join(c * 2 for c in h)
        return h.upper()
    m = re.match(r"rgba?\(\s*([\d.]+)[,\s]+([\d.]+)[,\s]+([\d.]+)(?:[,\s/]+([\d.]+%?))?", v)
    if m:
        if m.group(4) is not None:
            a = m.group(4)
            alpha = float(a[:-1]) / 100 if a.endswith("%") else float(a)
            if alpha < 0.05:
                return None
        return "%02X%02X%02X" % tuple(min(255, int(float(m.group(i)))) for i in (1, 2, 3))
    return None


def length_tw(v, ref=None, font_pt=11.0):
    """CSS-длина -> твипы (1/20 пт). ref — база для процентов (в твипах)."""
    if v is None:
        return None
    v = v.strip().lower().replace("!important", "").strip()
    if v in ("0", "auto", "normal", "initial", "inherit"):
        return 0 if v == "0" else None
    m = re.match(r"^(-?[\d.]+)\s*(pt|px|cm|mm|in|pc|em|rem|%|q)?$", v)
    if not m:
        return None
    n = float(m.group(1))
    unit = m.group(2) or "px"
    k = {"pt": 20, "px": 15, "cm": 566.929, "mm": 56.6929, "in": 1440, "pc": 240, "q": 14.173}.get(unit)
    if k:
        return int(round(n * k))
    if unit in ("em", "rem"):
        return int(round(n * font_pt * 20))
    if unit == "%":
        return int(round((ref or 0) * n / 100.0))
    return None


def font_pt(v, cur):
    v = (v or "").strip().lower()
    named = {"xx-small": 7, "x-small": 7.5, "small": 10, "medium": 12, "large": 13.5, "x-large": 18,
             "xx-large": 24, "smaller": cur * 0.83, "larger": cur * 1.2}
    if v in named:
        return named[v]
    m = re.match(r"^([\d.]+)\s*(pt|px|em|rem|%|cm|mm)?$", v)
    if not m:
        return None
    n = float(m.group(1))
    unit = m.group(2) or "px"
    return {"pt": n, "px": n * 0.75, "em": n * cur, "rem": n * 12, "%": cur * n / 100.0,
            "cm": n * 28.3465, "mm": n * 2.83465}[unit]


def classes(el):
    return (el.get("class") or "").split()


def cls_value(el, prefix):
    for c in classes(el):
        if c.startswith(prefix):
            return dec_id(c[len(prefix):])
    return None


def border_from_css(v):
    """'0.5pt solid #000' -> (val, sz, color) или 'nil'."""
    if v is None:
        return None
    v = v.strip().lower()
    if not v:
        return None
    if v in ("none", "0", "hidden") or v.startswith(("none", "hidden")) or " none" in v:
        return "nil"
    width, style, color = 0.5, "single", "auto"
    for tok in re.split(r"\s+(?![^(]*\))", v):
        w = length_tw(tok)
        if w is not None and re.match(r"^-?[\d.]", tok):
            width = w / 20.0
        elif tok in ("solid", "double", "dotted", "dashed", "groove", "ridge", "inset", "outset"):
            style = {"solid": "single", "double": "double", "dotted": "dotted", "dashed": "dashed"}.get(tok, "single")
        else:
            c = color_hex(tok)
            if c:
                color = c
    if width <= 0:
        return "nil"
    if style == "double":
        width = width / 3.0
    return (style, max(2, min(96, int(round(width * 8)))), color)


def border_el(side, spec):
    if spec == "nil":
        return mk(side, val="nil")
    val, sz, color = spec
    return mk(side, val=val, sz=sz, space=0, color=color)


# ------------------------------------------------------------------ нумерация
class NumberingWriter:
    def __init__(self, document):
        self.root = document.part.numbering_part.element
        nums = self.root.findall(W("num"))
        absn = self.root.findall(W("abstractNum"))
        self.existing = {n.get(W("numId")) for n in nums}
        self.next_abs = max([to_int(a.get(W("abstractNumId")), 0) for a in absn] + [0]) + 1
        self.next_num = max([to_int(n.get(W("numId")), 0) for n in nums] + [0]) + 1
        self.ours = {}
        self.left = {}

    def exists(self, nid):
        return nid in self.existing

    def level_left(self, nid, ilvl):
        return 720 * (ilvl + 1)

    @staticmethod
    def make_lvl(ilvl, ordered, css_type=None):
        lvl = mk("lvl", ilvl=ilvl)
        lvl.append(mk("start", val=1))
        if ordered:
            fmt = CSS_TO_NUMFMT.get(css_type or "", ["decimal", "lowerLetter", "lowerRoman"][ilvl % 3])
            lvl.append(mk("numFmt", val=fmt))
            lvl.append(mk("lvlText", val="%%%d." % (ilvl + 1)))
        else:
            ch = {"circle": "o", "square": "▪", "disc": "•"}.get(css_type or "", ["•", "o", "▪"][ilvl % 3])
            lvl.append(mk("numFmt", val="bullet"))
            lvl.append(mk("lvlText", val=ch))
        lvl.append(mk("lvlJc", val="left"))
        ppr = mk("pPr")
        ppr.append(mk("ind", left=720 * (ilvl + 1), hanging=360))
        lvl.append(ppr)
        if not ordered and lvl.find(W("lvlText")).get(WV) == "o":
            rpr = mk("rPr")
            rpr.append(mk("rFonts", ascii="Courier New", hAnsi="Courier New", cs="Courier New", hint="default"))
            lvl.append(rpr)
        return lvl

    def new_list(self, ordered, css_type):
        aid, nid = str(self.next_abs), str(self.next_num)
        self.next_abs += 1
        self.next_num += 1
        an = mk("abstractNum", abstractNumId=aid)
        an.append(mk("multiLevelType", val="hybridMultilevel"))
        for i in range(9):
            an.append(self.make_lvl(i, ordered, css_type if i == 0 else None))
        first_num = None
        for ch in self.root:
            if local(ch) in ("num", "numIdMacAtCleanup"):
                first_num = ch
                break
        if first_num is not None:
            first_num.addprevious(an)
        else:
            self.root.append(an)
        num = mk("num", numId=nid)
        num.append(mk("abstractNumId", val=aid))
        cleanup = self.root.find(W("numIdMacAtCleanup"))
        if cleanup is not None:
            cleanup.addprevious(num)
        else:
            self.root.append(num)
        self.existing.add(nid)
        self.ours[nid] = (an, ordered)
        return nid

    def set_level(self, nid, ilvl, ordered, css_type):
        if nid not in self.ours or not css_type:
            return
        an, _ = self.ours[nid]
        for lvl in an.findall(W("lvl")):
            if lvl.get(W("ilvl")) == str(ilvl):
                lvl.addprevious(self.make_lvl(ilvl, ordered, css_type))
                an.remove(lvl)
                break


# ------------------------------------------------------------------- писатель
class HtmlToDocx:
    def __init__(self, document, page):
        self.doc = document
        self.st = Styles(document)
        self.body = document.element.body
        self.page = page or {}
        self.numw = NumberingWriter(document)
        self.warnings = set()
        self.pending_pb = False
        self.last_body_p = None
        self.list_style = self.st.by_name("List Paragraph")
        w = to_int(self.page.get("w"), 11906)
        self.text_width = max(2000, w - to_int(self.page.get("left"), 1701) - to_int(self.page.get("right"), 850))

    # ----------------------------------------------------------- точка входа
    def run(self, html_text):
        sect = None
        for ch in list(self.body):
            if ch.tag == W("sectPr"):
                sect = ch
            else:
                self.body.remove(ch)
        if sect is None:
            sect = mk("sectPr")
            self.body.append(sect)
        self.sect = sect
        root = lxml.html.fragment_fromstring(html_text or "", create_parent="div")
        self.walk_blocks(root, self.body, {"props": {}, "pcss": {}})
        if self.pending_pb:
            self.emit_simple_p(self.body, pb=True)
        children = [c for c in self.body if c.tag != W("sectPr")]
        if not children or children[-1].tag != W("p"):
            self.emit_simple_p(self.body)
        self.apply_page()
        self.finish()

    def append(self, container, el):
        if container is self.body:
            self.sect.addprevious(el)
            if el.tag == W("p"):
                self.last_body_p = el
        else:
            container.append(el)

    def emit_simple_p(self, container, pb=False):
        p = mk("p")
        if pb:
            r = mk("r")
            r.append(mk("br", type="page"))
            p.append(r)
        self.append(container, p)
        return p

    # ----------------------------------------------------------------- блоки
    def walk_blocks(self, el, container, ctx):
        pending = []

        def flush():
            if any(self.meaningful(n) for n in pending):
                self.paragraph(container, None, list(pending), ctx)
            pending.clear()

        if el.text:
            pending.append(el.text)
        for ch in el:
            if not isinstance(ch.tag, str):
                if ch.tag is etree.Comment and "pagebreak" in (ch.text or ""):
                    flush()
                    self.pending_pb = True
            else:
                tag = ch.tag.lower()
                if tag in SKIP_TAGS:
                    pass
                elif tag in BLOCK_TAGS or self.has_block(ch):
                    flush()
                    self.block(ch, container, ctx)
                else:
                    pending.append(ch)
            if ch.tail:
                pending.append(ch.tail)
        flush()

    @staticmethod
    def meaningful(n):
        if isinstance(n, str):
            return bool(n.strip(" \t\r\n"))
        return True

    @staticmethod
    def has_block(el):
        for d in el.iterdescendants():
            if isinstance(d.tag, str) and d.tag.lower() in BLOCK_TAGS:
                return True
        return False

    def block(self, el, container, ctx):
        tag = el.tag.lower()
        cl = classes(el)
        if tag == "table":
            self.table(el, container, ctx)
        elif tag in ("ul", "ol"):
            self.walk_list(el, container, ctx, None)
        elif tag == "hr":
            p = self.emit_simple_p(container)
            ppr = mk("pPr")
            bdr = mk("pBdr")
            bdr.append(mk("bottom", val="single", sz=6, space=1, color="auto"))
            ppr.append(bdr)
            p.insert(0, ppr)
        elif "docx-section" in cl:
            self.section_break(el, container)
        elif "docx-raw" in cl and el.get("data-kind") == "b":
            raw = self.raw_xml(el)
            if raw is not None:
                self.append(container, raw)
        elif tag in PARA_TAGS and not self.has_block(el):
            self.paragraph(container, el, None, ctx)
        else:
            self.walk_blocks(el, container, self.inherit(ctx, el))

    def inherit(self, ctx, el):
        css = parse_style(el.get("style"))
        pcss = dict(ctx.get("pcss", {}))
        if "text-align" in css:
            pcss["text-align"] = css["text-align"]
        if el.get("align"):
            pcss["text-align"] = el.get("align")
        props = self.apply_element(ctx.get("props", {}), el) if el.tag.lower() not in BLOCK_TAGS else \
            ctx.get("props", {})
        return {"props": props, "pcss": pcss, "pre": ctx.get("pre") or el.tag.lower() == "pre"}

    def section_break(self, el, container):
        raw = self.b64xml(el.get("data-sectpr"))
        if raw is None or container is not self.body:
            return
        p = self.last_body_p
        if p is None or p.getparent() is None:
            p = self.emit_simple_p(container)
        ppr = p.find(W("pPr"))
        if ppr is None:
            ppr = mk("pPr")
            p.insert(0, ppr)
        insert_ordered(ppr, raw, PPR_ORDER)
        self.last_body_p = None

    # ----------------------------------------------------------------- списки
    def walk_list(self, lst, container, ctx, parent):
        ordered = lst.tag.lower() == "ol"
        css = parse_style(lst.get("style"))
        css_type = (css.get("list-style-type") or "").strip().lower() or None
        nid, level = None, 0
        for c in classes(lst):
            m = re.match(r"lw-n(\w+)-l(\d+)$", c)
            if m and self.numw.exists(m.group(1)):
                nid, level = m.group(1), int(m.group(2))
        if nid is None:
            if parent is not None:
                nid, level = parent["nid"], parent["level"] + 1
                self.numw.set_level(nid, level, ordered, css_type)
            else:
                nid = self.numw.new_list(ordered, css_type)
        elif parent is not None and parent["nid"] == nid and level != parent["level"] + 1:
            level = parent["level"] + 1
        level = min(level, 8)
        info = {"nid": nid, "level": level, "ordered": ordered}
        for li in lst:
            if not isinstance(li.tag, str):
                continue
            tag = li.tag.lower()
            if tag in ("ul", "ol"):
                self.walk_list(li, container, ctx, info)
                continue
            if tag != "li":
                self.block(li, container, ctx) if tag in BLOCK_TAGS else None
                continue
            self.list_item(li, container, ctx, info)

    def list_item(self, li, container, ctx, info):
        pending = []
        state = {"first": True}
        li_ctx = self.inherit(ctx, li)

        def emit(nodes, el=None):
            num = (info["nid"], info["level"]) if state["first"] else None
            self.paragraph(container, el, nodes, li_ctx, li=li, num=num, cont_level=info)
            state["first"] = False

        if li.text:
            pending.append(li.text)
        for ch in li:
            if isinstance(ch.tag, str):
                tag = ch.tag.lower()
                if tag in ("ul", "ol"):
                    if any(self.meaningful(n) for n in pending) or state["first"]:
                        emit(list(pending))
                    pending.clear()
                    self.walk_list(ch, container, ctx, info)
                elif tag in BLOCK_TAGS or self.has_block(ch):
                    if any(self.meaningful(n) for n in pending):
                        emit(list(pending))
                    pending.clear()
                    if tag in PARA_TAGS and not self.has_block(ch):
                        emit(None, ch)
                    else:
                        self.block(ch, container, li_ctx)
                else:
                    pending.append(ch)
            elif ch.tag is etree.Comment and "pagebreak" in (ch.text or ""):
                self.pending_pb = True
            if ch.tail:
                pending.append(ch.tail)
        if any(self.meaningful(n) for n in pending) or state["first"]:
            emit(list(pending))

    # ----------------------------------------------------------------- абзац
    def paragraph(self, container, el, nodes, ctx, li=None, num=None, cont_level=None):
        pcss = dict(ctx.get("pcss", {}))
        props = dict(ctx.get("props", {}))
        sid = None
        owner = el if el is not None else li
        if li is not None:
            sid = cls_value(li, "s-") or self.list_style
            pcss.update({k: v for k, v in parse_style(li.get("style")).items()
                         if k not in ("margin-left", "text-indent", "list-style-position", "margin-right")})
        if el is not None:
            tag = el.tag.lower()
            sid = cls_value(el, "s-") or sid
            if sid is None and re.match(r"h[1-6]$", tag):
                sid = self.st.by_name("heading " + tag[1])
            ecss = parse_style(el.get("style"))
            if el.get("align"):
                ecss.setdefault("text-align", el.get("align"))
            pcss.update(ecss)
            props = self.apply_element(props, el, block=True)
        if sid and sid not in self.st.styles:
            sid = None
        pre = ctx.get("pre") or (el is not None and el.tag.lower() == "pre")
        items = []
        if el is not None:
            self.walk_inline(el, props, items, None, pre)
        else:
            for n in nodes or []:
                if isinstance(n, str):
                    self.add_text(items, n, props, None, pre)
                else:
                    self.walk_node(n, props, items, None, pre)

        # разрывы страниц внутри абзаца -> несколько абзацев
        groups = [[]]
        for it in items:
            if it[0] == "pb":
                groups.append([])
            else:
                groups[-1].append(it)
        for gi, group in enumerate(groups):
            if gi > 0:
                self.pending_pb = True
                if not group and gi == len(groups) - 1:
                    continue
            self.write_p(container, owner, sid, pcss, props, group, li, num if gi == 0 else None, cont_level)

    def write_p(self, container, owner, sid, pcss, props, items, li, num, cont_level):
        p = mk("p")
        ppr = None
        raw_ppr = owner.get("data-ppr") if owner is not None else None
        if raw_ppr:
            ppr = self.b64xml(raw_ppr)
        if ppr is None:
            ppr = mk("pPr")
        if sid and sid != self.st.default.get("paragraph"):
            insert_ordered(ppr, mk("pStyle", val=sid), PPR_ORDER)
        if num is not None:
            np_ = mk("numPr")
            np_.append(mk("ilvl", val=num[1]))
            np_.append(mk("numId", val=num[0]))
            insert_ordered(ppr, np_, PPR_ORDER)
        fpt = props.get("size") or 11.0
        # интервалы
        spacing = {}
        for side, keys in (("before", ("padding-top", "margin-top")), ("after", ("padding-bottom", "margin-bottom"))):
            vals = [length_tw(pcss[k], font_pt=fpt) for k in keys if k in pcss]
            vals = [v for v in vals if v is not None]
            if vals:
                spacing[side] = max(0, sum(vals))
        lh = pcss.get("line-height")
        if lh:
            line = self.line_from_css(lh, fpt)
            if line:
                spacing["line"], spacing["lineRule"] = line
        if spacing:
            insert_ordered(ppr, mk("spacing", **spacing), PPR_ORDER)
        # отступы
        ind = {}
        if li is not None:
            if num is not None and li.get("data-ind") and li.get("data-lvl") == str(num[1]):
                parts = [to_int(x, 0) for x in li.get("data-ind").split(",")]
                left, first, right = (parts + [0, 0, 0])[:3]
                ind["left"] = left
                if first < 0:
                    ind["hanging"] = -first
                else:
                    ind["firstLine"] = first
                if right:
                    ind["right"] = right
            elif num is None and cont_level is not None:
                ind["left"] = self.numw.level_left(cont_level["nid"], cont_level["level"])
        else:
            ref = self.text_width
            for key, prop in (("left", "margin-left"), ("right", "margin-right")):
                if prop in pcss:
                    v = length_tw(pcss[prop], ref, fpt)
                    if v is not None:
                        ind[key] = v
            if "padding-left" in pcss:
                v = length_tw(pcss["padding-left"], ref, fpt)
                if v:
                    ind["left"] = ind.get("left", 0) + v
            if "text-indent" in pcss:
                v = length_tw(pcss["text-indent"], ref, fpt)
                if v is not None:
                    if v < 0:
                        ind["hanging"] = -v
                    else:
                        ind["firstLine"] = v
        if ind:
            insert_ordered(ppr, mk("ind", **ind), PPR_ORDER)
        al = (pcss.get("text-align") or "").lower()
        jc = {"center": "center", "right": "right", "justify": "both", "left": "left", "start": "left",
              "end": "right"}.get(al)
        if jc:
            insert_ordered(ppr, mk("jc", val=jc), PPR_ORDER)
        bg = color_hex(pcss.get("background-color") or pcss.get("background"))
        if bg:
            insert_ordered(ppr, mk("shd", val="clear", color="auto", fill=bg), PPR_ORDER)
        if owner is not None:
            mark = self.direct_props(owner)
            if mark:
                insert_ordered(ppr, self.build_rpr(mark), PPR_ORDER)
        if len(ppr):
            p.append(ppr)
        if self.pending_pb:
            r = mk("r")
            r.append(mk("br", type="page"))
            p.append(r)
            self.pending_pb = False
        self.write_runs(p, items)
        self.append(container, p)

    def direct_props(self, el):
        """Свойства знака абзаца из inline-стиля самого блока."""
        css = parse_style(el.get("style"))
        keys = ("font-family", "font-size", "color", "font-weight", "font-style")
        if not any(k in css for k in keys):
            return {}
        fake = lxml.html.Element("span")
        fake.set("style", "; ".join("%s: %s" % (k, css[k]) for k in keys if k in css))
        return self.apply_element({}, fake)

    @staticmethod
    def line_from_css(v, fpt):
        v = v.strip().lower()
        if v == "normal":
            return 240, "auto"
        m = re.match(r"max\(\s*([\d.]+)\s*(pt|px)", v)
        if m:
            n = float(m.group(1)) * (0.75 if m.group(2) == "px" else 1)
            return int(round(n * 20)), "atLeast"
        m = re.match(r"^([\d.]+)\s*(%|em)?$", v)
        if m:
            n = float(m.group(1)) / (100.0 if m.group(2) == "%" else 1)
            return max(1, int(round(n / CSS_LINE_FACTOR * 240))), "auto"
        tw = length_tw(v, font_pt=fpt)
        if tw:
            return tw, "exact"
        return None

    # --------------------------------------------------------- строчные узлы
    def apply_element(self, props, el, block=False):
        p = dict(props)
        tag = el.tag.lower() if isinstance(el.tag, str) else ""
        if not block:
            if tag in ("b", "strong"):
                p["b"] = True
            elif tag in ("i", "em", "cite", "dfn", "var"):
                p["i"] = True
            elif tag in ("u", "ins"):
                p["u"] = "single"
            elif tag in ("s", "strike", "del"):
                p["strike"] = True
            elif tag == "sup":
                p["va"] = "superscript"
            elif tag == "sub":
                p["va"] = "subscript"
            elif tag in ("code", "kbd", "samp", "tt"):
                p["font"] = "Courier New"
            elif tag == "mark":
                p["bg"] = "FFFF00"
            elif tag == "font":
                if el.get("face"):
                    p["font"] = el.get("face").split(",")[0].strip().strip("'\"")
                if color_hex(el.get("color")):
                    p["color"] = color_hex(el.get("color"))
                if el.get("size") in FONT_SIZES:
                    p["size"] = FONT_SIZES[el.get("size")]
            rs = cls_value(el, "c-")
            if rs and rs in self.st.styles:
                p["rStyle"] = rs
        elif tag == "pre":
            p["font"] = "Courier New"
        css = parse_style(el.get("style"))
        if not css:
            return p
        if "font-family" in css:
            fam = css["font-family"].split(",")[0].strip().strip("'\"")
            if fam and fam.lower() not in ("inherit", "initial", "serif", "sans-serif", "monospace"):
                p["font"] = fam
        if "font-size" in css:
            v = font_pt(css["font-size"], p.get("size") or 11.0)
            if v:
                p["size"] = round(v * 2) / 2.0
        fw = css.get("font-weight", "").lower()
        if fw in ("bold", "bolder") or (fw.isdigit() and int(fw) >= 600):
            p["b"] = True
        elif fw in ("normal", "lighter") or (fw.isdigit() and int(fw) < 600):
            p["b"] = False
        fs = css.get("font-style", "").lower()
        if fs in ("italic", "oblique"):
            p["i"] = True
        elif fs == "normal":
            p["i"] = False
        deco = " ".join(css.get(k, "") for k in ("text-decoration", "text-decoration-line")).lower()
        if "underline" in deco:
            dstyle = deco + " " + css.get("text-decoration-style", "").lower()
            p["u"] = ("double" if "double" in dstyle else "dotted" if "dotted" in dstyle else
                      "dash" if "dashed" in dstyle else "wave" if "wavy" in dstyle else "single")
        if "line-through" in deco:
            p["strike"] = True
        c = color_hex(css.get("color"))
        if c:
            p["color"] = c
        if not block:
            bg = color_hex(css.get("background-color") or css.get("background"))
            if bg:
                p["bg"] = bg
        va = css.get("vertical-align", "").lower()
        if va in ("super", "sub"):
            p["va"] = "superscript" if va == "super" else "subscript"
        tt = css.get("text-transform", "").lower()
        if tt == "uppercase":
            p["caps"] = True
        elif tt == "none":
            p["caps"] = False
        if "small-caps" in css.get("font-variant", "").lower():
            p["smallCaps"] = True
        if "letter-spacing" in css:
            v = length_tw(css["letter-spacing"], font_pt=p.get("size") or 11.0)
            if v is not None:
                p["spacing"] = v
        if css.get("display", "").lower() == "none":
            p["vanish"] = True
        return p

    def add_text(self, items, text, props, link, pre):
        if not text:
            return
        if pre:
            lines = text.split("\n")
            for i, line in enumerate(lines):
                if i:
                    items.append(("br", props, link))
                if line:
                    items.append(("text", line.replace("\r", ""), props, link))
            return
        text = re.sub(r"[ \t\r\n\f]+", " ", text)
        items.append(("text", text, props, link))

    def walk_inline(self, el, props, items, link, pre):
        if el.text:
            self.add_text(items, el.text, props, link, pre)
        for ch in el:
            self.walk_node(ch, props, items, link, pre)
            if ch.tail:
                self.add_text(items, ch.tail, props, link, pre)

    def walk_node(self, el, props, items, link, pre):
        if not isinstance(el.tag, str):
            if el.tag is etree.Comment and "pagebreak" in (el.text or ""):
                items.append(("pb",))
            return
        tag = el.tag.lower()
        cl = classes(el)
        if tag in SKIP_TAGS:
            return
        if "docx-tab" in cl:
            items.append(("tab", props, link))
            return
        if "docx-raw" in cl:
            items.append(("raw", el, props, link))
            return
        if tag == "br":
            items.append(("br", props, link))
            return
        if tag == "img":
            items.append(("img", el, props, link))
            return
        if "mce-pagebreak" in cl:
            items.append(("pb",))
            return
        if tag in ("input", "button", "select", "textarea", "iframe", "video", "audio", "svg", "object"):
            return
        new_link = link
        if tag == "a" and el.get("href"):
            new_link = el.get("href")
        self.walk_inline(el, self.apply_element(props, el), items, new_link, pre or tag == "pre")

    # ---------------------------------------------------------------- прогоны
    @staticmethod
    def props_key(props):
        return tuple(sorted((k, v) for k, v in props.items()))

    def write_runs(self, p, items):
        # обрезаем пробелы по краям и превращаем «технические» nbsp в обычные пробелы
        while items and items[-1][0] == "br":
            items.pop()
        texts = [i for i, it in enumerate(items) if it[0] == "text"]
        if texts:
            first, last = texts[0], texts[-1]
            if first == 0:
                items[first] = ("text", items[first][1].lstrip(" "),) + items[first][2:]
            if last == len(items) - 1:
                items[last] = ("text", items[last][1].rstrip(" "),) + items[last][2:]
        flat = []
        for i, it in enumerate(items):
            if it[0] == "text":
                flat.extend((i, j) for j in range(len(it[1])))
            else:
                flat.append((i, None))
        chars = {}
        for idx, (i, j) in enumerate(flat):
            if j is None:
                continue
            ch = items[i][1][j]
            if ch != " ":
                continue
            prev = flat[idx - 1] if idx > 0 else None
            nxt = flat[idx + 1] if idx + 1 < len(flat) else None

            def char_at(pos):
                if pos is None or pos[1] is None:
                    return None
                return items[pos[0]][1][pos[1]]

            pc, nc = char_at(prev), char_at(nxt)
            if pc in (None, " ", " ") or nc in (None, " ", " "):
                chars[(i, j)] = " "
        if chars:
            for i, it in enumerate(items):
                if it[0] == "text":
                    s = list(it[1])
                    for j in range(len(s)):
                        if (i, j) in chars:
                            s[j] = " "
                    items[i] = ("text", "".join(s)) + it[2:]

        cur_link, link_el = None, None
        last_run, last_key = None, None
        for it in items:
            kind = it[0]
            link = it[-1] if kind != "raw" else it[3]
            if link != cur_link:
                cur_link = link
                link_el = self.make_link(p, link) if link else None
                last_run, last_key = None, None
            target = link_el if link_el is not None else p
            if kind == "raw":
                el = self.raw_xml(it[1])
                if el is None:
                    continue
                if it[1].get("data-kind") == "p" or el.tag != W("r"):
                    p.append(el)
                else:
                    target.append(el)
                last_run = None
                continue
            if kind == "img":
                r = self.image_run(it[1], it[2])
                if r is not None:
                    target.append(r)
                last_run = None
                continue
            if kind == "text" and not it[1]:
                continue
            props = dict(it[2] if kind == "text" else it[1])
            if link and "rStyle" not in props:
                hl = self.ensure_char_style("Hyperlink")
                if hl:
                    props["rStyle"] = hl
            key = self.props_key(props)
            if last_run is None or key != last_key:
                last_run = mk("r")
                rpr = self.build_rpr(props)
                if len(rpr):
                    last_run.append(rpr)
                target.append(last_run)
                last_key = key
            if kind == "text":
                self.add_text_children(last_run, it[1])
            elif kind == "tab":
                last_run.append(mk("tab"))
            elif kind == "br":
                last_run.append(mk("br"))

    def make_link(self, p, href):
        h = mk("hyperlink")
        if href.startswith("#"):
            h.set(W("anchor"), href[1:])
        else:
            try:
                rid = self.doc.part.relate_to(href, RT.HYPERLINK, is_external=True)
                h.set("{%s}id" % R_NS, rid)
            except Exception:
                return None
        h.set(W("history"), "1")
        p.append(h)
        return h

    @staticmethod
    def add_text_children(r, text):
        text = BAD_XML.sub("", text)
        buf = []

        def flush():
            if buf:
                t = etree.SubElement(r, W("t"))
                t.text = "".join(buf)
                if t.text != t.text.strip():
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                buf.clear()

        for ch in text:
            if ch == "\t":
                flush()
                r.append(mk("tab"))
            elif ch == "‑":
                flush()
                r.append(mk("noBreakHyphen"))
            elif ch == "­":
                flush()
                r.append(mk("softHyphen"))
            else:
                buf.append(ch)
        flush()

    def build_rpr(self, props):
        rpr = mk("rPr")
        if props.get("rStyle"):
            insert_ordered(rpr, mk("rStyle", val=props["rStyle"]), RPR_ORDER)
        if props.get("font"):
            f = props["font"]
            insert_ordered(rpr, mk("rFonts", ascii=f, hAnsi=f, cs=f), RPR_ORDER)
        for k in ("b", "i", "caps", "smallCaps", "strike", "vanish"):
            if k in props:
                insert_ordered(rpr, mk(k) if props[k] else mk(k, val=0), RPR_ORDER)
        if props.get("b") is True:
            insert_ordered(rpr, mk("bCs"), RPR_ORDER)
        if props.get("color"):
            insert_ordered(rpr, mk("color", val=props["color"]), RPR_ORDER)
        if props.get("spacing") is not None:
            insert_ordered(rpr, mk("spacing", val=props["spacing"]), RPR_ORDER)
        if props.get("size"):
            hp = int(round(props["size"] * 2))
            insert_ordered(rpr, mk("sz", val=hp), RPR_ORDER)
            insert_ordered(rpr, mk("szCs", val=hp), RPR_ORDER)
        bg = props.get("bg")
        if bg:
            name = HIGHLIGHT_BY_RGB.get("#" + bg)
            if name:
                insert_ordered(rpr, mk("highlight", val=name), RPR_ORDER)
            else:
                insert_ordered(rpr, mk("shd", val="clear", color="auto", fill=bg), RPR_ORDER)
        if props.get("u"):
            insert_ordered(rpr, mk("u", val=props["u"]), RPR_ORDER)
        if props.get("va"):
            insert_ordered(rpr, mk("vertAlign", val=props["va"]), RPR_ORDER)
        return rpr

    # ------------------------------------------------------- сырые объекты/картинки
    @staticmethod
    def b64xml(s):
        if not s:
            return None
        try:
            return etree.fromstring(base64.b64decode(s))
        except Exception:
            return None

    def raw_xml(self, el):
        return self.b64xml(el.get("data-raw"))

    @staticmethod
    def img_size(el):
        css = parse_style(el.get("style"))

        def px(v):
            tw = length_tw(v)
            return tw / 15.0 if tw else None

        w = px(css.get("width")) or px(el.get("width") or "")
        h = px(css.get("height")) or px(el.get("height") or "")
        return w, h

    def image_run(self, el, props):
        w, h = self.img_size(el)
        raw = self.raw_xml(el)
        if raw is not None:
            ow, oh = to_int(el.get("data-w"), 0), to_int(el.get("data-h"), 0)
            if w and h and (abs(w - ow) > 1 or abs(h - oh) > 1):
                self.resize_raw(raw, w, h)
            return raw
        src = el.get("src") or ""
        m = re.match(r"data:([\w/+.-]+);base64,(.*)$", src, re.S)
        if not m:
            self.warnings.add("Часть картинок по внешним ссылкам не сохранена — вставьте их как файлы.")
            return None
        data = base64.b64decode(m.group(2))
        try:
            if not (w and h):
                inline = self.doc.part.new_pic_inline(io.BytesIO(data))
            else:
                inline = self.doc.part.new_pic_inline(io.BytesIO(data), Emu(int(w * 9525)), Emu(int(h * 9525)))
        except Exception:
            self.warnings.add("Картинка в неподдерживаемом формате пропущена (%s)." % m.group(1))
            return None
        r = mk("r")
        rpr = self.build_rpr({k: v for k, v in props.items() if k in ("rStyle",)})
        if len(rpr):
            r.append(rpr)
        drawing = etree.SubElement(r, W("drawing"))
        drawing.append(inline)
        return r

    @staticmethod
    def resize_raw(r, w, h):
        cx, cy = str(int(w * 9525)), str(int(h * 9525))
        for ext in r.iter("{%s}extent" % NS["wp"]):
            ext.set("cx", cx)
            ext.set("cy", cy)
        for xfrm in r.iter("{%s}xfrm" % NS["a"]):
            ext = xfrm.find("{%s}ext" % NS["a"])
            if ext is not None:
                ext.set("cx", cx)
                ext.set("cy", cy)
        for shape in r.iter("{%s}shape" % NS["v"]):
            st = shape.get("style", "")
            st = re.sub(r"width\s*:\s*[\d.]+\w*", "width:%.2fpt" % (w * 0.75), st)
            st = re.sub(r"height\s*:\s*[\d.]+\w*", "height:%.2fpt" % (h * 0.75), st)
            shape.set("style", st)

    def ensure_char_style(self, name):
        sid = self.st.by_name(name, "character")
        if sid:
            return sid
        if name != "Hyperlink":
            return None
        el = etree.fromstring(
            '<w:style xmlns:w="%s" w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>'
            '<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/>'
            '</w:rPr></w:style>' % NS["w"])
        self.doc.styles.element.append(el)
        self.st = Styles(self.doc)
        return "Hyperlink"

    def ensure_table_grid(self):
        sid = self.st.by_name("Table Grid", "table")
        if sid:
            return sid
        el = etree.fromstring(
            '<w:style xmlns:w="%s" w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>'
            '<w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            '<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>'
            '</w:style>' % NS["w"])
        self.doc.styles.element.append(el)
        self.st = Styles(self.doc)
        return "TableGrid"

    # --------------------------------------------------------------- таблицы
    def table(self, tel, container, ctx):
        rows = []
        for part in tel:
            if not isinstance(part.tag, str):
                continue
            t = part.tag.lower()
            if t == "tr":
                rows.append(part)
            elif t in ("thead", "tbody", "tfoot"):
                rows.extend(r for r in part if isinstance(r.tag, str) and r.tag.lower() == "tr")
        if not rows:
            return
        occ, grid_cells = {}, []
        for ri, tr in enumerate(rows):
            col, row_cells = 0, []
            for td in tr:
                if not isinstance(td.tag, str) or td.tag.lower() not in ("td", "th"):
                    continue
                while (ri, col) in occ:
                    col += 1
                cs = max(1, min(63, to_int(td.get("colspan"), 1) or 1))
                rs = max(1, min(len(rows) - ri, to_int(td.get("rowspan"), 1) or 1))
                cell = {"el": td, "row": ri, "col": col, "cs": cs, "rs": rs}
                for dr in range(rs):
                    for dc in range(cs):
                        occ[(ri + dr, col + dc)] = cell
                row_cells.append(cell)
                col += cs
            grid_cells.append(row_cells)
        ncols = max(c + 1 for (_, c) in occ) if occ else 1
        grid = self.compute_grid(tel, ncols, grid_cells)

        tbl = mk("tbl")
        tblpr = self.b64xml(tel.get("data-tblpr"))
        if tblpr is None:
            tblpr = mk("tblPr")
        tsid = cls_value(tel, "t-")
        if tsid not in self.st.styles:
            tsid = None
        tcss = parse_style(tel.get("style"))
        if tsid is None and not tel.get("data-tblpr"):
            tsid = self.ensure_table_grid()
            if tblpr.find(W("tblLook")) is None:
                insert_ordered(tblpr, mk("tblLook", val="04A0", firstRow=1, lastRow=0, firstColumn=1,
                                         lastColumn=0, noHBand=0, noVBand=1), TBLPR_ORDER)
            ml, mr = tcss.get("margin-left", ""), tcss.get("margin-right", "")
            if ml == "auto" and mr == "auto":
                insert_ordered(tblpr, mk("jc", val="center"), TBLPR_ORDER)
        if tsid:
            insert_ordered(tblpr, mk("tblStyle", val=tsid), TBLPR_ORDER)
        insert_ordered(tblpr, mk("tblW", w=sum(grid), type="dxa"), TBLPR_ORDER)
        if tblpr.find(W("tblLayout")) is None:
            insert_ordered(tblpr, mk("tblLayout", type="fixed"), TBLPR_ORDER)
        tbl.append(tblpr)
        tg = mk("tblGrid")
        for w in grid:
            tg.append(mk("gridCol", w=w))
        tbl.append(tg)
        table_border = to_int(tel.get("border"), 0) or 0
        ctx2 = {"props": ctx.get("props", {}), "pcss": {}}
        for ri, tr_el in enumerate(rows):
            tr = mk("tr")
            trpr = self.b64xml(tr_el.get("data-trpr"))
            if trpr is not None:
                tr.append(trpr)
            col = 0
            while col < ncols:
                cell = occ.get((ri, col))
                if cell is None:
                    tc = self.make_tc(None, grid[col:col + 1], 1, None, table_border)
                    tc.append(mk("p"))
                    tr.append(tc)
                    col += 1
                    continue
                widths = grid[cell["col"]:cell["col"] + cell["cs"]]
                if cell["row"] == ri:
                    vm = "restart" if cell["rs"] > 1 else None
                    tc = self.make_tc(cell["el"], widths, cell["cs"], vm, table_border)
                    cctx = self.inherit(ctx2, cell["el"])
                    self.walk_blocks(cell["el"], tc, cctx)
                else:
                    tc = self.make_tc(cell["el"], widths, cell["cs"], "continue", table_border)
                if not len(tc) or tc[-1].tag != W("p"):
                    tc.append(mk("p"))
                tr.append(tc)
                col = cell["col"] + cell["cs"]
            tbl.append(tr)
        self.append(container, tbl)

    def make_tc(self, td, widths, span, vmerge, table_border):
        tc = mk("tc")
        tcpr = self.b64xml(td.get("data-tcpr")) if td is not None else None
        if tcpr is None:
            tcpr = mk("tcPr")
        insert_ordered(tcpr, mk("tcW", w=sum(widths), type="dxa"), TCPR_ORDER)
        if span > 1:
            insert_ordered(tcpr, mk("gridSpan", val=span), TCPR_ORDER)
        if vmerge == "restart":
            insert_ordered(tcpr, mk("vMerge", val="restart"), TCPR_ORDER)
        elif vmerge == "continue":
            insert_ordered(tcpr, mk("vMerge"), TCPR_ORDER)
        css = parse_style(td.get("style")) if td is not None else {}
        sides = {}
        if any(k.startswith("border") for k in css):
            whole = css.get("border")
            for side in ("top", "left", "bottom", "right"):
                v = css.get("border-" + side, whole)
                if v is None and any(k in css for k in ("border-width", "border-style", "border-color")):
                    v = " ".join(css.get(k, "") for k in ("border-width", "border-style", "border-color"))
                spec = border_from_css(v) if v else "nil"
                sides[side] = spec or "nil"
        elif table_border > 0:
            sides = {s: ("single", 4, "auto") for s in ("top", "left", "bottom", "right")}
        else:
            sides = {s: "nil" for s in ("top", "left", "bottom", "right")}
        bd = mk("tcBorders")
        for side in BORDER_ORDER:
            if side in sides:
                bd.append(border_el(side, sides[side]))
        insert_ordered(tcpr, bd, TCPR_ORDER)
        bg = color_hex(css.get("background-color") or css.get("background"))
        if bg:
            insert_ordered(tcpr, mk("shd", val="clear", color="auto", fill=bg), TCPR_ORDER)
        va = css.get("vertical-align", "").lower() or (td.get("valign") or "").lower() if td is not None else ""
        if va in ("middle", "bottom"):
            insert_ordered(tcpr, mk("vAlign", val="center" if va == "middle" else "bottom"), TCPR_ORDER)
        tc.append(tcpr)
        return tc

    def compute_grid(self, tel, ncols, grid_cells):
        tcss = parse_style(tel.get("style"))
        table_w = length_tw(tcss.get("width"), self.text_width) or length_tw(tel.get("width") or "", self.text_width)
        ref = table_w or self.text_width
        widths = [None] * ncols
        cols = []
        for cg in tel:
            if isinstance(cg.tag, str) and cg.tag.lower() == "colgroup":
                for col in cg:
                    if isinstance(col.tag, str) and col.tag.lower() == "col":
                        cols.extend([col] * (to_int(col.get("span"), 1) or 1))
            elif isinstance(cg.tag, str) and cg.tag.lower() == "col":
                cols.append(cg)
        for i, col in enumerate(cols[:ncols]):
            v = parse_style(col.get("style")).get("width") or col.get("width")
            widths[i] = length_tw(v, ref) if v else None
        if None in widths:
            for row in grid_cells:
                for c in row:
                    if c["cs"] == 1 and widths[c["col"]] is None:
                        v = parse_style(c["el"].get("style")).get("width") or c["el"].get("width")
                        if v:
                            widths[c["col"]] = length_tw(v, ref)
        missing = [i for i, w in enumerate(widths) if not w]
        if missing:
            known = sum(w for w in widths if w)
            rest = max(ref - known, 500 * len(missing))
            for i in missing:
                widths[i] = int(rest / len(missing))
        return [max(100, int(w)) for w in widths]

    # -------------------------------------------------------------- страница
    def apply_page(self):
        pg = self.page
        if not pg:
            return
        sect = self.sect
        pgsz = sect.find(W("pgSz"))
        if pgsz is None:
            pgsz = insert_ordered(sect, mk("pgSz"), SECTPR_ORDER)
        w, h = to_int(pg.get("w"), 11906), to_int(pg.get("h"), 16838)
        pgsz.set(W("w"), str(w))
        pgsz.set(W("h"), str(h))
        if w > h:
            pgsz.set(W("orient"), "landscape")
        elif W("orient") in pgsz.attrib:
            del pgsz.attrib[W("orient")]
        mar = sect.find(W("pgMar"))
        if mar is None:
            mar = insert_ordered(sect, mk("pgMar"), SECTPR_ORDER)
        for k, default in (("top", 1134), ("right", 850), ("bottom", 1134), ("left", 1701),
                           ("header", 708), ("footer", 708), ("gutter", 0)):
            v = to_int(pg.get(k))
            if v is None:
                v = to_int(mar.get(W(k)), default)
            mar.set(W(k), str(v))

    def finish(self):
        # уникальные id рисунков
        n = 0
        for docpr in self.doc.element.body.iter("{%s}docPr" % NS["wp"]):
            n += 1
            docpr.set("id", str(n))
        # убираем ссылки на удалённые картинки и гиперссылки
        used = set()
        for el in self.doc.element.iter():
            for k, v in el.attrib.items():
                if k.startswith("{%s}" % R_NS):
                    used.add(v)
        for rid, rel in list(self.doc.part.rels.items()):
            if rel.reltype in (RT.IMAGE, RT.HYPERLINK) and rid not in used:
                try:
                    self.doc.part.drop_rel(rid)
                except Exception:
                    pass


def convert(payload, original_bytes=None):
    data = original_bytes
    if data is None and payload.get("original"):
        data = base64.b64decode(payload["original"])
    if data:
        check_zip(data)
        document = Document(io.BytesIO(data))
    else:
        document = Document(TEMPLATE)
    writer = HtmlToDocx(document, payload.get("page"))
    writer.run(payload.get("html") or "")
    buf = io.BytesIO()
    document.save(buf)
    return buf.getvalue(), sorted(writer.warnings)
