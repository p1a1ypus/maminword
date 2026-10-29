"""DOCX -> HTML для редактора.

Стили Word превращаются в CSS-классы (s-* абзацные, c-* знаковые, t-* табличные),
прямое форматирование — в inline-стили. Всё, что редактор не умеет показывать
(сноски, диаграммы, надписи, формулы...), сохраняется как исходный XML в data-raw
и при сохранении возвращается в документ без изменений.
"""
import base64
import copy
import html
import io
import re

from docx import Document
from lxml import etree

from docx_common import (NS, W, WV, R_NS, HIGHLIGHT, Styles, border_css, check_zip, css_font, enc_id, fmt_num,
                         local, ppr_css, read_ppr, read_rpr, rpr_css, to_int, ui_name, wget)

PB = object()  # маркер разрыва страницы внутри абзаца
TAB_HTML = '<span class="docx-tab" contenteditable="false">\t</span>'
BROWSER_IMAGES = {"image/png", "image/jpeg", "image/gif", "image/bmp", "image/svg+xml", "image/webp"}

# символы из шрифтов Symbol/Wingdings, которые Word использует как маркеры списков
PUA_BULLETS = {"": "•", "": "▪", "": "➢", "": "✓", "": "❖", "": "◆",
               "": "⇒", "": "➔", "": "●", "": "■", "": "❑", "": "◊",
               "": "–", "": "—"}
NUMFMT_CSS = {"decimal": "decimal", "lowerLetter": "lower-alpha", "upperLetter": "upper-alpha",
              "lowerRoman": "lower-roman", "upperRoman": "upper-roman", "decimalZero": "decimal-leading-zero",
              "russianLower": "lw-ru-lower", "russianUpper": "lw-ru-upper", "none": "none"}


def b64(s):
    if isinstance(s, str):
        s = s.encode("utf-8")
    return base64.b64encode(s).decode("ascii")


def xml_b64(el):
    return b64(etree.tostring(el, encoding="utf-8"))


def esc(s):
    return html.escape(s, quote=True)


def text_html(s):
    s = esc(s or "")
    s = re.sub(r" {2,}", lambda m: " " * (len(m.group(0)) - 1) + " ", s)
    return s.replace("­", "&shy;")


def placeholder_svg(label, w=120, h=60):
    svg = ('<svg xmlns="http://www.w3.org/2000/svg" width="%d" height="%d" viewBox="0 0 %d %d">'
           '<rect x="0.5" y="0.5" width="%d" height="%d" fill="#f1f3f5" stroke="#adb5bd" stroke-dasharray="4 3"/>'
           '<text x="50%%" y="50%%" dominant-baseline="middle" text-anchor="middle" font-family="Segoe UI,Arial"'
           ' font-size="12" fill="#495057">%s</text></svg>') % (w, h, w, h, w - 1, h - 1, esc(label))
    return "data:image/svg+xml;base64," + b64(svg)


def all_text(el):
    return "".join(t.text or "" for t in el.iter(W("t")))


class Numbering:
    def __init__(self, document):
        self.abstract, self.nums, self.style_links = {}, {}, {}
        try:
            root = document.part.numbering_part.element
        except Exception:
            return
        for an in root.findall(W("abstractNum")):
            aid = an.get(W("abstractNumId"))
            levels = {}
            for lvl in an.findall(W("lvl")):
                levels[to_int(lvl.get(W("ilvl")), 0)] = self._lvl(lvl)
            self.abstract[aid] = levels
            link = an.find(W("numStyleLink"))
            if link is not None and not levels:
                self.style_links[aid] = link.get(WV)
        for num in root.findall(W("num")):
            ov = {}
            for lo in num.findall(W("lvlOverride")):
                entry = {}
                so = lo.find(W("startOverride"))
                if so is not None:
                    entry["start"] = to_int(so.get(WV), 1)
                lv = lo.find(W("lvl"))
                if lv is not None:
                    entry["lvl"] = self._lvl(lv)
                ov[to_int(lo.get(W("ilvl")), 0)] = entry
            self.nums[num.get(W("numId"))] = (wget(num.find(W("abstractNumId")), "val"), ov)
        self.document = document

    @staticmethod
    def _lvl(lvl):
        rpr = lvl.find(W("rPr"))
        font = None
        if rpr is not None and rpr.find(W("rFonts")) is not None:
            font = rpr.find(W("rFonts")).get(W("ascii"))
        return {"fmt": wget(lvl.find(W("numFmt")), "val", "decimal"),
                "text": wget(lvl.find(W("lvlText")), "val", ""),
                "start": to_int(wget(lvl.find(W("start")), "val"), 1),
                "ppr": read_ppr(lvl.find(W("pPr"))), "font": font}

    def _levels(self, aid, styles, depth=0):
        if aid in self.style_links and depth < 3:
            sid = self.style_links[aid]
            st = styles.styles.get(sid)
            nid = st["ppr"].get("numId") if st else None
            if nid in self.nums:
                return self._levels(self.nums[nid][0], styles, depth + 1)
        return self.abstract.get(aid, {})

    def level(self, num_id, ilvl, styles):
        if num_id not in self.nums:
            return None
        aid, ov = self.nums[num_id]
        lv = dict(self._levels(aid, styles).get(ilvl) or {"fmt": "decimal", "text": "%%%d." % (ilvl + 1),
                                                           "start": 1, "ppr": {}, "font": None})
        o = ov.get(ilvl)
        if o:
            if "lvl" in o:
                lv.update(o["lvl"])
            if "start" in o:
                lv["start"] = o["start"]
        return lv


class Acc:
    """Собирает HTML строк одного абзаца, разрезая его по разрывам страниц."""

    def __init__(self):
        self.pieces = [[]]
        self.link = None

    def add(self, s):
        self.pieces[-1].append(s)

    def open_link(self, href):
        self.close_link()
        self.link = href
        self.add('<a href="%s">' % esc(href))

    def close_link(self):
        if self.link is not None:
            self.add("</a>")
            self.link = None

    def page_break(self):
        link = self.link
        self.close_link()
        self.pieces.append(PB)
        self.pieces.append([])
        if link:
            self.open_link(link)

    def result(self):
        self.close_link()
        return [p if p is PB else "".join(p) for p in self.pieces]


class ListCtx:
    def __init__(self):
        self.stack = []

    def close_top(self, out):
        top = self.stack.pop()
        if top["li"]:
            out.append("</li>")
        out.append("</%s>" % top["tag"])

    def close_all(self, out):
        while self.stack:
            self.close_top(out)


class DocxToHtml:
    def __init__(self, document):
        self.doc = document
        self.st = Styles(document)
        self.num = Numbering(document)
        self.counts = {}
        self.fields = []
        self.fn_no = 0
        self.en_no = 0
        self.list_css = {}
        self.used_styles = set()
        self.fonts = set()
        self.warnings = set()

    # ------------------------------------------------------------- общий вход
    def convert(self):
        out = []
        self.blocks(self.doc.element.body, out, None)
        return {
            "html": "".join(out),
            "css": self.build_css(),
            "page": self.page_info(),
            "tab": fmt_num(self.default_tab() / 20.0),
            "styles": self.style_list(),
            "defaultStyle": enc_id(self.st.default.get("paragraph") or "Normal"),
            "nextStyles": {enc_id(s["id"]): enc_id(s["next"]) for s in self.st.styles.values()
                           if s["type"] == "paragraph" and s["next"] and s["next"] in self.st.styles},
            "fonts": sorted(self.fonts),
            "warnings": sorted(self.warnings),
        }

    def default_tab(self):
        try:
            el = self.doc.settings.element.find(W("defaultTabStop"))
            return to_int(el.get(WV), 720) or 720
        except Exception:
            return 720

    def page_info(self):
        sect = self.doc.element.body.find(W("sectPr"))
        pg = sect.find(W("pgSz")) if sect is not None else None
        mar = sect.find(W("pgMar")) if sect is not None else None
        d = {"w": to_int(wget(pg, "w"), 11906), "h": to_int(wget(pg, "h"), 16838),
             "orient": wget(pg, "orient", "portrait")}
        for k, default in (("top", 1134), ("right", 850), ("bottom", 1134), ("left", 1701),
                           ("header", 708), ("footer", 708), ("gutter", 0)):
            d[k] = abs(to_int(wget(mar, k), default) or 0)
        return d

    # ------------------------------------------------------------------ блоки
    def blocks(self, container, out, tbl_style, lists=None):
        own = lists is None
        lists = lists or ListCtx()
        for ch in container:
            t = local(ch)
            if t == "p":
                self.para(ch, out, lists, tbl_style)
            elif t == "tbl":
                lists.close_all(out)
                self.table(ch, out)
            elif t in ("sdt", "customXml"):
                inner = ch.find(W("sdtContent")) if t == "sdt" else ch
                if inner is not None:
                    self.blocks(inner, out, tbl_style, lists)
            elif t == "altChunk":
                lists.close_all(out)
                out.append('<p><span class="docx-raw" contenteditable="false" data-kind="b" data-raw="%s">'
                           '[Вложенный документ]</span></p>' % xml_b64(ch))
        if own:
            lists.close_all(out)

    def para(self, p, out, lists, tbl_style):
        ppr = p.find(W("pPr"))
        direct = read_ppr(ppr)
        sid = direct.get("pStyle")
        if sid not in self.st.styles:
            sid = self.st.default.get("paragraph")
        self.used_styles.add(sid)
        style_ppr = self.st.merged(sid, "ppr") if sid else {}
        num_id = direct.get("numId", style_ppr.get("numId"))
        ilvl = direct.get("ilvl", style_ppr.get("ilvl", 0)) or 0
        lvl = self.num.level(num_id, ilvl, self.st) if num_id and num_id != "0" else None

        acc = Acc()
        self.inline(p, acc)
        pieces = acc.result()

        attrs = []
        cls = "s-" + enc_id(sid) if sid else ""
        css = ppr_css({k: v for k, v in direct.items() if k not in ("pBdr",)}, with_indent=lvl is None)
        if ppr is not None and ppr.find(W("pBdr")) is not None:
            css += ppr_css({"pBdr": ppr.find(W("pBdr"))})
        keep = self.preserved_ppr(ppr)
        if keep is not None:
            attrs.append('data-ppr="%s"' % xml_b64(keep))

        if direct.get("pbb"):
            lists.close_all(out)
            out.append("<!-- pagebreak -->")

        texts = [x for x in pieces if x is not PB]
        only_breaks = len(pieces) > 1 and all(not x for x in texts)
        for i, piece in enumerate(pieces):
            if piece is PB:
                lists.close_all(out)
                out.append("<!-- pagebreak -->")
                continue
            last = i == len(pieces) - 1
            if not piece and (only_breaks or (len(pieces) > 1 and (last or i == 0))):
                continue
            content = piece
            pcss = list(css)
            if not content:
                content = "<br>"
                mark = read_rpr(ppr.find(W("rPr")) if ppr is not None else None, self.st.theme)
                if mark.get("size"):
                    pcss.append("font-size: %spt" % fmt_num(mark["size"]))
            if lvl is not None:
                self.list_item(out, lists, num_id, ilvl, lvl, direct, style_ppr, cls, pcss, attrs, content)
            else:
                lists.close_all(out)
                out.append("<p%s>%s</p>" % (self.attr_str(cls, pcss, attrs), content))

        sect = ppr.find(W("sectPr")) if ppr is not None else None
        if sect is not None:
            lists.close_all(out)
            kind = wget(sect.find(W("type")), "val", "nextPage")
            label = {"continuous": "Разрыв раздела (на текущей странице)", "evenPage": "Разрыв раздела (с чётной)",
                     "oddPage": "Разрыв раздела (с нечётной)"}.get(kind, "Разрыв раздела (со следующей страницы)")
            out.append('<div class="docx-section" contenteditable="false" data-sectpr="%s">%s</div>'
                       % (xml_b64(sect), label))

    @staticmethod
    def attr_str(cls, css, attrs):
        s = ""
        if cls:
            s += ' class="%s"' % cls
        if css:
            s += ' style="%s"' % esc("; ".join(css))
        if attrs:
            s += " " + " ".join(attrs)
        return s

    @staticmethod
    def preserved_ppr(ppr):
        if ppr is None:
            return None
        keep = copy.deepcopy(ppr)
        for ch in list(keep):
            if local(ch) in ("pStyle", "numPr", "jc", "ind", "spacing", "shd", "rPr", "sectPr",
                             "pageBreakBefore", "pPrChange"):
                keep.remove(ch)
        return keep if len(keep) else None

    # ---------------------------------------------------------------- списки
    def list_item(self, out, lists, num_id, ilvl, lvl, direct, style_ppr, cls, css, attrs, content):
        is_bullet = lvl["fmt"] == "bullet"
        tag = "ul" if is_bullet else "ol"
        stack = lists.stack
        while stack and (stack[-1]["lvl"] > ilvl or (stack[-1]["lvl"] == ilvl and
                                                       (stack[-1]["num"] != num_id or stack[-1]["tag"] != tag))):
            lists.close_top(out)
        # отступы: прямые > уровень нумерации > стиль
        ind = {}
        for src in (style_ppr, lvl["ppr"], direct):
            for k in ("ind_left", "ind_first", "ind_right"):
                if k in src:
                    ind[k] = src[k]
        left, first = ind.get("ind_left", 720 * (ilvl + 1)), ind.get("ind_first", -360)
        key = (num_id, ilvl)
        self.counts[key] = self.counts.get(key, 0) + 1
        for (n, l) in list(self.counts):
            if n == num_id and l > ilvl:
                del self.counts[(n, l)]
        if stack and stack[-1]["lvl"] == ilvl:
            top = stack[-1]
            if top["li"]:
                out.append("</li>")
        else:
            lcls = self.list_class(num_id, ilvl, lvl)
            lcss = ["padding-left: 0"]
            if not is_bullet:
                lcss.append("counter-reset: lwl%d %d" % (ilvl, lvl["start"] - 1 + self.counts[key] - 1))
            out.append('<%s class="%s" style="%s">' % (tag, lcls, "; ".join(lcss)))
            top = {"tag": tag, "num": num_id, "lvl": ilvl, "li": False, "left": left}
            stack.append(top)
        parent_left = stack[-2]["left"] if len(stack) > 1 else 0
        top["left"] = left
        li_css = [c for c in css if not c.startswith(("margin-left", "text-indent"))]
        li_css.append("margin-left: %spt" % fmt_num((left - parent_left) / 20.0))
        if first >= 0:
            li_css.append("list-style-position: inside")
        li_css.append("text-indent: %spt" % fmt_num(max(first, 0) / 20.0))
        li_attrs = list(attrs) + ['data-ind="%d,%d,%d"' % (left, first, ind.get("ind_right", 0)),
                                  'data-lvl="%d"' % ilvl]
        out.append("<li%s>%s" % (self.attr_str(cls, li_css, li_attrs), content))
        top["li"] = True

    def list_class(self, num_id, ilvl, lvl):
        name = "lw-n%s-l%d" % (re.sub(r"[^0-9A-Za-z]", "", str(num_id)), ilvl)
        if name in self.list_css:
            return name
        rules = []
        fmt, text = lvl["fmt"], lvl["text"] or ""
        if fmt == "bullet":
            ch = text[:1] or "•"
            ch = PUA_BULLETS.get(ch, ch)
            if ch in ("•", "·", "●"):
                lst = "disc"
            elif ch in ("o", "◦", "○"):
                lst = "circle"
            elif ch in ("▪", "■", "§", ""):
                lst = "square"
            elif "" <= ch <= "":
                lst = "disc"
            else:
                lst = '"%s  "' % ch.replace('"', "")
            rules.append(":where(ul.%s) { list-style-type: %s; }" % (name, lst))
        else:
            base = NUMFMT_CSS.get(fmt, "decimal")
            rules.append(":where(ol.%s) { list-style-type: %s; }" % (name, "none" if base == "none" else
                                                                     base if not base.startswith("lw-") else "decimal"))
            rules.append(":where(ol.%s > li) { counter-increment: lwl%d; }" % (name, ilvl))
            parts = re.split(r"(%\d)", text)
            content = []
            for part in parts:
                m = re.match(r"%(\d)", part)
                if m:
                    k = int(m.group(1)) - 1
                    kfmt = fmt if k == ilvl else (self.num.level(num_id, k, self.st) or {}).get("fmt", "decimal")
                    kcss = NUMFMT_CSS.get(kfmt, "decimal")
                    if kcss == "none":
                        continue
                    content.append("counter(lwl%d, %s)" % (k, kcss))
                elif part:
                    content.append('"%s"' % part.replace("\\", "\\\\").replace('"', '\\"'))
            content.append('"\\00a0\\00a0"')
            if base != "none":
                rules.append("ol.%s > li::marker { content: %s; }" % (name, " ".join(content)))
        self.list_css[name] = "\n".join(rules)
        return name

    # ------------------------------------------------------------ содержимое
    def field_visible(self):
        return all(f["state"] == "result" for f in self.fields)

    def inline(self, el, acc):
        for ch in el:
            t = local(ch)
            if ch.tag == W("r"):
                self.run(ch, acc)
            elif t == "hyperlink" and ch.tag == W("hyperlink"):
                rid = ch.get("{%s}id" % R_NS)
                href = None
                if rid and rid in self.doc.part.rels:
                    href = self.doc.part.rels[rid].target_ref
                elif ch.get(W("anchor")):
                    href = "#" + ch.get(W("anchor"))
                if href and self.field_visible():
                    acc.open_link(href)
                    self.inline(ch, acc)
                    acc.close_link()
                else:
                    self.inline(ch, acc)
            elif t in ("ins", "moveTo", "smartTag", "customXml", "fldSimple", "dir", "bdo") and \
                    ch.tag.startswith("{%s}" % NS["w"]):
                self.inline(ch, acc)
            elif t == "sdt":
                c = ch.find(W("sdtContent"))
                if c is not None:
                    self.inline(c, acc)
            elif ch.tag in ("{%s}oMath" % NS["m"], "{%s}oMathPara" % NS["m"]):
                label = "".join(x.text or "" for x in ch.iter("{%s}t" % NS["m"])) or "формула"
                acc.add('<span class="docx-raw docx-math" contenteditable="false" data-kind="p" data-raw="%s"'
                        ' title="Формула (сохраняется без изменений)">%s</span>' % (xml_b64(ch), esc(label)))
            elif t == "AlternateContent":
                self.raw_run_html(acc, ch, None, "Объект")

    def run_wrap(self, d):
        css, opens, closes = [], [], []
        if d.get("font"):
            self.fonts.add(d["font"])
            css.append("font-family: " + css_font(d["font"]))
        if d.get("size"):
            css.append("font-size: %spt" % fmt_num(d["size"]))
        if d.get("b") is False:
            css.append("font-weight: normal")
        if d.get("i") is False:
            css.append("font-style: normal")
        u = d.get("u")
        if u and u != "none":
            style = {"double": " double", "dotted": " dotted", "dottedHeavy": " dotted", "dash": " dashed",
                     "dashedHeavy": " dashed", "dashLong": " dashed", "wave": " wavy", "wavyDouble": " wavy",
                     "wavyHeavy": " wavy"}.get(u, "")
            css.append("text-decoration: underline" + style)
        if d.get("color") and d["color"] != "auto":
            css.append("color: " + d["color"])
        bg = d.get("hl") or d.get("shd")
        if bg:
            css.append("background-color: " + bg)
        if d.get("caps"):
            css.append("text-transform: uppercase")
        elif d.get("caps") is False:
            css.append("text-transform: none")
        if d.get("smallCaps"):
            css.append("font-variant: small-caps")
        if d.get("spacing"):
            css.append("letter-spacing: %spt" % fmt_num(d["spacing"]))
        if d.get("vanish"):
            css.append("display: none")
        tags = []
        if d.get("b"):
            tags.append("strong")
        if d.get("i"):
            tags.append("em")
        if d.get("strike"):
            tags.append("s")
        if d.get("va") == "superscript":
            tags.append("sup")
        elif d.get("va") == "subscript":
            tags.append("sub")
        for tg in tags:
            opens.append("<%s>" % tg)
            closes.insert(0, "</%s>" % tg)
        rs = d.get("rStyle")
        cls = ""
        if rs and rs in self.st.styles:
            cls = ' class="c-%s"' % enc_id(rs)
            self.used_styles.add(rs)
        if css or cls:
            opens.append('<span%s%s>' % (cls, ' style="%s"' % esc("; ".join(css)) if css else ""))
            closes.insert(0, "</span>")
        return "".join(opens), "".join(closes)

    def run(self, r, acc):
        rpr = r.find(W("rPr"))
        d = read_rpr(rpr, self.st.theme)
        op, cl = self.run_wrap(d)
        buf = []

        def flush():
            if buf:
                acc.add(op + "".join(buf) + cl)
                buf.clear()

        for ch in r:
            t = local(ch)
            if t == "fldChar":
                kind = ch.get(W("fldCharType"))
                if kind == "begin":
                    self.fields.append({"state": "instr", "instr": "", "link": False})
                elif kind == "separate" and self.fields:
                    f = self.fields[-1]
                    f["state"] = "result"
                    m = re.match(r'\s*HYPERLINK\s+(?:\\l\s+)?"([^"]+)"', f["instr"])
                    if m and self.field_visible():
                        flush()
                        url = m.group(1)
                        if "\\l" in f["instr"].split('"')[0]:
                            url = "#" + url
                        acc.open_link(url)
                        f["link"] = True
                elif kind == "end" and self.fields:
                    f = self.fields.pop()
                    if f["link"]:
                        flush()
                        acc.close_link()
                continue
            if t == "instrText":
                if self.fields:
                    self.fields[-1]["instr"] += ch.text or ""
                continue
            if not self.field_visible():
                continue
            if t == "t":
                buf.append(text_html(ch.text))
            elif t in ("tab", "ptab"):
                buf.append(TAB_HTML)
            elif t == "br":
                if ch.get(W("type")) == "page":
                    flush()
                    acc.page_break()
                else:
                    buf.append("<br>")
            elif t == "cr":
                buf.append("<br>")
            elif t == "noBreakHyphen":
                buf.append("‑")
            elif t == "softHyphen":
                buf.append("&shy;")
            elif t == "sym":
                try:
                    code = int(ch.get(W("char")) or "0", 16)
                except ValueError:
                    code = 0
                c = chr(code) if code else "?"
                flush()
                self.raw_run_html(acc, ch, rpr, PUA_BULLETS.get(c, c if code < 0xF000 else "□"), chip=False)
            elif t == "drawing":
                flush()
                acc.add(self.drawing(ch, rpr))
            elif t in ("pict", "object"):
                flush()
                acc.add(self.vml(ch, rpr))
            elif t == "footnoteReference":
                flush()
                self.fn_no += 1
                self.raw_run_html(acc, ch, rpr, str(self.fn_no), cls="docx-fn", tag="sup",
                                  title="Сноска (текст сноски сохраняется)")
            elif t == "endnoteReference":
                flush()
                self.en_no += 1
                self.raw_run_html(acc, ch, rpr, "i" * min(self.en_no, 3), cls="docx-fn", tag="sup",
                                  title="Концевая сноска")
            elif t == "AlternateContent":
                flush()
                self.raw_run_html(acc, ch, rpr, "Объект")
        flush()

    def raw_run(self, child, rpr):
        r = etree.Element(W("r"))
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
        r.append(copy.deepcopy(child))
        return r

    def raw_run_html(self, acc, child, rpr, label, cls="", tag="span", title=None, chip=True):
        txt = all_text(child).strip()
        if label == "Объект":
            if child.find(".//{%s}txbx" % NS["wps"]) is not None or child.find(".//" + W("txbxContent")) is not None:
                label = "Надпись" + (": " + txt[:60] if txt else "")
            elif child.find(".//{%s}chart" % "http://schemas.openxmlformats.org/drawingml/2006/chart") is not None:
                label = "Диаграмма"
            else:
                label = "Фигура" + (": " + txt[:60] if txt else "")
        classes = "docx-raw" + (" docx-chip" if chip else "") + (" " + cls if cls else "")
        acc.add('<%s class="%s" contenteditable="false" data-kind="r" data-raw="%s" title="%s">%s</%s>'
                % (tag, classes, xml_b64(self.raw_run(child, rpr)),
                   esc(title or "Объект Word — сохраняется без изменений"), esc(label), tag))

    def image_src(self, rid, w_px, h_px):
        part = self.doc.part.related_parts.get(rid) if rid else None
        if part is None:
            return placeholder_svg("нет рисунка", w_px, h_px)
        ct = getattr(part, "content_type", "") or ""
        if ct in BROWSER_IMAGES:
            return "data:%s;base64,%s" % (ct, b64(part.blob))
        ext = (ct.split("/")[-1].replace("x-", "") or "IMG").upper()
        self.warnings.add("Некоторые рисунки (%s) браузер не показывает — они сохранятся в файле как были." % ext)
        return placeholder_svg(ext, max(40, w_px), max(24, h_px))

    def drawing(self, dr, rpr):
        holder = dr.find("{%s}inline" % NS["wp"])
        anchor = holder is None
        if anchor:
            holder = dr.find("{%s}anchor" % NS["wp"])
        blip = dr.find(".//{%s}blip" % NS["a"])
        complex_obj = (dr.find(".//{%s}txbx" % NS["wps"]) is not None or
                       dr.find(".//" + W("txbxContent")) is not None or
                       dr.find(".//{%s}wgp" % "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup") is not None)
        if holder is None or blip is None or complex_obj:
            acc = Acc()
            self.raw_run_html(acc, dr, rpr, "Объект")
            return "".join(acc.result()[0])
        ext = holder.find("{%s}extent" % NS["wp"])
        cx, cy = to_int(wget_plain(ext, "cx"), 914400), to_int(wget_plain(ext, "cy"), 914400)
        w_px, h_px = max(1, round(cx / 9525)), max(1, round(cy / 9525))
        rid = blip.get("{%s}embed" % R_NS) or blip.get("{%s}link" % R_NS)
        src = self.image_src(rid, w_px, h_px)
        cls = ' class="docx-anchor" title="Плавающий рисунок: положение на странице сохранится"' if anchor else ""
        return ('<img src="%s" width="%d" height="%d"%s data-raw="%s" data-w="%d" data-h="%d" alt="">'
                % (src, w_px, h_px, cls, xml_b64(self.raw_run(dr, rpr)), w_px, h_px))

    def vml(self, el, rpr):
        imd = el.find(".//{%s}imagedata" % NS["v"])
        shape = el.find(".//{%s}shape" % NS["v"])
        if imd is None or shape is None:
            acc = Acc()
            self.raw_run_html(acc, el, rpr, "Объект")
            return "".join(acc.result()[0])
        style = shape.get("style", "")

        def dim(name, default):
            m = re.search(name + r"\s*:\s*([\d.]+)(pt|px|in|cm|mm)?", style)
            if not m:
                return default
            v = float(m.group(1))
            return v * {"pt": 4 / 3, "px": 1, "in": 96, "cm": 37.8, "mm": 3.78, None: 4 / 3}[m.group(2)]

        w_px, h_px = round(dim("width", 100)), round(dim("height", 100))
        src = self.image_src(imd.get("{%s}id" % R_NS), w_px, h_px)
        return ('<img src="%s" width="%d" height="%d" class="docx-vml" data-raw="%s" data-w="%d" data-h="%d" alt="">'
                % (src, w_px, h_px, xml_b64(self.raw_run(el, rpr)), w_px, h_px))

    # ---------------------------------------------------------------- таблицы
    def table(self, tbl, out):
        tblpr = tbl.find(W("tblPr"))
        tsid = wget(tblpr.find(W("tblStyle")) if tblpr is not None else None, "val")
        if tsid not in self.st.styles:
            tsid = self.st.default.get("table")
        if tsid:
            self.used_styles.add(tsid)
        grid_el = tbl.find(W("tblGrid"))
        grid = [to_int(g.get(W("w")), 0) or 0 for g in grid_el.findall(W("gridCol"))] if grid_el is not None else []

        borders = dict(self.st.table_borders(tsid)) if tsid else {}
        if tblpr is not None and tblpr.find(W("tblBorders")) is not None:
            for b in tblpr.find(W("tblBorders")):
                borders[local(b)] = b

        # строим сетку: (tc, row, col, span), объединения по вертикали -> rowspan
        rows = tbl.findall(W("tr"))
        cells = []
        for ri, tr in enumerate(rows):
            trpr = tr.find(W("trPr"))
            col = to_int(wget(trpr.find(W("gridBefore")) if trpr is not None else None, "val"), 0) or 0
            row_cells = []
            for tc in tr.findall(W("tc")):
                tcpr = tc.find(W("tcPr"))
                span = to_int(wget(tcpr.find(W("gridSpan")) if tcpr is not None else None, "val"), 1) or 1
                vm = tcpr.find(W("vMerge")) if tcpr is not None else None
                vmerge = None if vm is None else (vm.get(WV) or "continue")
                row_cells.append({"tc": tc, "col": col, "span": span, "vm": vmerge, "rowspan": 1})
                col += span
            cells.append(row_cells)
        ncols = max([len(grid)] + [c["col"] + c["span"] for rc in cells for c in rc] or [1])
        while len(grid) < ncols:
            grid.append(grid[-1] if grid else 2000)
        for ri, rc in enumerate(cells):
            for c in rc:
                if c["vm"] == "restart":
                    n = ri + 1
                    while n < len(cells):
                        below = [x for x in cells[n] if x["col"] == c["col"] and x["vm"] == "continue"]
                        if not below:
                            break
                        below[0]["skip"] = True
                        c["rowspan"] += 1
                        n += 1

        tcss = ["border-collapse: collapse", "table-layout: fixed",
                "width: %spt" % fmt_num(sum(grid[:ncols]) / 20.0)]
        if tblpr is not None:
            jc = wget(tblpr.find(W("jc")), "val")
            ind = to_int(wget(tblpr.find(W("tblInd")), "w"), 0) or 0
            if jc == "center":
                tcss += ["margin-left: auto", "margin-right: auto"]
            elif jc in ("right", "end"):
                tcss += ["margin-left: auto", "margin-right: 0"]
            elif ind:
                tcss.append("margin-left: %spt" % fmt_num(ind / 20.0))
        keep = None
        if tblpr is not None:
            keep = copy.deepcopy(tblpr)
            for ch in list(keep):
                if local(ch) in ("tblStyle", "tblW", "tblPrChange"):
                    keep.remove(ch)
        attrs = ' class="t-%s"' % enc_id(tsid) if tsid else ""
        attrs += ' style="%s"' % "; ".join(tcss)
        if keep is not None and len(keep):
            attrs += ' data-tblpr="%s"' % xml_b64(keep)
        out.append("<table%s><colgroup>" % attrs)
        for w in grid[:ncols]:
            out.append('<col style="width: %spt">' % fmt_num(w / 20.0))
        out.append("</colgroup><tbody>")
        last_row = len(cells) - 1
        for ri, (tr, rc) in enumerate(zip(rows, cells)):
            trpr = tr.find(W("trPr"))
            tr_attrs = ""
            if trpr is not None:
                h = to_int(wget(trpr.find(W("trHeight")), "val"), 0)
                if h:
                    tr_attrs += ' style="height: %spt"' % fmt_num(h / 20.0)
                tr_attrs += ' data-trpr="%s"' % xml_b64(trpr)
            out.append("<tr%s>" % tr_attrs)
            for c in rc:
                if c.get("skip"):
                    continue
                out.append(self.cell(c, borders, ri, ri + c["rowspan"] - 1 == last_row or
                                     ri + c["rowspan"] - 1 >= last_row, ncols, tsid))
            out.append("</tr>")
        out.append("</tbody></table>")

    def cell(self, c, borders, ri, is_last_row, ncols, tsid):
        tc = c["tc"]
        tcpr = tc.find(W("tcPr"))
        own = {}
        if tcpr is not None and tcpr.find(W("tcBorders")) is not None:
            for b in tcpr.find(W("tcBorders")):
                own[local(b)] = b
        sides = {
            "top": own.get("top", borders.get("top") if ri == 0 else borders.get("insideH")),
            "bottom": own.get("bottom", borders.get("bottom") if is_last_row else borders.get("insideH")),
            "left": own.get("left", own.get("start", borders.get("left", borders.get("start")) if c["col"] == 0
                                    else borders.get("insideV"))),
            "right": own.get("right", own.get("end", borders.get("right", borders.get("end"))
                                      if c["col"] + c["span"] >= ncols else borders.get("insideV"))),
        }
        css = []
        values = {k: border_css(v) for k, v in sides.items()}
        if any(v and v != "none" for v in values.values()):
            for k, v in values.items():
                css.append("border-%s: %s" % (k, v or "none"))
        if tcpr is not None:
            shd = tcpr.find(W("shd"))
            fill = shd.get(W("fill")) if shd is not None else None
            if fill and fill.lower() != "auto":
                css.append("background-color: #" + fill.upper())
            va = wget(tcpr.find(W("vAlign")), "val")
            if va in ("center", "bottom"):
                css.append("vertical-align: " + ("middle" if va == "center" else "bottom"))
        attrs = ""
        if c["span"] > 1:
            attrs += ' colspan="%d"' % c["span"]
        if c["rowspan"] > 1:
            attrs += ' rowspan="%d"' % c["rowspan"]
        if css:
            attrs += ' style="%s"' % "; ".join(css)
        if tcpr is not None:
            keep = copy.deepcopy(tcpr)
            for ch in list(keep):
                if local(ch) in ("tcW", "gridSpan", "vMerge", "hMerge", "tcBorders", "shd", "vAlign", "tcPrChange"):
                    keep.remove(ch)
            if len(keep):
                attrs += ' data-tcpr="%s"' % xml_b64(keep)
        inner = []
        self.blocks(tc, inner, tsid)
        if not inner:
            inner = ["<p><br></p>"]
        return "<td%s>%s</td>" % (attrs, "".join(inner))

    # -------------------------------------------------------------------- CSS
    def build_css(self):
        st = self.st
        css = []
        base = dict(st.def_rpr)
        base.setdefault("font", "Times New Roman")
        base.setdefault("size", 10)
        self.fonts.add(base["font"])
        css.append(".lw-doc { %s; }" % "; ".join(rpr_css(base)))
        css.append(":where(.lw-doc p, .lw-doc li) { %s }" % "; ".join(ppr_css(st.def_ppr) + [""]))

        def para_rules(sid, extra_sel=()):
            s = st.styles[sid]
            pp = st.merged(sid, "ppr")
            rp = st.merged(sid, "rpr")
            for k in ("font",):
                if rp.get(k):
                    self.fonts.add(rp[k])
            e = enc_id(sid)
            sel = ["p.s-%s" % e, "li.s-%s" % e] + list(extra_sel)
            body = ppr_css(pp, with_indent=False) + rpr_css(rp)
            if body:
                css.append(":where(%s) { %s; }" % (", ".join(sel), "; ".join(body)))
            ind = [c for c in ppr_css(pp) if c.startswith(("margin-", "text-indent"))]
            psel = [x for x in sel if not x.startswith("li")]
            if ind:
                css.append(":where(%s) { %s; }" % (", ".join(psel), "; ".join(ind)))
            if pp.get("ctx"):
                css.append(":where(.s-%s + .s-%s) { padding-top: 0; }" % (e, e))
                css.append(":where(.s-%s:has(+ .s-%s)) { padding-bottom: 0; }" % (e, e))
            if s.get("type") == "paragraph" and rp.get("vanish"):
                pass

        default = st.default.get("paragraph")
        if default in st.styles:
            para_rules(default, ('p:not([class*="s-"])',))
        for sid, s in st.styles.items():
            if s["type"] != "table":
                continue
            e = enc_id(sid)
            pp = st.merged(sid, "ppr")
            rp = st.merged(sid, "rpr")
            body = rpr_css(rp)
            pad = {}
            for tbl_st in st.chain(sid):
                tp = tbl_st["tblPr"]
                mar = tp.find(W("tblCellMar")) if tp is not None else None
                if mar is not None:
                    for side, el in (("left", mar.find(W("left")) if mar.find(W("left")) is not None else mar.find(W("start"))),
                                     ("right", mar.find(W("right")) if mar.find(W("right")) is not None else mar.find(W("end"))),
                                     ("top", mar.find(W("top"))), ("bottom", mar.find(W("bottom")))):
                        if el is not None:
                            pad[side] = "padding-%s: %spt" % (side, fmt_num((to_int(el.get(W("w")), 0) or 0) / 20.0))
            if body or pad:
                css.append(":where(table.t-%s) > tbody > tr > :where(td, th) { %s; }"
                           % (e, "; ".join(body + list(pad.values()))))
            pbody = ppr_css(pp)
            if pbody:
                css.append(":where(table.t-%s td p, table.t-%s td li) { %s; }" % (e, e, "; ".join(pbody)))
        for sid, s in st.styles.items():
            if s["type"] == "paragraph" and sid != default:
                extra = ()
                m = re.match(r"heading (\d)$", s["name"].lower())
                if m:
                    extra = ("h%s:not([class])" % m.group(1),)
                para_rules(sid, extra)
        for sid, s in st.styles.items():
            if s["type"] == "character":
                body = rpr_css(st.merged(sid, "rpr"))
                if body:
                    css.append(":where(span.c-%s) { %s; }" % (enc_id(sid), "; ".join(body)))
        css.extend(self.list_css.values())
        return "\n".join(css)

    def style_list(self):
        out = []
        for sid, s in self.st.styles.items():
            if s["type"] not in ("paragraph",):
                continue
            used = sid in self.used_styles
            if not used and (s["hidden"] or not s["qformat"]):
                continue
            out.append({"id": enc_id(sid), "name": s["name"], "ui": ui_name(s["name"]),
                        "prio": s["prio"] if not used else min(s["prio"], 50), "used": used})
        default = enc_id(self.st.default.get("paragraph") or "")
        out.sort(key=lambda x: (x["id"] != default, x["prio"], x["ui"]))
        return out


def wget_plain(el, name):
    return None if el is None else el.get(name)


def convert(data):
    check_zip(data)
    try:
        document = Document(io.BytesIO(data))
    except Exception as e:
        raise ValueError("Не удалось открыть файл как документ Word (.docx). "
                         "Если это старый формат .doc — пересохраните его в Word как .docx. (%s)" % e)
    return DocxToHtml(document).convert()
