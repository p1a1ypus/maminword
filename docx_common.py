"""Общие вещи для чтения/записи DOCX: пространства имён, единицы, стили Word."""
import io
import re
import zipfile

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "v": "urn:schemas-microsoft-com:vml",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}
W_NS = NS["w"]
R_NS = NS["r"]


def W(tag):
    return "{%s}%s" % (W_NS, tag)


WV = W("val")


def local(el):
    tag = el.tag
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


def is_w(el, name):
    return el.tag == W(name)


def wget(el, name, default=None):
    if el is None:
        return default
    v = el.get(W(name))
    return default if v is None else v


def is_on(el):
    """Значение «переключателя» вида <w:b/>, <w:b w:val="0"/>."""
    v = el.get(WV)
    return v is None or v.lower() not in ("0", "false", "off", "none")


def to_int(v, default=None):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return default


def check_zip(data, max_unpacked=400 * 1024 * 1024, max_files=10000):
    """Защита от «zip-бомб»: сервер доступен по ссылке, файлы присылают чужие люди."""
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            infos = z.infolist()
    except zipfile.BadZipFile:
        raise ValueError("Файл повреждён или это не документ .docx.")
    if len(infos) > max_files or sum(i.file_size for i in infos) > max_unpacked:
        raise ValueError("Документ слишком большой для обработки.")


# ---------------------------------------------------------------- единицы
def tw2pt(tw):
    return round(tw / 20.0, 2)


def fmt_num(x):
    s = ("%.2f" % x).rstrip("0").rstrip(".")
    return s if s not in ("-0", "") else "0"


CSS_LINE_FACTOR = 1.15  # «одинарный» интервал Word ≈ 1.15 высоты шрифта в CSS

HIGHLIGHT = {
    "yellow": "#FFFF00", "green": "#00FF00", "cyan": "#00FFFF", "magenta": "#FF00FF",
    "blue": "#0000FF", "red": "#FF0000", "darkBlue": "#000080", "darkCyan": "#008080",
    "darkGreen": "#008000", "darkMagenta": "#800080", "darkRed": "#800000",
    "darkYellow": "#808000", "darkGray": "#808080", "lightGray": "#C0C0C0",
    "black": "#000000", "white": "#FFFFFF",
}
HIGHLIGHT_BY_RGB = {v: k for k, v in HIGHLIGHT.items()}

SERIF_HINTS = ("times", "cambria", "georgia", "garamond", "book", "serif", "roman", "palatino", "bodoni")


def css_font(name):
    generic = "serif" if any(h in name.lower() for h in SERIF_HINTS) else "sans-serif"
    return "'%s', %s" % (name.replace("'", ""), generic)


# ------------------------------------------------- имена CSS-классов <-> styleId
def enc_id(sid):
    out = []
    for ch in sid:
        if (ch.isalnum() or ch == "-") and ch != "_":
            out.append(ch)
        else:
            out.append("_%x_" % ord(ch))
    return "".join(out)


def dec_id(s):
    return re.sub(r"_([0-9a-f]+)_", lambda m: chr(int(m.group(1), 16)), s)


# ------------------------------------------- порядок дочерних элементов (схема)
PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
             "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
             "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
             "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
             "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
             "sectPr", "pPrChange"]
RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
             "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
             "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText",
             "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"]
TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
              "textDirection", "tcFitText", "vAlign", "hideMark"]
TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
               "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar",
               "tblLook", "tblCaption", "tblDescription"]
SECTPR_ORDER = ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar",
                "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote",
                "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"]
BORDER_ORDER = ["top", "left", "start", "bottom", "right", "end", "insideH", "insideV", "between", "bar"]


def insert_ordered(parent, child, order):
    """Вставляет child в parent, соблюдая порядок элементов по схеме OOXML."""
    name = local(child)
    old = parent.find(W(name))
    if old is not None and name not in ("headerReference", "footerReference"):
        parent.replace(old, child)
        return child
    try:
        idx = order.index(name)
    except ValueError:
        parent.append(child)
        return child
    later = set(order[idx + 1:])
    for i, sib in enumerate(parent):
        if local(sib) in later:
            parent.insert(i, child)
            return child
    parent.append(child)
    return child


def mk(tag, **attrs):
    el = etree.Element(W(tag))
    for k, v in attrs.items():
        el.set(W(k), str(v))
    return el


# ------------------------------------------------------------- чтение свойств
def theme_fonts(document):
    try:
        part = document.part.part_related_by(RT.THEME)
        root = etree.fromstring(part.blob)
    except Exception:
        return {}
    a = NS["a"]
    out = {}
    for key, tag in (("major", "majorFont"), ("minor", "minorFont")):
        el = root.find(".//{%s}%s/{%s}latin" % (a, tag, a))
        if el is not None and el.get("typeface"):
            out[key] = el.get("typeface")
    return out


def font_from_rfonts(el, theme):
    for attr in ("hAnsiTheme", "hAnsi", "asciiTheme", "ascii", "cs"):
        v = el.get(W(attr))
        if not v:
            continue
        if attr.endswith("Theme"):
            f = theme.get("major" if v.startswith("major") else "minor")
            if f:
                return f
            continue
        return v
    return None


def read_rpr(rpr, theme):
    d = {}
    if rpr is None:
        return d
    for ch in rpr:
        t = local(ch)
        if t == "rFonts":
            f = font_from_rfonts(ch, theme)
            if f:
                d["font"] = f
        elif t == "sz":
            v = to_int(ch.get(WV))
            if v:
                d["size"] = v / 2.0
        elif t in ("b", "i", "strike", "caps", "smallCaps", "vanish"):
            d[t] = is_on(ch)
        elif t == "dstrike":
            if is_on(ch):
                d["strike"] = True
        elif t == "u":
            d["u"] = ch.get(WV) or "single"
        elif t == "color":
            v = ch.get(WV)
            d["color"] = "auto" if (not v or v == "auto") else "#" + v.upper()
        elif t == "highlight":
            v = ch.get(WV)
            d["hl"] = HIGHLIGHT.get(v)
        elif t == "shd":
            fill = ch.get(W("fill"))
            if fill and fill.lower() != "auto":
                d["shd"] = "#" + fill.upper()
        elif t == "vertAlign":
            d["va"] = ch.get(WV)
        elif t == "spacing":
            v = to_int(ch.get(WV))
            if v is not None:
                d["spacing"] = v / 20.0
        elif t == "rStyle":
            d["rStyle"] = ch.get(WV)
    return d


def read_ppr(ppr):
    d = {}
    if ppr is None:
        return d
    for ch in ppr:
        t = local(ch)
        if t == "pStyle":
            d["pStyle"] = ch.get(WV)
        elif t == "jc":
            d["jc"] = ch.get(WV)
        elif t == "ind":
            for attrs, key in ((("left", "start"), "ind_left"), (("right", "end"), "ind_right")):
                for a in attrs:
                    v = to_int(ch.get(W(a)))
                    if v is not None:
                        d[key] = v
                        break
            v = to_int(ch.get(W("hanging")))
            if v is not None:
                d["ind_first"] = -v
            else:
                v = to_int(ch.get(W("firstLine")))
                if v is not None:
                    d["ind_first"] = v
        elif t == "spacing":
            for a in ("before", "after", "line"):
                v = to_int(ch.get(W(a)))
                if v is not None:
                    d["sp_" + a] = v
            if ch.get(W("lineRule")):
                d["sp_rule"] = ch.get(W("lineRule"))
            if ch.get(W("beforeAutospacing")) in ("1", "true", "on"):
                d["sp_before"] = 280
            if ch.get(W("afterAutospacing")) in ("1", "true", "on"):
                d["sp_after"] = 280
        elif t == "contextualSpacing":
            d["ctx"] = is_on(ch)
        elif t == "pageBreakBefore":
            d["pbb"] = is_on(ch)
        elif t == "numPr":
            il = ch.find(W("ilvl"))
            ni = ch.find(W("numId"))
            if il is not None:
                d["ilvl"] = to_int(il.get(WV), 0)
            if ni is not None:
                d["numId"] = ni.get(WV)
        elif t == "shd":
            fill = ch.get(W("fill"))
            if fill and fill.lower() != "auto":
                d["shd"] = "#" + fill.upper()
        elif t == "pBdr":
            d["pBdr"] = ch
    return d


def border_css(el):
    """<w:top w:val="single" w:sz="4" w:color="auto"/> -> CSS-значение рамки."""
    if el is None:
        return None
    val = (el.get(WV) or "none").lower()
    if val in ("nil", "none"):
        return "none"
    sz = to_int(el.get(W("sz")), 4) or 4
    width = max(0.5, sz / 8.0)
    color = el.get(W("color"))
    color = "#000000" if not color or color == "auto" else "#" + color.upper()
    style = "solid"
    if val == "double":
        style = "double"
        width = max(width * 3, 2.25)
    elif val.startswith("dot"):
        style = "dotted"
    elif "dash" in val:
        style = "dashed"
    return "%spt %s %s" % (fmt_num(width), style, color)


def rpr_css(d):
    css = []
    if d.get("font"):
        css.append("font-family: " + css_font(d["font"]))
    if d.get("size"):
        css.append("font-size: %spt" % fmt_num(d["size"]))
    if "b" in d:
        css.append("font-weight: " + ("bold" if d["b"] else "normal"))
    if "i" in d:
        css.append("font-style: " + ("italic" if d["i"] else "normal"))
    deco = []
    if d.get("u") and d["u"] != "none":
        deco.append("underline")
    if d.get("strike"):
        deco.append("line-through")
    if deco:
        css.append("text-decoration: " + " ".join(deco))
    if d.get("color") and d["color"] != "auto":
        css.append("color: " + d["color"])
    if d.get("caps"):
        css.append("text-transform: uppercase")
    if d.get("smallCaps"):
        css.append("font-variant: small-caps")
    if d.get("spacing"):
        css.append("letter-spacing: %spt" % fmt_num(d["spacing"]))
    return css


def line_css(line, rule):
    if line is None:
        return None
    rule = rule or "auto"
    if rule == "auto":
        return fmt_num(line / 240.0 * CSS_LINE_FACTOR)
    if rule == "exact":
        return "%spt" % fmt_num(line / 20.0)
    return "max(%spt, %sem)" % (fmt_num(line / 20.0), fmt_num(CSS_LINE_FACTOR))


def ppr_css(d, with_indent=True):
    css = []
    jc = d.get("jc")
    if jc:
        css.append("text-align: " + {"both": "justify", "distribute": "justify", "center": "center",
                                     "right": "right", "end": "right"}.get(jc, "left"))
    if with_indent:
        if "ind_left" in d:
            css.append("margin-left: %spt" % fmt_num(d["ind_left"] / 20.0))
        if "ind_right" in d:
            css.append("margin-right: %spt" % fmt_num(d["ind_right"] / 20.0))
        if "ind_first" in d:
            css.append("text-indent: %spt" % fmt_num(d["ind_first"] / 20.0))
    if "sp_before" in d:
        css.append("padding-top: %spt" % fmt_num(d["sp_before"] / 20.0))
    if "sp_after" in d:
        css.append("padding-bottom: %spt" % fmt_num(d["sp_after"] / 20.0))
    lh = line_css(d.get("sp_line"), d.get("sp_rule"))
    if lh:
        css.append("line-height: " + lh)
    if d.get("shd"):
        css.append("background-color: " + d["shd"])
    bdr = d.get("pBdr")
    if bdr is not None:
        for side in ("top", "left", "bottom", "right"):
            b = border_css(bdr.find(W(side)))
            if b and b != "none":
                css.append("border-%s: %s" % (side, b))
    return css


# ----------------------------------------------------------------- стили Word
RU_NAMES = {
    "normal": "Обычный", "no spacing": "Без интервала", "title": "Название", "subtitle": "Подзаголовок",
    "quote": "Цитата 2", "intense quote": "Выделенная цитата", "list paragraph": "Абзац списка",
    "caption": "Название объекта", "header": "Верхний колонтитул", "footer": "Нижний колонтитул",
    "footnote text": "Текст сноски", "hyperlink": "Гиперссылка", "strong": "Строгий",
    "emphasis": "Выделение", "toc heading": "Заголовок оглавления", "list bullet": "Маркированный список",
    "list number": "Нумерованный список", "body text": "Основной текст", "table grid": "Сетка таблицы",
}


def ui_name(name):
    low = name.lower()
    m = re.match(r"heading (\d)$", low)
    if m:
        return "Заголовок " + m.group(1)
    m = re.match(r"toc (\d)$", low)
    if m:
        return "Оглавление " + m.group(1)
    return RU_NAMES.get(low, name)


class Styles:
    def __init__(self, document):
        self.document = document
        self.theme = theme_fonts(document)
        root = document.styles.element
        dd = root.find(W("docDefaults"))
        rpr = ppr = None
        if dd is not None:
            rpr = dd.find("%s/%s" % (W("rPrDefault"), W("rPr")))
            ppr = dd.find("%s/%s" % (W("pPrDefault"), W("pPr")))
        self.def_rpr = read_rpr(rpr, self.theme)
        self.def_ppr = read_ppr(ppr)
        self.styles = {}
        self.default = {}
        for s in root.findall(W("style")):
            sid = s.get(W("styleId"))
            if not sid:
                continue
            typ = s.get(W("type"), "paragraph")
            name_el = s.find(W("name"))
            info = {
                "id": sid, "type": typ,
                "name": name_el.get(WV) if name_el is not None else sid,
                "basedOn": wget(s.find(W("basedOn")), "val"),
                "next": wget(s.find(W("next")), "val"),
                "ppr": read_ppr(s.find(W("pPr"))),
                "rpr": read_rpr(s.find(W("rPr")), self.theme),
                "tblPr": s.find(W("tblPr")),
                "qformat": s.find(W("qFormat")) is not None,
                "hidden": s.find(W("semiHidden")) is not None or s.find(W("hidden")) is not None,
                "prio": to_int(wget(s.find(W("uiPriority")), "val"), 99),
            }
            self.styles[sid] = info
            if s.get(W("default")) in ("1", "true", "on"):
                self.default[typ] = sid
        if "paragraph" not in self.default:
            self.default["paragraph"] = "Normal" if "Normal" in self.styles else None

    def chain(self, sid):
        out, seen = [], set()
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid)
            out.append(self.styles[sid])
            sid = self.styles[sid]["basedOn"]
        return list(reversed(out))

    def merged(self, sid, key):
        d = {}
        for st in self.chain(sid):
            d.update(st[key])
        return d

    def by_name(self, name, typ="paragraph"):
        low = name.lower()
        for sid, st in self.styles.items():
            if st["type"] == typ and st["name"].lower() == low:
                return sid
        return None

    def table_borders(self, sid):
        """Рамки таблицы из цепочки табличных стилей: {side: element}."""
        out = {}
        for st in self.chain(sid):
            tp = st["tblPr"]
            if tp is None:
                continue
            tb = tp.find(W("tblBorders"))
            if tb is None:
                continue
            for b in tb:
                out[local(b)] = b
        return out
