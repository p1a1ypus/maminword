"""Создаёт templates/blank.docx — шаблон нового документа (A4, поля по ГОСТ, Calibri 11)."""
import os

from docx import Document
from lxml import etree

from docx_common import NS, W, mk

HERE = os.path.dirname(os.path.abspath(__file__))


def main():
    doc = Document()
    styles = doc.styles.element
    dd = styles.find(W("docDefaults"))
    rpr = dd.find("%s/%s" % (W("rPrDefault"), W("rPr")))
    fonts = rpr.find(W("rFonts"))
    for attr in ("asciiTheme", "hAnsiTheme", "cstheme"):
        fonts.attrib.pop(W(attr), None)
    for attr in ("ascii", "hAnsi", "cs"):
        fonts.set(W(attr), "Calibri")
    lang = rpr.find(W("lang"))
    lang.set(W("val"), "ru-RU")
    ppr = dd.find("%s/%s" % (W("pPrDefault"), W("pPr")))
    sp = ppr.find(W("spacing"))
    sp.set(W("after"), "160")
    sp.set(W("line"), "259")

    if doc.styles.element.find(".//" + W("style") + "[@" + W("styleId") + "='Hyperlink']") is None:
        styles.append(etree.fromstring(
            '<w:style xmlns:w="%s" w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>'
            '<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/>'
            '<w:u w:val="single"/></w:rPr></w:style>' % NS["w"]))

    sec = doc.sections[0]
    sp_ = sec._sectPr
    pgsz = sp_.find(W("pgSz"))
    pgsz.set(W("w"), "11906")
    pgsz.set(W("h"), "16838")
    mar = sp_.find(W("pgMar"))
    for k, v in (("top", 1134), ("right", 850), ("bottom", 1134), ("left", 1701),
                 ("header", 708), ("footer", 708), ("gutter", 0)):
        mar.set(W(k), str(v))

    tab = doc.settings.element.find(W("defaultTabStop"))
    if tab is not None:
        tab.set(W("val"), "708")

    pkg = doc.part.package
    for rid, rel in list(pkg.rels.items()):
        if rel.reltype.endswith("/thumbnail"):
            pkg.rels.pop(rid)
    doc.core_properties.author = ""
    doc.core_properties.title = ""
    doc.core_properties.comments = ""
    doc.core_properties.last_modified_by = ""

    body = doc.element.body
    for ch in list(body):
        if ch.tag != W("sectPr"):
            body.remove(ch)
    body.insert(0, mk("p"))

    os.makedirs(os.path.join(HERE, "templates"), exist_ok=True)
    out = os.path.join(HERE, "templates", "blank.docx")
    doc.save(out)
    print("saved", out)


if __name__ == "__main__":
    main()
