"""Готовит файлы для статического хостинга (GitHub Pages) и для сервера:

  static/tinymce-bundle.js — все скрипты TinyMCE одним файлом (один запрос вместо ~17);
  static/blank.json        — пустой документ, уже переведённый в HTML (новый документ открывается мгновенно).

Запускать после обновления TinyMCE, шаблона или конвертера:  python build_static.py
"""
import json
import os

import docx2html

HERE = os.path.dirname(os.path.abspath(__file__))
STATIC = os.path.join(HERE, "static")
TINY = os.path.join(STATIC, "vendor", "tinymce")
# порядок важен: ядро, модель, тема, иконки, плагины, язык
BUNDLE_PARTS = ["tinymce.min.js", "models/dom/model.min.js", "themes/silver/theme.min.js",
                "icons/default/icons.min.js"] + \
               ["plugins/%s/plugin.min.js" % p for p in
                ("advlist", "lists", "table", "image", "link", "searchreplace", "pagebreak", "charmap",
                 "wordcount", "visualchars", "visualblocks", "nonbreaking", "autolink")] + ["langs/ru.js"]


def main():
    parts = []
    for rel in BUNDLE_PARTS:
        with open(os.path.join(TINY, rel), "rb") as f:
            parts.append(f.read())
    with open(os.path.join(STATIC, "tinymce-bundle.js"), "wb") as f:
        f.write(b"\n;\n".join(parts))
    with open(os.path.join(HERE, "templates", "blank.docx"), "rb") as f:
        blank = docx2html.convert(f.read())
    with open(os.path.join(STATIC, "blank.json"), "w", encoding="utf-8") as f:
        json.dump(blank, f, ensure_ascii=False)
    print("static/tinymce-bundle.js и static/blank.json готовы")


if __name__ == "__main__":
    main()
