/* МаминWord — интерфейс редактора (TinyMCE + конвертация DOCX на сервере). */
(function () {
  'use strict';

  var TW_PER_CM = 566.929;
  var LINE_FACTOR = 1.15; // «одинарный» интервал Word ≈ 1.15 высоты шрифта
  var PX_PER_PT = 96 / 72;
  var TAB_HTML = '<span class="docx-tab" contenteditable="false">\t</span>';

  var PAPER = [
    { id: 'A4', name: 'A4 (21 × 29,7 см)', w: 11906, h: 16838 },
    { id: 'A5', name: 'A5 (14,8 × 21 см)', w: 8391, h: 11906 },
    { id: 'A3', name: 'A3 (29,7 × 42 см)', w: 16838, h: 23811 },
    { id: 'Letter', name: 'Letter (21,59 × 27,94 см)', w: 12240, h: 15840 },
    { id: 'Legal', name: 'Legal (21,59 × 35,56 см)', w: 12240, h: 20160 }
  ];
  var MARGINS = [
    { id: 'normal', name: 'Обычные: 2 / 2 / 3 / 1,5 см', top: 1134, bottom: 1134, left: 1701, right: 850 },
    { id: 'narrow', name: 'Узкие: 1,27 см', top: 720, bottom: 720, left: 720, right: 720 },
    { id: 'moderate', name: 'Средние: 2,54 / 1,91 см', top: 1440, bottom: 1440, left: 1080, right: 1080 },
    { id: 'wide', name: 'Широкие: 2,54 / 5,08 см', top: 1440, bottom: 1440, left: 2880, right: 2880 }
  ];
  var FONTS = ['Times New Roman', 'Arial', 'Calibri', 'Calibri Light', 'Cambria', 'Aptos', 'Georgia', 'Verdana',
    'Tahoma', 'Segoe UI', 'Trebuchet MS', 'Century Gothic', 'Arial Narrow', 'Garamond', 'Book Antiqua',
    'Courier New', 'Consolas', 'PT Sans', 'PT Serif'];
  var FONT_STEPS = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
  var COLOR_MAP = [
    '000000', 'Чёрный', '404040', 'Тёмно-серый', '7F7F7F', 'Серый', 'BFBFBF', 'Светло-серый', 'FFFFFF', 'Белый',
    '44546A', 'Сине-серый', '4472C4', 'Синий (тема)', 'ED7D31', 'Оранжевый (тема)', '70AD47', 'Зелёный (тема)',
    '5B9BD5', 'Голубой (тема)',
    'C00000', 'Тёмно-красный', 'FF0000', 'Красный', 'FFC000', 'Оранжевый', 'FFFF00', 'Жёлтый', '92D050',
    'Светло-зелёный', '00B050', 'Зелёный', '00B0F0', 'Голубой', '0070C0', 'Синий', '002060', 'Тёмно-синий',
    '7030A0', 'Фиолетовый',
    '00FF00', 'Ярко-зелёный (выделение)', '00FFFF', 'Бирюзовый (выделение)', 'FF00FF', 'Розовый (выделение)',
    '0000FF', 'Синий (выделение)', '000080', 'Тёмно-синий (выделение)', '008080', 'Сине-зелёный (выделение)',
    '008000', 'Зелёный (выделение)', '800080', 'Фиолетовый (выделение)', '800000', 'Бордовый (выделение)',
    '808000', 'Оливковый (выделение)', '808080', 'Серый 50% (выделение)', 'C0C0C0', 'Серый 25% (выделение)'
  ];

  var BASE_CSS = [
    'html { background: #e9eaec; }',
    'body.lw-doc { margin: 24px auto 72px; background: #fff; color: #000; box-sizing: content-box;',
    '  box-shadow: 0 0 0 1px #d4d7dc, 0 3px 12px rgba(16,24,40,.14); overflow-wrap: break-word; position: relative; }',
    ':where(.lw-doc p, .lw-doc li, .lw-doc h1, .lw-doc h2, .lw-doc h3, .lw-doc h4, .lw-doc h5, .lw-doc h6,',
    '  .lw-doc div, .lw-doc blockquote, .lw-doc pre) { margin-top: 0; margin-bottom: 0; }',
    ':where(.lw-doc h1, .lw-doc h2, .lw-doc h3, .lw-doc h4, .lw-doc h5, .lw-doc h6) { font-size: inherit; font-weight: inherit; }',
    ':where(.lw-doc ul, .lw-doc ol) { margin: 0; padding-left: 36pt; }',
    ':where(.lw-doc li) { margin-left: 0; text-indent: 0; }',
    ':where(.lw-doc table) { border-collapse: collapse; }',
    ':where(.lw-doc td, .lw-doc th) { padding: 0 5.4pt; vertical-align: top; }',
    ':where(.lw-doc table[border]:not([border="0"]) > * > tr > td, .lw-doc table[border]:not([border="0"]) > * > tr > th) { border: 0.5pt solid #000; }',
    ':where(.lw-doc a) { color: #0563C1; text-decoration: underline; }',
    '.lw-doc img { max-width: none; }',
    '.lw-doc img.docx-anchor, .lw-doc img.docx-vml { outline: 1px dashed #7d93b2; outline-offset: 1px; }',
    // TinyMCE схлопывает \\t внутри метки в пробел, поэтому сам знак табуляции рисуем через ::before
    '.lw-doc .docx-tab { white-space: pre; font-size: 0; }',
    '.lw-doc .docx-tab::before { content: "\\9"; font-size: 1rem; }',
    '.lw-doc .docx-chip { display: inline-block; max-width: 100%; overflow: hidden; text-overflow: ellipsis; vertical-align: baseline;',
    '  white-space: nowrap; background: #eef2f7; border: 1px dashed #9aa7b8; border-radius: 3px; padding: 0 5px; text-indent: 0;',
    '  font: 9pt/1.5 "Segoe UI", Arial, sans-serif; color: #2f3b4d; font-weight: normal; font-style: normal; }',
    '.lw-doc .docx-fn { color: #0b5cad; cursor: default; }',
    '.lw-doc .docx-math { background: #f6f2ff; border: 1px dashed #b9a6e0; border-radius: 3px; padding: 0 3px; font-family: "Cambria Math", Cambria, serif; }',
    '.lw-doc .docx-section { margin: 4pt 0; padding: 1px 0; border-top: 1px dashed #8b9bb4; border-bottom: 1px dashed #8b9bb4;',
    '  color: #5a6b85; font: 8pt "Segoe UI", Arial, sans-serif; text-align: center; text-indent: 0; user-select: none; }',
    '.lw-marks p::after, .lw-marks li::after, .lw-marks h1::after, .lw-marks h2::after, .lw-marks h3::after {',
    '  content: "\\00b6"; color: #7d93b2; font: normal 400 1em Arial, sans-serif; text-decoration: none; }',
    '.lw-marks .docx-tab { position: relative; }',
    '.lw-marks .docx-tab::after { content: "\\2192"; position: absolute; left: 1px; top: 0; color: #7d93b2; font: normal 400 11px Arial; }',
    '#lw-pages { position: absolute; left: 0; top: 0; width: 0; height: 0; pointer-events: none; z-index: 5; }',
    '#lw-pages .lw-pb { position: absolute; height: 0; border-top: 1px dashed #9fb2cf; }',
    '#lw-pages .lw-pb span { position: absolute; right: -6px; top: -9px; transform: translateX(100%); padding: 1px 6px;',
    '  font: 10px/14px "Segoe UI", Arial, sans-serif; color: #5d6f8c; background: #e9eaec; white-space: nowrap; }',
    '#lw-pages .lw-pb.manual { border-top-color: #6f8bb8; }',
    '@counter-style lw-ru-lower { system: alphabetic; symbols: "а" "б" "в" "г" "д" "е" "ж" "з" "и" "к" "л" "м" "н" "о" "п" "р" "с" "т" "у" "ф" "х" "ц" "ч" "ш" "щ" "э" "ю" "я"; }',
    '@counter-style lw-ru-upper { system: alphabetic; symbols: "А" "Б" "В" "Г" "Д" "Е" "Ж" "З" "И" "К" "Л" "М" "Н" "О" "П" "Р" "С" "Т" "У" "Ф" "Х" "Ц" "Ч" "Ш" "Щ" "Э" "Ю" "Я"; }'
  ].join('\n');

  var state = {
    page: { w: 11906, h: 16838, top: 1134, right: 850, bottom: 1134, left: 1701, header: 708, footer: 708, gutter: 0 },
    styles: [], defaultStyle: 'Normal', nextStyles: {}, tab: 35.4,
    original: null, originalB64: null, zoom: 1, marks: [], pages: 1, painter: null, ready: false
  };
  var editor = null;
  var $ = function (id) { return document.getElementById(id); };

  // ------------------------------------------------------------ утилиты
  function num(v) { var n = parseFloat(String(v).replace(',', '.')); return isFinite(n) ? n : NaN; }
  function fmt(n, d) { return (Math.round(n * Math.pow(10, d || 2)) / Math.pow(10, d || 2)).toString().replace('.', ','); }
  function twToCm(tw) { return fmt(tw / TW_PER_CM, 2); }
  function cmToTw(cm) { return Math.round(num(cm) * TW_PER_CM); }
  function pxToPt(px) { return px * 0.75; }
  function ptStr(tw) { return (tw / 20).toFixed(2) + 'pt'; }
  function escHtml(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function debounce(fn, ms) { var t; return function () { clearTimeout(t); t = setTimeout(fn, ms); }; }
  function fontsFmt(list) { return list.map(function (f) { return f + '=' + f; }).join('; '); }
  function rgbToHex(c) {
    var m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c || '');
    if (!m) { return '#000000'; }
    return '#' + [1, 2, 3].map(function (i) { return ('0' + parseInt(m[i], 10).toString(16)).slice(-2); }).join('').toUpperCase();
  }
  function notify(text, type, timeout) {
    if (editor) { editor.notificationManager.open({ text: text, type: type || 'info', timeout: timeout === undefined ? 6000 : timeout }); }
    else { alert(text); }
  }
  function showBusy(text) { $('busyText').textContent = text; $('busy').hidden = false; }
  function hideBusy() { $('busy').hidden = true; }

  function blocks() {
    var list = editor.selection.getSelectedBlocks().filter(function (b) { return /^(P|LI|H[1-6]|DIV|BLOCKQUOTE|PRE)$/.test(b.nodeName) && !b.classList.contains('docx-section'); });
    if (!list.length) {
      var b = editor.dom.getParent(editor.selection.getNode(), 'p,li,h1,h2,h3,h4,h5,h6,div,blockquote,pre');
      if (b && b !== editor.getBody()) { list = [b]; }
    }
    return list;
  }
  function styleOf(el) {
    var cls = el && el.className ? String(el.className).split(/\s+/) : [];
    for (var i = 0; i < cls.length; i++) { if (cls[i].indexOf('s-') === 0) { return cls[i].slice(2); } }
    if (el && /^H[1-6]$/.test(el.nodeName)) {
      var want = 'heading ' + el.nodeName[1];
      var s = state.styles.filter(function (x) { return x.name.toLowerCase() === want; })[0];
      if (s) { return s.id; }
    }
    return el && el.nodeName === 'LI' ? null : state.defaultStyle;
  }
  function styleName(id) {
    var s = state.styles.filter(function (x) { return x.id === id; })[0];
    return s ? s.ui : (id || 'Абзац списка');
  }
  function setStyles(el, obj) {
    Object.keys(obj).forEach(function (k) { editor.dom.setStyle(el, k, obj[k]); });
    if (!el.getAttribute('style')) { el.removeAttribute('style'); el.removeAttribute('data-mce-style'); }
  }

  // ------------------------------------------------------- CSS страницы
  function injectStyle(id, css) {
    var doc = editor.getDoc();
    var el = doc.getElementById(id);
    if (!el) { el = doc.createElement('style'); el.id = id; doc.head.appendChild(el); }
    el.textContent = css;
  }
  function applyPageCss() {
    var p = state.page, z = state.zoom;
    var css = [
      'body.lw-doc { width: ' + ptStr(p.w - p.left - p.right) + '; min-height: ' + ptStr(p.h - p.top - p.bottom) + ';',
      '  padding: ' + [p.top, p.right, p.bottom, p.left].map(ptStr).join(' ') + '; zoom: ' + z + ';',
      '  tab-size: ' + state.tab + 'pt; -moz-tab-size: ' + state.tab + 'pt; }',
      '@page { size: ' + ptStr(p.w) + ' ' + ptStr(p.h) + '; margin: ' + [p.top, p.right, p.bottom, p.left].map(ptStr).join(' ') + '; }',
      '@media print { html { background: none !important; } body.lw-doc { width: auto; min-height: 0; padding: 0; margin: 0; box-shadow: none; zoom: 1; }',
      '  #lw-pages, .docx-section { display: none !important; } img.mce-pagebreak { break-after: page; border: 0 !important; height: 0 !important; margin: 0 !important; } }'
    ].join('\n');
    injectStyle('lw-page-css', css);
    var paper = PAPER.filter(function (x) { return (x.w === p.w && x.h === p.h) || (x.w === p.h && x.h === p.w); })[0];
    $('stPageSize').textContent = (paper ? paper.id : twToCm(p.w) + ' × ' + twToCm(p.h) + ' см') + ', ' + (p.w > p.h ? 'альбомная' : 'книжная');
    schedulePages();
  }
  function setZoom(z) {
    state.zoom = Math.max(0.5, Math.min(2, Math.round(z * 10) / 10));
    $('zoomRange').value = Math.round(state.zoom * 100);
    $('zoomVal').textContent = Math.round(state.zoom * 100) + '%';
    if (editor && state.ready) { applyPageCss(); }
  }

  // ------------------------------------------------- разметка страниц
  function flatBlocks(body) {
    var out = [];
    Array.prototype.forEach.call(body.children, function (el) {
      if (el.nodeName === 'TABLE') {
        Array.prototype.forEach.call(el.querySelectorAll(':scope > tbody > tr, :scope > thead > tr, :scope > tfoot > tr, :scope > tr'), function (tr) { out.push(tr); });
      } else if (el.nodeName === 'UL' || el.nodeName === 'OL') {
        Array.prototype.forEach.call(el.children, function (li) { out.push(li); });
      } else if (el.querySelector && el.querySelector('img.mce-pagebreak')) {
        out.push(el.querySelector('img.mce-pagebreak'));
      } else { out.push(el); }
    });
    return out;
  }
  function computePages() {
    if (!editor || !state.ready) { return; }
    var doc = editor.getDoc(), win = editor.getWin(), body = editor.getBody();
    var p = state.page, z = state.zoom;
    var contentH = (p.h - p.top - p.bottom) / 20 * PX_PER_PT * z;
    var sy = win.scrollY || doc.documentElement.scrollTop || 0, sx = win.scrollX || 0;
    var br = body.getBoundingClientRect();
    var start = br.top + sy + p.top / 20 * PX_PER_PT * z, limit = start + contentH;
    var marks = [];
    flatBlocks(body).forEach(function (el) {
      var r = el.getBoundingClientRect();
      var y0 = r.top + sy, y1 = r.bottom + sy;
      if (el.classList && el.classList.contains('mce-pagebreak')) {
        marks.push({ y: y0 + r.height / 2, manual: true }); start = y1; limit = start + contentH; return;
      }
      if (!r.height) { return; }
      var guard = 0;
      while (y1 > limit + 0.5 && guard++ < 400) {
        var cut = (y0 > start + 1 && y0 < limit && (y1 - y0) < contentH * 0.3) ? y0 : limit;
        if (cut <= start) { cut = limit; }
        marks.push({ y: cut }); start = cut; limit = start + contentH;
      }
    });
    state.marks = marks;
    state.pages = marks.length + 1;
    var layer = doc.getElementById('lw-pages');
    if (!layer) { layer = doc.createElement('div'); layer.id = 'lw-pages'; doc.documentElement.appendChild(layer); }
    layer.innerHTML = marks.map(function (m, i) {
      return '<div class="lw-pb' + (m.manual ? ' manual' : '') + '" style="top:' + m.y.toFixed(1) + 'px;left:' + (br.left + sx).toFixed(1) +
        'px;width:' + br.width.toFixed(1) + 'px"><span>' + (m.manual ? 'Разрыв · ' : '') + 'стр. ' + (i + 2) + '</span></div>';
    }).join('');
    updateStatus();
  }
  var schedulePages = debounce(computePages, 250);

  function currentPage() {
    try {
      var rng = editor.selection.getRng();
      var rect = rng.getBoundingClientRect();
      if (!rect || (!rect.height && !rect.top)) {
        var n = editor.selection.getNode();
        rect = n && n.getBoundingClientRect ? n.getBoundingClientRect() : null;
      }
      if (!rect) { return 1; }
      var y = rect.top + (editor.getWin().scrollY || 0);
      return 1 + state.marks.filter(function (m) { return m.y <= y; }).length;
    } catch (e) { return 1; }
  }
  function updateStatus() {
    if (!editor || !state.ready) { return; }
    $('stPage').textContent = 'Страница ' + Math.min(currentPage(), state.pages) + ' из ' + state.pages;
    try { $('stWords').textContent = 'Число слов: ' + editor.plugins.wordcount.body.getWordCount(); } catch (e) { /* нет плагина */ }
    var b = blocks()[0];
    $('stStyle').textContent = 'Стиль: ' + styleName(b ? styleOf(b) : state.defaultStyle);
    $('dirtyMark').hidden = !editor.isDirty();
  }

  // ------------------------------------------------------------- стили
  function applyStyle(id) {
    editor.undoManager.transact(function () {
      blocks().forEach(function (el) {
        if (/^(H[1-6]|DIV|BLOCKQUOTE|PRE)$/.test(el.nodeName)) { el = editor.dom.rename(el, 'p'); }
        String(el.className || '').split(/\s+/).forEach(function (c) { if (c.indexOf('s-') === 0) { editor.dom.removeClass(el, c); } });
        editor.dom.addClass(el, 's-' + id);
        setStyles(el, { 'padding-top': '', 'padding-bottom': '', 'line-height': '', 'font-size': '', 'font-family': '', color: '' });
        Array.prototype.forEach.call(el.querySelectorAll('span'), function (s) {
          if (s.classList.contains('docx-tab') || s.classList.contains('docx-raw')) { return; }
          setStyles(s, { 'font-size': '', 'font-family': '', color: '' });
        });
      });
    });
    editor.nodeChanged();
    editor.focus();
  }
  function styleItems() {
    var cur = (blocks()[0] && styleOf(blocks()[0])) || state.defaultStyle;
    return state.styles.map(function (s) {
      return { type: 'togglemenuitem', text: s.ui, active: s.id === cur, onAction: function () { applyStyle(s.id); } };
    });
  }

  // ------------------------------------------------------- интервалы
  function lineCss(mult) { return String(Math.round(mult * LINE_FACTOR * 10000) / 10000); }
  function currentLine(b) {
    var inl = (b.style.lineHeight || '').trim();
    var cs = editor.getWin().getComputedStyle(b);
    var fs = parseFloat(cs.fontSize) || 16;
    if (/^max\(/.test(inl)) { return { type: 'atLeast', val: num(inl.replace(/^max\(\s*/, '')) }; }
    if (/(pt|px)$/.test(inl)) { return { type: 'exact', val: /px$/.test(inl) ? pxToPt(num(inl)) : num(inl) }; }
    var lh = cs.lineHeight === 'normal' ? fs * LINE_FACTOR : parseFloat(cs.lineHeight);
    var m = Math.round(lh / (fs * LINE_FACTOR) * 100) / 100;
    var type = Math.abs(m - 1) < 0.02 ? 'single' : Math.abs(m - 1.5) < 0.02 ? 'onehalf' : Math.abs(m - 2) < 0.02 ? 'double' : 'multiple';
    return { type: type, val: m };
  }
  function setLine(mult) {
    editor.undoManager.transact(function () { blocks().forEach(function (b) { setStyles(b, { 'line-height': lineCss(mult) }); }); });
    editor.nodeChanged(); schedulePages();
  }
  function changeIndent(dir) {
    var b = blocks();
    if (b.some(function (x) { return x.nodeName === 'LI'; })) { editor.execCommand(dir > 0 ? 'Indent' : 'Outdent'); return; }
    var step = 1.25 * TW_PER_CM / 20;
    editor.undoManager.transact(function () {
      b.forEach(function (el) {
        var cur = pxToPt(parseFloat(editor.getWin().getComputedStyle(el).marginLeft) || 0);
        var next = Math.max(0, Math.round((cur + dir * step) / step) * step);
        setStyles(el, { 'margin-left': next ? next.toFixed(2) + 'pt' : '0pt' });
      });
    });
    editor.nodeChanged(); schedulePages();
  }
  function growFont(dir) {
    var n = editor.selection.getNode();
    var cur = pxToPt(parseFloat(editor.getWin().getComputedStyle(n).fontSize) || 14.67);
    var next = cur;
    if (dir > 0) { next = FONT_STEPS.filter(function (s) { return s > cur + 0.01; })[0] || Math.ceil(cur / 10 + 1) * 10; }
    else { next = FONT_STEPS.filter(function (s) { return s < cur - 0.01; }).pop() || Math.max(1, cur - 1); }
    editor.execCommand('FontSize', false, next + 'pt');
  }

  // ------------------------------------------------------------ диалоги
  function paraDialog() {
    var list = blocks();
    if (!list.length) { return; }
    var b = list[0], isLi = list.some(function (x) { return x.nodeName === 'LI'; });
    var cs = editor.getWin().getComputedStyle(b);
    var ti = parseFloat(cs.textIndent) || 0;
    var line = currentLine(b);
    var al = cs.textAlign;
    var init = {
      align: al === 'center' ? 'center' : al === 'right' || al === 'end' ? 'right' : al === 'justify' ? 'justify' : 'left',
      left: isLi ? '—' : fmt(pxToPt(parseFloat(cs.marginLeft) || 0) * 20 / TW_PER_CM),
      right: fmt(pxToPt(parseFloat(cs.marginRight) || 0) * 20 / TW_PER_CM),
      firstType: isLi ? 'none' : ti > 0.5 ? 'first' : ti < -0.5 ? 'hanging' : 'none',
      firstVal: isLi ? '' : ti ? fmt(Math.abs(pxToPt(ti)) * 20 / TW_PER_CM) : '1,25',
      before: fmt(pxToPt(parseFloat(cs.paddingTop) || 0), 1),
      after: fmt(pxToPt(parseFloat(cs.paddingBottom) || 0), 1),
      lineType: line.type,
      lineVal: line.type === 'single' ? '1' : line.type === 'onehalf' ? '1,5' : line.type === 'double' ? '2' : fmt(line.val, 2)
    };
    editor.windowManager.open({
      title: 'Абзац',
      size: 'normal',
      body: {
        type: 'panel', items: [
          { type: 'selectbox', name: 'align', label: 'Выравнивание', items: [
            { value: 'left', text: 'По левому краю' }, { value: 'center', text: 'По центру' },
            { value: 'right', text: 'По правому краю' }, { value: 'justify', text: 'По ширине' }] },
          { type: 'label', label: 'Отступ' + (isLi ? ' (для списка задаётся уровнем: Tab / Shift+Tab)' : ''), items: [
            { type: 'grid', columns: 2, items: [
              { type: 'input', name: 'left', label: 'Слева, см', inputMode: 'decimal', enabled: !isLi },
              { type: 'input', name: 'right', label: 'Справа, см', inputMode: 'decimal' },
              { type: 'selectbox', name: 'firstType', label: 'Первая строка', enabled: !isLi, items: [
                { value: 'none', text: '(нет)' }, { value: 'first', text: 'Отступ' }, { value: 'hanging', text: 'Выступ' }] },
              { type: 'input', name: 'firstVal', label: 'на, см', inputMode: 'decimal', enabled: !isLi }] }] },
          { type: 'label', label: 'Интервал', items: [
            { type: 'grid', columns: 2, items: [
              { type: 'input', name: 'before', label: 'Перед, пт', inputMode: 'decimal' },
              { type: 'input', name: 'after', label: 'После, пт', inputMode: 'decimal' },
              { type: 'selectbox', name: 'lineType', label: 'Междустрочный', items: [
                { value: 'single', text: 'Одинарный' }, { value: 'onehalf', text: '1,5 строки' },
                { value: 'double', text: 'Двойной' }, { value: 'atLeast', text: 'Минимум' },
                { value: 'exact', text: 'Точно' }, { value: 'multiple', text: 'Множитель' }] },
              { type: 'input', name: 'lineVal', label: 'Значение (пт или множитель)', inputMode: 'decimal' }] }] }
        ]
      },
      initialData: init,
      buttons: [{ type: 'cancel', text: 'Отмена' }, { type: 'submit', text: 'ОК', primary: true }],
      onChange: function (api, d) {
        if (d.name === 'lineType') {
          var t = api.getData().lineType;
          var v = { single: '1', onehalf: '1,5', double: '2', multiple: '1,15', atLeast: '12', exact: '12' }[t];
          api.setData({ lineVal: v });
        }
      },
      onSubmit: function (api) {
        var d = api.getData();
        var css = {};
        if (d.align !== init.align) { css['text-align'] = d.align; }
        if (!isLi && d.left !== init.left && isFinite(num(d.left))) { css['margin-left'] = (cmToTw(d.left) / 20).toFixed(2) + 'pt'; }
        if (d.right !== init.right && isFinite(num(d.right))) { css['margin-right'] = (cmToTw(d.right) / 20).toFixed(2) + 'pt'; }
        if (!isLi && (d.firstType !== init.firstType || d.firstVal !== init.firstVal)) {
          var fv = isFinite(num(d.firstVal)) ? cmToTw(d.firstVal) / 20 : 0;
          css['text-indent'] = d.firstType === 'none' ? '0pt' : ((d.firstType === 'hanging' ? -fv : fv).toFixed(2) + 'pt');
          if (d.firstType === 'hanging' && !css['margin-left']) {
            var curLeft = pxToPt(parseFloat(cs.marginLeft) || 0);
            if (curLeft < fv) { css['margin-left'] = fv.toFixed(2) + 'pt'; }
          }
        }
        if (d.before !== init.before && isFinite(num(d.before))) { css['padding-top'] = num(d.before) + 'pt'; }
        if (d.after !== init.after && isFinite(num(d.after))) { css['padding-bottom'] = num(d.after) + 'pt'; }
        if (d.lineType !== init.lineType || d.lineVal !== init.lineVal) {
          var lv = num(d.lineVal);
          if (d.lineType === 'single') { css['line-height'] = lineCss(1); }
          else if (d.lineType === 'onehalf') { css['line-height'] = lineCss(1.5); }
          else if (d.lineType === 'double') { css['line-height'] = lineCss(2); }
          else if (isFinite(lv) && lv > 0) {
            css['line-height'] = d.lineType === 'exact' ? lv + 'pt' : d.lineType === 'atLeast' ? 'max(' + lv + 'pt, ' + LINE_FACTOR + 'em)' : lineCss(lv);
          }
        }
        editor.undoManager.transact(function () { list.forEach(function (el) { setStyles(el, css); }); });
        api.close();
        editor.nodeChanged(); schedulePages();
      }
    });
  }

  function fontDialog() {
    var n = editor.selection.getNode();
    var cs = editor.getWin().getComputedStyle(n);
    var fam = (cs.fontFamily || '').split(',')[0].replace(/["']/g, '').trim();
    var fonts = FONTS.slice();
    if (fam && fonts.indexOf(fam) < 0) { fonts.unshift(fam); }
    var bold = parseInt(cs.fontWeight, 10) >= 600 || cs.fontWeight === 'bold';
    var italic = cs.fontStyle === 'italic';
    var deco = (cs.textDecorationLine || cs.textDecoration || '');
    var f = editor.formatter;
    var init = {
      family: fam || 'Calibri',
      size: fmt(pxToPt(parseFloat(cs.fontSize) || 14.67), 1),
      style: bold && italic ? 'bi' : bold ? 'b' : italic ? 'i' : 'r',
      color: rgbToHex(cs.color),
      underline: deco.indexOf('underline') >= 0 ? (cs.textDecorationStyle === 'double' ? 'double' : cs.textDecorationStyle === 'dotted' ? 'dotted' : cs.textDecorationStyle === 'dashed' ? 'dashed' : cs.textDecorationStyle === 'wavy' ? 'wavy' : 'single') : 'none',
      strike: f.match('strikethrough') || deco.indexOf('line-through') >= 0,
      sup: f.match('superscript'), sub: f.match('subscript'),
      smallcaps: (cs.fontVariant || '').indexOf('small-caps') >= 0,
      allcaps: cs.textTransform === 'uppercase',
      hidden: false,
      spacing: fmt(pxToPt(parseFloat(cs.letterSpacing) || 0), 1)
    };
    editor.windowManager.open({
      title: 'Шрифт',
      size: 'normal',
      body: { type: 'panel', items: [
        { type: 'grid', columns: 2, items: [
          { type: 'listbox', name: 'family', label: 'Шрифт', items: fonts.map(function (x) { return { text: x, value: x }; }) },
          { type: 'input', name: 'size', label: 'Размер, пт', inputMode: 'decimal' },
          { type: 'selectbox', name: 'style', label: 'Начертание', items: [
            { value: 'r', text: 'Обычный' }, { value: 'i', text: 'Курсив' }, { value: 'b', text: 'Полужирный' }, { value: 'bi', text: 'Полужирный курсив' }] },
          { type: 'colorinput', name: 'color', label: 'Цвет текста' },
          { type: 'selectbox', name: 'underline', label: 'Подчёркивание', items: [
            { value: 'none', text: '(нет)' }, { value: 'single', text: 'Одинарное' }, { value: 'double', text: 'Двойное' },
            { value: 'dotted', text: 'Точечное' }, { value: 'dashed', text: 'Штриховое' }, { value: 'wavy', text: 'Волнистое' }] },
          { type: 'input', name: 'spacing', label: 'Интервал между знаками, пт', inputMode: 'decimal' }] },
        { type: 'label', label: 'Видоизменение', items: [
          { type: 'grid', columns: 2, items: [
            { type: 'checkbox', name: 'strike', label: 'Зачёркнутый' },
            { type: 'checkbox', name: 'smallcaps', label: 'Малые прописные' },
            { type: 'checkbox', name: 'sup', label: 'Надстрочный' },
            { type: 'checkbox', name: 'allcaps', label: 'Все прописные' },
            { type: 'checkbox', name: 'sub', label: 'Подстрочный' },
            { type: 'checkbox', name: 'hidden', label: 'Скрытый' }] }] }
      ] },
      initialData: init,
      buttons: [{ type: 'cancel', text: 'Отмена' }, { type: 'submit', text: 'ОК', primary: true }],
      onSubmit: function (api) {
        var d = api.getData();
        editor.undoManager.transact(function () {
          if (d.family !== init.family) { editor.execCommand('FontName', false, d.family); }
          if (d.size !== init.size && num(d.size) > 0) { editor.execCommand('FontSize', false, num(d.size) + 'pt'); }
          if (d.style !== init.style) {
            var wantB = d.style.indexOf('b') >= 0, wantI = d.style.indexOf('i') >= 0;
            if (wantB !== bold) { f[wantB ? 'apply' : 'remove']('bold'); }
            if (wantI !== italic) { f[wantI ? 'apply' : 'remove']('italic'); }
          }
          if (d.color && d.color.toUpperCase() !== init.color) { editor.execCommand('ForeColor', false, d.color); }
          if (d.underline !== init.underline) {
            f.remove('underline'); f.remove('lw_u');
            if (d.underline === 'single') { f.apply('underline'); }
            else if (d.underline !== 'none') { f.register('lw_u', { inline: 'span', styles: { 'text-decoration': 'underline ' + d.underline } }); f.apply('lw_u'); }
          }
          ['strike', 'sup', 'sub'].forEach(function (k) {
            var name = { strike: 'strikethrough', sup: 'superscript', sub: 'subscript' }[k];
            if (d[k] !== init[k]) { f[d[k] ? 'apply' : 'remove'](name); }
          });
          var custom = [['smallcaps', 'font-variant', 'small-caps'], ['allcaps', 'text-transform', 'uppercase'], ['hidden', 'display', 'none']];
          custom.forEach(function (c) {
            if (d[c[0]] !== init[c[0]]) {
              var st = {}; st[c[1]] = c[2];
              f.register('lw_' + c[0], { inline: 'span', styles: st });
              f[d[c[0]] ? 'apply' : 'remove']('lw_' + c[0]);
            }
          });
          if (d.spacing !== init.spacing && isFinite(num(d.spacing))) {
            f.register('lw_sp', { inline: 'span', styles: { 'letter-spacing': num(d.spacing) + 'pt' } });
            f.apply('lw_sp');
          }
        });
        api.close();
        editor.nodeChanged();
      }
    });
  }

  function pageDialog() {
    var p = state.page;
    var land = p.w > p.h;
    var paper = PAPER.filter(function (x) { return (x.w === p.w && x.h === p.h) || (x.w === p.h && x.h === p.w); })[0];
    var init = {
      paper: paper ? paper.id : 'custom', orient: land ? 'landscape' : 'portrait',
      width: twToCm(p.w), height: twToCm(p.h),
      top: twToCm(p.top), bottom: twToCm(p.bottom), left: twToCm(p.left), right: twToCm(p.right),
      header: twToCm(p.header || 708), footer: twToCm(p.footer || 708)
    };
    editor.windowManager.open({
      title: 'Параметры страницы',
      size: 'normal',
      body: { type: 'panel', items: [
        { type: 'label', label: 'Бумага', items: [
          { type: 'grid', columns: 2, items: [
            { type: 'selectbox', name: 'paper', label: 'Размер', items: PAPER.map(function (x) { return { value: x.id, text: x.name }; }).concat([{ value: 'custom', text: 'Другой' }]) },
            { type: 'selectbox', name: 'orient', label: 'Ориентация', items: [{ value: 'portrait', text: 'Книжная' }, { value: 'landscape', text: 'Альбомная' }] },
            { type: 'input', name: 'width', label: 'Ширина, см', inputMode: 'decimal' },
            { type: 'input', name: 'height', label: 'Высота, см', inputMode: 'decimal' }] }] },
        { type: 'label', label: 'Поля', items: [
          { type: 'grid', columns: 2, items: [
            { type: 'input', name: 'top', label: 'Верхнее, см', inputMode: 'decimal' },
            { type: 'input', name: 'bottom', label: 'Нижнее, см', inputMode: 'decimal' },
            { type: 'input', name: 'left', label: 'Левое, см', inputMode: 'decimal' },
            { type: 'input', name: 'right', label: 'Правое, см', inputMode: 'decimal' }] }] },
        { type: 'label', label: 'Колонтитулы (расстояние от края)', items: [
          { type: 'grid', columns: 2, items: [
            { type: 'input', name: 'header', label: 'Верхнего, см', inputMode: 'decimal' },
            { type: 'input', name: 'footer', label: 'Нижнего, см', inputMode: 'decimal' }] }] }
      ] },
      initialData: init,
      buttons: [{ type: 'cancel', text: 'Отмена' }, { type: 'submit', text: 'ОК', primary: true }],
      onChange: function (api, det) {
        var d = api.getData();
        if (det.name === 'paper' && d.paper !== 'custom') {
          var x = PAPER.filter(function (q) { return q.id === d.paper; })[0];
          var w = d.orient === 'landscape' ? x.h : x.w, h = d.orient === 'landscape' ? x.w : x.h;
          api.setData({ width: twToCm(w), height: twToCm(h) });
        } else if (det.name === 'orient') {
          var a = num(d.width), b = num(d.height);
          if ((d.orient === 'landscape') !== (a > b)) { api.setData({ width: d.height, height: d.width }); }
        } else if (det.name === 'width' || det.name === 'height') {
          api.setData({ paper: 'custom' });
        }
      },
      onSubmit: function (api) {
        var d = api.getData();
        var np = {};
        var ok = ['width', 'height', 'top', 'bottom', 'left', 'right', 'header', 'footer'].every(function (k) { return isFinite(num(d[k])) && num(d[k]) >= 0; });
        if (!ok) { notify('Проверьте значения: нужны числа в сантиметрах.', 'error'); return; }
        np.w = cmToTw(d.width); np.h = cmToTw(d.height);
        np.top = cmToTw(d.top); np.bottom = cmToTw(d.bottom); np.left = cmToTw(d.left); np.right = cmToTw(d.right);
        np.header = cmToTw(d.header); np.footer = cmToTw(d.footer); np.gutter = state.page.gutter || 0;
        if (np.w - np.left - np.right < 1000 || np.h - np.top - np.bottom < 1000) { notify('Поля слишком большие для такого размера страницы.', 'error'); return; }
        setPage(np);
        api.close();
      }
    });
  }
  function setPage(np) {
    state.page = Object.assign({}, state.page, np);
    applyPageCss();
    editor.setDirty(true);
    updateStatus();
  }

  // ------------------------------------------------------- формат по образцу
  function captureFormat() {
    var node = editor.selection.getStart();
    if (node.nodeType === 3) { node = node.parentNode; }
    var body = editor.getBody();
    var block = editor.dom.getParent(node, editor.dom.isBlock) || body;
    var res = { tags: [], styles: {} };
    var props = ['font-family', 'font-size', 'color', 'background-color', 'text-decoration', 'letter-spacing', 'text-transform', 'font-variant', 'font-weight', 'font-style'];
    for (var n = node; n && n !== block && n !== body; n = n.parentNode) {
      var t = n.nodeName;
      var map = { STRONG: 'bold', B: 'bold', EM: 'italic', I: 'italic', U: 'underline', S: 'strikethrough', STRIKE: 'strikethrough', DEL: 'strikethrough', SUP: 'superscript', SUB: 'subscript' };
      if (map[t] && res.tags.indexOf(map[t]) < 0) { res.tags.push(map[t]); }
      if (n.style) {
        props.forEach(function (pr) { var v = n.style.getPropertyValue(pr); if (v && !(pr in res.styles)) { res.styles[pr] = v; } });
      }
    }
    return res;
  }
  function applyCaptured(fmtObj) {
    editor.undoManager.transact(function () {
      editor.execCommand('RemoveFormat');
      fmtObj.tags.forEach(function (t) { editor.formatter.apply(t); });
      if (Object.keys(fmtObj.styles).length) {
        editor.formatter.register('lw_paint', { inline: 'span', styles: fmtObj.styles });
        editor.formatter.apply('lw_paint');
      }
    });
  }

  // ------------------------------------------------------------ списки
  function normalizeLists() {
    var body = editor.getBody();
    Array.prototype.forEach.call(body.querySelectorAll('li[data-lvl]'), function (li) {
      var depth = 0, top = null;
      for (var n = li.parentNode; n && n !== body; n = n.parentNode) { if (n.nodeName === 'UL' || n.nodeName === 'OL') { depth++; top = n; } }
      var base = 0;
      var m = top && /lw-n\w+-l(\d+)/.exec(top.className || '');
      if (m) { base = parseInt(m[1], 10); }
      if (String(base + depth - 1) !== li.getAttribute('data-lvl')) {
        li.removeAttribute('data-lvl'); li.removeAttribute('data-ind');
        setStyles(li, { 'margin-left': '', 'text-indent': '', 'list-style-position': '' });
      }
    });
    Array.prototype.forEach.call(body.querySelectorAll('p[data-ind], p[data-lvl]'), function (p) {
      p.removeAttribute('data-lvl'); p.removeAttribute('data-ind');
      setStyles(p, { 'margin-left': '', 'text-indent': '', 'list-style-position': '' });
    });
  }

  // ------------------------------------------- конвертация: сервер или Python в браузере
  // Запуск с компьютера (server.py) — конвертирует сервер. На GitHub Pages сервера нет —
  // тот же Python-код работает в браузере (Pyodide в static/worker.mjs).
  var DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  var conv = { mode: null, worker: null, seq: 0, pending: {}, progress: '' };
  function detectMode() {
    if (conv.mode) { return Promise.resolve(conv.mode); }
    return fetch('api/ping', { cache: 'no-store' })
      .then(function (r) { return r.ok ? r.json() : null; })
      .catch(function () { return null; })
      .then(function (j) {
        conv.mode = j && j.ok ? 'server' : 'browser';
        if (conv.mode === 'browser') { startWorker(); }
        return conv.mode;
      });
  }
  function startWorker() {
    if (conv.worker) { return; }
    conv.worker = new Worker('static/worker.mjs', { type: 'module' });
    conv.worker.onmessage = function (e) {
      var m = e.data;
      if (m.progress !== undefined) {
        conv.progress = m.progress;
        if (m.progress && !$('busy').hidden) { $('busyText').textContent = m.progress; }
        return;
      }
      var p = conv.pending[m.id];
      if (!p) { return; }
      delete conv.pending[m.id];
      if (m.error) { p.reject(new Error(m.error)); } else { p.resolve(m.result); }
    };
    conv.worker.onerror = function (e) {
      var err = new Error('модуль Word не загрузился' + (e && e.message ? ': ' + e.message : '') + '. Обновите страницу.');
      Object.keys(conv.pending).forEach(function (id) { conv.pending[id].reject(err); delete conv.pending[id]; });
    };
    workerCall({ cmd: 'warmup' }).catch(function () { /* повторим при открытии файла */ });
  }
  function workerCall(msg, transfer) {
    return new Promise(function (resolve, reject) {
      var id = ++conv.seq;
      conv.pending[id] = { resolve: resolve, reject: reject };
      msg.id = id;
      conv.worker.postMessage(msg, transfer || []);
    });
  }
  function convertOpen(buf) {
    return detectMode().then(function (mode) {
      if (mode === 'browser') {
        if (conv.progress) { $('busyText').textContent = conv.progress; }
        var copy = buf.slice(0);
        return workerCall({ cmd: 'open', data: copy }, [copy]);
      }
      return fetch('api/open', { method: 'POST', headers: { 'Content-Type': 'application/octet-stream' }, body: buf })
        .then(function (res) {
          return res.json().then(function (j) { if (!res.ok || j.error) { throw new Error(j.error || res.statusText); } return j; });
        });
    });
  }
  function postSave(html, withOriginal) {
    var prep = withOriginal && state.original && !state.originalB64 ?
      bufToB64(state.original).then(function (b) { state.originalB64 = b; }) : Promise.resolve();
    return prep.then(function () {
      return fetch('api/save', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ html: html, page: state.page, token: state.token, filename: fileNameValue(),
          original: withOriginal ? state.originalB64 : null })
      });
    });
  }
  function convertSave(html) {
    return detectMode().then(function (mode) {
      if (mode === 'browser') {
        if (conv.progress) { $('busyText').textContent = conv.progress; }
        var orig = state.original ? state.original.slice(0) : null;
        return workerCall({ cmd: 'save', html: html, page: state.page, original: orig }, orig ? [orig] : [])
          .then(function (r) { return { blob: new Blob([r.bytes], { type: DOCX_MIME }), warnings: r.warnings || [] }; });
      }
      // сервер помнит исходный файл по токену; если забыл (перезапуск) — отвечает 409 и мы шлём файл сами
      return postSave(html, !!state.original && !state.token).then(function (res) {
        return res.status === 409 && state.original ? postSave(html, true) : res;
      }).then(function (res) {
        if (!res.ok) { return res.json().catch(function () { return {}; }).then(function (j) { throw new Error(j.error || res.statusText); }); }
        var warn = [];
        try { warn = JSON.parse(decodeURIComponent(res.headers.get('X-MaminWord-Warnings') || '%5B%5D')); } catch (e) { /* ignore */ }
        return res.blob().then(function (blob) { return { blob: blob, warnings: warn }; });
      });
    });
  }

  // -------------------------------------------------------- открыть / сохранить
  function bufToB64(buf) {
    return new Promise(function (resolve, reject) {
      var r = new FileReader();
      r.onload = function () { resolve(String(r.result).split(',')[1] || ''); };
      r.onerror = reject;
      r.readAsDataURL(new Blob([buf]));
    });
  }
  function loadDoc(data, name, buf) {
    state.page = data.page; state.styles = data.styles || []; state.defaultStyle = data.defaultStyle || 'Normal';
    state.nextStyles = data.nextStyles || {}; state.tab = parseFloat(String(data.tab).replace(',', '.')) || 35.4;
    state.original = buf || null; state.originalB64 = null; state.token = data.token || null;
    var fonts = FONTS.slice();
    (data.fonts || []).forEach(function (f) { if (fonts.indexOf(f) < 0) { fonts.push(f); } });
    fonts.sort(function (a, b) { return a.localeCompare(b); });
    editor.options.set('font_family_formats', fontsFmt(fonts));
    injectStyle('lw-doc-css', data.css || '');
    applyPageCss();
    editor.resetContent(data.html || '<p><br></p>'); // содержимое + история отмены + флаг изменений
    $('fileName').value = name;
    document.title = name.replace(/\.docx$/i, '') + ' — МаминWord';
    $('welcome').hidden = true;
    editor.getWin().scrollTo(0, 0);
    var first = editor.getBody().firstElementChild;
    if (first) { try { editor.selection.setCursorLocation(first, 0); } catch (e) { /* ignore */ } }
    editor.focus();
    // первый focus() в TinyMCE шлёт «change» без реальной правки — сбрасываем, пока пользователь ничего не сделал
    editor.setDirty(false);
    setTimeout(function () { if (!editor.undoManager.hasUndo()) { editor.setDirty(false); updateStatus(); } }, 1000);
    (data.warnings || []).forEach(function (w) { notify(w, 'warning', 10000); });
    schedulePages();
    setTimeout(schedulePages, 900);
  }
  function confirmDiscard() {
    return !editor.isDirty() || confirm('Документ изменён и не сохранён. Продолжить без сохранения?');
  }
  function openFile(file) {
    if (!file) { return; }
    if (!/\.docx$/i.test(file.name)) {
      notify(/\.doc$/i.test(file.name) ? 'Старый формат .doc не поддерживается. Откройте файл в Word и сохраните как .docx.' : 'Нужен файл Word в формате .docx.', 'error');
      return;
    }
    if (!confirmDiscard()) { return; }
    showBusy('Открываю «' + file.name + '»…');
    var buf;
    file.arrayBuffer().then(function (b) {
      buf = b;
      return convertOpen(b);
    }).then(function (data) { loadDoc(data, file.name, buf); })
      .catch(function (e) { notify('Не удалось открыть документ: ' + e.message, 'error', 0); })
      .then(hideBusy);
  }
  function newDoc(skipConfirm) {
    if (!skipConfirm && !confirmDiscard()) { return Promise.resolve(); }
    showBusy('Создаю документ…');
    return fetch('static/blank.json').then(function (r) { return r.json(); })
      .then(function (data) { if (data.error) { throw new Error(data.error); } loadDoc(data, 'Новый документ.docx', null); })
      .catch(function (e) { notify('Ошибка: ' + e.message, 'error', 0); })
      .then(hideBusy);
  }
  function toDataUrl(blob) {
    return new Promise(function (resolve, reject) {
      var r = new FileReader(); r.onload = function () { resolve(r.result); }; r.onerror = reject; r.readAsDataURL(blob);
    });
  }
  function toPng(src) {
    return new Promise(function (resolve) {
      var img = new Image();
      img.onload = function () {
        var c = document.createElement('canvas');
        c.width = img.naturalWidth || 300; c.height = img.naturalHeight || 150;
        c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
        try { resolve(c.toDataURL('image/png')); } catch (e) { resolve(null); }
      };
      img.onerror = function () { resolve(null); };
      img.src = src;
    });
  }
  function prepareHtml(html) {
    var doc = new DOMParser().parseFromString('<!doctype html><html><body>' + html + '</body></html>', 'text/html');
    var jobs = [];
    Array.prototype.forEach.call(doc.querySelectorAll('img'), function (img) {
      if (img.classList.contains('mce-pagebreak')) { return; }
      if (img.getAttribute('data-raw')) { img.setAttribute('src', ''); return; }
      var src = img.getAttribute('src') || '';
      if (/^blob:/.test(src)) {
        jobs.push(fetch(src).then(function (r) { return r.blob(); }).then(toDataUrl).then(function (u) { return /^data:image\/(png|jpe?g|gif|bmp)/.test(u) ? u : toPng(u); })
          .then(function (u) { if (u) { img.setAttribute('src', u); } }).catch(function () { /* оставим как есть */ }));
      } else if (/^data:image\/(webp|svg\+xml|avif|tiff|x-icon)/.test(src)) {
        jobs.push(toPng(src).then(function (u) { if (u) { img.setAttribute('src', u); } }));
      }
    });
    return Promise.all(jobs).then(function () { return doc.body.innerHTML; });
  }
  function fileNameValue() {
    var n = ($('fileName').value || 'Документ').trim().replace(/[\\/:*?"<>|]+/g, '_');
    return /\.docx$/i.test(n) ? n : n + '.docx';
  }
  function saveDoc() {
    showBusy('Сохраняю документ…');
    prepareHtml(editor.getContent()).then(convertSave).then(function (r) {
      var a = document.createElement('a');
      a.href = URL.createObjectURL(r.blob);
      a.download = fileNameValue();
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 30000);
      editor.setDirty(false); updateStatus();
      notify('Файл «' + fileNameValue() + '» сохранён в загрузки браузера.', 'success', 4000);
      r.warnings.forEach(function (w) { notify(w, 'warning', 10000); });
    }).catch(function (e) { notify('Не удалось сохранить: ' + e.message, 'error', 0); })
      .then(hideBusy);
  }
  function printDoc() {
    computePages();
    editor.getWin().focus();
    editor.getWin().print();
  }

  // ------------------------------------------------------- настройка редактора
  function setupEditor(ed) {
    editor = ed;
    var reg = ed.ui.registry;

    reg.addMenuButton('lwstyles', { text: 'Стили', tooltip: 'Стили абзаца', fetch: function (cb) { cb(styleItems()); } });
    reg.addNestedMenuItem('lwstylesmenu', { text: 'Стиль абзаца', getSubmenuItems: styleItems });
    reg.addButton('lwgrow', { icon: 'text-size-increase', tooltip: 'Увеличить размер шрифта (Ctrl+])', onAction: function () { growFont(1); } });
    reg.addButton('lwshrink', { icon: 'text-size-decrease', tooltip: 'Уменьшить размер шрифта (Ctrl+[)', onAction: function () { growFont(-1); } });
    reg.addButton('lwindent', { icon: 'indent', tooltip: 'Увеличить отступ', onAction: function () { changeIndent(1); } });
    reg.addButton('lwoutdent', { icon: 'outdent', tooltip: 'Уменьшить отступ', onAction: function () { changeIndent(-1); } });
    reg.addButton('lwpara', { text: 'Абзац…', tooltip: 'Отступы и интервалы', onAction: paraDialog });
    reg.addMenuItem('lwpara', { text: 'Абзац…', icon: 'line-height', onAction: paraDialog });
    reg.addMenuItem('lwfont', { text: 'Шрифт…', icon: 'text-color', onAction: fontDialog });
    reg.addButton('lwfont', { icon: 'text-color', tooltip: 'Шрифт… (все параметры)', onAction: fontDialog });
    reg.addMenuButton('lwlinespacing', {
      icon: 'line-height', tooltip: 'Интервал между строками и абзацами',
      fetch: function (cb) {
        var b = blocks()[0];
        var cur = b ? currentLine(b) : { type: 'single', val: 1 };
        var items = [1, 1.15, 1.5, 2, 2.5, 3].map(function (v) {
          return { type: 'togglemenuitem', text: fmt(v, 2), active: Math.abs(cur.val - v) < 0.02 && cur.type !== 'exact' && cur.type !== 'atLeast', onAction: function () { setLine(v); } };
        });
        items.push({ type: 'separator' });
        items.push({ type: 'menuitem', text: 'Добавить интервал перед абзацем (12 пт)', onAction: function () {
          editor.undoManager.transact(function () { blocks().forEach(function (x) { setStyles(x, { 'padding-top': '12pt' }); }); }); schedulePages();
        } });
        items.push({ type: 'menuitem', text: 'Удалить интервал после абзаца', onAction: function () {
          editor.undoManager.transact(function () { blocks().forEach(function (x) { setStyles(x, { 'padding-bottom': '0pt' }); }); }); schedulePages();
        } });
        items.push({ type: 'separator' });
        items.push({ type: 'menuitem', text: 'Другие варианты интервалов…', onAction: paraDialog });
        cb(items);
      }
    });
    reg.addToggleButton('lwpainter', {
      icon: 'format-painter', tooltip: 'Формат по образцу: выделите образец, нажмите, затем выделите текст',
      onAction: function (api) {
        if (state.painter) { state.painter = null; api.setActive(false); return; }
        state.painter = captureFormat(); state.painterApi = api; api.setActive(true);
      }
    });
    reg.addToggleButton('lwmarks', {
      icon: 'visualchars', tooltip: 'Отобразить все знаки (¶)',
      onAction: function (api) { var on = !editor.getBody().classList.contains('lw-marks'); editor.getBody().classList.toggle('lw-marks', on); api.setActive(on); },
      onSetup: function (api) { api.setActive(editor.getBody() && editor.getBody().classList.contains('lw-marks')); return function () {}; }
    });
    reg.addMenuItem('lwmarks', { text: 'Отобразить все знаки', icon: 'visualchars', onAction: function () { editor.getBody().classList.toggle('lw-marks'); } });
    reg.addButton('lwpagesetup', { icon: 'document-properties', tooltip: 'Параметры страницы (поля, размер, ориентация)', onAction: pageDialog });
    reg.addMenuItem('lwpagesetup', { text: 'Параметры страницы…', icon: 'document-properties', onAction: pageDialog });
    reg.addNestedMenuItem('lwmargins', { text: 'Поля', getSubmenuItems: function () {
      return MARGINS.map(function (m) { return { type: 'menuitem', text: m.name, onAction: function () { setPage({ top: m.top, bottom: m.bottom, left: m.left, right: m.right }); } }; })
        .concat([{ type: 'separator' }, { type: 'menuitem', text: 'Настраиваемые поля…', onAction: pageDialog }]);
    } });
    reg.addNestedMenuItem('lworient', { text: 'Ориентация', getSubmenuItems: function () {
      var p = state.page;
      return [
        { type: 'togglemenuitem', text: 'Книжная', active: p.w <= p.h, onAction: function () { if (p.w > p.h) { setPage({ w: p.h, h: p.w }); } } },
        { type: 'togglemenuitem', text: 'Альбомная', active: p.w > p.h, onAction: function () { if (p.w <= p.h) { setPage({ w: p.h, h: p.w }); } } }
      ];
    } });
    reg.addNestedMenuItem('lwsize', { text: 'Размер', getSubmenuItems: function () {
      var p = state.page, land = p.w > p.h;
      return PAPER.map(function (x) {
        return { type: 'togglemenuitem', text: x.name, active: (x.w === p.w && x.h === p.h) || (x.w === p.h && x.h === p.w),
          onAction: function () { setPage(land ? { w: x.h, h: x.w } : { w: x.w, h: x.h }); } };
      });
    } });
    reg.addNestedMenuItem('lwzoom', { text: 'Масштаб', getSubmenuItems: function () {
      return [0.5, 0.75, 1, 1.25, 1.5, 2].map(function (z) { return { type: 'togglemenuitem', text: Math.round(z * 100) + '%', active: Math.abs(state.zoom - z) < 0.01, onAction: function () { setZoom(z); } }; });
    } });
    reg.addMenuItem('lwnew', { text: 'Создать', icon: 'new-document', onAction: function () { newDoc(false); } });
    reg.addMenuItem('lwopen', { text: 'Открыть… (Ctrl+O)', icon: 'upload', onAction: function () { $('fileInput').click(); } });
    reg.addMenuItem('lwsave', { text: 'Скачать .docx (Ctrl+S)', icon: 'save', onAction: saveDoc });
    reg.addMenuItem('lwprint', { text: 'Печать / PDF (Ctrl+P)', icon: 'print', onAction: printDoc });
    reg.addMenuItem('lwtab', { text: 'Знак табуляции (Tab)', onAction: function () { editor.insertContent(TAB_HTML); } });

    // TinyMCE переводит px в pt через parseInt и для 11pt (14.67px) показывает 10.5pt — отдаём размер сами
    ed.on('init', function () {
      ed.addQueryValueHandler('FontSize', function () {
        var n = ed.selection.getNode();
        if (n && n.nodeType !== 1) { n = n.parentNode; }
        var px = n ? parseFloat(ed.getWin().getComputedStyle(n).fontSize) : 0;
        return px ? (Math.round(px * 0.75 * 2) / 2) + 'pt' : '';
      });
    });

    ed.addShortcut('meta+s', 'Сохранить', saveDoc);
    ed.addShortcut('meta+o', 'Открыть', function () { $('fileInput').click(); });
    ed.addShortcut('meta+p', 'Печать', printDoc);
    ed.addShortcut('meta+13', 'Разрыв страницы', function () { editor.execCommand('mcePageBreak'); });
    ed.addShortcut('meta+shift+32', 'Неразрывный пробел', function () { editor.insertContent(' '); });
    ed.addShortcut('meta+e', 'По центру', function () { editor.execCommand('JustifyCenter'); });
    ed.addShortcut('meta+l', 'По левому краю', function () { editor.execCommand('JustifyLeft'); });
    ed.addShortcut('meta+r', 'По правому краю', function () { editor.execCommand('JustifyRight'); });
    ed.addShortcut('meta+j', 'По ширине', function () { editor.execCommand('JustifyFull'); });
    ed.addShortcut('meta+1', 'Одинарный интервал', function () { setLine(1); });
    ed.addShortcut('meta+5', 'Полуторный интервал', function () { setLine(1.5); });
    ed.addShortcut('meta+2', 'Двойной интервал', function () { setLine(2); });
    ed.addShortcut('meta+221', 'Крупнее', function () { growFont(1); });
    ed.addShortcut('meta+219', 'Мельче', function () { growFont(-1); });
    ed.addShortcut('meta+d', 'Шрифт', fontDialog);

    ed.on('keydown', function (e) {
      if (e.key === 'Tab' && !e.ctrlKey && !e.altKey && !e.metaKey) {
        if (ed.dom.getParent(ed.selection.getNode(), 'li,td,th')) { return; }
        e.preventDefault();
        if (!e.shiftKey) { ed.insertContent(TAB_HTML); }
      }
      if (e.key === 'Escape' && state.painter) { state.painter = null; if (state.painterApi) { state.painterApi.setActive(false); } }
    });
    ed.on('NewBlock', function (e) {
      var nb = e.newBlock;
      if (!nb || nb.nodeName === 'LI') { return; }
      var prev = nb.previousElementSibling;
      if (!prev || (nb.textContent || '').replace(/[​﻿]/g, '') !== '') { return; }
      var cur = styleOf(prev), next = cur && state.nextStyles[cur];
      if (next && next !== cur) {
        String(nb.className || '').split(/\s+/).forEach(function (c) { if (c.indexOf('s-') === 0) { ed.dom.removeClass(nb, c); } });
        ed.dom.addClass(nb, 's-' + next);
        setStyles(nb, { 'padding-top': '', 'padding-bottom': '', 'line-height': '', 'font-size': '', 'font-family': '', color: '', 'text-align': '' });
        nb.innerHTML = '<br data-mce-bogus="1">';
        ed.selection.setCursorLocation(nb, 0);
      }
    });
    ed.on('ExecCommand', function (e) {
      if (/^(Indent|Outdent|InsertUnorderedList|InsertOrderedList|RemoveList|mceListProps|mceToggleList)$/i.test(e.command)) { normalizeLists(); }
      schedulePages();
    });
    ed.on('mouseup keyup', function () {
      if (state.painter && !ed.selection.isCollapsed()) {
        applyCaptured(state.painter); state.painter = null;
        if (state.painterApi) { state.painterApi.setActive(false); }
      }
    });
    ed.on('input SetContent Undo Redo ObjectResized', schedulePages);
    ed.on('NodeChange', debounce(updateStatus, 80));
    ed.on('dirty change', function () { $('dirtyMark').hidden = !ed.isDirty(); });
    ed.on('ScrollContent', debounce(updateStatus, 120));
    ed.on('dragover', function (e) { if (hasDocx(e.dataTransfer)) { e.preventDefault(); } });
    ed.on('drop', function (e) {
      var f = docxFrom(e.dataTransfer);
      if (f) { e.preventDefault(); e.stopImmediatePropagation(); openFile(f); }
    });
  }

  function hasDocx(dt) {
    if (!dt || !dt.items) { return false; }
    return Array.prototype.some.call(dt.items, function (i) { return i.kind === 'file'; }) &&
      Array.prototype.some.call(dt.types || [], function (t) { return t === 'Files'; });
  }
  function docxFrom(dt) {
    if (!dt || !dt.files) { return null; }
    return Array.prototype.filter.call(dt.files, function (f) { return /\.docx?$/i.test(f.name); })[0] || null;
  }

  function onInit(ed) {
    state.ready = true;
    detectMode(); // в режиме «без сервера» модуль Word начинает грузиться сразу, пока человек выбирает файл
    ed.getDoc().addEventListener('load', schedulePages, true);
    ed.getWin().addEventListener('resize', schedulePages);
    applyPageCss();
    newDoc(true).then(function () { $('welcome').hidden = false; });
  }

  // --------------------------------------------------------- оболочка
  function initShell() {
    $('btnOpen').onclick = $('wOpen').onclick = function () { $('fileInput').click(); };
    $('btnNew').onclick = function () { newDoc(false); };
    $('wNew').onclick = function () { $('welcome').hidden = true; editor.focus(); };
    $('btnSave').onclick = saveDoc;
    $('btnPrint').onclick = printDoc;
    $('fileInput').onchange = function () { var f = this.files[0]; this.value = ''; openFile(f); };
    $('fileName').addEventListener('change', function () { $('fileName').value = fileNameValue(); });
    $('zoomRange').oninput = function () { setZoom(this.value / 100); };
    $('zoomIn').onclick = function () { setZoom(state.zoom + 0.1); };
    $('zoomOut').onclick = function () { setZoom(state.zoom - 0.1); };
    window.addEventListener('resize', schedulePages);
    window.addEventListener('keydown', function (e) {
      if (!(e.ctrlKey || e.metaKey) || !editor) { return; }
      var k = e.key.toLowerCase();
      if (k === 's' || k === 'ы') { e.preventDefault(); saveDoc(); }
      else if (k === 'o' || k === 'щ') { e.preventDefault(); $('fileInput').click(); }
      else if (k === 'p' || k === 'з') { e.preventDefault(); printDoc(); }
    });
    var dragDepth = 0;
    window.addEventListener('dragenter', function (e) { if (hasDocx(e.dataTransfer)) { dragDepth++; $('dropZone').hidden = false; } });
    window.addEventListener('dragleave', function () { dragDepth = Math.max(0, dragDepth - 1); if (!dragDepth) { $('dropZone').hidden = true; } });
    window.addEventListener('dragover', function (e) { if (hasDocx(e.dataTransfer)) { e.preventDefault(); } });
    window.addEventListener('drop', function (e) {
      dragDepth = 0; $('dropZone').hidden = true;
      var f = docxFrom(e.dataTransfer);
      if (f) { e.preventDefault(); openFile(f); }
    });
    window.addEventListener('beforeunload', function (e) {
      if (editor && editor.isDirty()) { e.preventDefault(); e.returnValue = ''; }
    });
  }

  initShell();
  tinymce.init({
    selector: '#editor',
    license_key: 'gpl',
    // скрипты TinyMCE приходят одним файлом (static/tinymce-bundle.js), поэтому путь к скинам указываем явно;
    // относительно страницы — сайт может жить в подпапке (GitHub Pages: /maminword/)
    base_url: new URL('static/vendor/tinymce', location.href).href,
    suffix: '.min',
    language: 'ru',
    height: '100%',
    resize: false,
    branding: false,
    promotion: false,
    statusbar: false,
    menubar: 'file edit view insert format layout table',
    menu: {
      file: { title: 'Файл', items: 'lwnew lwopen lwsave | lwpagesetup | lwprint' },
      edit: { title: 'Правка', items: 'undo redo | cut copy paste pastetext | selectall | searchreplace' },
      view: { title: 'Вид', items: 'lwmarks visualblocks | lwzoom' },
      insert: { title: 'Вставка', items: 'image link inserttable | lwtab pagebreak | charmap nonbreaking' },
      format: { title: 'Формат', items: 'lwfont lwpara | lwstylesmenu | bold italic underline strikethrough superscript subscript | forecolor backcolor | removeformat' },
      layout: { title: 'Макет', items: 'lwpagesetup | lwmargins lworient lwsize' },
      table: { title: 'Таблица', items: 'inserttable | cell row column | tableprops deletetable' }
    },
    plugins: 'advlist lists table image link searchreplace pagebreak charmap wordcount visualchars visualblocks nonbreaking autolink',
    toolbar: 'undo redo lwpainter | lwstyles | fontfamily fontsize lwgrow lwshrink lwfont | bold italic underline strikethrough subscript superscript | forecolor backcolor removeformat | alignleft aligncenter alignright alignjustify | bullist numlist lwoutdent lwindent lwlinespacing lwpara | table image link pagebreak | lwpagesetup searchreplace lwmarks',
    toolbar_mode: window.innerWidth < 1000 ? 'sliding' : 'wrap', // в узком окне панель сворачивается в «…»
    content_css: false,
    content_style: BASE_CSS,
    body_class: 'lw-doc',
    font_family_formats: fontsFmt(FONTS),
    // TinyMCE показывает размер в пт, только если он есть в списке — поэтому список подробный
    font_size_formats: '7pt 7.5pt 8pt 8.5pt 9pt 9.5pt 10pt 10.5pt 11pt 11.5pt 12pt 12.5pt 13pt 13.5pt 14pt 15pt 16pt 18pt 20pt 22pt 24pt 26pt 28pt 32pt 36pt 40pt 48pt 56pt 72pt',
    font_size_input_default_unit: 'pt',
    color_map: COLOR_MAP,
    color_cols: 10,
    custom_colors: true,
    valid_elements: '*[*]',
    verify_html: false,
    entity_encoding: 'raw',
    convert_urls: false,
    paste_data_images: true,
    automatic_uploads: true,
    images_upload_handler: function (blobInfo) { return Promise.resolve('data:' + blobInfo.blob().type + ';base64,' + blobInfo.base64()); },
    images_file_types: 'jpg,jpeg,png,gif,bmp,webp',
    image_uploadtab: true,
    image_description: false,
    file_picker_types: 'image',
    file_picker_callback: function (cb) {
      var inp = document.createElement('input');
      inp.type = 'file'; inp.accept = 'image/*';
      inp.onchange = function () {
        var f = inp.files[0];
        if (f) { toDataUrl(f).then(function (u) { cb(u, { title: f.name }); }); }
      };
      inp.click();
    },
    object_resizing: 'img',
    table_default_attributes: { border: '1' },
    table_default_styles: { 'border-collapse': 'collapse', width: '100%' },
    table_sizing_mode: 'fixed',
    table_use_colgroups: true,
    pagebreak_separator: '<!-- pagebreak -->',
    pagebreak_split_block: true,
    indent_use_margin: true,
    lists_indent_on_tab: true,
    advlist_bullet_styles: 'default,circle,square',
    advlist_number_styles: 'default,lower-alpha,upper-alpha,lower-roman,upper-roman',
    link_title: false,
    link_target_list: false,
    link_assume_external_targets: 'https',
    browser_spellcheck: true,
    contextmenu: 'link image table',
    setup: setupEditor,
    init_instance_callback: onInit
  });
})();
