/*
 * 研修報告書 Excel(.xlsx) 生成モジュール
 * 外部ライブラリ不要。既存の「研修記録」Excelと同じ書式（A〜AL列の方眼、MS明朝/MSゴシック、
 * グレー見出し）で、1シートの報告書を作成します。
 */
(function (global) {
  'use strict';

  // ---------- 文字幅・折り返し ----------
  // 全角=2, 半角=1 として幅を数える
  function charWidth(ch) {
    var c = ch.codePointAt(0);
    if (c <= 0x7e) return 1;
    if (c >= 0xff61 && c <= 0xff9f) return 1; // 半角カナ
    return 2;
  }
  function strWidth(s) {
    var w = 0;
    for (var ch of s) w += charWidth(ch);
    return w;
  }

  // 行頭に来てはいけない文字（行頭禁則）
  var NO_START = '、。，．,.）)」』】〕］]｝}〉》・ー～ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ！？!?：；:;％%';

  /**
   * テキストを指定幅で折り返し、行の配列を返す。
   * - 入力時の改行はそのまま維持
   * - 行頭の空白（字下げ）は折り返し後の行にも引き継ぐ
   * - 行頭禁則文字は前の行にぶら下げる
   */
  function wrapText(text, maxWidth) {
    var out = [];
    var src = String(text || '').replace(/\r\n?/g, '\n').split('\n');
    // 末尾の空行は削除
    while (src.length && src[src.length - 1].trim() === '') src.pop();
    src.forEach(function (line) {
      line = line.replace(/\t/g, '　');
      if (strWidth(line) <= maxWidth) { out.push(line); return; }
      var m = line.match(/^[ 　]*/)[0];
      var indent = m;
      // 箇条書き（・や数字）で始まる行は、その記号の分だけ続きの行を下げる
      var rest = line.slice(m.length);
      if (/^[・●○◆■□▶-]/.test(rest)) indent = m + '　';
      if (strWidth(indent) > maxWidth / 3) indent = '';
      var chars = Array.from(line);
      var cur = '', curW = 0, first = true;
      for (var i = 0; i < chars.length; i++) {
        var ch = chars[i], w = charWidth(ch);
        if (curW + w > maxWidth && cur.trim() !== '') {
          // 禁則：次の文字が行頭禁則なら最大2文字までぶら下げ
          var hang = 0;
          while (i < chars.length && NO_START.indexOf(chars[i]) >= 0 && hang < 2) {
            cur += chars[i]; i++; hang++;
          }
          out.push(cur);
          first = false;
          cur = indent; curW = strWidth(indent);
          if (i >= chars.length) { cur = ''; break; }
          ch = chars[i]; w = charWidth(ch);
          // 続きの行の先頭の空白は詰める
          if (ch === ' ' || ch === '　') continue;
        }
        cur += ch; curW += w;
      }
      if (cur !== '' && cur !== indent) out.push(cur);
    });
    return out;
  }

  // ---------- XML ユーティリティ ----------
  function esc(s) {
    return String(s)
      .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }
  function colName(n) { // 1 -> A
    var s = '';
    while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  var LAST_COL = 38; // AL

  // ---------- スタイル定義 ----------
  // fonts: 0=明朝11, 1=明朝10, 2=ゴシック18, 3=ゴシック11, 4=ゴシック10 白
  // fills: 0=none, 1=gray125, 2=グレー(#595959)
  // borders: 下記 BORDERS の順
  var BORDERS = [''];
  var styleKeys = []; // "font|fill|border|h|v|wrap"
  function sid(font, fill, border, h, v) {
    var key = [font, fill, border, h || '', v || ''].join('|');
    var i = styleKeys.indexOf(key);
    if (i < 0) { styleKeys.push(key); i = styleKeys.length - 1; }
    return i + 1; // 0 は既定スタイル
  }
  function bIdx(sides) {
    sides = sides.split('').sort(function (a, b) { return 'ltrb'.indexOf(a) - 'ltrb'.indexOf(b); }).join('');
    var i = BORDERS.indexOf(sides);
    if (i < 0) { BORDERS.push(sides); i = BORDERS.length - 1; }
    return i;
  }

  function stylesXml() {
    var fonts =
      '<fonts count="5">' +
      '<font><sz val="11"/><name val="ＭＳ 明朝"/><family val="1"/><charset val="128"/></font>' +
      '<font><sz val="10"/><name val="ＭＳ 明朝"/><family val="1"/><charset val="128"/></font>' +
      '<font><sz val="18"/><name val="ＭＳ ゴシック"/><family val="3"/><charset val="128"/></font>' +
      '<font><sz val="11"/><name val="ＭＳ ゴシック"/><family val="3"/><charset val="128"/></font>' +
      '<font><sz val="10"/><color rgb="FFFFFFFF"/><name val="ＭＳ ゴシック"/><family val="3"/><charset val="128"/></font>' +
      '</fonts>';
    var fills =
      '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FF595959"/><bgColor indexed="64"/></patternFill></fill></fills>';
    var borders = '<borders count="' + BORDERS.length + '">' + BORDERS.map(function (s) {
      function side(tag, ch) { return s.indexOf(ch) >= 0 ? '<' + tag + ' style="thin"><color indexed="64"/></' + tag + '>' : '<' + tag + '/>'; }
      return '<border>' + side('left', 'l') + side('right', 'r') + side('top', 't') + side('bottom', 'b') + '<diagonal/></border>';
    }).join('') + '</borders>';
    var xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf>'];
    styleKeys.forEach(function (k) {
      var p = k.split('|');
      var al = '<alignment' + (p[3] ? ' horizontal="' + p[3] + '"' : '') + (p[4] ? ' vertical="' + p[4] + '"' : '') + '/>';
      xfs.push('<xf numFmtId="0" fontId="' + p[0] + '" fillId="' + p[1] + '" borderId="' + p[2] + '" xfId="0"' +
        (p[0] !== '0' ? ' applyFont="1"' : '') + (p[1] !== '0' ? ' applyFill="1"' : '') +
        (p[2] !== '0' ? ' applyBorder="1"' : '') + ' applyAlignment="1">' + al + '</xf>');
    });
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      fonts + fills + borders +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="' + xfs.length + '">' + xfs.join('') + '</cellXfs>' +
      '<cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>';
  }

  // ---------- シート構築 ----------
  function Sheet() {
    this.rows = {};      // r -> {height, cells: {c: {v, s}}}
    this.merges = [];
  }
  Sheet.prototype.row = function (r, height) {
    if (!this.rows[r]) this.rows[r] = { height: height, cells: {} };
    if (height) this.rows[r].height = height;
    return this.rows[r];
  };
  Sheet.prototype.set = function (r, c, v, s) {
    var cell = this.row(r).cells[c] || (this.row(r).cells[c] = {});
    if (v !== undefined && v !== null) cell.v = v;
    if (s) cell.s = s;
  };
  // 範囲 [r1,c1]-[r2,c2] に枠線を付けつつ結合（外周のみ線）
  Sheet.prototype.box = function (r1, c1, r2, c2, font, fill, h, v, edges, merge) {
    edges = edges || 'ltrb';
    for (var r = r1; r <= r2; r++) {
      for (var c = c1; c <= c2; c++) {
        var sides = '';
        if (c === c1 && edges.indexOf('l') >= 0) sides += 'l';
        if (c === c2 && edges.indexOf('r') >= 0) sides += 'r';
        if (r === r1 && edges.indexOf('t') >= 0) sides += 't';
        if (r === r2 && edges.indexOf('b') >= 0) sides += 'b';
        this.set(r, c, undefined, sid(font, fill, bIdx(sides), h, v));
      }
    }
    if (merge !== false && (r1 !== r2 || c1 !== c2)) this.merges.push(colName(c1) + r1 + ':' + colName(c2) + r2);
  };
  Sheet.prototype.xml = function () {
    var cols = '<cols>';
    for (var c = 1; c <= LAST_COL + 1; c++) {
      var w = c <= LAST_COL ? 2.875 : 2.625;
      cols += '<col min="' + c + '" max="' + c + '" width="' + w + '" customWidth="1"/>';
    }
    cols += '</cols>';
    var rowNums = Object.keys(this.rows).map(Number).sort(function (a, b) { return a - b; });
    var maxRow = rowNums[rowNums.length - 1] || 1;
    var data = '<sheetData>';
    var self = this;
    rowNums.forEach(function (r) {
      var row = self.rows[r];
      data += '<row r="' + r + '"' + (row.height ? ' ht="' + row.height + '" customHeight="1"' : '') + '>';
      Object.keys(row.cells).map(Number).sort(function (a, b) { return a - b; }).forEach(function (c) {
        var cell = row.cells[c], ref = colName(c) + r;
        var s = cell.s ? ' s="' + cell.s + '"' : '';
        if (cell.v === undefined || cell.v === '') data += '<c r="' + ref + '"' + s + '/>';
        else if (typeof cell.v === 'number') data += '<c r="' + ref + '"' + s + '><v>' + cell.v + '</v></c>';
        else data += '<c r="' + ref + '"' + s + ' t="inlineStr"><is><t xml:space="preserve">' + esc(cell.v) + '</t></is></c>';
      });
      data += '</row>';
    });
    data += '</sheetData>';
    var merges = this.merges.length ? '<mergeCells count="' + this.merges.length + '">' +
      this.merges.map(function (m) { return '<mergeCell ref="' + m + '"/>'; }).join('') + '</mergeCells>' : '';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' +
      '<dimension ref="A1:' + colName(LAST_COL) + maxRow + '"/>' +
      '<sheetViews><sheetView showGridLines="0" tabSelected="1" workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>' +
      '<sheetFormatPr baseColWidth="8" defaultColWidth="2.625" defaultRowHeight="21.95" customHeight="1"/>' +
      cols + data + merges +
      '<pageMargins left="0.31496062992125984" right="0.2755905511811024" top="0.3937007874015748" bottom="0.3937007874015748" header="0.5118110236220472" footer="0.5118110236220472"/>' +
      '<pageSetup paperSize="9" fitToHeight="0" orientation="portrait"/>' +
      '</worksheet>';
  };

  // ---------- 報告書レイアウト ----------
  var BODY_WIDTH = 116;   // 本文1行の幅（半角換算）。既存の報告書の最長行に合わせた値
  var FIELD_WIDTH = 104;  // テーマ・講師・目的欄の幅
  var MIN_LAST_ROW = 40;  // A4 1枚分になるよう最低この行まで枠を作る

  var F_M11 = 0, F_M10 = 1, F_G18 = 2, F_G11 = 3, F_LBL = 4;

  function buildSheet(rep) {
    styleKeys = [];
    var sh = new Sheet();
    var full = LAST_COL;

    // 1行目：タイトル・報告日
    sh.row(1, 30);
    for (var c = 1; c <= full; c++) sh.set(1, c, undefined, sid(F_M11, 0, bIdx('b'), '', 'center'));
    sh.set(1, 2, '研修報告書', sid(F_G18, 0, bIdx('b'), '', 'center'));
    var d = rep.reportDate ? rep.reportDate.split('-').map(Number) : null;
    function ymd(c1, c2, val, label, lc) {
      sh.box(1, c1, 1, c2, F_M11, 0, 'center', 'center', 'b');
      if (val !== null && val !== undefined) sh.set(1, c1, val);
      sh.set(1, lc, label, sid(F_G11, 0, bIdx('b'), 'center', 'center'));
    }
    ymd(26, 30, d ? d[0] : null, '年', 31);   // Z1:AD1, AE1
    ymd(32, 33, d ? d[1] : null, '月', 34);   // AF1:AG1, AH1
    ymd(35, 36, d ? d[2] : null, '日', 37);   // AI1:AJ1, AK1

    sh.row(2, 9.95);
    // 3行目：氏名
    sh.row(3, 24.95);
    sh.box(3, 1, 3, 3, F_LBL, 2, 'center', 'center');
    sh.set(3, 1, '氏名');
    sh.box(3, 4, 3, full, F_M11, 0, 'center', 'center');
    sh.set(3, 4, rep.name || '');
    sh.row(4, 9.95); sh.row(5, 18);

    // 6行目：日時・場所
    sh.row(6, 21.95);
    sh.box(6, 1, 6, 3, F_LBL, 2, 'center', 'center'); sh.set(6, 1, '日　時');
    sh.box(6, 4, 6, 19, F_M11, 0, 'center', 'center'); sh.set(6, 4, rep.dateText || '');
    sh.box(6, 20, 6, 22, F_LBL, 2, 'center', 'center'); sh.set(6, 20, '場所');
    sh.box(6, 23, 6, full, F_M11, 0, 'center', 'center'); sh.set(6, 23, rep.place || '');

    // 7・8行目：テーマ・講師
    [[7, 'テーマ', rep.theme], [8, '講　師', rep.lecturer]].forEach(function (x) {
      sh.row(x[0], 21.95);
      sh.box(x[0], 1, x[0], 3, F_LBL, 2, 'center', 'center'); sh.set(x[0], 1, x[1]);
      sh.box(x[0], 4, x[0], full, F_M10, 0, '', 'center');
      sh.set(x[0], 4, String(x[2] || '').replace(/\s*\n\s*/g, '　').trim());
    });

    // 9・10行目：目的（2行）
    sh.row(9, 21.95); sh.row(10, 21.95);
    sh.box(9, 1, 10, 3, F_LBL, 2, 'center', 'center'); sh.set(9, 1, '目　的');
    var purpose = wrapText(rep.purpose, FIELD_WIDTH);
    if (purpose.length > 2) purpose = [purpose[0], purpose.slice(1).join('')];
    sh.box(9, 4, 9, full, F_M10, 0, '', 'center', 'ltr'); sh.set(9, 4, purpose[0] || '');
    sh.box(10, 4, 10, full, F_M10, 0, '', 'center', 'lrb'); sh.set(10, 4, purpose[1] || '');
    sh.row(11, 9.95); sh.row(12, 18);

    // 本文セクション
    var r = 13;
    function header(title) {
      sh.row(r, 18);
      sh.box(r, 1, r, full, F_LBL, 2, 'center', 'center');
      sh.set(r, 1, title);
      r++;
    }
    function bodyLines(lines, last) {
      lines.forEach(function (text, i) {
        sh.row(r, 21.95);
        var edges = 'lr' + (i === 0 ? 't' : '') + (last && i === lines.length - 1 ? 'b' : '');
        sh.box(r, 1, r, full, F_M10, 0, '', 'center', edges);
        sh.set(r, 1, text);
        r++;
      });
    }
    var content = wrapText(rep.content, BODY_WIDTH);
    var impression = wrapText(rep.impression, BODY_WIDTH);
    var notes = wrapText(rep.notes, BODY_WIDTH);
    while (content.length < 5) content.push('');
    // 最終行 = 13 + 1 + C + 1 + I + 1 + N - 1 → A4 1枚に収まるよう所感欄で調整
    var minI = Math.max(5, MIN_LAST_ROW - (15 + content.length + Math.max(3, notes.length)));
    while (impression.length < minI) impression.push('');
    while (notes.length < 3) notes.push('');

    header('研修内容'); bodyLines(content);
    header('所　　感'); bodyLines(impression);
    header('備　　考'); bodyLines(notes, true);

    return sh;
  }

  // ---------- ZIP（無圧縮）----------
  var CRC_TABLE = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(buf) {
    var c = 0xffffffff;
    for (var i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }
  function zip(files) { // files: [{name, data: Uint8Array}]
    var enc = new TextEncoder();
    var now = new Date();
    var dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    var dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    var parts = [], central = [], offset = 0;
    files.forEach(function (f) {
      var name = enc.encode(f.name), data = f.data, crc = crc32(data);
      var h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true);
      h.setUint16(8, 0, true); h.setUint16(10, dosTime, true); h.setUint16(12, dosDate, true);
      h.setUint32(14, crc, true); h.setUint32(18, data.length, true); h.setUint32(22, data.length, true);
      h.setUint16(26, name.length, true); h.setUint16(28, 0, true);
      parts.push(new Uint8Array(h.buffer), name, data);
      var ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true);
      ch.setUint16(8, 0x0800, true); ch.setUint16(10, 0, true); ch.setUint16(12, dosTime, true);
      ch.setUint16(14, dosDate, true); ch.setUint32(16, crc, true); ch.setUint32(20, data.length, true);
      ch.setUint32(24, data.length, true); ch.setUint16(28, name.length, true);
      ch.setUint32(42, offset, true);
      central.push(new Uint8Array(ch.buffer), name);
      offset += 30 + name.length + data.length;
    });
    var cdSize = central.reduce(function (s, p) { return s + p.length; }, 0);
    var end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true);
    end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
    end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
    var all = parts.concat(central, [new Uint8Array(end.buffer)]);
    var total = all.reduce(function (s, p) { return s + p.length; }, 0);
    var out = new Uint8Array(total), pos = 0;
    all.forEach(function (p) { out.set(p, pos); pos += p.length; });
    return out;
  }

  // ---------- 公開関数 ----------
  /**
   * rep: {name, reportDate:'YYYY-MM-DD', dateText, place, theme, lecturer, purpose, content, impression, notes}
   * 戻り値: Uint8Array (.xlsx)
   */
  function buildXlsx(rep) {
    var sheet = buildSheet(rep);
    var sheetXml = sheet.xml();
    var enc = new TextEncoder();
    var iso = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    var files = {
      '[Content_Types].xml':
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
        '</Types>',
      '_rels/.rels':
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
        '</Relationships>',
      'docProps/core.xml':
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
        '<dc:title>研修報告書</dc:title><dc:creator>' + esc(rep.name || '') + '</dc:creator>' +
        '<dcterms:created xsi:type="dcterms:W3CDTF">' + iso + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + iso + '</dcterms:modified>' +
        '</cp:coreProperties>',
      'docProps/app.xml':
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>',
      'xl/workbook.xml':
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<bookViews><workbookView/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels':
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '</Relationships>',
      'xl/worksheets/sheet1.xml': sheetXml,
      'xl/styles.xml': stylesXml()
    };
    return zip(Object.keys(files).map(function (n) { return { name: n, data: enc.encode(files[n]) }; }));
  }

  // プレビュー用：各欄の行分割結果を返す
  function layoutPreview(rep) {
    return {
      content: wrapText(rep.content, BODY_WIDTH),
      impression: wrapText(rep.impression, BODY_WIDTH),
      notes: wrapText(rep.notes, BODY_WIDTH),
      purpose: wrapText(rep.purpose, FIELD_WIDTH)
    };
  }

  global.KenshuXlsx = { buildXlsx: buildXlsx, wrapText: wrapText, strWidth: strWidth, layoutPreview: layoutPreview };
})(typeof window !== 'undefined' ? window : this);
