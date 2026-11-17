(function () {
  'use strict';
  var APP_VERSION = '1.0.0';
  var STORE_KEY = 'kenshu.records.v1';
  var SETTINGS_KEY = 'kenshu.settings.v1';
  var MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  var FIELDS = ['theme', 'trainDate', 'timeFrom', 'timeTo', 'place', 'lecturer', 'purpose',
    'content', 'impression', 'notes', 'name', 'reportDate', 'shortTitle'];

  var $ = function (id) { return document.getElementById(id); };

  // ---------- 保存 ----------
  function load(key, def) {
    try { var v = localStorage.getItem(key); return v ? JSON.parse(v) : def; } catch (e) { return def; }
  }
  function save(key, val) {
    try { localStorage.setItem(key, JSON.stringify(val)); return true; } catch (e) { return false; }
  }
  var records = load(STORE_KEY, []);
  var settings = load(SETTINGS_KEY, { name: '', place: '事務所' });
  var current = null;

  function persist() {
    if (!save(STORE_KEY, records)) toast('保存できませんでした（プライベートブラウズでは保存できません）');
  }

  // ---------- 日付 ----------
  function today() {
    var d = new Date();
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
  }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function parseYmd(s) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || '');
    return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
  }
  var WEEK = '日月火水木金土';
  function jpDate(s) {
    var d = parseYmd(s);
    if (!d) return '';
    return d.getFullYear() + '年' + (d.getMonth() + 1) + '月' + d.getDate() + '日（' + WEEK[d.getDay()] + '）';
  }
  function dateText(r) {
    var t = jpDate(r.trainDate);
    if (r.timeFrom || r.timeTo) t += '　' + (r.timeFrom || '') + '～' + (r.timeTo || '');
    return t;
  }

  // ---------- 報告書データ・ファイル名 ----------
  function toReport(r) {
    return {
      name: r.name, reportDate: r.reportDate || r.trainDate, dateText: dateText(r),
      place: r.place, theme: r.theme, lecturer: r.lecturer, purpose: r.purpose,
      content: r.content, impression: r.impression, notes: r.notes
    };
  }
  function fileName(r) {
    var d = parseYmd(r.trainDate);
    var ds = d ? String(d.getFullYear()).slice(2) + '.' + (d.getMonth() + 1) + '.' + d.getDate() : '日付未定';
    var title = (r.shortTitle || r.theme || '研修').trim();
    var family = (r.name || '').trim().split(/[ 　]+/)[0];
    var n = '研修記録（' + ds + ' ' + title + '）' + family;
    return n.replace(/[\\\/:*?"<>|\n\r\t]/g, '').slice(0, 100) + '.xlsx';
  }

  // ---------- 画面切替 ----------
  var views = ['list', 'edit', 'preview', 'settings'];
  function show(name) {
    views.forEach(function (v) { $('view-' + v).classList.toggle('hidden', v !== name); });
    window.scrollTo(0, 0);
  }

  // ---------- 一覧 ----------
  function renderList() {
    var ul = $('list');
    ul.innerHTML = '';
    var sorted = records.slice().sort(function (a, b) {
      return (b.trainDate || '').localeCompare(a.trainDate || '') || b.updated - a.updated;
    });
    $('list-empty').classList.toggle('hidden', sorted.length > 0);
    $('list-title').classList.toggle('hidden', sorted.length === 0);
    sorted.forEach(function (r) {
      var li = document.createElement('li');
      var b = document.createElement('button');
      var t = document.createElement('span'); t.className = 't'; t.textContent = r.theme || '（テーマ未入力）';
      var m = document.createElement('span'); m.className = 'm';
      m.textContent = (jpDate(r.trainDate) || '日付未入力') + (r.lecturer ? '　' + r.lecturer : '');
      b.appendChild(t); b.appendChild(m);
      b.addEventListener('click', function () { openRecord(r.id); });
      li.appendChild(b); ul.appendChild(li);
    });
  }

  // ---------- 編集 ----------
  function newRecord(base) {
    var r = {
      id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
      created: Date.now(), updated: Date.now(),
      theme: '', trainDate: today(), timeFrom: '', timeTo: '', place: settings.place || '',
      lecturer: '', purpose: '', content: '', impression: '', notes: '',
      name: settings.name || '', reportDate: '', shortTitle: ''
    };
    if (base) FIELDS.forEach(function (k) { if (base[k] !== undefined) r[k] = base[k]; });
    records.push(r); persist();
    return r;
  }
  function openRecord(id) {
    current = records.filter(function (r) { return r.id === id; })[0];
    if (!current) return;
    FIELDS.forEach(function (k) { var el = $('f-' + k); if (el) el.value = current[k] || ''; });
    $('saved-msg').textContent = '';
    updateEditHeader();
    show('edit');
  }
  function updateEditHeader() {
    $('edit-title').textContent = current.theme || '新しい記録';
    $('fname-hint').textContent = 'ファイル名：' + fileName(current);
  }
  var saveTimer = null;
  function onInput(e) {
    var k = e.target.getAttribute('data-k');
    if (!k || !current) return;
    current[k] = e.target.value;
    current.updated = Date.now();
    updateEditHeader();
    clearTimeout(saveTimer);
    saveTimer = setTimeout(function () {
      persist();
      var d = new Date();
      $('saved-msg').textContent = '自動保存しました ' + pad(d.getHours()) + ':' + pad(d.getMinutes());
    }, 400);
  }
  function flush() { if (saveTimer) { clearTimeout(saveTimer); saveTimer = null; persist(); } }

  // カーソル位置に文字を挿入（行頭用ボタン）
  function insertAtCursor(ta, text) {
    var s = ta.selectionStart, e = ta.selectionEnd, v = ta.value;
    ta.value = v.slice(0, s) + text + v.slice(e);
    ta.selectionStart = ta.selectionEnd = s + text.length;
    ta.dispatchEvent(new Event('input', { bubbles: true }));
  }

  // ---------- Excel 出力 ----------
  function exportExcel() {
    if (!current) return;
    flush();
    var bytes;
    try { bytes = KenshuXlsx.buildXlsx(toReport(current)); }
    catch (e) { toast('Excelの作成に失敗しました：' + e.message); return; }
    shareFile(bytes, fileName(current), MIME);
  }
  function shareFile(bytes, name, type) {
    var blob = new Blob([bytes], { type: type });
    var file = null;
    try { file = new File([blob], name, { type: type }); } catch (e) { /* 古いブラウザ */ }
    if (file && navigator.canShare && navigator.canShare({ files: [file] })) {
      navigator.share({ files: [file] }).catch(function (err) {
        if (err && err.name === 'AbortError') return; // キャンセル
        download(blob, name);
      });
    } else {
      download(blob, name);
    }
  }
  function download(blob, name) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 10000);
    toast('ファイルを保存しました');
  }

  // ---------- プレビュー ----------
  function renderPreview() {
    var rep = toReport(current);
    var lay = KenshuXlsx.layoutPreview(rep);
    var d = parseYmd(rep.reportDate);
    var el = $('sheet');
    el.innerHTML = '';
    function h(tag, cls, text) {
      var x = document.createElement(tag);
      if (cls) x.className = cls;
      if (text !== undefined) x.textContent = text;
      return x;
    }
    var top = h('div', 'top');
    top.appendChild(h('b', '', '研修報告書'));
    top.appendChild(h('span', '', d ? d.getFullYear() + '年　' + (d.getMonth() + 1) + '月　' + d.getDate() + '日' : '　年　月　日'));
    el.appendChild(top);

    var t1 = h('table');
    var tr = h('tr'); tr.appendChild(h('td', 'l', '氏名')); tr.appendChild(h('td', 'c', rep.name || '')); t1.appendChild(tr);
    el.appendChild(t1);

    var t2 = h('table');
    tr = h('tr');
    tr.appendChild(h('td', 'l', '日時'));
    var dt = h('td', 'c', rep.dateText); dt.style.width = '42%'; tr.appendChild(dt);
    tr.appendChild(h('td', 'l', '場所'));
    tr.appendChild(h('td', 'c', rep.place || ''));
    t2.appendChild(tr);
    [['テーマ', rep.theme], ['講師', rep.lecturer], ['目的', lay.purpose.join('\n')]].forEach(function (x) {
      var r = h('tr'); r.appendChild(h('td', 'l', x[0]));
      var v = h('td', '', x[1] || ''); v.colSpan = 3; v.style.whiteSpace = 'pre-wrap';
      r.appendChild(v); t2.appendChild(r);
    });
    el.appendChild(t2);

    [['研修内容', lay.content, 5], ['所　　感', lay.impression, 5], ['備　　考', lay.notes, 3]].forEach(function (s) {
      var sec = h('div', 'sec');
      sec.appendChild(h('h3', '', s[0]));
      var lines = s[1].slice();
      while (lines.length < s[2]) lines.push('');
      lines.forEach(function (line) { sec.appendChild(h('div', '', line)); });
      el.appendChild(sec);
    });
    show('preview');
    // 1行（全角約60文字）が画面幅に収まるよう文字サイズを調整
    var inner = el.clientWidth - parseFloat(getComputedStyle(el).paddingLeft) * 2 - 10;
    el.style.fontSize = Math.max(5, Math.min(12, inner / 60)) + 'px';
  }

  // ---------- バックアップ ----------
  function backup() {
    flush();
    var data = JSON.stringify({ app: 'kenshu-houkoku', version: 1, exported: new Date().toISOString(), records: records }, null, 1);
    var bytes = new TextEncoder().encode(data);
    shareFile(bytes, '研修報告_バックアップ_' + today() + '.json', 'application/json');
  }
  function restore(file) {
    var reader = new FileReader();
    reader.onload = function () {
      try {
        var data = JSON.parse(reader.result);
        var list = Array.isArray(data) ? data : data.records;
        if (!Array.isArray(list)) throw new Error('形式が違います');
        var ids = {}; records.forEach(function (r) { ids[r.id] = r; });
        var added = 0, updated = 0;
        list.forEach(function (r) {
          if (!r || !r.id) return;
          if (!ids[r.id]) { records.push(r); added++; }
          else if ((r.updated || 0) > (ids[r.id].updated || 0)) { Object.assign(ids[r.id], r); updated++; }
        });
        persist(); renderList();
        toast('読み込みました（追加 ' + added + '件・更新 ' + updated + '件）');
      } catch (e) { toast('読み込めませんでした：' + e.message); }
    };
    reader.readAsText(file);
  }

  // ---------- トースト ----------
  var toastTimer;
  function toast(msg) {
    var t = $('toast');
    t.textContent = msg; t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.classList.remove('show'); }, 2800);
  }

  // ---------- イベント ----------
  $('btn-new').addEventListener('click', function () { openRecord(newRecord().id); setTimeout(function () { $('f-theme').focus(); }, 50); });
  $('btn-back').addEventListener('click', function () {
    flush();
    // 何も入力されていない記録は破棄
    if (current && !['theme', 'lecturer', 'purpose', 'content', 'impression', 'notes'].some(function (k) { return (current[k] || '').trim(); })) {
      records = records.filter(function (r) { return r !== current; }); persist();
    }
    current = null; renderList(); show('list');
  });
  document.querySelectorAll('#view-edit [data-k]').forEach(function (el) { el.addEventListener('input', onInput); });
  document.querySelectorAll('.tools').forEach(function (bar) {
    var ta = $(bar.getAttribute('data-for'));
    bar.querySelectorAll('button').forEach(function (b) {
      // フォーカスを奪わない（キーボードが閉じないように）
      b.addEventListener('pointerdown', function (e) { e.preventDefault(); });
      b.addEventListener('mousedown', function (e) { e.preventDefault(); });
      b.addEventListener('click', function () {
        if (document.activeElement !== ta) { ta.focus(); ta.selectionStart = ta.selectionEnd = ta.value.length; }
        insertAtCursor(ta, b.getAttribute('data-ins'));
      });
    });
  });
  $('btn-export').addEventListener('click', exportExcel);
  $('btn-preview-export').addEventListener('click', exportExcel);
  $('btn-preview').addEventListener('click', function () { flush(); renderPreview(); });
  $('btn-preview-back').addEventListener('click', function () { show('edit'); });
  $('btn-dup').addEventListener('click', function () {
    flush();
    var copy = newRecord(current);
    copy.theme = (current.theme || '') + '（コピー）';
    copy.trainDate = today(); copy.reportDate = '';
    persist(); openRecord(copy.id); toast('複製しました');
  });
  $('btn-delete').addEventListener('click', function () {
    if (!confirm('「' + (current.theme || '無題') + '」を削除します。よろしいですか？')) return;
    records = records.filter(function (r) { return r !== current; });
    persist(); current = null; renderList(); show('list'); toast('削除しました');
  });

  $('btn-settings').addEventListener('click', function () {
    $('s-name').value = settings.name || '';
    $('s-place').value = settings.place || '';
    show('settings');
  });
  ['name', 'place'].forEach(function (k) {
    $('s-' + k).addEventListener('input', function (e) { settings[k] = e.target.value; save(SETTINGS_KEY, settings); });
  });
  $('btn-settings-back').addEventListener('click', function () { renderList(); show('list'); });
  $('btn-backup').addEventListener('click', backup);
  $('btn-restore').addEventListener('click', function () { $('file-restore').click(); });
  $('file-restore').addEventListener('change', function (e) {
    if (e.target.files[0]) restore(e.target.files[0]);
    e.target.value = '';
  });
  $('version').textContent = 'バージョン ' + APP_VERSION;

  window.addEventListener('pagehide', flush);
  document.addEventListener('visibilitychange', function () { if (document.visibilityState === 'hidden') flush(); });

  // 初回：氏名が未設定なら設定画面へ案内
  renderList();
  if (!settings.name && records.length === 0) {
    setTimeout(function () { toast('最初に「設定」で氏名を登録してください'); }, 600);
  }

  // オフライン対応
  if ('serviceWorker' in navigator && location.protocol === 'https:') {
    navigator.serviceWorker.register('sw.js').catch(function () {});
  }
})();
