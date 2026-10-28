/* Edición de las fincas desde su ficha: botón «Editar», formulario, cambios guardados en este móvil,
   y exportar / importar para llevarlos de un dispositivo a otro. No modifica los datos originales del Excel. */
(function () {
  'use strict';
  window.VT = window.VT || {};
  const VT = window.VT;
  const $ = id => document.getElementById(id);
  const KEY = 'fincas-edits-v1';
  const norm = s => String(s == null ? '' : s).trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const toNum = v => { const n = parseFloat(String(v).replace(',', '.').replace(/[^\d.\-]/g, '')); return isFinite(n) ? n : null; };
  const SNAP = ['estado', 'notas', 'tel', 'precio', 'aforoDet', 'lat', 'lon', 'located', 'approx'];

  let edits = {};
  try { edits = JSON.parse(localStorage.getItem(KEY) || '{}') || {}; } catch (e) { edits = {}; }
  const persist = () => { try { localStorage.setItem(KEY, JSON.stringify(edits)); } catch (e) { /* sin almacenamiento */ } updateCount(); };
  const keyOf = p => norm(p.name) + '|' + norm(p.addr);
  const fixup = (p, e) => {
    if (('lat' in e || 'lon' in e) && isFinite(p.lat) && isFinite(p.lon) && p.lat != null && p.lon != null) { p.located = true; p.approx = false; }
  };

  // La aplicación llama a esto con cada finca al cargar los datos
  function apply(p) {
    if (!p._key) { p._key = keyOf(p); p._orig = {}; SNAP.forEach(k => { p._orig[k] = p[k]; }); }
    const e = edits[p._key];
    if (e) { Object.assign(p, e); fixup(p, e); }
    p._edited = !!(e && Object.keys(e).length);
  }
  const count = () => Object.keys(edits).filter(k => Object.keys(edits[k]).length).length;
  VT.edits = { apply: apply, count: count, data: () => edits };

  /* ---------- estilos ---------- */
  const st = document.createElement('style');
  st.textContent = '.actions .editbtn{border:1px solid var(--accent);color:var(--accent);font-weight:600;font-size:13px;padding:7px 14px}' +
    '.vtf label{display:block;margin:10px 0 3px;font-size:13px;color:var(--muted)}' +
    '.vtf input[type=text],.vtf input[type=url],.vtf input[type=tel],.vtf textarea{width:100%;padding:9px 11px;border:1px solid var(--line);border-radius:10px;background:transparent;color:var(--ink);font:inherit;font-size:16px}' +
    '.vtf textarea{min-height:64px;resize:vertical}.vtf .chk{display:flex;align-items:center;gap:10px;font-size:15px;color:var(--ink);margin:8px 0}' +
    '.vtf .chk input{width:20px;height:20px}.vtf h3{margin:18px 0 0;font-size:14px}.vtf .two{display:grid;grid-template-columns:1fr 1fr;gap:10px}' +
    '.vtf .btns{display:flex;gap:8px;flex-wrap:wrap;margin-top:16px;padding:10px 0 2px;position:sticky;bottom:0;background:var(--panel)}' +
    '.vtf .btns .main{background:var(--accent);color:#fff;border-color:var(--accent);font-weight:600}.vtf .hint{font-size:12.5px;color:var(--muted);margin:6px 0 0}' +
    '#vtCross{position:fixed;left:50%;top:50%;width:30px;height:30px;margin:-15px 0 0 -15px;z-index:8;pointer-events:none;display:none}' +
    '#vtCross:before,#vtCross:after{content:"";position:absolute;background:#d93025;box-shadow:0 0 0 1.5px #fff}#vtCross:before{left:14px;top:0;width:2px;height:30px}#vtCross:after{top:14px;left:0;height:2px;width:30px}' +
    '#vtMoveBar{position:fixed;left:10px;right:10px;bottom:calc(14px + env(safe-area-inset-bottom,0px));z-index:8;padding:10px 12px;display:none;gap:8px;align-items:center;flex-wrap:wrap;background:var(--panel);border:1px solid var(--line);border-radius:14px}' +
    '#vtMoveBar span{flex:1;min-width:150px}';
  document.head.appendChild(st);

  /* ---------- formulario ---------- */
  // mismos 4 estados y mismo orden que ESTADO_LABEL en app.js (duplicado a propósito: archivos independientes)
  const ESTADO_OPTIONS = [['pendiente', 'Pendiente'], ['esperando', 'Esperando confirmación'], ['descartada', 'Descartada'], ['visitada', 'Visitada']];
  const SECTIONS = [
    ['Vuestra valoración', [['sel', 'estado', 'Estado', ESTADO_OPTIONS], ['area', 'notas', 'Notas propias']]],
    ['Datos de contacto', [['tel', 'tel', 'Teléfono'], ['text', 'precio', 'Precio orientativo'], ['text', 'aforoDet', 'Aforo (detalle)']]]
  ];

  function field(f, v) {
    const t = f[0], k = f[1], l = f[2], id = 'vtf_' + k;
    if (t === 'check') return '<label class="chk"><input type="checkbox" data-f="' + k + '"> ' + esc(l) + '</label>';
    let ctl;
    if (t === 'area') ctl = '<textarea id="' + id + '" data-f="' + k + '"></textarea>';
    else if (t === 'sel') ctl = '<select id="' + id + '" data-f="' + k + '">' + f[3].map(o => '<option value="' + esc(o[0]) + '">' + esc(o[1]) + '</option>').join('') + '</select>';
    else ctl = '<input id="' + id + '" data-f="' + k + '" type="' + (t === 'tel' ? 'tel' : t === 'url' ? 'url' : 'text') + '">';
    return '<label for="' + id + '">' + esc(l) + '</label>' + ctl;
  }

  function openEditor(idx, draft) {
    const A = VT.app; if (!A) return;
    const p = A.places()[idx]; if (!p) return;
    const card = $('card');
    let html = '<div class="grab"></div><button class="close" type="button" data-close aria-label="Cerrar">✕</button><h2>Editar</h2><p class="zone">' + esc(p.name) + '</p><div class="vtf">';
    SECTIONS.forEach(s => { html += '<h3>' + s[0] + '</h3>' + s[1].map(f => field(f, p[f[1]])).join(''); });
    html += '<h3>Posición en el mapa</h3><div class="two"><div><label for="vtf_lat">Latitud</label><input id="vtf_lat" data-f="lat" type="text" inputmode="decimal"></div>' +
      '<div><label for="vtf_lon">Longitud</label><input id="vtf_lon" data-f="lon" type="text" inputmode="decimal"></div></div>' +
      '<div class="btns"><button type="button" id="vtMove">Mover la chincheta en el mapa</button></div>' +
      '<p class="hint">' + (p.approx ? 'Esta posición es aproximada: muévela hasta el sitio exacto.' : 'Mueve el mapa hasta el sitio exacto y fija la posición.') + '</p>' +
      '<div class="btns"><button type="button" id="vtSave" class="main">Guardar</button><button type="button" id="vtCancel">Cancelar</button>' +
      (p._edited ? '<button type="button" id="vtReset">Volver al original</button>' : '') + '</div>' +
      '<p class="hint">Los cambios se guardan en este móvil. En Opciones puedes exportarlos para llevarlos a otro dispositivo.</p></div>';
    card.innerHTML = html;
    const set = (k, v) => { const el = card.querySelector('[data-f="' + k + '"]'); if (!el) return; if (el.type === 'checkbox') el.checked = !!v; else el.value = v == null ? '' : v; };
    const val = k => (draft && k in draft) ? draft[k] : p[k];
    SECTIONS.forEach(s => s[1].forEach(f => set(f[1], val(f[1]))));
    set('lat', val('lat')); set('lon', val('lon'));
    $('vtSave').addEventListener('click', () => save(p));
    $('vtCancel').addEventListener('click', () => A.select(p, false));
    $('vtMove').addEventListener('click', () => startMove(p, idx));
    const r = $('vtReset'); if (r) r.addEventListener('click', () => restore(p));
    card.scrollTop = 0;
  }

  function read(p) {
    const card = $('card'), v = {};
    SECTIONS.forEach(s => s[1].forEach(f => {
      const t = f[0], k = f[1], el = card.querySelector('[data-f="' + k + '"]'); if (!el) return;
      if (t === 'check') v[k] = el.checked;
      else v[k] = el.value.trim();
    }));
    ['lat', 'lon'].forEach(k => { const el = card.querySelector('[data-f="' + k + '"]'); if (el) v[k] = el.value.trim() === '' ? null : toNum(el.value); });
    return v;
  }
  function draftOf(p) { return read(p); }

  const same = (a, b) => (typeof a === 'boolean' || typeof b === 'boolean') ? !!a === !!b : String(a == null ? '' : a) === String(b == null ? '' : b);
  function save(p) {
    const A = VT.app, v = read(p), e = edits[p._key] ? Object.assign({}, edits[p._key]) : {};
    if ((v.lat == null) !== (v.lon == null)) { alert('Escribe la latitud y la longitud, o deja las dos vacías.'); return; }
    if (v.lat != null && (v.lat < 39.5 || v.lat > 41.2 || v.lon < -4.6 || v.lon > -3)) { alert('Esa posición parece quedar fuera de la zona del mapa. Revisa la latitud y la longitud.'); return; }
    const touched = [];
    Object.keys(v).forEach(k => {
      if (same(p[k], v[k])) return;
      touched.push(k); p[k] = v[k];
      if (same(p._orig[k], v[k])) delete e[k]; else e[k] = v[k];
    });
    if (touched.indexOf('lat') >= 0 || touched.indexOf('lon') >= 0) { p.located = p.lat != null && p.lon != null; p.approx = false; }
    fixup(p, e);
    if (Object.keys(e).length) edits[p._key] = e; else delete edits[p._key];
    p._edited = !!edits[p._key];
    persist(); A.refresh(); A.select(p, false);
  }
  function restore(p) {
    const A = VT.app, e = edits[p._key] || {};
    Object.keys(e).forEach(k => { p[k] = p._orig[k]; });
    if ('lat' in e || 'lon' in e) { p.located = !!p._orig.located; p.approx = !!p._orig.approx; if (!(p.lat != null && p.lon != null)) p.located = false; }
    delete edits[p._key]; p._edited = false; persist(); A.refresh(); A.select(p, false);
    if (!p.located && A.geocode) A.geocode();
  }

  /* ---------- mover la chincheta con una mira en el centro del mapa ---------- */
  let cross = null, bar = null;
  function startMove(p, idx) {
    const A = VT.app, card = $('card'), draft = draftOf(p);
    if (!cross) {
      cross = document.createElement('div'); cross.id = 'vtCross'; document.body.appendChild(cross);
      bar = document.createElement('div'); bar.id = 'vtMoveBar'; bar.className = 'panel'; bar.style.position = 'fixed';
      bar.innerHTML = '<span>Mueve el mapa hasta que la mira quede en el sitio exacto.</span><button type="button" id="vtFix" class="main" style="background:var(--accent);color:#fff;border-color:var(--accent)">Fijar aquí</button><button type="button" id="vtMoveCancel">Cancelar</button>';
      document.body.appendChild(bar);
    }
    const end = () => { cross.style.display = 'none'; bar.style.display = 'none'; card.style.visibility = ''; };
    $('vtFix').onclick = () => { const c = A.map.getCenter(); end(); draft.lat = Math.round(c.lat * 1e6) / 1e6; draft.lon = Math.round(c.lng * 1e6) / 1e6; A.select(p, false); openEditor(idx, draft); };
    $('vtMoveCancel').onclick = () => { end(); A.select(p, false); openEditor(idx, draft); };
    card.style.visibility = 'hidden'; cross.style.display = 'block'; bar.style.display = 'flex';
    if (p.lat != null && p.lon != null) A.map.jumpTo({ center: [p.lon, p.lat], zoom: Math.max(A.map.getZoom(), 15.5) });
  }

  /* ---------- exportar e importar ---------- */
  function download(name, text, type) {
    const a = document.createElement('a'), url = URL.createObjectURL(new Blob([text], { type: type || 'application/json' }));
    a.href = url; a.download = name; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1500);
  }

  /* ---------- exportar a Excel (datos originales + tus ediciones) ---------- */
  const ESTADO_LABEL = {};
  ESTADO_OPTIONS.forEach(o => { ESTADO_LABEL[o[0]] = o[1]; });

  function placeToRow(p, COL) {
    const row = {};
    Object.keys(COL).forEach(k => { if (k !== 'categoria') row[COL[k]] = p[k] == null ? '' : p[k]; });
    row['Estado (en la app)'] = ESTADO_LABEL[p.estado] || p.estado || '';
    return row;
  }

  async function exportExcel() {
    const A = VT.app; if (!A) return;
    try { await A.ensureXLSX(); } catch (e) { alert('No se pudo cargar el lector de Excel.'); return; }
    const header = Object.keys(A.COL).filter(k => k !== 'categoria').map(k => A.COL[k]).concat(['Estado (en la app)']);
    const bySheet = {};
    A.places().forEach(p => {
      const sheet = Object.keys(A.SHEETS).find(s => A.SHEETS[s] === p.categoria) || Object.keys(A.SHEETS)[0];
      (bySheet[sheet] = bySheet[sheet] || []).push(placeToRow(p, A.COL));
    });
    const wb = XLSX.utils.book_new();
    Object.keys(A.SHEETS).forEach(sheet => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(bySheet[sheet] || [], { header: header }), sheet);
    });
    XLSX.writeFile(wb, 'fincas-export.xlsx');
  }
  let countEl = null;
  function updateCount() { if (countEl) countEl.textContent = count() ? count() + ' fincas con cambios en este móvil' : 'Sin cambios todavía'; }

  const host = $('hudbody');
  if (host) {
    const wrap = document.createElement('div');
    wrap.innerHTML = '<div class="row" style="display:block"><b>Mis cambios</b><div id="vtEditCount" style="color:var(--muted);font-size:12.5px"></div></div>' +
      '<div class="row"><button type="button" id="vtExp">Exportar cambios</button><button type="button" id="vtImp">Importar cambios</button><input type="file" id="vtImpFile" accept=".json,application/json" hidden></div>' +
      '<div class="row"><button type="button" id="vtExpXlsx">Exportar a Excel</button></div>' +
      '<p class="hint" style="font-size:12.5px;color:var(--muted);margin:2px 0 0">Descarga un Excel con los datos originales más tus ediciones (estado, notas, posición…), con las mismas hojas y columnas que el Excel de origen.</p>';
    host.appendChild(wrap);
    countEl = $('vtEditCount'); updateCount();
    $('vtExp').addEventListener('click', () => download('fincas-cambios.json', JSON.stringify({ version: 1, cambios: edits }, null, 1)));
    $('vtExpXlsx').addEventListener('click', () => exportExcel());
    $('vtImp').addEventListener('click', () => $('vtImpFile').click());
    $('vtImpFile').addEventListener('change', async ev => {
      const f = ev.target.files && ev.target.files[0]; if (!f) return;
      try {
        const j = JSON.parse(await f.text()), c = j && j.cambios;
        if (!c || typeof c !== 'object') throw new Error('formato');
        Object.keys(c).forEach(k => { edits[k] = Object.assign(edits[k] || {}, c[k]); });
        persist();
        const A = VT.app; if (A) { A.places().forEach(p => { if (p._orig) { SNAP.forEach(s => { p[s] = p._orig[s]; }); apply(p); } }); A.refresh(); if (A.geocode) A.geocode(); }
        alert('Cambios importados.');
      } catch (e) { alert('No se pudo leer ese archivo de cambios.'); }
      ev.target.value = '';
    });
  }

  document.addEventListener('click', e => { const b = e.target.closest && e.target.closest('[data-edit]'); if (b) openEditor(+b.getAttribute('data-edit')); });
})();
