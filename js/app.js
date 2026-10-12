/* ====== Auditorías 5'S — aplicación ====== */
(function () {
  'use strict';

  const SECS = window.SECCIONES;
  const X = window.Xlsx5S;
  const MAX_FOTOS = X.MAX_PER_ITEM || 6;

  /* ---------- íconos ---------- */
  const I = {
    back: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 18l-6-6 6-6"/></svg>',
    more: '<svg viewBox="0 0 24 24" fill="currentColor"><circle cx="12" cy="5" r="2"/><circle cx="12" cy="12" r="2"/><circle cx="12" cy="19" r="2"/></svg>',
    plus: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>',
    check: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>',
    x: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round"><path d="M6 6l12 12M18 6L6 18"/></svg>',
    download: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 4v11M7 10l5 5 5-5M5 20h14"/></svg>',
    share: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="18" cy="5" r="2.5"/><circle cx="6" cy="12" r="2.5"/><circle cx="18" cy="19" r="2.5"/><path d="M8.2 10.8l7.6-4.4M8.2 13.2l7.6 4.4"/></svg>',
    copy: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 12a8 8 0 0 1 14-5.3L20 9M20 4v5h-5M20 12a8 8 0 0 1-14 5.3L4 15M4 20v-5h5"/></svg>',
    trash: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/></svg>',
    camera: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="3.5"/></svg>',
    image: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="2"/><path d="M21 16l-5-5-9 9"/></svg>',
    search: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg>',
    zip: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 3h9l4 4v14H6z"/><path d="M11 3v2M11 7v2M11 11v2M10 15h2v3h-2z"/></svg>',
    upload: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20V9M7 14l5-5 5 5M5 4h14"/></svg>',
    save: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 4h11l3 3v13H5z"/><path d="M8 4v5h7V4M8 20v-6h8v6"/></svg>',
    list: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M9 6h11M9 12h11M9 18h11M4.5 6h.01M4.5 12h.01M4.5 18h.01"/></svg>',
    next: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 18l6-6-6-6"/></svg>',
    edit: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 20h4L19 9l-4-4L4 16z"/></svg>',
    sun: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg>',
    moon: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 14.5A8 8 0 0 1 9.5 4a8 8 0 1 0 10.5 10.5z"/></svg>',
    clip: '<svg viewBox="0 0 120 120" fill="none" stroke="currentColor" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"><rect x="26" y="18" width="68" height="90" rx="10" fill="currentColor" fill-opacity=".08"/><rect x="44" y="10" width="32" height="16" rx="6" fill="currentColor" fill-opacity=".15"/><path d="M40 50l6 6 10-12M40 74l6 6 10-12"/><path d="M66 52h16M66 76h16"/></svg>',
  };

  /* ---------- utilidades ---------- */
  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];
  const h = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  const todayISO = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; };
  const pctTxt = p => Math.round(p * 100) + '%';
  const ringColor = p => p <= 0.5 ? 'var(--bad)' : p <= 0.8 ? 'var(--warn)' : 'var(--ok)';
  const nivel = p => p <= 0.5 ? ['bad', 'Crítico'] : p <= 0.8 ? ['warn', 'En mejora'] : ['ok', 'Satisfactorio'];
  const secOf = n => SECS.find(s => s.items.some(i => i.n === n));
  const ORDER = ['datos', ...SECS.map(s => s.key), 'resumen'];
  const MESES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
  const fechaCorta = iso => { if (!iso) return 'Sin fecha'; const [y, m, d] = iso.split('-'); return `${+d} ${MESES[+m - 1]} ${y}`; };

  function toast(msg, ms = 2400) {
    const t = $('#toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toast._t);
    toast._t = setTimeout(() => t.classList.remove('show'), ms);
  }

  const isQuota = e => e && (e.name === 'QuotaExceededError' || e.code === 22 || /quota|space|espacio/i.test(e.message || ''));
  async function storageInfo(el) {
    if (!el || !navigator.storage || !navigator.storage.estimate) return;
    try {
      const { usage = 0, quota = 0 } = await navigator.storage.estimate();
      if (!quota) return;
      const mb = b => (b / 1048576).toFixed(b < 10485760 ? 1 : 0) + ' MB';
      const pct = Math.min(100, usage / quota * 100);
      el.innerHTML = `Almacenamiento usado en este dispositivo: <b>${mb(usage)}</b> de ${mb(quota)} disponibles<div class="meter"><i style="width:${pct.toFixed(1)}%"></i></div>`;
    } catch (e) { /* sin datos */ }
  }

  /* ---------- tema claro / oscuro ---------- */
  const THEME_KEY = 'a5s-theme';
  const darkMQ = window.matchMedia ? matchMedia('(prefers-color-scheme: dark)') : null;
  function getTheme() { try { return localStorage.getItem(THEME_KEY) || 'auto'; } catch (e) { return 'auto'; } }
  const isDark = () => { const t = getTheme(); return t === 'dark' || (t === 'auto' && !!darkMQ && darkMQ.matches); };
  function applyTheme() {
    const t = getTheme();
    if (t === 'auto') document.documentElement.removeAttribute('data-theme');
    else document.documentElement.setAttribute('data-theme', t);
    const meta = document.querySelector('meta[name="theme-color"]');
    if (meta) meta.setAttribute('content', isDark() ? '#121a24' : '#1f3a5f');
    const btn = document.getElementById('themeBtn');
    if (btn) { btn.innerHTML = isDark() ? I.sun : I.moon; btn.setAttribute('aria-label', isDark() ? 'Cambiar a modo claro' : 'Cambiar a modo oscuro'); }
  }
  function setTheme(t) { try { if (t === 'auto') localStorage.removeItem(THEME_KEY); else localStorage.setItem(THEME_KEY, t); } catch (e) { /* sin almacenamiento */ } applyTheme(); }
  const toggleTheme = () => setTheme(isDark() ? 'light' : 'dark');
  if (darkMQ && darkMQ.addEventListener) darkMQ.addEventListener('change', applyTheme);
  const themeBtnHTML = () => `<button class="icon-btn" id="themeBtn" aria-label="${isDark() ? 'Cambiar a modo claro' : 'Cambiar a modo oscuro'}">${isDark() ? I.sun : I.moon}</button>`;

  /* ---------- hoja inferior (diálogo) ---------- */
  const sheet = $('#sheet');
  function openSheet(html, mount) {
    sheet.innerHTML = `<div class="grab"></div><div class="body">${html}</div>`;
    if (!sheet.open) sheet.showModal();
    if (mount) mount(sheet);
  }
  function closeSheet() { if (sheet.open) sheet.close(); }
  sheet.addEventListener('click', e => { if (e.target === sheet) closeSheet(); }); // tocar fuera cierra
  function confirmSheet({ title, text, ok = 'Aceptar', danger = false }) {
    return new Promise(resolve => {
      openSheet(`<h2>${h(title)}</h2><p class="muted">${text}</p>
        <div class="actions"><button class="btn ${danger ? 'danger' : 'primary'} block" data-ok>${h(ok)}</button>
        <button class="btn ghost block" data-cancel>Cancelar</button></div>`, s => {
        let done = false;
        const finish = v => { if (done) return; done = true; closeSheet(); resolve(v); };
        $('[data-ok]', s).onclick = () => finish(true);
        $('[data-cancel]', s).onclick = () => finish(false);
        s.addEventListener('close', () => finish(false), { once: true });
      });
    });
  }

  /* ---------- descarga / compartir ---------- */
  function saveBlob(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
  }
  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  async function photosForExport(auditId) {
    const ps = (await DB.photosFor(auditId)).sort((a, b) => a.createdAt - b.createdAt);
    const out = [];
    for (const p of ps) out.push({ item: p.item, w: p.w, h: p.h, ext: 'jpeg', bytes: new Uint8Array(await p.blob.arrayBuffer()) });
    return out;
  }
  async function makeXlsxBlob(audit) {
    const bytes = X.buildXlsx(audit, await photosForExport(audit.id));
    return new Blob([bytes], { type: XLSX_MIME });
  }

  function problems(audit) {
    const sc = X.score(audit);
    const issues = [];
    if (!audit.proceso?.trim()) issues.push('Falta el <b>proceso</b>.');
    if (sc.pend) issues.push(`<b>${sc.pend}</b> criterio${sc.pend > 1 ? 's' : ''} sin responder (en el Excel cuentan como no cumplidos).`);
    const sinCm = SECS.flatMap(s => s.items).filter(i => audit.items[i.n]?.estado === 'NC' && !audit.items[i.n]?.contramedida?.trim()).length;
    if (sinCm) issues.push(`<b>${sinCm}</b> no conformidad${sinCm > 1 ? 'es' : ''} sin contramedida.`);
    return issues;
  }

  async function downloadAudit(audit, { share = false, skipCheck = false } = {}) {
    const issues = problems(audit);
    if (!skipCheck && issues.length) {
      const go = await confirmSheet({ title: 'Antes de descargar', text: issues.join('<br>') + '<br><br>¿Descargar de todos modos?', ok: share ? 'Compartir igual' : 'Descargar igual' });
      if (!go) return;
    }
    try {
      const blob = await makeXlsxBlob(audit);
      const name = X.fileName(audit);
      if (share) {
        const file = new File([blob], name, { type: XLSX_MIME });
        if (navigator.canShare && navigator.canShare({ files: [file] })) {
          try { await navigator.share({ files: [file], title: name, text: `Auditoría 5'S — ${audit.proceso || ''}` }); return; }
          catch (e) { if (e.name === 'AbortError') return; }
        }
        toast('Este navegador no permite compartir archivos; se descargará.');
      }
      saveBlob(blob, name);
      toast('Excel descargado');
    } catch (e) {
      console.error(e);
      toast('No se pudo generar el Excel: ' + e.message, 4000);
    }
  }

  async function downloadAll(audits) {
    if (!audits.length) return toast('No hay auditorías para descargar');
    toast('Generando archivos…', 6000);
    const used = new Set();
    const files = [];
    for (const a of audits) {
      let name = X.fileName(a), i = 2;
      while (used.has(name)) name = X.fileName(a).replace('.xlsx', `_${i++}.xlsx`);
      used.add(name);
      files.push({ name, data: X.buildXlsx(a, await photosForExport(a.id)) });
    }
    saveBlob(new Blob([Zip.zipStore(files)], { type: 'application/zip' }), `Auditorias_5S_${todayISO()}.zip`);
    toast(`${files.length} auditorías descargadas (.zip)`);
  }

  /* ---------- respaldo JSON ---------- */
  const blobToB64 = b => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1]); r.onerror = rej; r.readAsDataURL(b); });
  const b64ToBlob = (b64, type) => { const bin = atob(b64); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return new Blob([u], { type }); };

  async function backupExport() {
    const audits = await DB.allAudits();
    const photos = [];
    for (const p of await DB.allPhotos()) {
      const { blob, ...rest } = p;
      photos.push({ ...rest, type: blob.type, data: await blobToB64(blob) });
    }
    const json = JSON.stringify({ app: 'auditorias5s', version: 1, exportedAt: new Date().toISOString(), audits, photos });
    saveBlob(new Blob([json], { type: 'application/json' }), `respaldo_auditorias_5S_${todayISO()}.json`);
    toast('Respaldo descargado');
  }
  async function backupImport(file) {
    try {
      const data = JSON.parse(await file.text());
      if (data.app !== 'auditorias5s' || !Array.isArray(data.audits)) throw new Error('El archivo no es un respaldo de esta aplicación.');
      for (const a of data.audits) await DB.putAudit(a);
      for (const p of data.photos || []) {
        const { data: b64, type, ...rest } = p;
        await DB.putPhoto({ ...rest, blob: b64ToBlob(b64, type || 'image/jpeg') });
      }
      toast(`${data.audits.length} auditorías importadas`);
      render();
    } catch (e) {
      toast('No se pudo importar: ' + e.message, 4000);
    }
  }
  $('#importFile').addEventListener('change', e => { const f = e.target.files[0]; e.target.value = ''; if (f) backupImport(f); });

  /* ---------- modelo ---------- */
  function newAudit(base) {
    return {
      id: uid(),
      proceso: base?.proceso || '',
      fecha: todayISO(),
      auditores: base?.auditores?.length ? [...base.auditores] : [''],
      items: {},
      prevId: base?.id || null,
      createdAt: Date.now(),
      updatedAt: Date.now(),
    };
  }

  function iterationIndex(audit, all) {
    const same = all.filter(a => (a.proceso || '').trim().toLowerCase() === (audit.proceso || '').trim().toLowerCase())
      .sort((a, b) => (a.fecha || '').localeCompare(b.fecha || '') || a.createdAt - b.createdAt);
    return same.findIndex(a => a.id === audit.id) + 1;
  }

  /* ============================================================
     VISTA: LISTA
     ============================================================ */
  const listState = { q: '', proceso: '' };

  async function renderList() {
    const all = (await DB.allAudits()).sort((a, b) => (b.fecha || '').localeCompare(a.fecha || '') || b.updatedAt - a.updatedAt);
    document.title = "Auditorías 5'S";

    $('#appbar').innerHTML = `
      <div class="title brand"><div class="brand-mark">5S</div><div class="grow"><h1>Auditorías 5'S</h1><div class="sub">Registro de auditorías · Planta</div></div></div>
      ${themeBtnHTML()}
      <button class="icon-btn" id="listMenu" aria-label="Más opciones">${I.more}</button>`;
    $('#themeBtn').onclick = toggleTheme;
    $('#listMenu').onclick = () => openSheet(`
      <h2>Opciones</h2>
      <div class="menu">
        <button data-a="zip">${I.zip}<span>Descargar todas en Excel (.zip)</span></button>
        <button data-a="backup">${I.save}<span>Guardar respaldo (.json)</span></button>
        <button data-a="import">${I.upload}<span>Restaurar respaldo</span></button>
      </div>
      <div class="small muted" style="margin-top:16px">Apariencia</div>
      <div class="theme-choice" role="group" aria-label="Apariencia">
        <button data-t="light" aria-pressed="${getTheme() === 'light'}">Claro</button>
        <button data-t="dark" aria-pressed="${getTheme() === 'dark'}">Oscuro</button>
        <button data-t="auto" aria-pressed="${getTheme() === 'auto'}">Automático</button>
      </div>
      <div id="storeInfo" class="small muted" style="margin-top:14px"></div>
      <p class="muted small">Las auditorías se guardan solo en este dispositivo y navegador. Usa el respaldo para pasarlas a otro equipo.</p>`, s => {
      storageInfo($('#storeInfo', s));
      $$('[data-t]', s).forEach(b => b.onclick = () => { setTheme(b.dataset.t); $$('[data-t]', s).forEach(x => x.setAttribute('aria-pressed', x === b)); });
      $('[data-a=zip]', s).onclick = () => { closeSheet(); downloadAll(filtered()); };
      $('[data-a=backup]', s).onclick = () => { closeSheet(); backupExport(); };
      $('[data-a=import]', s).onclick = () => { closeSheet(); $('#importFile').click(); };
    });

    const procesos = [...new Set(all.map(a => (a.proceso || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es'));
    if (listState.proceso && !procesos.includes(listState.proceso)) listState.proceso = '';

    function filtered() {
      const q = listState.q.trim().toLowerCase();
      return all.filter(a => (!listState.proceso || (a.proceso || '').trim() === listState.proceso) &&
        (!q || [a.proceso, X.fmtDMY(a.fecha), ...(a.auditores || [])].join(' ').toLowerCase().includes(q)));
    }

    const view = $('#view');
    view.className = 'view';
    $('#bottom').innerHTML = `<div class="inner"><button class="btn primary" id="bNew">${I.plus} Nueva auditoría</button></div>`;
    $('#bNew').onclick = () => newAuditSheet();

    if (!all.length) {
      view.innerHTML = `
        <div class="card empty">
          <h2>No hay auditorías registradas</h2>
          <p class="muted">Registre una nueva auditoría 5'S. Al finalizar podrá descargarla en el formato Excel oficial.</p>
          <button class="btn primary" id="emptyNew">${I.plus} Nueva auditoría</button>
        </div>`;
      $('#emptyNew').onclick = () => newAuditSheet();
      $('#bottom').innerHTML = '';
      return;
    }

    const scores = all.map(a => X.score(a));
    const done = scores.filter(s => s.pend === 0);
    const avg = done.length ? done.reduce((s, x) => s + x.pct, 0) / done.length : null;
    const ym = todayISO().slice(0, 7);
    const thisMonth = all.filter(a => (a.fecha || '').startsWith(ym)).length;

    view.innerHTML = `
      <div class="kpi-strip">
        <div><b>${all.length}</b><span>Auditorías</span></div>
        <div><b>${avg == null ? '—' : pctTxt(avg)}</b><span>Promedio</span></div>
        <div><b>${thisMonth}</b><span>Este mes</span></div>
      </div>
      <label class="search">${I.search}<input id="q" type="search" placeholder="Buscar por proceso, auditor o fecha" value="${h(listState.q)}" autocomplete="off"></label>
      ${procesos.length > 1 ? `<div class="filter"><select id="fProc" aria-label="Filtrar por proceso">
        <option value="">Todos los procesos</option>
        ${procesos.map(p => `<option value="${h(p)}" ${listState.proceso === p ? 'selected' : ''}>${h(p)}</option>`).join('')}
      </select></div>` : ''}
      <h2 class="section-title">Registro de auditorías</h2>
      <div class="audit-list" id="alist" style="margin-top:0"></div>`;

    $('#q').addEventListener('input', e => { listState.q = e.target.value; drawCards(); });
    const fp = $('#fProc');
    if (fp) fp.onchange = () => { listState.proceso = fp.value; drawCards(); };

    function drawCards() {
      const items = filtered();
      const el = $('#alist');
      if (!items.length) { el.innerHTML = `<p class="muted" style="padding:16px;margin:0">Sin resultados.</p>`; return; }
      el.innerHTML = items.map(a => {
        const sc = X.score(a);
        const it = iterationIndex(a, all);
        const complete = sc.pend === 0;
        return `<article class="audit">
          <button class="open" data-open="${a.id}">
            <div class="score" style="--sc:${ringColor(sc.pct)}"><b>${pctTxt(sc.pct)}</b><span>Cumpl.</span></div>
            <div class="info">
              <h3>${h(a.proceso || 'Sin proceso')}</h3>
              <div class="meta">${X.fmtDMY(a.fecha) || 'Sin fecha'} · ${h((a.auditores || []).filter(Boolean).join(', ') || 'Sin auditor')}</div>
              <div class="tags">
                ${complete ? '<span class="status ok">Completa</span>' : `<span class="status warn">En curso · ${sc.pend} pendientes</span>`}
                ${a.proceso ? `<span>Iteración ${it}</span>` : ''}
                ${sc.nc ? `<span>${sc.nc} NC</span>` : ''}
                ${complete && a.fecha ? `<span>Próxima: ${X.fmtDMY(X.proximaAuditoria(a))}</span>` : ''}
              </div>
            </div>
          </button>
          <button class="icon-btn more" data-menu="${a.id}" aria-label="Opciones de la auditoría">${I.more}</button>
        </article>`;
      }).join('');
      $$('[data-open]', el).forEach(b => b.onclick = () => go(`#/a/${b.dataset.open}`));
      $$('[data-menu]', el).forEach(b => b.onclick = () => auditMenu(all.find(a => a.id === b.dataset.menu)));
    }
    drawCards();

  }

  function auditMenu(a) {
    openSheet(`
      <h2>${h(a.proceso || 'Sin proceso')}</h2>
      <p class="muted small" style="margin:0">${fechaCorta(a.fecha)} · ${pctTxt(X.score(a).pct)} de cumplimiento</p>
      <div class="menu">
        <button data-a="open">${I.edit}<span>Abrir y editar</span></button>
        <button data-a="dl">${I.download}<span>Descargar Excel</span></button>
        <button data-a="share">${I.share}<span>Compartir Excel</span></button>
        <button data-a="iter">${I.copy}<span>Nueva iteración de este proceso</span></button>
        <button data-a="del" class="danger">${I.trash}<span>Eliminar</span></button>
      </div>`, s => {
      $('[data-a=open]', s).onclick = () => { closeSheet(); go(`#/a/${a.id}`); };
      $('[data-a=dl]', s).onclick = () => { closeSheet(); downloadAudit(a); };
      $('[data-a=share]', s).onclick = () => { closeSheet(); downloadAudit(a, { share: true }); };
      $('[data-a=iter]', s).onclick = () => { closeSheet(); newAuditSheet(a); };
      $('[data-a=del]', s).onclick = async () => {
        closeSheet();
        const ok = await confirmSheet({ title: '¿Eliminar auditoría?', text: `Se borrará “${h(a.proceso || 'Sin proceso')}” del ${fechaCorta(a.fecha)} y sus fotos. No se puede deshacer.`, ok: 'Eliminar', danger: true });
        if (ok) { await DB.deleteAudit(a.id); toast('Auditoría eliminada'); render(); }
      };
    });
  }

  async function newAuditSheet(base) {
    const all = await DB.allAudits();
    const procesos = [...new Set(all.map(a => (a.proceso || '').trim()).filter(Boolean))];
    const a = newAudit(base);
    openSheet(`
      <h2>${base ? 'Nueva iteración' : 'Nueva auditoría'}</h2>
      ${base ? `<p class="muted small" style="margin:0 0 6px">Copia el proceso y el equipo auditor. Las respuestas empiezan en blanco y verás las contramedidas anteriores en el criterio 32.</p>` : ''}
      <form id="nf" class="stack" style="margin-top:12px">
        <label class="field"><span>Proceso</span>
          <input class="input" name="proceso" list="procList" required placeholder="Ej: Taller Mantenimiento" value="${h(a.proceso)}" autocomplete="off">
          <datalist id="procList">${procesos.map(p => `<option value="${h(p)}">`).join('')}</datalist>
        </label>
        <label class="field"><span>Fecha de auditoría</span>
          <input class="input" type="date" name="fecha" required value="${a.fecha}">
        </label>
        <div class="field"><span>Equipo auditor</span><div id="audRows"></div>
          <button type="button" class="btn sm ghost" id="addAud" style="margin-top:8px">${I.plus} Agregar auditor</button>
        </div>
        <div class="actions"><button class="btn primary block" type="submit">Empezar auditoría ${I.next}</button></div>
      </form>`, s => {
      const rows = $('#audRows', s);
      const draw = () => {
        rows.innerHTML = a.auditores.map((v, i) => `<div class="auditor-row">
          <input class="input" data-i="${i}" placeholder="Nombre del auditor" value="${h(v)}" autocomplete="name">
          ${a.auditores.length > 1 ? `<button type="button" class="icon-btn" data-rm="${i}" aria-label="Quitar">${I.x}</button>` : ''}</div>`).join('');
        $$('input', rows).forEach(inp => inp.oninput = () => { a.auditores[+inp.dataset.i] = inp.value; });
        $$('[data-rm]', rows).forEach(b => b.onclick = () => { a.auditores.splice(+b.dataset.rm, 1); draw(); });
      };
      draw();
      $('#addAud', s).onclick = () => { a.auditores.push(''); draw(); $$('input', rows).pop().focus(); };
      $('#nf', s).onsubmit = async e => {
        e.preventDefault();
        const fd = new FormData(e.target);
        a.proceso = String(fd.get('proceso')).trim();
        a.fecha = String(fd.get('fecha'));
        a.auditores = a.auditores.map(x => x.trim()).filter(Boolean);
        if (!a.auditores.length) a.auditores = [''];
        await DB.putAudit(a);
        closeSheet();
        go(`#/a/${a.id}/${SECS[0].key}`);
      };
    });
  }

  /* ============================================================
     VISTA: EDITOR
     ============================================================ */
  const ed = { audit: null, photos: [], prev: null, urls: [], all: [] };
  let saveTimer = null;

  function setSaveState(txt) { const el = $('#saveState'); if (el) el.textContent = txt; }
  function saveSoon(delay = 400) {
    setSaveState('Guardando…');
    clearTimeout(saveTimer);
    saveTimer = setTimeout(saveNow, delay);
  }
  async function saveNow() {
    clearTimeout(saveTimer);
    if (!ed.audit) return;
    ed.audit.updatedAt = Date.now();
    await DB.putAudit(ed.audit);
    setSaveState('Guardado');
  }
  window.addEventListener('pagehide', () => { if (saveTimer) saveNow(); });
  document.addEventListener('visibilitychange', () => { if (document.hidden && saveTimer) saveNow(); });

  function revokeUrls() { ed.urls.forEach(u => URL.revokeObjectURL(u)); ed.urls = []; }
  function photoUrl(p) { if (!p._url) { p._url = URL.createObjectURL(p.blob); ed.urls.push(p._url); } return p._url; }

  async function renderEditor(id, secKey) {
    if (!ed.audit || ed.audit.id !== id) {
      if (saveTimer) await saveNow();
      revokeUrls();
      const a = await DB.getAudit(id);
      if (!a) { toast('La auditoría no existe'); return go('#/', true); }
      a.items = a.items || {};
      ed.audit = a;
      ed.photos = (await DB.photosFor(id)).sort((x, y) => x.createdAt - y.createdAt);
      ed.prev = a.prevId ? await DB.getAudit(a.prevId) : null;
      ed.all = await DB.allAudits();
    }
    const a = ed.audit;
    if (!ORDER.includes(secKey)) {
      const firstPending = SECS.find(s => s.items.some(i => !a.items[i.n]?.estado));
      secKey = !a.proceso ? 'datos' : firstPending ? firstPending.key : 'resumen';
    }
    ed.sec = secKey;
    document.title = `${a.proceso || 'Auditoría'} · 5'S`;

    $('#appbar').innerHTML = `
      <button class="icon-btn" id="back" aria-label="Volver a la lista">${I.back}</button>
      <div class="title"><h1>${h(a.proceso || 'Nueva auditoría')}</h1><div class="sub">${fechaCorta(a.fecha)}</div></div>
      <span class="save-state" id="saveState">Guardado</span>
      ${themeBtnHTML()}
      <button class="icon-btn" id="dlTop" aria-label="Descargar Excel">${I.download}</button>`;
    $('#back').onclick = async () => { await saveNow(); go('#/'); };
    $('#themeBtn').onclick = toggleTheme;
    $('#dlTop').onclick = async () => { await saveNow(); downloadAudit(a); };

    const view = $('#view');
    view.className = 'view';
    view.innerHTML = `
      <div class="progress-wrap" id="prog"></div>
      <div id="secBody"></div>`;
    drawProgress();
    if (secKey === 'datos') drawDatos();
    else if (secKey === 'resumen') drawResumen();
    else drawSection(SECS.find(s => s.key === secKey));
    drawBottom();
    window.scrollTo(0, 0);
    const cur = $('.tab[aria-current="true"]');
    if (cur) cur.scrollIntoView({ inline: 'center', block: 'nearest' });
  }

  function drawProgress() {
    const a = ed.audit;
    const sc = X.score(a);
    const answered = sc.c + sc.nc;
    const tabs = [
      { key: 'datos', label: 'Datos generales', color: 'var(--primary)', cnt: '', done: !!a.proceso },
      ...SECS.map(s => {
        const n = s.items.filter(i => a.items[i.n]?.estado).length;
        return { key: s.key, label: s.corto, color: s.color, cnt: `${n}/${s.items.length}`, done: n === s.items.length };
      }),
      { key: 'resumen', label: 'Resumen', color: 'var(--text)', cnt: '', done: sc.pend === 0 },
    ];
    $('#prog').innerHTML = `
      <div class="progress-line"><span><b>${answered}</b> de ${sc.total} respondidos</span><span>Cumplimiento <b>${pctTxt(sc.pct)}</b></span></div>
      <div class="bar" aria-hidden="true"><i class="c" style="width:${sc.c / sc.total * 100}%"></i><i class="nc" style="width:${sc.nc / sc.total * 100}%"></i></div>
      <nav class="tabs" aria-label="Secciones">
        ${tabs.map(t => `<button class="tab ${t.done ? 'done' : ''}" data-k="${t.key}" style="--sc:${t.color}" aria-current="${t.key === ed.sec}"><span class="dot"></span>${h(t.label)}<span class="cnt">${t.cnt}</span></button>`).join('')}
      </nav>`;
    $$('.tab', $('#prog')).forEach(t => t.onclick = () => goSec(t.dataset.k));
  }

  function goSec(key) { saveNow(); go(`#/a/${ed.audit.id}/${key}`, true); }

  function drawBottom() {
    const i = ORDER.indexOf(ed.sec);
    const prev = ORDER[i - 1], next = ORDER[i + 1];
    const label = k => k === 'datos' ? 'Datos' : k === 'resumen' ? 'Resumen' : SECS.find(s => s.key === k).corto;
    let right;
    if (next) right = `<button class="btn primary" id="bNext">${h(label(next))} ${I.next}</button>`;
    else right = `<button class="btn primary" id="bDl">${I.download} Descargar Excel</button>`;
    $('#bottom').innerHTML = `<div class="inner">
      ${prev ? `<button class="btn" id="bPrev" style="flex:0 0 auto">${I.back}<span class="sr">${h(label(prev))}</span></button>` : ''}
      ${right}</div>`;
    if (prev) $('#bPrev').onclick = () => goSec(prev);
    if (next) $('#bNext').onclick = () => goSec(next);
    else $('#bDl').onclick = async () => { await saveNow(); downloadAudit(ed.audit); };
  }

  /* ----- sección: datos generales ----- */
  function drawDatos() {
    const a = ed.audit;
    const procesos = [...new Set(ed.all.map(x => (x.proceso || '').trim()).filter(Boolean))];
    const body = $('#secBody');
    body.innerHTML = `
      <div class="sec-head" style="--sc:var(--primary)"><div class="sec-code">0</div><div class="grow"><h2>Datos generales</h2><div class="small">Encabezado del formato</div></div></div>
      <div class="card pad stack" style="margin-top:12px">
        <label class="field"><span>Proceso</span>
          <input class="input" id="fProceso" list="procList2" placeholder="Ej: Taller Mantenimiento" value="${h(a.proceso)}" autocomplete="off">
          <datalist id="procList2">${procesos.map(p => `<option value="${h(p)}">`).join('')}</datalist></label>
        <label class="field"><span>Fecha de auditoría</span><input class="input" type="date" id="fFecha" value="${h(a.fecha)}"></label>
        <div class="field"><span>Equipo auditor</span><div id="audRows"></div>
          <button type="button" class="btn sm ghost" id="addAud" style="margin-top:8px">${I.plus} Agregar auditor</button></div>
      </div>
      ${ed.prev ? `<div class="card pad" style="margin-top:12px"><div class="small muted">Iteración anterior</div>
        <div class="row" style="margin-top:6px"><div class="grow"><b>${h(ed.prev.proceso)}</b><div class="small muted">${fechaCorta(ed.prev.fecha)} · ${pctTxt(X.score(ed.prev).pct)} cumplimiento</div></div>
        <button class="btn sm" id="openPrev">Ver</button></div></div>` : ''}`;
    $('#fProceso').oninput = e => { a.proceso = e.target.value; $('#appbar h1').textContent = a.proceso || 'Nueva auditoría'; saveSoon(); };
    $('#fFecha').onchange = e => { a.fecha = e.target.value; $('#appbar .sub').textContent = fechaCorta(a.fecha); saveSoon(); };
    if (ed.prev) $('#openPrev').onclick = async () => { await saveNow(); go(`#/a/${ed.prev.id}`); };
    const rows = $('#audRows');
    if (!a.auditores?.length) a.auditores = [''];
    const draw = () => {
      rows.innerHTML = a.auditores.map((v, i) => `<div class="auditor-row">
        <input class="input" data-i="${i}" placeholder="Nombre del auditor" value="${h(v)}" autocomplete="name">
        ${a.auditores.length > 1 ? `<button type="button" class="icon-btn" data-rm="${i}" aria-label="Quitar">${I.x}</button>` : ''}</div>`).join('');
      $$('input', rows).forEach(inp => inp.oninput = () => { a.auditores[+inp.dataset.i] = inp.value; saveSoon(); });
      $$('[data-rm]', rows).forEach(b => b.onclick = () => { a.auditores.splice(+b.dataset.rm, 1); draw(); saveSoon(); });
    };
    draw();
    $('#addAud').onclick = () => { a.auditores.push(''); draw(); $$('input', rows).pop().focus(); };
  }

  /* ----- sección de criterios ----- */
  function drawSection(sec) {
    const a = ed.audit;
    const body = $('#secBody');
    const pend = sec.items.filter(i => !a.items[i.n]?.estado).length;
    body.innerHTML = `
      <div class="sec-head" style="--sc:${sec.color}">
        <div class="sec-code">${SECS.indexOf(sec) || 'S'}</div>
        <div class="grow"><h2>${h(sec.nombre.replace(/^\d+\.\s*/, ''))}</h2><div class="small">${sec.items.length} criterios · ${pend ? pend + ' pendientes' : 'completa'}</div></div>
        ${pend ? `<button class="btn sm" id="allC">Marcar pendientes C</button>` : ''}
      </div>
      <div id="crits">${sec.items.map(it => critHTML(sec, it)).join('')}</div>`;
    if (pend) $('#allC').onclick = () => {
      let n = 0;
      sec.items.forEach(it => { const d = a.items[it.n] = a.items[it.n] || {}; if (!d.estado) { d.estado = 'C'; n++; } });
      saveSoon(0);
      toast(`${n} criterio${n > 1 ? 's' : ''} marcado${n > 1 ? 's' : ''} como Cumple`);
      drawSection(sec); drawProgress();
    };
    sec.items.forEach(it => bindCrit(sec, it));
  }

  function critHTML(sec, it) {
    const d = ed.audit.items[it.n] || {};
    const ps = ed.photos.filter(p => p.item === it.n);
    const isNC = d.estado === 'NC';
    const responsables = [...new Set(ed.all.concat([ed.audit]).flatMap(x => Object.values(x.items || {}).map(v => v.responsable?.trim())).filter(Boolean))];
    const prevHint = it.n === 32 && ed.prev ? prevNCHint() : '';
    return `<article class="card crit ${d.estado === 'C' ? 'is-c' : isNC ? 'is-nc' : ''}" id="crit-${it.n}" style="--sc:${sec.color}">
      <div class="top"><div class="num">${it.n}</div><p class="grow">${h(it.texto)}</p></div>
      ${prevHint}
      <div class="answer" role="group" aria-label="Resultado del criterio ${it.n}">
        <button class="opt c" data-v="C" aria-pressed="${d.estado === 'C'}"><abbr>C</abbr> Cumple</button>
        <button class="opt nc" data-v="NC" aria-pressed="${isNC}"><abbr>NC</abbr> No cumple</button>
      </div>
      ${isNC ? `<div class="details">
        <label class="field"><span>Descripción de la anomalía</span><textarea class="input" data-f="desc" rows="2" placeholder="¿Qué se encontró?">${h(d.desc)}</textarea></label>
        <label class="field"><span>Contramedida</span><textarea class="input" data-f="contramedida" rows="2" placeholder="¿Qué se va a hacer?">${h(d.contramedida)}</textarea></label>
        <div class="grid2" style="margin-top:12px">
          <label class="field"><span>Responsable</span><input class="input" data-f="responsable" list="resp-${it.n}" value="${h(d.responsable)}" placeholder="Nombre" autocomplete="off">
            <datalist id="resp-${it.n}">${responsables.map(r => `<option value="${h(r)}">`).join('')}</datalist></label>
          <label class="field"><span>Plazo</span><input class="input" type="date" data-f="plazo" value="${h(d.plazo)}">
            <div class="quick">${[7, 15, 30].map(n => `<button type="button" data-plus="${n}">+${n} días</button>`).join('')}</div></label>
        </div>
      </div>` : ''}
      <div class="photos ${isNC ? 'in-details' : ''}">
        ${ps.map(p => `<button class="thumb" data-ph="${p.id}" aria-label="Ver foto"><img src="${photoUrl(p)}" alt="Evidencia criterio ${it.n}"></button>`).join('')}
        ${ps.length < MAX_FOTOS ? `<button class="add-photo" data-add>${I.camera} ${ps.length ? 'Agregar foto' : 'Evidencia fotográfica'}</button>` : ''}
        ${ps.length ? `<span class="photo-count">${ps.length}/${MAX_FOTOS}</span>` : ''}
      </div>
    </article>`;
  }

  function prevNCHint() {
    const p = ed.prev;
    const ncs = SECS.flatMap(s => s.items).filter(i => p.items?.[i.n]?.estado === 'NC');
    if (!ncs.length) return `<div class="hint">La auditoría anterior (${fechaCorta(p.fecha)}) no tuvo no conformidades.</div>`;
    return `<div class="hint"><b>Contramedidas de la auditoría anterior (${fechaCorta(p.fecha)}):</b><ul>
      ${ncs.map(i => { const d = p.items[i.n]; return `<li><b>#${i.n}</b> ${h(d.contramedida || d.desc || 'Sin contramedida registrada')}${d.responsable ? ` — <i>${h(d.responsable)}</i>` : ''}${d.plazo ? ` (${X.fmtDMY(d.plazo)})` : ''}</li>`; }).join('')}
    </ul></div>`;
  }

  function bindCrit(sec, it) {
    const el = $(`#crit-${it.n}`);
    const a = ed.audit;
    const d = () => (a.items[it.n] = a.items[it.n] || {});
    $$('.opt', el).forEach(b => b.onclick = () => {
      const v = b.dataset.v;
      d().estado = d().estado === v ? null : v;
      if (navigator.vibrate) navigator.vibrate(10);
      saveSoon(0);
      redrawCrit(sec, it);
      drawProgress();
      const head = $('.sec-head .small');
      const pend = sec.items.filter(i => !a.items[i.n]?.estado).length;
      if (head) head.textContent = `${sec.items.length} criterios · ${pend ? pend + ' pendientes' : 'completa'}`;
      if (!pend) { const all = $('#allC'); if (all) all.remove(); }
    });
    $$('[data-f]', el).forEach(inp => {
      const ev = inp.type === 'date' ? 'change' : 'input';
      inp.addEventListener(ev, () => { d()[inp.dataset.f] = inp.value; saveSoon(); });
    });
    $$('[data-plus]', el).forEach(b => b.onclick = () => {
      const v = X.addDaysISO(a.fecha || todayISO(), +b.dataset.plus);
      d().plazo = v; $('[data-f=plazo]', el).value = v; saveSoon();
    });
    const add = $('[data-add]', el);
    if (add) add.onclick = () => photoSource(it.n, sec, it);
    $$('[data-ph]', el).forEach(b => b.onclick = () => viewPhoto(b.dataset.ph, sec, it));
  }

  function redrawCrit(sec, it) {
    const el = $(`#crit-${it.n}`);
    el.outerHTML = critHTML(sec, it);
    bindCrit(sec, it);
  }

  /* ----- fotos ----- */
  let photoTarget = null;
  function photoSource(n, sec, it) {
    photoTarget = { n, sec, it };
    openSheet(`<h2>Evidencia fotográfica · criterio ${n}</h2>
      <div class="menu">
        <button data-a="cam">${I.camera}<span>Tomar foto</span></button>
        <button data-a="gal">${I.image}<span>Elegir de la galería</span></button>
      </div>
      <p class="muted small">Se incluyen en la hoja EVIDENCIAS del Excel. Máximo ${MAX_FOTOS} fotos por criterio; en la galería puede seleccionar varias a la vez.</p>`, s => {
      $('[data-a=cam]', s).onclick = () => { closeSheet(); $('#photoCam').click(); };
      $('[data-a=gal]', s).onclick = () => { closeSheet(); $('#photoGal').click(); };
    });
  }
  async function onPhotoFiles(e) {
    const files = [...e.target.files];
    e.target.value = '';
    if (!files.length || !photoTarget || !ed.audit) return;
    const { n, sec, it } = photoTarget;
    const free = MAX_FOTOS - ed.photos.filter(p => p.item === n).length;
    if (free <= 0) return toast(`Este criterio ya tiene el máximo de ${MAX_FOTOS} fotos`, 3500);
    let ok = 0, err = null;
    for (const f of files.slice(0, free)) {
      try {
        const { blob, w, h: hh } = await compressImage(f);
        const p = { id: uid(), auditId: ed.audit.id, item: n, blob, w, h: hh, createdAt: Date.now() };
        await DB.putPhoto(p);
        ed.photos.push(p);
        ok++;
      } catch (e2) { console.error(e2); err = e2; if (isQuota(e2)) break; }
    }
    if ($(`#crit-${n}`)) redrawCrit(sec, it);
    if (err && isQuota(err)) {
      openSheet(`<h2>Sin espacio para más fotos</h2>
        <p class="muted">El navegador no tiene espacio disponible para guardar más fotos en este dispositivo.${ok ? ` Se guardaron ${ok} foto${ok > 1 ? 's' : ''}.` : ''}</p>
        <p class="muted small">Para liberar espacio: descargue en Excel las auditorías terminadas y elimínelas de la lista, libere memoria del teléfono, o use el navegador normal (no el modo incógnito).</p>
        <div id="storeInfo2" class="small muted"></div>
        <div class="actions"><button class="btn primary block" data-close>Entendido</button></div>`, s => { storageInfo($('#storeInfo2', s)); $('[data-close]', s).onclick = closeSheet; });
      return;
    }
    if (err) toast(ok ? `Se guardaron ${ok} fotos; una no se pudo leer` : 'No se pudo leer la imagen. Intente con otra foto.', 3500);
    else if (files.length > free) toast(`Se agregaron ${ok} fotos. Máximo ${MAX_FOTOS} por criterio.`, 3500);
    else toast(ok > 1 ? `${ok} fotos guardadas` : 'Foto guardada');
  }
  $('#photoCam').addEventListener('change', onPhotoFiles);
  $('#photoGal').addEventListener('change', onPhotoFiles);

  async function compressImage(file, max = 1600, q = 0.7) {
    const url = URL.createObjectURL(file);
    try {
      const img = new Image();
      img.src = url;
      await img.decode();
      const k = Math.min(1, max / Math.max(img.naturalWidth, img.naturalHeight));
      const w = Math.round(img.naturalWidth * k), hh = Math.round(img.naturalHeight * k);
      const c = document.createElement('canvas');
      c.width = w; c.height = hh;
      const ctx = c.getContext('2d');
      ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, w, hh);
      ctx.drawImage(img, 0, 0, w, hh);
      const blob = await new Promise((res, rej) => c.toBlob(b => b ? res(b) : rej(new Error('toBlob')), 'image/jpeg', q));
      return { blob, w, h: hh };
    } finally { URL.revokeObjectURL(url); }
  }

  function viewPhoto(id, sec, it) {
    const p = ed.photos.find(x => x.id === id);
    if (!p) return;
    openSheet(`<div class="viewer"><img src="${photoUrl(p)}" alt="Evidencia"></div>
      <div class="actions"><button class="btn danger block" data-del>${I.trash} Eliminar foto</button>
      <button class="btn ghost block" data-close>Cerrar</button></div>`, s => {
      $('[data-close]', s).onclick = closeSheet;
      $('[data-del]', s).onclick = async () => {
        await DB.deletePhoto(id);
        ed.photos = ed.photos.filter(x => x.id !== id);
        closeSheet();
        redrawCrit(sec, it);
        toast('Foto eliminada');
      };
    });
  }

  /* ----- resumen ----- */
  function drawResumen() {
    const a = ed.audit;
    const sc = X.score(a);
    const all = SECS.flatMap(s => s.items);
    const pend = all.filter(i => !a.items[i.n]?.estado);
    const ncs = all.filter(i => a.items[i.n]?.estado === 'NC');
    const dias = X.diasProxima(sc.pct);
    const issues = problems(a);
    const body = $('#secBody');
    body.innerHTML = `
      <div class="card pad">
        <div class="result-head">
          <div class="big-score" style="--sc:${ringColor(sc.pct)}"><b>${pctTxt(sc.pct)}</b><span>Calificación</span></div>
          <div class="grow">
            <div style="font-weight:600;font-size:1.05rem">${h(a.proceso || 'Sin proceso')}</div>
            <div class="small muted">${X.fmtDMY(a.fecha)} · ${h((a.auditores || []).filter(Boolean).join(', '))}</div>
            <div class="small" style="margin-top:6px"><span class="status ${nivel(sc.pct)[0]}">${nivel(sc.pct)[1]}</span></div>
          </div>
        </div>
      </div>
      <div class="card" style="margin-top:12px"><table class="table">
        <tr><th>Indicador</th><th style="text-align:right">Valor</th></tr>
        <tr><td>Criterios que cumplen (C)</td><td class="n">${sc.c}</td></tr>
        <tr><td>Criterios que no cumplen (NC)</td><td class="n">${sc.nc}</td></tr>
        <tr><td>Criterios pendientes</td><td class="n">${sc.pend}</td></tr>
        <tr><td>Total de criterios</td><td class="n">${sc.total}</td></tr>
        ${a.fecha ? `<tr><td>Próxima auditoría <span class="muted small">(${dias} días)</span></td><td class="n"><b>${X.fmtDMY(X.proximaAuditoria(a))}</b></td></tr>` : ''}
      </table></div>

      ${issues.length ? `<div class="notice" style="margin-top:12px">${issues.join('<br>')}</div>` : ''}

      <h2 class="section-title">Resultado por sección</h2>
      <div class="card"><table class="table">
        <tr><th>Sección</th><th style="text-align:right">C</th><th style="text-align:right">NC</th><th style="text-align:right">Pend.</th></tr>
        ${SECS.map(s => { const c = s.items.filter(i => a.items[i.n]?.estado === 'C').length; const n = s.items.filter(i => a.items[i.n]?.estado === 'NC').length;
          return `<tr class="link" data-sec="${s.key}"><td><span class="num" style="--sc:${s.color};display:inline-grid;margin-right:8px;min-width:22px;height:20px;font-size:.7rem">${SECS.indexOf(s) || 'S'}</span>${h(s.corto)}</td><td class="n">${c}</td><td class="n">${n}</td><td class="n">${s.items.length - c - n}</td></tr>`; }).join('')}
      </table></div>

      ${pend.length ? `<h2 class="section-title">Pendientes (${pend.length})</h2>
      <div class="card"><ul class="list">${pend.map(i => itemLink(i, '')).join('')}</ul></div>` : ''}

      <h2 class="section-title">No conformidades (${ncs.length})</h2>
      ${ncs.length ? `<div class="card"><ul class="list">${ncs.map(i => {
        const d = a.items[i.n];
        const extra = `<div class="small" style="margin-top:4px">${d.desc ? h(d.desc) + '<br>' : ''}${d.contramedida ? `<b>Contramedida:</b> ${h(d.contramedida)}` : '<span class="warn-text">Sin contramedida</span>'}
          ${d.responsable || d.plazo ? `<br><span class="muted">${h(d.responsable || 'Sin responsable')}${d.plazo ? ' · ' + X.fmtDMY(d.plazo) : ''}</span>` : ''}</div>`;
        return itemLink(i, extra);
      }).join('')}</ul></div>` : `<div class="card pad muted">No se registraron no conformidades.</div>`}

      <div class="stack" style="margin-top:18px">
        <button class="btn primary block" id="sDl">${I.download} Descargar Excel</button>
        <button class="btn block" id="sShare">${I.share} Compartir Excel</button>
        <button class="btn block ghost" id="sIter">Programar nueva iteración del proceso</button>
      </div>`;
    $$('[data-sec]', body).forEach(b => b.onclick = () => goSec(b.dataset.sec));
    $$('[data-go]', body).forEach(b => b.onclick = () => {
      const n = +b.dataset.go;
      goSec(secOf(n).key);
      setTimeout(() => { const el = $(`#crit-${n}`); if (el) { el.scrollIntoView({ block: 'center' }); el.animate([{ outline: '3px solid var(--primary-2)' }, { outline: '3px solid transparent' }], { duration: 1400 }); } }, 60);
    });
    $('#sDl').onclick = async () => { await saveNow(); downloadAudit(a); };
    $('#sShare').onclick = async () => { await saveNow(); downloadAudit(a, { share: true }); };
    $('#sIter').onclick = async () => { await saveNow(); newAuditSheet(a); };
  }
  function itemLink(i, extra) {
    const s = secOf(i.n);
    return `<li><button class="link" data-go="${i.n}"><span class="num" style="--sc:${s.color}">${i.n}</span><span class="grow"><span class="small">${h(i.texto)}</span>${extra}</span></button></li>`;
  }

  /* ============================================================
     RUTAS
     ============================================================ */
  function go(hash, replace = false) {
    if (location.hash === hash) return render();
    if (replace) { history.replaceState(null, '', hash); render(); }
    else location.hash = hash;
  }

  async function render() {
    closeSheet();
    const parts = location.hash.replace(/^#\/?/, '').split('/');
    try {
      if (parts[0] === 'a' && parts[1]) await renderEditor(parts[1], parts[2]);
      else {
        if (saveTimer) await saveNow();
        ed.audit = null; revokeUrls();
        await renderList();
      }
    } catch (e) {
      console.error(e);
      $('#view').innerHTML = `<div class="card pad"><b>Ocurrió un error.</b><p class="muted">${h(e.message)}</p></div>`;
    }
  }
  window.addEventListener('hashchange', render);

  /* ---------- arranque ---------- */
  if (!('indexedDB' in window)) {
    $('#view').innerHTML = '<div class="card pad">Este navegador no permite guardar datos. Prueba con Chrome, Edge o Safari actualizados.</div>';
    return;
  }
  applyTheme();
  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
  if ('serviceWorker' in navigator && location.protocol === 'https:') {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
  render();
})();
