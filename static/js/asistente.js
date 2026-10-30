/* AD17 · Asistente IA (interfaz).
   Habla con /asistente/chat (texto + archivos) y /asistente/accion.

   /asistente/chat
     petición  multipart: mensaje, historial (JSON), archivos[]
     respuesta { ok, respuesta (formato compacto, ver markdownToHtml), acciones: [...], adjuntos: [...],
                 navegar: { url, titulo, resultados } | null, cache? }
   `navegar` (herramienta abrir_en_panel): al terminar la respuesta se abre esa
   vista del panel y, en escritorio, el chat se vuelve a abrir al llegar.
   Las acciones NUNCA se ejecutan solas: el usuario confirma en la tarjeta y se
   envían a /asistente/accion.

   El HTML de la IA pasa por sanitize(): solo etiquetas, clases y atributos de
   una lista permitida; todo lo demás se descarta (scripts, estilos, enlaces). */
(function () {
  'use strict';
  var panel = document.getElementById('aiPanel');
  if (!panel) return;

  var body = document.getElementById('aiBody');
  var form = document.getElementById('aiForm');
  var input = document.getElementById('aiInput');
  var send = document.getElementById('aiSend');
  var statusEl = document.getElementById('aiStatus');
  var welcome = document.getElementById('aiWelcome');
  var suggestWrap = document.getElementById('aiSuggestWrap');
  var fileInput = document.getElementById('aiFiles');
  var tray = document.getElementById('aiTray');
  var drop = document.getElementById('aiDrop');
  var STORE = 'ai-chat-v2';
  var MAX_FILES = 5, MAX_MB = 10;
  var EXT_OK = ['png', 'jpg', 'jpeg', 'webp', 'gif', 'pdf', 'xlsx', 'docx', 'csv', 'txt', 'md', 'json'];
  var history = [];
  var pending = [];   // archivos elegidos para el próximo mensaje
  var busy = false;

  /* ---------- Abrir / cerrar / pantalla completa ---------- */
  function open() {
    document.body.classList.add('ai-open');
    panel.setAttribute('aria-hidden', 'false');
    setTimeout(function () { input.focus(); }, 260);
  }
  function close() {
    document.body.classList.remove('ai-open');
    panel.setAttribute('aria-hidden', 'true');
  }
  function setFull(on) {
    document.body.classList.toggle('ai-full', on);
    var b = document.getElementById('aiExpand');
    b.innerHTML = on ? '<i class="fa-solid fa-compress"></i>' : '<i class="fa-solid fa-expand"></i>';
    b.title = on ? 'Salir de pantalla completa' : 'Pantalla completa';
    try { localStorage.setItem('ai-full', on ? '1' : '0'); } catch (e) {}
  }
  document.getElementById('aiExpand').addEventListener('click', function () {
    setFull(!document.body.classList.contains('ai-full'));
  });
  try { if (localStorage.getItem('ai-full') === '1') setFull(true); } catch (e) {}

  document.querySelectorAll('[data-ai-open]').forEach(function (b) { b.addEventListener('click', open); });
  document.querySelectorAll('[data-ai-close]').forEach(function (b) { b.addEventListener('click', close); });
  document.addEventListener('keydown', function (e) {
    if ((e.ctrlKey || e.metaKey) && (e.key === 'k' || e.key === 'K')) {
      e.preventDefault();
      document.body.classList.contains('ai-open') ? close() : open();
    } else if (e.key === 'Escape' && document.body.classList.contains('ai-open')) {
      if (document.body.classList.contains('ai-full')) setFull(false); else close();
    }
  });

  /* ---------- Estado de conexión ---------- */
  function setStatus(online, text) {
    statusEl.classList.toggle('online', !!online);
    statusEl.querySelector('.ai-status-text').textContent = text;
  }
  fetch('/asistente/estado', { credentials: 'same-origin' })
    .then(function (r) { return r.json(); })
    .then(function (d) { setStatus(d.configured, d.configured ? 'En línea · ' + (d.modelo || 'IA') : 'Sin conexión'); })
    .catch(function () { setStatus(false, 'Sin conexión'); });

  /* ---------- HTML seguro ---------- */
  var TAGS = { p: 1, br: 1, strong: 1, b: 1, em: 1, i: 1, u: 1, ul: 1, ol: 1, li: 1, h3: 1, h4: 1, h5: 1, small: 1,
               table: 1, thead: 1, tbody: 1, tr: 1, th: 1, td: 1, div: 1, span: 1, a: 1, code: 1, hr: 1, blockquote: 1 };
  var DROP = { script: 1, style: 1, iframe: 1, object: 1, embed: 1, link: 1, meta: 1, img: 1, svg: 1, math: 1,
               form: 1, input: 1, button: 1, textarea: 1, select: 1, video: 1, audio: 1, template: 1 };
  var CLASSES = { 'ai-kpis': 1, 'ai-kpi': 1, 'ai-kpi-label': 1, 'ai-kpi-value': 1, 'ai-kpi-sub': 1, 'ai-table': 1,
                  num: 1, 'status-pill': 1, 'ai-callout': 1, 'ai-muted': 1, 'ai-list': 1, 'ai-item': 1,
                  'ai-item-main': 1, 'ai-item-side': 1, 'ai-chip': 1 };
  var TONES = { info: 1, success: 1, warning: 1, danger: 1 };

  function cleanNode(src, dst) {
    Array.prototype.forEach.call(src.childNodes, function (n) {
      if (n.nodeType === 3) { dst.appendChild(document.createTextNode(n.nodeValue)); return; }
      if (n.nodeType !== 1) return;
      var tag = n.nodeName.toLowerCase();
      if (DROP[tag]) return;
      if (!TAGS[tag]) { cleanNode(n, dst); return; }          // etiqueta desconocida: se conserva solo su texto
      var el = document.createElement(tag);
      var cls = (n.getAttribute('class') || '').split(/\s+/).filter(function (c) { return CLASSES[c]; });
      if (cls.length) el.className = cls.join(' ');
      var c = n.getAttribute('data-c');
      if (c && /^[a-z-]{3,30}$/.test(c)) el.setAttribute('data-c', c);
      var tone = n.getAttribute('data-tone');
      if (tone && TONES[tone]) el.setAttribute('data-tone', tone);
      ['colspan', 'rowspan'].forEach(function (a) {
        var v = n.getAttribute(a);
        if (v && /^\d{1,2}$/.test(v)) el.setAttribute(a, v);
      });
      if (tag === 'a') {
        var sid = n.getAttribute('data-sid');
        if (!sid || !/^\d{1,9}$/.test(sid)) { cleanNode(n, dst); return; }   // solo enlaces a solicitudes
        el.href = '/admin_dashboard?sid=' + sid;
        el.className = 'ai-sid';
      }
      cleanNode(n, el);
      if (tag === 'table') {
        var wrap = document.createElement('div');
        wrap.className = 'md-table';
        wrap.appendChild(el);
        dst.appendChild(wrap);
      } else {
        dst.appendChild(el);
      }
    });
  }
  function sanitize(html) {
    var tpl = document.createElement('template');   // contenido inerte: no ejecuta ni carga nada
    tpl.innerHTML = html;
    var frag = document.createDocumentFragment();
    cleanNode(tpl.content, frag);
    return frag;
  }

  /* ---------- Formato compacto de la IA -> HTML con diseño ----------
     La IA escribe Markdown (párrafos, **negritas**, listas, ####, tablas) más:
       ::kpi Etiqueta | valor | detalle              (líneas seguidas = fila de indicadores)
       ::item Principal | detalle | monto [[Estado]] (líneas seguidas = lista de solicitudes)
       ::aviso tono | texto      ::nota texto
       [[Estado]] = etiqueta de estado      [#2620] = enlace a la solicitud
     Cuesta mucho menos que pedirle el HTML; el resultado pasa igual por sanitize(). */
  function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function slug(s) {
    return s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim().replace(/\s+/g, '-');
  }
  function inline(s) {
    return s.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>').replace(/`([^`]+)`/g, '<code>$1</code>')
      .replace(/\[\[([^\[\]]{2,40})\]\]/g, function (_, e) { return '<span class="status-pill" data-c="' + slug(e) + '">' + e + '</span>'; })
      .replace(/\[\s*#\s*(\d{1,9})\s*\]/g, '<a data-sid="$1">#$1</a>');
  }
  function directiva(tipo, partes, crudo) {
    var p = partes.map(function (x) { return inline(x.trim()); });
    if (tipo === 'kpi') {
      return '<div class="ai-kpi"><span class="ai-kpi-label">' + (p[0] || '') + '</span><span class="ai-kpi-value">' +
        (p[1] || '') + '</span>' + (p[2] ? '<span class="ai-kpi-sub">' + p.slice(2).join(' · ') + '</span>' : '') + '</div>';
    }
    if (tipo === 'item') {
      return '<div class="ai-item"><div class="ai-item-main"><strong>' + (p[0] || '') + '</strong>' +
        (p[1] ? '<small class="ai-muted">' + p[1] + '</small>' : '') + '</div><div class="ai-item-side">' + p.slice(2).join(' ') + '</div></div>';
    }
    if (tipo === 'aviso') {
      var tono = (partes[0] || '').trim().toLowerCase();
      return TONES[tono] ? '<div class="ai-callout" data-tone="' + tono + '">' + p.slice(1).join(' | ') + '</div>'
                         : '<div class="ai-callout" data-tone="info">' + inline(crudo) + '</div>';
    }
    return '<p><small class="ai-muted">' + inline(crudo) + '</small></p>';
  }
  function celdas(ln) {
    return ln.trim().replace(/^\|/, '').replace(/\|$/, '').split('|').map(function (c) { return c.trim(); });
  }
  function renderTabla(filas) {
    var cuerpo = filas.filter(function (f) { return !/^\|?\s*:?-{2,}/.test(f.trim()); });
    if (!cuerpo.length) return '';
    var out = '<table class="ai-table"><thead><tr>';
    celdas(cuerpo[0]).forEach(function (c) { out += '<th>' + inline(c) + '</th>'; });
    out += '</tr></thead><tbody>';
    cuerpo.slice(1).forEach(function (f) {
      out += '<tr>' + celdas(f).map(function (c) {
        return '<td' + (/^[-−]?\$|^[-−]?\d[\d,.]*%?$/.test(c) ? ' class="num"' : '') + '>' + inline(c) + '</td>';
      }).join('') + '</tr>';
    });
    return out + '</tbody></table>';
  }
  function markdownToHtml(texto) {
    var lineas = escapeHtml(texto).split(/\r?\n/);
    var html = [], lista = null, tabla = [], grupo = null;
    function cerrarLista() { if (lista) { html.push('</' + lista + '>'); lista = null; } }
    function cerrarTabla() { if (tabla.length) { html.push(renderTabla(tabla)); tabla = []; } }
    function cerrarGrupo() { if (grupo) { html.push('</div>'); grupo = null; } }
    lineas.forEach(function (ln) {
      var d = ln.match(/^\s*::(kpi|item|aviso|nota)\s+(.*)$/i);
      if (d) {
        var tipo = d[1].toLowerCase();
        cerrarLista(); cerrarTabla();
        if (tipo === 'kpi' || tipo === 'item') {
          if (grupo !== tipo) { cerrarGrupo(); html.push('<div class="' + (tipo === 'kpi' ? 'ai-kpis' : 'ai-list') + '">'); grupo = tipo; }
        } else {
          cerrarGrupo();
        }
        html.push(directiva(tipo, d[2].split('|'), d[2]));
        return;
      }
      cerrarGrupo();
      if (/^\s*\|.*\|\s*$/.test(ln)) { cerrarLista(); tabla.push(ln); return; }
      cerrarTabla();
      var ul = ln.match(/^\s*[-*•]\s+(.*)$/), ol = ln.match(/^\s*\d+[.)]\s+(.*)$/), h = ln.match(/^\s*#{1,4}\s+(.*)$/);
      if (ul || ol) {
        var tipo = ul ? 'ul' : 'ol';
        if (lista !== tipo) { cerrarLista(); html.push('<' + tipo + '>'); lista = tipo; }
        html.push('<li>' + inline((ul || ol)[1]) + '</li>');
      } else {
        cerrarLista();
        if (h) html.push('<h4>' + inline(h[1]) + '</h4>');
        else if (ln.trim()) html.push('<p>' + inline(ln) + '</p>');
      }
    });
    cerrarLista(); cerrarTabla(); cerrarGrupo();
    return html.join('');
  }
  function renderBot(texto) {
    var t = String(texto || '').trim().replace(/^```[a-z]*\s*/i, '').replace(/```\s*$/, '');
    var esHtml = /<(p|div|table|ul|ol|strong|span|h4|br)\b/i.test(t);
    return sanitize(esHtml ? t : markdownToHtml(t));
  }

  /* ---------- Render de mensajes ---------- */
  function scrollDown() { body.scrollTop = body.scrollHeight; }
  function hideIntro() {
    if (welcome) welcome.style.display = 'none';
    if (suggestWrap) suggestWrap.style.display = 'none';
  }
  function fileIcon(nombre, tipo) {
    var ext = (String(nombre).split('.').pop() || '').toLowerCase();
    if (tipo === 'imagen') return 'fa-file-image';
    return { pdf: 'fa-file-pdf', xlsx: 'fa-file-excel', csv: 'fa-file-csv', docx: 'fa-file-word',
             png: 'fa-file-image', jpg: 'fa-file-image', jpeg: 'fa-file-image', webp: 'fa-file-image', gif: 'fa-file-image' }[ext] || 'fa-file-lines';
  }
  function chipsAdjuntos(adjuntos, urls) {
    var wrap = document.createElement('div');
    wrap.className = 'ai-msg-files';
    adjuntos.forEach(function (a, i) {
      var chip = document.createElement('span');
      chip.className = 'ai-file-chip';
      if (urls && urls[i]) {
        var im = document.createElement('img'); im.src = urls[i]; im.alt = ''; chip.appendChild(im);
      } else {
        var ic = document.createElement('i'); ic.className = 'fa-solid ' + fileIcon(a.nombre, a.tipo); chip.appendChild(ic);
      }
      var tx = document.createElement('span'); tx.textContent = a.nombre; chip.appendChild(tx);
      wrap.appendChild(chip);
    });
    return wrap;
  }
  function bubble(role, text, opts) {
    opts = opts || {};
    hideIntro();
    var wrap = document.createElement('div');
    wrap.className = 'ai-msg ' + role + (opts.error ? ' error' : '');
    if (role === 'bot') {
      var orb = document.createElement('span');
      orb.className = 'mini-orb';
      orb.innerHTML = '<i class="fa-solid fa-wand-magic-sparkles"></i>';
      wrap.appendChild(orb);
    }
    var b = document.createElement('div');
    b.className = 'bubble';
    if (role === 'bot') {
      b.classList.add('rich');
      b.appendChild(renderBot(text));
      if (opts.nav) b.appendChild(navButton(opts.nav));
    }
    else {
      if (opts.adjuntos && opts.adjuntos.length) b.appendChild(chipsAdjuntos(opts.adjuntos, opts.urls));
      if (text) { var t = document.createElement('div'); t.textContent = text; b.appendChild(t); }
    }
    wrap.appendChild(b);
    body.appendChild(wrap);
    scrollDown();
    return b;
  }
  /* ---------- Ir al panel con un filtro (herramienta abrir_en_panel) ---------- */
  function navSegura(nav) {   // solo vistas del panel de este mismo sitio
    return nav && /^\/admin_dashboard(\?[^#\s]*)?$/.test(nav.url || '') ? nav : null;
  }
  function irAlPanel(nav) {
    try { sessionStorage.setItem('ai-reopen', '1'); } catch (e) {}
    if (document.body.classList.contains('ai-full')) setFull(false);   // que se vea el panel al llegar
    location.href = nav.url;
  }
  function navButton(nav) {
    var a = document.createElement('a');
    a.className = 'ai-nav';
    a.href = nav.url;
    a.innerHTML = '<i class="fa-solid fa-table-list"></i><span class="ai-nav-txt"><strong></strong><small></small></span>' +
      '<i class="fa-solid fa-arrow-right"></i>';
    a.querySelector('strong').textContent = nav.titulo || 'Ver en el panel';
    a.querySelector('small').textContent = nav.resultados
      ? nav.resultados + (nav.resultados === 1 ? ' solicitud' : ' solicitudes') + ' en el panel' : 'Abrir en el panel';
    a.addEventListener('click', function (e) { e.preventDefault(); irAlPanel(nav); });
    return a;
  }

  function typing(conArchivos) {
    hideIntro();
    var wrap = document.createElement('div');
    wrap.className = 'ai-msg bot';
    wrap.innerHTML = '<span class="mini-orb"><i class="fa-solid fa-wand-magic-sparkles"></i></span>' +
      '<div class="bubble"><span class="ai-typing"><span></span><span></span><span></span></span>' +
      (conArchivos ? '<small class="ai-typing-label">Leyendo archivos…</small>' : '') + '</div>';
    body.appendChild(wrap);
    scrollDown();
    return wrap;
  }
  function money(n) { return window.UI ? UI.money(n) : '$' + n; }

  // Tarjeta de confirmación para una acción propuesta por la IA
  function actionCard(accion) {
    var card = document.createElement('div');
    card.className = 'ai-action';
    card.innerHTML =
      '<div class="ai-action-head"><i class="fa-solid fa-pen-to-square"></i> Cambio propuesto</div>' +
      '<div class="ai-action-body"></div>' +
      '<div class="ai-action-foot"><button type="button" class="btn btn-primary btn-sm" data-ok><i class="fa-solid fa-check"></i> Confirmar</button>' +
      '<button type="button" class="btn btn-secondary btn-sm" data-cancel>Cancelar</button></div>';
    var info = card.querySelector('.ai-action-body');
    [['Solicitud', '#' + accion.solicitud_id + (accion.fp ? ' · FP ' + accion.fp : '')], ['Destinatario', accion.destinatario],
     ['Monto', money(accion.monto)], ['Estado', (accion.estado_actual || '—') + '  →  ' + accion.estado_nuevo],
     ['Motivo', accion.motivo || null]].filter(function (r) { return r[1] !== null && r[1] !== ''; }).forEach(function (r) {
      var row = document.createElement('div');
      row.innerHTML = '<span class="text-muted"></span> <strong></strong>';
      row.children[0].textContent = r[0] + ':';
      row.children[1].textContent = r[1] == null ? '—' : r[1];
      info.appendChild(row);
    });
    var foot = card.querySelector('.ai-action-foot');
    card.querySelector('[data-cancel]').addEventListener('click', function () {
      foot.innerHTML = '<span class="chip">Cancelado</span>';
    });
    card.querySelector('[data-ok]').addEventListener('click', function () {
      foot.innerHTML = '<span class="spinner"></span>';
      fetch('/asistente/accion', {
        method: 'POST', credentials: 'same-origin',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ accion: accion })
      })
        .then(function (r) { return r.json().then(function (d) { return { ok: r.ok && d.ok, d: d }; }); })
        .then(function (res) {
          foot.innerHTML = res.ok
            ? '<span class="chip chip-success"><i class="fa-solid fa-check"></i> Aplicado</span>' +
              '<a class="btn btn-ghost btn-sm" href="/admin_dashboard?sid=' + encodeURIComponent(accion.solicitud_id) + '">Ver solicitud</a>' +
              (location.pathname === '/admin_dashboard' ? '<button type="button" class="btn btn-ghost btn-sm" onclick="location.reload()">Actualizar panel</button>' : '')
            : '<span class="chip chip-warning"></span>';
          if (!res.ok) foot.querySelector('.chip').textContent = res.d.respuesta || res.d.error || 'No se pudo aplicar';
          if (res.ok && window.UI) UI.toast(res.d.respuesta || 'Cambio aplicado', 'success');
        })
        .catch(function () { foot.innerHTML = '<span class="chip chip-warning">Error de red</span>'; });
    });
    body.appendChild(card);
    scrollDown();
  }

  /* ---------- Persistencia de la conversación (por pestaña) ---------- */
  // Se recorta de 16 en 16 (no uno por uno): el servidor arma la ventana de
  // contexto por posición y así el inicio no se mueve y OpenAI lo reutiliza del caché.
  function save() {
    while (history.length > 40) history = history.slice(16);
    try { sessionStorage.setItem(STORE, JSON.stringify(history)); } catch (e) {}
  }
  function restore() {
    try { history = JSON.parse(sessionStorage.getItem(STORE) || '[]'); } catch (e) { history = []; }
    history.forEach(function (m) { bubble(m.rol, m.texto, { error: m.error, adjuntos: m.adjuntos, nav: navSegura(m.nav) }); });
  }
  document.getElementById('aiReset').addEventListener('click', function () {
    history = []; save();
    body.querySelectorAll('.ai-msg, .ai-action').forEach(function (n) { n.remove(); });
    if (welcome) welcome.style.display = '';
    if (suggestWrap) suggestWrap.style.display = '';
    clearPending();
    input.focus();
  });

  /* ---------- Archivos pendientes ---------- */
  function humanSize(b) { return b > 1048576 ? (b / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(b / 1024)) + ' KB'; }
  function notify(msg) { if (window.UI) UI.toast(msg, 'warning'); else alert(msg); }
  function addFiles(lista) {
    Array.prototype.forEach.call(lista, function (f) {
      var ext = (f.name.split('.').pop() || '').toLowerCase();
      if (EXT_OK.indexOf(ext) < 0) { notify('No puedo leer "' + f.name + '". Usa imágenes, PDF, Excel, Word, CSV o TXT.'); return; }
      if (f.size > MAX_MB * 1048576) { notify('"' + f.name + '" pesa más de ' + MAX_MB + ' MB.'); return; }
      if (pending.length >= MAX_FILES) { notify('Puedes adjuntar hasta ' + MAX_FILES + ' archivos por mensaje.'); return; }
      pending.push({ file: f, url: /^image\//.test(f.type) ? URL.createObjectURL(f) : null });
    });
    renderTray();
    input.focus();
  }
  function clearPending() {
    pending.forEach(function (p) { if (p.url) URL.revokeObjectURL(p.url); });
    pending = [];
    renderTray();
  }
  function renderTray() {
    tray.innerHTML = '';
    tray.hidden = !pending.length;
    pending.forEach(function (p, i) {
      var chip = document.createElement('div');
      chip.className = 'ai-tray-item';
      chip.innerHTML = (p.url ? '<img alt="">' : '<i class="fa-solid ' + fileIcon(p.file.name) + '"></i>') +
        '<span class="nm"></span><span class="sz"></span>' +
        '<button type="button" aria-label="Quitar"><i class="fa-solid fa-xmark"></i></button>';
      if (p.url) chip.querySelector('img').src = p.url;
      chip.querySelector('.nm').textContent = p.file.name;
      chip.querySelector('.sz').textContent = humanSize(p.file.size);
      chip.querySelector('button').addEventListener('click', function () {
        if (p.url) URL.revokeObjectURL(p.url);
        pending.splice(i, 1);
        renderTray();
      });
      tray.appendChild(chip);
    });
    updateSend();
  }
  document.getElementById('aiAttach').addEventListener('click', function () { fileInput.click(); });
  fileInput.addEventListener('change', function () { addFiles(fileInput.files); fileInput.value = ''; });
  body.querySelectorAll('[data-attach]').forEach(function (b) { b.addEventListener('click', function () { fileInput.click(); }); });
  input.addEventListener('paste', function (e) {
    var files = Array.prototype.filter.call((e.clipboardData || {}).files || [], function (f) { return /^image\//.test(f.type); });
    if (files.length) { e.preventDefault(); addFiles(files); }
  });
  var dragDepth = 0;
  panel.addEventListener('dragenter', function (e) {
    if (!e.dataTransfer || Array.prototype.indexOf.call(e.dataTransfer.types || [], 'Files') < 0) return;
    e.preventDefault(); dragDepth++; drop.hidden = false;
  });
  panel.addEventListener('dragover', function (e) { if (!drop.hidden) e.preventDefault(); });
  panel.addEventListener('dragleave', function () { if (--dragDepth <= 0) { dragDepth = 0; drop.hidden = true; } });
  panel.addEventListener('drop', function (e) {
    if (drop.hidden) return;
    e.preventDefault(); dragDepth = 0; drop.hidden = true;
    addFiles(e.dataTransfer.files);
  });

  /* ---------- Enviar ---------- */
  function updateSend() { send.disabled = busy || (!input.value.trim() && !pending.length); }

  function ask(text) {
    text = (text || '').trim();
    if ((!text && !pending.length) || busy) return;
    busy = true;
    updateSend();
    var archivos = pending.slice();
    pending = [];
    renderTray();

    var adjLocal = archivos.map(function (p) { return { nombre: p.file.name, tipo: /^image\//.test(p.file.type) ? 'imagen' : 'documento' }; });
    bubble('user', text, { adjuntos: adjLocal, urls: archivos.map(function (p) { return p.url; }) });
    var entrada = { rol: 'user', texto: text, adjuntos: [] };
    input.value = '';
    autosize();
    var t = typing(archivos.length > 0);

    var fd = new FormData();
    fd.append('mensaje', text);
    fd.append('historial', JSON.stringify(history));   // el servidor elige la ventana (asistente_ia.ventana)
    archivos.forEach(function (p) { fd.append('archivos', p.file, p.file.name); });

    fetch('/asistente/chat', { method: 'POST', credentials: 'same-origin', body: fd })
      .then(function (r) { return r.json().then(function (d) { return { ok: r.ok && d.ok, d: d }; }); })
      .then(function (res) {
        t.remove();
        var d = res.d || {};
        entrada.adjuntos = (d.adjuntos || []).map(function (a) { return { id: a.id, nombre: a.nombre, tipo: a.tipo }; });
        history.push(entrada);
        var msg = d.respuesta || (res.ok ? '' : 'No pude procesar tu mensaje.');
        var nav = res.ok ? navSegura(d.navegar) : null;
        if (nav) nav = { url: nav.url, titulo: nav.titulo, resultados: nav.resultados };
        bubble('bot', msg, { error: !res.ok, nav: nav });
        history.push(nav ? { rol: 'bot', texto: msg, error: false, nav: nav } : { rol: 'bot', texto: msg, error: !res.ok });
        save();
        if (d.configured === false) setStatus(false, 'Sin conexión');
        (d.acciones || []).forEach(actionCard);
        if (nav) setTimeout(function () { irAlPanel(nav); }, 900);   // deja leer la respuesta un momento
      })
      .catch(function () {
        t.remove();
        bubble('bot', 'No hay conexión con el servidor. Revisa tu internet e inténtalo de nuevo.', { error: true });
      })
      .finally(function () {
        archivos.forEach(function (p) { if (p.url) setTimeout(function () { URL.revokeObjectURL(p.url); }, 60000); });
        busy = false;
        updateSend();
        input.focus();
      });
  }

  function autosize() {
    input.style.height = 'auto';
    input.style.height = Math.min(input.scrollHeight, 160) + 'px';
  }
  input.addEventListener('input', function () { autosize(); updateSend(); });
  input.addEventListener('keydown', function (e) {
    if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) { e.preventDefault(); ask(input.value); }
  });
  form.addEventListener('submit', function (e) { e.preventDefault(); ask(input.value); });
  body.querySelectorAll('.ai-suggestion[data-prompt]').forEach(function (b) {
    b.addEventListener('click', function () { ask(b.getAttribute('data-prompt')); });
  });

  restore();
  // Se llegó aquí desde el chat (abrir_en_panel): se baja al listado filtrado y,
  // en escritorio, el chat se reabre al lado.
  try {
    if (sessionStorage.getItem('ai-reopen') === '1') {
      sessionStorage.removeItem('ai-reopen');
      var lista = document.querySelector('.list-card');
      if (lista) window.scrollTo({ top: lista.getBoundingClientRect().top + window.scrollY - 80 });
      if (window.matchMedia('(min-width: 1024px)').matches) { open(); scrollDown(); }
    }
  } catch (e) {}
})();
