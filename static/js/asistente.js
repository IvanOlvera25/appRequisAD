/* AD17 · Asistente IA (interfaz).
   Habla con /asistente/chat y /asistente/accion. Mientras el servidor no tenga
   configurada la IA (OPENAI_API_KEY), esos endpoints responden "no configurado"
   y aquí se muestra el aviso correspondiente.

   Contrato esperado de /asistente/chat (cuando se conecte la IA):
     petición  { mensaje: str, historial: [{rol: 'user'|'bot', texto: str}] }
     respuesta { ok: true, respuesta: str,
                 acciones?: [{ id, tipo: 'cambiar_estado', solicitud_id, fp,
                               destinatario, monto, estado_actual, estado_nuevo }] }
   Las acciones NUNCA se ejecutan solas: el usuario confirma en la tarjeta y se
   envían a /asistente/accion. */
(function () {
  'use strict';
  var panel = document.getElementById('aiPanel');
  if (!panel) return;

  var body = document.getElementById('aiBody');
  var form = document.getElementById('aiForm');
  var input = document.getElementById('aiInput');
  var send = document.getElementById('aiSend');
  var statusEl = document.getElementById('aiStatus');
  var notice = document.getElementById('aiNotice');
  var welcome = document.getElementById('aiWelcome');
  var suggestWrap = document.getElementById('aiSuggestWrap');
  var STORE = 'ai-chat-v1';
  var history = [];
  var busy = false;

  /* ---------- Abrir / cerrar ---------- */
  function open() {
    document.body.classList.add('ai-open');
    panel.setAttribute('aria-hidden', 'false');
    setTimeout(function () { input.focus(); }, 260);
  }
  function close() {
    document.body.classList.remove('ai-open');
    panel.setAttribute('aria-hidden', 'true');
  }
  document.querySelectorAll('[data-ai-open]').forEach(function (b) { b.addEventListener('click', open); });
  document.querySelectorAll('[data-ai-close]').forEach(function (b) { b.addEventListener('click', close); });
  document.addEventListener('keydown', function (e) {
    if ((e.ctrlKey || e.metaKey) && (e.key === 'k' || e.key === 'K')) {
      e.preventDefault();
      document.body.classList.contains('ai-open') ? close() : open();
    } else if (e.key === 'Escape' && document.body.classList.contains('ai-open')) {
      close();
    }
  });

  /* ---------- Estado de conexión ---------- */
  function setStatus(online, text) {
    statusEl.classList.toggle('online', !!online);
    statusEl.querySelector('.ai-status-text').textContent = text;
    if (notice) notice.hidden = !!online;
  }
  fetch('/asistente/estado', { credentials: 'same-origin' })
    .then(function (r) { return r.json(); })
    .then(function (d) { setStatus(d.configured, d.configured ? 'En línea' : 'Sin conexión a la IA'); })
    .catch(function () { setStatus(false, 'Sin conexión a la IA'); });

  /* ---------- Render ---------- */
  function scrollDown() { body.scrollTop = body.scrollHeight; }
  function hideIntro() {
    if (welcome) welcome.style.display = 'none';
    if (suggestWrap) suggestWrap.style.display = 'none';
  }
  function bubble(role, text, extraClass) {
    hideIntro();
    var wrap = document.createElement('div');
    wrap.className = 'ai-msg ' + role + (extraClass ? ' ' + extraClass : '');
    if (role === 'bot') {
      var orb = document.createElement('span');
      orb.className = 'mini-orb';
      orb.innerHTML = '<i class="fa-solid fa-wand-magic-sparkles"></i>';
      wrap.appendChild(orb);
    }
    var b = document.createElement('div');
    b.className = 'bubble';
    b.textContent = text;
    wrap.appendChild(b);
    body.appendChild(wrap);
    scrollDown();
    return b;
  }
  function typing() {
    hideIntro();
    var wrap = document.createElement('div');
    wrap.className = 'ai-msg bot';
    wrap.innerHTML = '<span class="mini-orb"><i class="fa-solid fa-wand-magic-sparkles"></i></span>' +
      '<div class="bubble"><span class="ai-typing"><span></span><span></span><span></span></span></div>';
    body.appendChild(wrap);
    scrollDown();
    return wrap;
  }
  function money(n) { return window.UI ? UI.money(n) : '$' + n; }

  // Tarjeta de confirmación para una acción propuesta por la IA
  function actionCard(target, accion) {
    var card = document.createElement('div');
    card.className = 'ai-action';
    card.innerHTML =
      '<div class="ai-action-head"><i class="fa-solid fa-pen-to-square"></i> Cambio propuesto</div>' +
      '<div class="ai-action-body"></div>' +
      '<div class="ai-action-foot"><button type="button" class="btn btn-primary btn-sm" data-ok><i class="fa-solid fa-check"></i> Confirmar</button>' +
      '<button type="button" class="btn btn-secondary btn-sm" data-cancel>Cancelar</button></div>';
    var info = card.querySelector('.ai-action-body');
    [['FP', accion.fp], ['Destinatario', accion.destinatario], ['Monto', money(accion.monto)],
     ['Estado', (accion.estado_actual || '—') + '  →  ' + accion.estado_nuevo]].forEach(function (r) {
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
            ? '<span class="chip chip-success"><i class="fa-solid fa-check"></i> Aplicado</span>'
            : '<span class="chip chip-warning"></span>';
          if (!res.ok) foot.querySelector('.chip').textContent = res.d.respuesta || res.d.error || 'No se pudo aplicar';
          if (res.ok && window.UI) UI.toast(res.d.respuesta || 'Cambio aplicado', 'success');
        })
        .catch(function () { foot.innerHTML = '<span class="chip chip-warning">Error de red</span>'; });
    });
    target.parentNode.appendChild(card);
    scrollDown();
  }

  /* ---------- Persistencia de la conversación (por pestaña) ---------- */
  function save() { try { sessionStorage.setItem(STORE, JSON.stringify(history.slice(-30))); } catch (e) {} }
  function restore() {
    try { history = JSON.parse(sessionStorage.getItem(STORE) || '[]'); } catch (e) { history = []; }
    history.forEach(function (m) { bubble(m.rol, m.texto, m.error ? 'error' : ''); });
  }
  document.getElementById('aiReset').addEventListener('click', function () {
    history = []; save();
    body.querySelectorAll('.ai-msg, .ai-action').forEach(function (n) { n.remove(); });
    if (welcome) welcome.style.display = '';
    if (suggestWrap) suggestWrap.style.display = '';
    input.focus();
  });

  /* ---------- Enviar ---------- */
  function ask(text) {
    text = (text || '').trim();
    if (!text || busy) return;
    busy = true;
    send.disabled = true;
    bubble('user', text);
    history.push({ rol: 'user', texto: text });
    save();
    input.value = '';
    autosize();
    var t = typing();

    fetch('/asistente/chat', {
      method: 'POST', credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ mensaje: text, historial: history.slice(-12) })
    })
      .then(function (r) { return r.json().then(function (d) { return { ok: r.ok && d.ok, d: d }; }); })
      .then(function (res) {
        t.remove();
        var d = res.d || {};
        var msg = d.respuesta || (res.ok ? '' : 'No pude procesar tu mensaje.');
        var b = bubble('bot', msg, res.ok ? '' : 'error');
        history.push({ rol: 'bot', texto: msg, error: !res.ok });
        save();
        if (d.configured === false) setStatus(false, 'Sin conexión a la IA');
        (d.acciones || []).forEach(function (a) { actionCard(b, a); });
      })
      .catch(function () {
        t.remove();
        bubble('bot', 'No hay conexión con el servidor. Revisa tu internet e inténtalo de nuevo.', 'error');
      })
      .finally(function () {
        busy = false;
        send.disabled = !input.value.trim();
        input.focus();
      });
  }

  function autosize() {
    input.style.height = 'auto';
    input.style.height = Math.min(input.scrollHeight, 140) + 'px';
  }
  input.addEventListener('input', function () { autosize(); send.disabled = busy || !input.value.trim(); });
  input.addEventListener('keydown', function (e) {
    if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) { e.preventDefault(); ask(input.value); }
  });
  form.addEventListener('submit', function (e) { e.preventDefault(); ask(input.value); });
  body.querySelectorAll('.ai-suggestion').forEach(function (b) {
    b.addEventListener('click', function () { ask(b.getAttribute('data-prompt')); });
  });

  restore();
})();
