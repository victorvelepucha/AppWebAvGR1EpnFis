/* ==========================================================================
   PoliMusic - Logica de cliente
   Reproductor persistente, cola, favoritos, menus, modales y peticiones AJAX
   ========================================================================== */
(function () {
  "use strict";

  var PM = (window.PM = window.PM || {});
  var root = PM.appRoot || "/";
  var STORE = "pm.player.v1";
  var FAVS = "pm.favs.v1";

  /* ------------------------------------------------------------ utilidades */
  function $(sel, ctx) { return (ctx || document).querySelector(sel); }
  function $$(sel, ctx) { return Array.prototype.slice.call((ctx || document).querySelectorAll(sel)); }
  function fmtTime(s) {
    if (!isFinite(s) || s < 0) s = 0;
    s = Math.floor(s);
    return Math.floor(s / 60) + ":" + ("0" + (s % 60)).slice(-2);
  }
  function esc(t) {
    return String(t == null ? "" : t).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function readJSON(key, def) {
    try { var v = sessionStorage.getItem(key) || localStorage.getItem(key); return v ? JSON.parse(v) : def; }
    catch (e) { return def; }
  }
  function writeJSON(key, val, local) {
    try { (local ? localStorage : sessionStorage).setItem(key, JSON.stringify(val)); } catch (e) { }
  }

  /* --------------------------------------------------------------- estado */
  var audio = document.getElementById("pmAudio");
  var state = {
    queue: [],
    index: -1,
    shuffle: false,
    repeat: 0,          /* 0 = off, 1 = todas, 2 = una */
    volume: 0.8,
    muted: false,
    sourceList: null
  };
  var favs = readJSON(FAVS, []) || [];

  /* ======================================================================
     LECTOR DE DATOS DESDE EL DOM
     Cada fila de cancion lleva sus datos en atributos data-song-*
     ====================================================================== */
  function songFromEl(el) {
    if (!el) return null;
    return {
      id: el.getAttribute("data-song-id"),
      title: el.getAttribute("data-song-title") || "Sin titulo",
      artist: el.getAttribute("data-song-artist") || "Artista desconocido",
      cover: el.getAttribute("data-song-cover") || "",
      src: el.getAttribute("data-song-src") || "",
      album: el.getAttribute("data-song-album") || "",
      href: el.getAttribute("data-song-href") || "#"
    };
  }
  function songsFromContainer(container) {
    return $$("[data-song-id]", container).map(songFromEl).filter(function (s) { return s && s.src; });
  }

  /* ======================================================================
     REPRODUCTOR
     ====================================================================== */
  var ui = {
    bar: document.getElementById("playerbar"),
    cover: document.getElementById("pbCover"),
    coverBox: document.querySelector(".pb-cover"),
    title: document.getElementById("pbTitle"),
    artist: document.getElementById("pbArtist"),
    like: document.getElementById("pbLike"),
    play: document.getElementById("pbPlay"),
    prev: document.getElementById("pbPrev"),
    next: document.getElementById("pbNext"),
    shuffle: document.getElementById("pbShuffle"),
    repeat: document.getElementById("pbRepeat"),
    seek: document.getElementById("pbSeek"),
    cur: document.getElementById("pbCur"),
    dur: document.getElementById("pbDur"),
    vol: document.getElementById("pbVol"),
    mute: document.getElementById("pbMute"),
    queueBtn: document.getElementById("pbQueueBtn"),
    queuePanel: document.getElementById("queuepanel"),
    queueClose: document.getElementById("pbQueueClose"),
    queueList: document.getElementById("queueList")
  };

  function setRangeFill(el) {
    if (!el) return;
    var min = parseFloat(el.min || 0), max = parseFloat(el.max || 100), val = parseFloat(el.value || 0);
    var pct = max > min ? ((val - min) / (max - min)) * 100 : 0;
    el.style.setProperty("--fill", pct + "%");
  }

  function currentSong() { return state.index >= 0 ? state.queue[state.index] : null; }

  function playIndex(i, opts) {
    opts = opts || {};
    if (!state.queue.length) return;
    if (i < 0) i = state.queue.length - 1;
    if (i >= state.queue.length) i = 0;
    state.index = i;
    var s = state.queue[i];
    if (!s) return;

    audio.src = s.src;
    if (opts.seek) { try { audio.currentTime = opts.seek; } catch (e) { } }
    var p = audio.play();
    if (p && p.catch) { p.catch(function () { /* autoplay bloqueado */ }); }

    renderNowPlaying(s);
    renderQueue();
    markCurrentRow(s);
    countPlay(s);
    saveState(true);
  }

  function renderNowPlaying(s) {
    if (ui.cover) ui.cover.src = s.cover || (root + "assets/img/fallback-cover.svg");
    if (ui.coverBox) ui.coverBox.classList.add("has-track");
    if (ui.title) { ui.title.textContent = s.title; ui.title.href = s.href || "#"; }
    if (ui.artist) { ui.artist.textContent = s.artist; }
    if (ui.like) ui.like.classList.toggle("is-on", favs.indexOf(String(s.id)) >= 0);
    if (ui.bar) ui.bar.classList.add("has-track");
    if (navigator.mediaSession) {
      try {
        navigator.mediaSession.metadata = new window.MediaMetadata({
          title: s.title, artist: s.artist,
          artwork: s.cover ? [{ src: s.cover }] : []
        });
      } catch (e) { }
    }
  }

  function markCurrentRow(s) {
    $$(".tl-row.is-current, .q-item.is-current, .mini-item.is-current").forEach(function (r) { r.classList.remove("is-current"); });
    $$('[data-song-id="' + s.id + '"]').forEach(function (r) { r.classList.add("is-current"); });
  }

  function setPlaying(b) {
    if (ui.play) ui.play.classList.toggle("is-playing", b);
    $$(".play-fab").forEach(function (b2) { b2.classList.toggle("is-playing", b); });
    if (ui.bar) ui.bar.classList.toggle("is-playing", b);
  }

  function countPlay(s) {
    if (!s || !s.id) return;
    var url = root + "api/play.asp?id=" + encodeURIComponent(s.id);
    if (window.fetch) { fetch(url, { credentials: "same-origin" }).catch(function () { }); }
    else {
      var img = new Image(); img.src = url;
    }
  }

  /* ------------------------------------------------------------- controles */
  function togglePlay() {
    if (!currentSong()) {
      var first = $$("[data-song-id]")[0];
      if (first) { playFromEl(first); }
      return;
    }
    if (audio.paused) { var p = audio.play(); if (p && p.catch) p.catch(function () { }); }
    else { audio.pause(); }
  }
  function nextIndex(dir) {
    if (!state.queue.length) return -1;
    if (state.shuffle && state.queue.length > 1) {
      var n = state.index;
      while (n === state.index) { n = Math.floor(Math.random() * state.queue.length); }
      return n;
    }
    return (state.index + dir + state.queue.length) % state.queue.length;
  }
  function playNext(manual) {
    if (!state.queue.length) return;
    if (state.repeat === 2 && !manual) { audio.currentTime = 0; audio.play(); return; }
    var i = nextIndex(1);
    if (!manual && i === 0 && state.repeat === 0 && !state.shuffle) {
      audio.pause();
      audio.currentTime = 0;
      return;
    }
    playIndex(i);
  }

  /* ------------------------------------------------------------- estado UI */
  function syncToggles() {
    if (ui.shuffle) ui.shuffle.classList.toggle("is-on", state.shuffle);
    if (ui.shuffle) ui.shuffle.setAttribute("aria-pressed", state.shuffle ? "true" : "false");
    if (ui.repeat) {
      ui.repeat.classList.toggle("is-on", state.repeat > 0);
      ui.repeat.setAttribute("aria-pressed", state.repeat > 0 ? "true" : "false");
      ui.repeat.title = state.repeat === 2 ? "Repetir una cancion" : (state.repeat === 1 ? "Repetir todo" : "Repetir");
    }
    if (ui.vol) { ui.vol.value = Math.round(state.volume * 100); setRangeFill(ui.vol); }
    if (ui.mute) ui.mute.innerHTML = (state.muted || state.volume === 0) ? PM_ICON_MUTE : PM_ICON_VOL;
  }

  /* Iconos inyectados por el servidor (evita duplicar SVG en JS) */
  var PM_ICON_VOL = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none"><path d="M4 9.5h3.2L12 5.4v13.2l-4.8-4.1H4z" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round"/><path d="M15.4 9.2a4 4 0 0 1 0 5.6M18.2 6.6a7.6 7.6 0 0 1 0 10.8" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"/></svg>';
  var PM_ICON_MUTE = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none"><path d="M4 9.5h3.2L12 5.4v13.2l-4.8-4.1H4z" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round"/><path d="m15.6 9.6 4.6 4.8M20.2 9.6l-4.6 4.8" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"/></svg>';

  function renderQueue() {
    if (!ui.queueList) return;
    if (!state.queue.length) { ui.queueList.innerHTML = '<p class="muted small">La cola esta vacia.</p>'; return; }
    var html = "";
    state.queue.forEach(function (s, i) {
      html += '<div class="q-item' + (i === state.index ? " is-current" : "") + '" data-q-index="' + i + '" data-song-id="' + esc(s.id) + '"' +
        ' data-song-title="' + esc(s.title) + '" data-song-artist="' + esc(s.artist) + '"' +
        ' data-song-cover="' + esc(s.cover) + '" data-song-src="' + esc(s.src) + '" data-song-href="' + esc(s.href) + '">' +
        '<img src="' + esc(s.cover) + '" alt="" loading="lazy">' +
        '<div class="q-item-body"><div class="q-item-title">' + esc(s.title) + '</div>' +
        '<div class="q-item-sub">' + esc(s.artist) + '</div></div>' +
        (i === state.index ? '<div class="q-eq"><span></span><span></span><span></span></div>' : '') +
        '</div>';
    });
    ui.queueList.innerHTML = html;
  }

  function saveState(playing) {
    var s = currentSong();
    writeJSON(STORE, {
      queue: state.queue, index: state.index,
      shuffle: state.shuffle, repeat: state.repeat, volume: state.volume,
      time: audio.currentTime || 0, playing: !!playing && !audio.paused,
      trackId: s ? s.id : null
    });
  }
  function restoreState() {
    var st = readJSON(STORE, null);
    if (!st || !st.queue || !st.queue.length) return;
    state.queue = st.queue;
    state.index = typeof st.index === "number" ? st.index : 0;
    state.shuffle = !!st.shuffle;
    state.repeat = st.repeat || 0;
    state.volume = typeof st.volume === "number" ? st.volume : 0.8;
    var s = currentSong();
    if (s) {
      audio.src = s.src;
      audio.volume = state.volume;
      renderNowPlaying(s);
      renderQueue();
      markCurrentRow(s);
      if (st.time > 0) { try { audio.currentTime = st.time; } catch (e) { } }
      if (st.playing) {
        var p = audio.play();
        if (p && p.catch) { p.catch(function () { setPlaying(false); }); }
      }
    }
  }

  /* ------------------------------------------------------- API publica */
  PM.playSong = function (el) { playFromEl(el); };

  function playFromEl(el) {
    var container = el.closest("[data-tracklist]") || document;
    var list = songsFromContainer(container);
    var id = el.getAttribute("data-song-id");
    var idx = 0;
    for (var i = 0; i < list.length; i++) { if (String(list[i].id) === String(id)) { idx = i; break; } }
    state.queue = list;
    state.sourceList = container;
    playIndex(idx);
    setPlaying(true);
  }

  function playList(selector) {
    var container = document.querySelector(selector);
    if (!container) return;
    var list = songsFromContainer(container);
    if (!list.length) return;
    state.queue = list;
    state.sourceList = container;
    state.index = -1;
    playIndex(state.shuffle ? Math.floor(Math.random() * list.length) : 0);
  }

  /* ======================================================================
     EVENTOS DEL REPRODUCTOR
     ====================================================================== */
  if (audio) {
    audio.volume = state.volume;

    audio.addEventListener("play", function () { setPlaying(true); saveState(true); });
    audio.addEventListener("pause", function () { setPlaying(false); saveState(false); });
    audio.addEventListener("ended", function () { playNext(false); });
    audio.addEventListener("loadedmetadata", function () {
      if (ui.dur) ui.dur.textContent = fmtTime(audio.duration);
    });
    audio.addEventListener("timeupdate", function () {
      if (audio.duration) {
        if (ui.seek) { ui.seek.value = Math.round((audio.currentTime / audio.duration) * 100); setRangeFill(ui.seek); }
        if (ui.cur) ui.cur.textContent = fmtTime(audio.currentTime);
      }
    });
    audio.addEventListener("error", function () {
      if (!audio.src) return;
      var s = currentSong();
      if (s) { s.src = ""; }
      var t = ui.title ? ui.title.textContent : "";
      if (window.console) console.warn("PoliMusic: no se pudo cargar el audio de \"" + t + "\"");
    });
  }

  if (ui.play) ui.play.addEventListener("click", togglePlay);
  if (ui.next) ui.next.addEventListener("click", function () { playNext(true); });
  if (ui.prev) ui.prev.addEventListener("click", function () {
    if (audio.currentTime > 3) { audio.currentTime = 0; return; }
    playIndex(nextIndex(-1));
  });
  if (ui.shuffle) ui.shuffle.addEventListener("click", function () {
    state.shuffle = !state.shuffle; syncToggles(); saveState(!audio.paused);
  });
  if (ui.repeat) ui.repeat.addEventListener("click", function () {
    state.repeat = (state.repeat + 1) % 3; syncToggles(); saveState(!audio.paused);
  });
  if (ui.seek) ui.seek.addEventListener("input", function () {
    if (!audio.duration) return;
    audio.currentTime = (ui.seek.value / 100) * audio.duration;
    setRangeFill(ui.seek);
  });
  if (ui.vol) ui.vol.addEventListener("input", function () {
    state.volume = ui.vol.value / 100;
    state.muted = state.volume === 0;
    audio.volume = state.volume;
    audio.muted = false;
    setRangeFill(ui.vol); syncToggles(); saveState(!audio.paused);
  });
  if (ui.mute) ui.mute.addEventListener("click", function () {
    state.muted = !state.muted;
    audio.muted = state.muted;
    syncToggles();
  });
  if (ui.queueBtn) ui.queueBtn.addEventListener("click", function () {
    ui.queuePanel.classList.toggle("is-open");
  });
  if (ui.queueClose) ui.queueClose.addEventListener("click", function () {
    ui.queuePanel.classList.remove("is-open");
  });
  if (ui.queueList) ui.queueList.addEventListener("click", function (e) {
    var item = e.target.closest("[data-q-index]");
    if (item) { playIndex(parseInt(item.getAttribute("data-q-index"), 10)); setPlaying(true); }
  });

  /* Favoritos */
  if (ui.like) ui.like.addEventListener("click", function () {
    var s = currentSong();
    if (!s) return;
    var i = favs.indexOf(String(s.id));
    if (i >= 0) favs.splice(i, 1); else favs.push(String(s.id));
    writeJSON(FAVS, favs, true);
    ui.like.classList.toggle("is-on", favs.indexOf(String(s.id)) >= 0);
  });

  /* ======================================================================
     DELEGACION GLOBAL DE CLICS
     ====================================================================== */
  document.addEventListener("click", function (e) {
    /* Reproducir una cancion concreta */
    var playBtn = e.target.closest("[data-play-song]");
    if (playBtn) { e.preventDefault(); playFromEl(playBtn); setPlaying(true); return; }

    /* Reproducir una lista completa en orden aleatorio */
    var playShuf = e.target.closest("[data-play-shuffle]");
    if (playShuf) {
      e.preventDefault();
      state.shuffle = true;
      syncToggles();
      playList(playShuf.getAttribute("data-play-shuffle"));
      setPlaying(true);
      return;
    }

    /* Reproducir una lista completa */
    var playAll = e.target.closest("[data-play-list]");
    if (playAll) {
      e.preventDefault();
      playList(playAll.getAttribute("data-play-list"));
      setPlaying(true);
      return;
    }

    /* Fila completa de la lista de canciones */
    var row = e.target.closest(".tl-row[data-song-id]");
    if (row && !e.target.closest("a") && !e.target.closest("button")) {
      playFromEl(row); setPlaying(true); return;
    }

    /* Boton de fila de mini lista */
    var mini = e.target.closest(".mini-item[data-song-id]");
    if (mini && !e.target.closest("a") && !e.target.closest("button[data-ajax]")) {
      playFromEl(mini); setPlaying(true); return;
    }

    /* Enlaces "reproducir album / playlist" */
    var pb = e.target.closest("[data-play-container]");
    if (pb) { e.preventDefault(); playList(pb.getAttribute("data-play-container")); setPlaying(true); return; }

    /* Selector de archivos de recursos */
    var pick = e.target.closest("[data-pick]");
    if (pick) {
      var modal = pick.closest(".modal");
      var target = null;
      if (modal) {
        var id = modal.id.replace(/^modalPick_/, "");
        target = document.querySelector('[data-file-input="' + id + '"]');
      }
      if (target) {
        target.value = pick.getAttribute("data-pick");
        target.dispatchEvent(new Event("change", { bubbles: true }));
      }
      if (modal) closeModal(modal);
      return;
    }
  });

  /* ======================================================================
     INTERFAZ GENERAL: menus, sidebar, modales, formularios
     ====================================================================== */

  /* Menu de usuario */
  var um = document.querySelector("[data-usermenu]");
  if (um) {
    var umBtn = um.querySelector("[data-usermenu-toggle]");
    umBtn.addEventListener("click", function (e) {
      e.stopPropagation();
      var open = um.classList.toggle("is-open");
      umBtn.setAttribute("aria-expanded", open ? "true" : "false");
    });
    document.addEventListener("click", function (e) {
      if (!um.contains(e.target)) { um.classList.remove("is-open"); umBtn.setAttribute("aria-expanded", "false"); }
    });
  }

  /* Sidebar movil */
  var sidebar = document.getElementById("sidebar");
  var backdrop = document.querySelector(".sidebar-backdrop");
  function toggleSidebar(open) {
    if (!sidebar) return;
    sidebar.classList.toggle("is-open", open);
    if (backdrop) backdrop.classList.toggle("is-open", open);
  }
  $$("[data-sidebar-open]").forEach(function (b) { b.addEventListener("click", function () { toggleSidebar(true); }); });
  $$("[data-sidebar-close]").forEach(function (b) { b.addEventListener("click", function () { toggleSidebar(false); }); });

  /* Modales genericos */
  function openModal(id) {
    var m = document.getElementById(id);
    if (m) { m.classList.add("is-open"); var f = m.querySelector("[autofocus],input,select,textarea"); if (f) setTimeout(function () { f.focus(); }, 60); }
  }
  function closeModal(m) { if (m) m.classList.remove("is-open"); }
  $$("[data-modal-open]").forEach(function (b) {
    b.addEventListener("click", function (e) { e.preventDefault(); openModal(b.getAttribute("data-modal-open")); });
  });
  $$("[data-modal-close]").forEach(function (b) {
    b.addEventListener("click", function (e) { e.preventDefault(); closeModal(b.closest(".modal")); });
  });
  $$(".modal").forEach(function (m) {
    m.addEventListener("click", function (e) { if (e.target === m) closeModal(m); });
  });
  document.addEventListener("keydown", function (e) {
    if (e.key === "Escape") {
      $$(".modal.is-open").forEach(closeModal);
      if (ui.queuePanel) ui.queuePanel.classList.remove("is-open");
      if (um) um.classList.remove("is-open");
    }
  });

  /* Confirmacion de borrado */
  $$("[data-confirm]").forEach(function (el) {
    el.addEventListener("submit", function (e) {
      if (!window.confirm(el.getAttribute("data-confirm"))) { e.preventDefault(); }
    });
    if (el.tagName === "A") {
      el.addEventListener("click", function (e) {
        if (!window.confirm(el.getAttribute("data-confirm"))) { e.preventDefault(); }
      });
    }
  });

  /* Tarjetas de radio */
  $$(".radio-card input[type=radio]").forEach(function (r) {
    function sync() {
      $$(".radio-card", r.closest(".radio-cards")).forEach(function (c) {
        c.classList.toggle("is-checked", c.querySelector("input").checked);
      });
    }
    r.addEventListener("change", sync); sync();
  });

  /* Relleno de sliders y selects */
  $$("input[type=range]").forEach(setRangeFill);
  $$("[data-autosubmit]").forEach(function (s) {
    s.addEventListener("change", function () { s.form.submit(); });
  });

  /* Atajos de teclado */
  document.addEventListener("keydown", function (e) {
    var tag = (e.target.tagName || "").toLowerCase();
    if (tag === "input" || tag === "textarea" || tag === "select" || e.target.isContentEditable) return;
    if (e.code === "Space") { e.preventDefault(); togglePlay(); }
    else if (e.key === "ArrowRight" && e.shiftKey) { playNext(true); }
    else if (e.key === "ArrowLeft" && e.shiftKey) { playIndex(nextIndex(-1)); }
    else if (e.key === "/") { var g = document.getElementById("globalSearch"); if (g) { e.preventDefault(); g.focus(); } }
  });

  /* ======================================================================
     AÑADIR CANCION A PLAYLIST (modal + AJAX)
     ====================================================================== */
  var addCtx = { songId: null };

  document.addEventListener("click", function (e) {
    var btn = e.target.closest("[data-add-to-playlist]");
    if (!btn) return;
    e.preventDefault();
    addCtx.songId = btn.getAttribute("data-song-id");
    var modal = document.getElementById("modalAddPlaylist");
    if (!modal) return;
    var listBox = modal.querySelector("#addPlaylistList");
    listBox.innerHTML = '<p class="muted small">Cargando playlists...</p>';
    openModal("modalAddPlaylist");
    fetch(root + "api/playlists.asp?forSong=" + encodeURIComponent(addCtx.songId), { credentials: "same-origin" })
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (!data.items || !data.items.length) {
          listBox.innerHTML = '<p class="muted small">No tienes playlists. Crea una primero.</p>';
          return;
        }
        var h = '<div class="mini-list">';
        data.items.forEach(function (p) {
          h += '<button type="button" class="mini-item' + (p.has ? " is-added" : "") + '" data-playlist-id="' + p.id + '"' + (p.has ? " disabled" : "") + '>' +
            '<img src="' + esc(p.photo) + '" alt="">' +
            '<span class="mini-item-body"><span class="mini-item-title">' + esc(p.title) + '</span>' +
            '<span class="mini-item-sub">' + esc(p.songs) + ' canciones &middot; ' + esc(p.type) + '</span></span>' +
            (p.has ? '<span class="badge badge-ok">Ya esta</span>' : '<span class="ico-add">+</span>') +
            '</button>';
        });
        h += "</div>";
        listBox.innerHTML = h;
      })
      .catch(function () { listBox.innerHTML = '<p class="text-err small">No se pudieron cargar las playlists.</p>'; });
  });

  var addList = document.getElementById("addPlaylistList");
  if (addList) {
    addList.addEventListener("click", function (e) {
      var item = e.target.closest("[data-playlist-id]");
      if (!item || item.disabled) return;
      var pid = item.getAttribute("data-playlist-id");
      var body = "playlistId=" + encodeURIComponent(pid) + "&songId=" + encodeURIComponent(addCtx.songId);
      fetch(root + "api/playlist_song.asp", {
        method: "POST", credentials: "same-origin",
        headers: { "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8" },
        body: body
      }).then(function (r) { return r.json(); }).then(function (d) {
        if (d.ok) {
          item.classList.add("is-added"); item.disabled = true;
          var badge = document.createElement("span");
          badge.className = "badge badge-ok"; badge.textContent = "Agregada";
          var old = item.querySelector(".ico-add"); if (old) old.replaceWith(badge);
        } else { window.alert(d.error || "No se pudo agregar la cancion."); }
      }).catch(function () { window.alert("Error de red al agregar la cancion."); });
    });
  }

  /* ======================================================================
     FILTRO DEL MODAL "AGREGAR CANCIONES"
     ====================================================================== */
  var addFilter = document.getElementById("addSongsFilter");
  if (addFilter) {
    addFilter.addEventListener("input", function () {
      var q = addFilter.value.toLowerCase();
      $$("#addSongsList .mini-form").forEach(function (fm) {
        var n = fm.getAttribute("data-name") || "";
        fm.style.display = (q.length === 0 || n.indexOf(q) >= 0) ? "" : "none";
      });
    });
  }

  /* ======================================================================
     INICIALIZACION
     ====================================================================== */
  function init() {
    syncToggles();
    setRangeFill(ui.seek);

    if (PM.autoPlay) {
      /* La pagina pidio reproducir una lista concreta al cargar */
      var c = document.querySelector(PM.autoPlay);
      if (c) {
        var list = songsFromContainer(c);
        if (list.length) {
          state.queue = list;
          state.index = -1;
          playIndex(0);
          return;
        }
      }
    }

    restoreState();

    if (ui.coverBox && !currentSong()) ui.coverBox.classList.remove("has-track");
    /* Resalta la fila que estaba sonando antes de navegar */
    var st = readJSON(STORE, null);
    if (st && st.trackId) {
      $$('[data-song-id="' + st.trackId + '"]').forEach(function (r) { r.classList.add("is-current"); });
    }
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();

  /* Exponer utilidades para las paginas */
  PM.openModal = openModal;
  PM.closeModal = closeModal;
  PM.fmtTime = fmtTime;
})();
