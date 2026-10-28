/* =========================================================================
   PoliMusic - Song.aspx
   Reproductor de la lista de canciones (un reproductor por fila).
   Se comunica con el servidor mediante un callback de ASP.NET para
   incrementar el contador de reproducciones de TBL_SONG.
   ========================================================================= */
(function () {
    'use strict';

    var VOLUME_KEY = 'polimusic.volume';
    var DEFAULT_VOLUME = 0.8;

    // Solo un audio puede estar sonando a la vez.
    var currentAudio = null;
    var toastTimer = null;

    /* ------------------------------ utilidades ------------------------------ */

    function formatTime(seconds) {
        if (!isFinite(seconds) || seconds < 0) {
            seconds = 0;
        }
        var minutes = Math.floor(seconds / 60);
        var rest = Math.floor(seconds % 60);
        return minutes + ':' + (rest < 10 ? '0' + rest : rest);
    }

    function showToast(message) {
        var toast = document.getElementById('toast');
        if (!toast) {
            return;
        }
        toast.innerHTML = message;
        toast.className = 'toast is-visible';
        window.clearTimeout(toastTimer);
        toastTimer = window.setTimeout(function () {
            toast.className = 'toast';
        }, 2600);
    }

    function readVolume() {
        try {
            var stored = window.localStorage.getItem(VOLUME_KEY);
            var value = stored === null ? DEFAULT_VOLUME : parseFloat(stored);
            return isNaN(value) ? DEFAULT_VOLUME : value;
        } catch (e) {
            return DEFAULT_VOLUME;
        }
    }

    function saveVolume(value) {
        try {
            window.localStorage.setItem(VOLUME_KEY, value);
        } catch (e) {
            // localStorage puede estar bloqueado: se ignora.
        }
    }

    function copyText(text) {
        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(text).then(function () {
                showToast('Ruta copiada');
            }, function () {
                showToast('No se pudo copiar la ruta');
            });
            return;
        }
        var area = document.createElement('textarea');
        area.value = text;
        area.style.position = 'fixed';
        area.style.opacity = '0';
        document.body.appendChild(area);
        area.select();
        try {
            document.execCommand('copy');
            showToast('Ruta copiada');
        } catch (e) {
            showToast('No se pudo copiar la ruta');
        }
        document.body.removeChild(area);
    }

    function closeMenus() {
        var opened = document.querySelectorAll('.menu.is-open');
        for (var i = 0; i < opened.length; i++) {
            opened[i].className = 'menu';
        }
    }

    function playNext(root) {
        var tracks = document.querySelectorAll('.track');
        for (var i = 0; i < tracks.length; i++) {
            if (tracks[i] === root) {
                var next = tracks[i + 1];
                if (next) {
                    var button = next.querySelector('.btn-play');
                    if (button) {
                        button.click();
                    }
                }
                return;
            }
        }
    }

    /* --------------------------- reproductor por fila --------------------------- */

    function setupTrack(root) {
        var src = root.getAttribute('data-src');
        var songId = root.getAttribute('data-id');

        var btnPlay = root.querySelector('.btn-play');
        var btnMute = root.querySelector('.btn-mute');
        var btnMore = root.querySelector('.btn-more');
        var menu = root.querySelector('.menu');
        var progress = root.querySelector('.progress');
        var fill = root.querySelector('.progress-fill');
        var timeCurrent = root.querySelector('.time-current');
        var timeTotal = root.querySelector('.time-total');
        var volume = root.querySelector('.volume');
        var playsCount = root.querySelector('.plays-count');

        if (!src) {
            root.classList.add('is-unavailable');
            timeTotal.innerHTML = '--:--';
            return;
        }

        var audio = document.createElement('audio');
        audio.preload = 'metadata';
        audio.src = src;

        var storedVolume = readVolume();
        audio.volume = storedVolume;
        volume.value = Math.round(storedVolume * 100);

        function startPlayback() {
            if (currentAudio && currentAudio !== audio) {
                currentAudio.pause();
            }
            currentAudio = audio;

            var promise = audio.play();
            if (promise && typeof promise.catch === 'function') {
                promise.catch(function () {
                    showToast('No se pudo reproducir: el archivo no esta disponible.');
                });
            }
        }

        /* -------- controles -------- */

        btnPlay.addEventListener('click', function () {
            if (audio.paused) {
                startPlayback();
            } else {
                audio.pause();
            }
        });

        progress.addEventListener('click', function (event) {
            if (!audio.duration) {
                return;
            }
            var box = progress.getBoundingClientRect();
            var ratio = (event.clientX - box.left) / box.width;
            if (ratio < 0) { ratio = 0; }
            if (ratio > 1) { ratio = 1; }
            audio.currentTime = ratio * audio.duration;
        });

        volume.addEventListener('input', function () {
            var value = volume.value / 100;
            audio.volume = value;
            audio.muted = false;
            root.classList.remove('is-muted');
            saveVolume(value);
        });

        btnMute.addEventListener('click', function () {
            audio.muted = !audio.muted;
            if (audio.muted) {
                root.classList.add('is-muted');
            } else {
                root.classList.remove('is-muted');
            }
        });

        btnMore.addEventListener('click', function (event) {
            event.stopPropagation();
            var wasOpen = menu.classList.contains('is-open');
            closeMenus();
            if (!wasOpen) {
                menu.classList.add('is-open');
            }
        });

        menu.addEventListener('click', function (event) {
            var action = event.target.getAttribute('data-action');
            if (action === 'play') {
                startPlayback();
            } else if (action === 'copy') {
                copyText(src);
            }
            closeMenus();
        });

        /* -------- eventos del audio -------- */

        audio.addEventListener('play', function () {
            var others = document.querySelectorAll('.track.is-active');
            for (var i = 0; i < others.length; i++) {
                others[i].classList.remove('is-active');
            }
            root.classList.add('is-playing');
            root.classList.add('is-active');
            notifyPlayCount(songId, root);
        });

        audio.addEventListener('pause', function () {
            root.classList.remove('is-playing');
        });

        audio.addEventListener('ended', function () {
            root.classList.remove('is-playing');
            fill.style.width = '0%';
            timeCurrent.innerHTML = '0:00';
            playNext(root);
        });

        audio.addEventListener('loadedmetadata', function () {
            root.classList.remove('is-unavailable');
            timeTotal.innerHTML = formatTime(audio.duration);
        });

        audio.addEventListener('timeupdate', function () {
            var percent = audio.duration ? (audio.currentTime / audio.duration) * 100 : 0;
            fill.style.width = percent + '%';
            timeCurrent.innerHTML = formatTime(audio.currentTime);
        });

        audio.addEventListener('error', function () {
            root.classList.add('is-unavailable');
            timeTotal.innerHTML = '--:--';
        });

        /* -------- contador de reproducciones -------- */

        function notifyPlayCount(id, row) {
            if (typeof window.notifyPlay !== 'function') {
                return;
            }
            try {
                window.notifyPlay(id, row);
            } catch (e) {
                // Si el callback falla la cancion se sigue reproduciendo.
            }
        }

        // La respuesta del servidor trae el nuevo numero de reproducciones.
        root.updatePlays = function (plays) {
            if (playsCount) {
                playsCount.innerHTML = plays;
            }
        };
    }

    /* --------------------------- respuesta del servidor --------------------------- */

    window.songPlayedCallback = function (result, context) {
        if (!result || !context || !context.updatePlays) {
            return;
        }
        context.updatePlays(result);
    };

    /* --------------------------- arranque --------------------------- */

    function init() {
        var tracks = document.querySelectorAll('.track');
        for (var i = 0; i < tracks.length; i++) {
            setupTrack(tracks[i]);
        }

        var btnPlayAll = document.getElementById('btnPlayAll');
        if (btnPlayAll && tracks.length > 0) {
            btnPlayAll.addEventListener('click', function () {
                var first = tracks[0].querySelector('.btn-play');
                if (first) {
                    first.click();
                }
            });
        }

        document.addEventListener('click', function (event) {
            if (!event.target.closest || !event.target.closest('.menu')) {
                closeMenus();
            }
        });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
