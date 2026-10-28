<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV2.Song" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>PoliMusic | Song list</title>
    <link href="Styles/Song.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
        <div class="app">

            <!-- ============================ SIDEBAR ============================ -->
            <aside class="sidebar">
                <div class="brand">
                    <span class="brand-mark">
                        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v10.6a4 4 0 1 0 2 3.4V7h4V3h-6z" /></svg>
                    </span>
                    <span class="brand-name">PoliMusic</span>
                </div>

                <nav class="nav">
                    <a class="nav-item is-active" href="Song.aspx">
                        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16v2H4V6zm0 5h16v2H4v-2zm0 5h10v2H4v-2z" /></svg>
                        <span>Song list</span>
                    </a>
                    <span class="nav-item is-disabled">
                        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M10 4a6 6 0 1 0 3.6 10.8l4.3 4.3 1.4-1.4-4.3-4.3A6 6 0 0 0 10 4zm0 2a4 4 0 1 1 0 8 4 4 0 0 1 0-8z" /></svg>
                        <span>Buscar</span>
                    </span>
                    <span class="nav-item is-disabled">
                        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 4h6v6H4V4zm10 0h6v6h-6V4zM4 14h6v6H4v-6zm10 0h6v6h-6v-6z" /></svg>
                        <span>Tu biblioteca</span>
                    </span>
                </nav>

                <div class="sidebar-card">
                    <p class="sidebar-card-title">Tu cuenta</p>
                    <div class="account">
                        <img id="imgSidebarPhoto" runat="server" class="account-photo" src="UserImages/default.jpg" alt="Foto del usuario" />
                        <div class="account-text">
                            <asp:Label ID="lblSidebarUser" runat="server" CssClass="account-name"></asp:Label>
                            <span class="account-status">Sesion activa</span>
                        </div>
                    </div>
                </div>
            </aside>

            <!-- ============================ MAIN ============================ -->
            <main class="main">

                <header class="topbar">
                    <div class="welcome">
                        <img id="imgUserPhoto" runat="server" class="welcome-photo" src="UserImages/default.jpg" alt="Foto del usuario" />
                        <div class="welcome-text">
                            <span class="welcome-label">Welcome</span>
                            <asp:Label ID="lblUserName" runat="server" CssClass="welcome-name"></asp:Label>
                        </div>
                    </div>

                    <a class="logout" href="Logout.aspx" title="Cerrar sesion">
                        <svg viewBox="0 0 24 24" aria-hidden="true">
                            <path d="M10 4H5v16h5" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" />
                            <path d="M14 8l4 4-4 4" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" />
                            <path d="M18 12H9" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" />
                        </svg>
                        <span>Logout</span>
                    </a>
                </header>

                <section class="hero">
                    <span class="hero-tag">Tu biblioteca</span>
                    <h1 class="hero-title">Song list</h1>
                    <p class="hero-sub"><asp:Label ID="lblSummary" runat="server"></asp:Label> &middot; disponibles para reproducir</p>
                    <div class="hero-actions">
                        <button type="button" class="btn-primary" id="btnPlayAll">
                            <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M8 5v14l11-7z" /></svg>
                            <span>Reproducir todo</span>
                        </button>
                    </div>
                </section>

                <section class="panel">

                    <div class="track-head">
                        <span class="c-id">ID</span>
                        <span class="c-name">Nombre</span>
                        <span class="c-path">Ruta</span>
                        <span class="c-plays">Reproducciones</span>
                        <span class="c-player">Reproducir</span>
                    </div>

                    <asp:Repeater ID="rptSongs" runat="server">
                        <ItemTemplate>
                            <div class="track" data-id='<%# Eval("ID_SONG") %>' data-src='<%# ResolveMediaUrl(Eval("SONG_PATH")) %>'>

                                <span class="c-id track-id"><%# Eval("ID_SONG") %></span>

                                <div class="c-name track-main">
                                    <span class="track-cover"><%# InitialOf(Eval("SONG_NAME")) %></span>
                                    <span class="track-text">
                                        <span class="track-name"><%# Server.HtmlEncode(Convert.ToString(Eval("SONG_NAME"))) %></span>
                                        <span class="track-artist">PoliMusic</span>
                                    </span>
                                </div>

                                <span class="c-path track-path"><%# Server.HtmlEncode(Convert.ToString(Eval("SONG_PATH"))) %></span>

                                <span class="c-plays track-plays">
                                    <span class="plays-count"><%# PlaysText(Eval("PLAYS")) %></span>
                                    <span class="plays-label">plays</span>
                                </span>

                                <div class="c-player track-player">
                                    <div class="player">

                                        <button type="button" class="btn-play" title="Reproducir" aria-label="Reproducir">
                                            <svg class="ico ico-play" viewBox="0 0 24 24" aria-hidden="true"><path d="M8 5v14l11-7z" /></svg>
                                            <svg class="ico ico-pause" viewBox="0 0 24 24" aria-hidden="true"><path d="M6 5h4v14H6zM14 5h4v14h-4z" /></svg>
                                        </button>

                                        <span class="time time-current">0:00</span>

                                        <div class="progress" title="Buscar dentro de la cancion">
                                            <div class="progress-fill"></div>
                                        </div>

                                        <span class="time time-total">0:00</span>

                                        <button type="button" class="btn-mute" title="Silenciar" aria-label="Silenciar">
                                            <svg class="ico ico-vol" viewBox="0 0 24 24" aria-hidden="true">
                                                <path d="M4 9v6h4l5 4V5L8 9H4z" />
                                                <path d="M16.5 8.6a4.8 4.8 0 0 1 0 6.8" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" />
                                            </svg>
                                            <svg class="ico ico-vol-muted" viewBox="0 0 24 24" aria-hidden="true">
                                                <path d="M4 9v6h4l5 4V5L8 9H4z" />
                                                <path d="M16 9l5 6M21 9l-5 6" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" />
                                            </svg>
                                        </button>

                                        <input type="range" class="volume" min="0" max="100" value="80" title="Volumen" aria-label="Volumen" />

                                        <button type="button" class="btn-more" title="Mas opciones" aria-label="Mas opciones">
                                            <svg viewBox="0 0 24 24" aria-hidden="true">
                                                <circle cx="12" cy="5" r="1.6" />
                                                <circle cx="12" cy="12" r="1.6" />
                                                <circle cx="12" cy="19" r="1.6" />
                                            </svg>
                                        </button>

                                        <div class="menu">
                                            <button type="button" class="menu-item" data-action="play">Reproducir</button>
                                            <button type="button" class="menu-item" data-action="copy">Copiar ruta</button>
                                        </div>

                                    </div>
                                </div>

                            </div>
                        </ItemTemplate>
                    </asp:Repeater>

                    <asp:Panel ID="pnlEmpty" runat="server" CssClass="empty" Visible="false">
                        <p class="empty-title">Todavia no hay canciones</p>
                        <p class="empty-sub">La tabla TBL_SONG no tiene registros para mostrar.</p>
                    </asp:Panel>

                </section>

            </main>
        </div>

        <div class="toast" id="toast" role="status" aria-live="polite"></div>
    </form>

    <script type="text/javascript" src="Scripts/Song.js"></script>
</body>
</html>
