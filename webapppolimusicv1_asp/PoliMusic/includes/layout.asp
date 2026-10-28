<%
' ============================================================================
'  PoliMusic - Layout compartido (sidebar + topbar + reproductor)
'
'  Uso en las paginas:
'      PageHeader "Titulo", "home", ""
'      ... contenido ...
'      PageFooter
'
'  Para las pantallas de acceso:
'      PageHeaderAuth "Iniciar sesion", ""
'      ... contenido ...
'      PageFooterAuth
' ============================================================================

' ============================================================================
'  PageHeader - abre el documento, dibuja la barra lateral y la barra superior
'  NOTA: sTitle se inserta tal cual, por lo que los datos provenientes del
'  usuario deben llegar ya escapados con H().
' ============================================================================
Sub PageHeader(sTitle, sActive, sExtraHead)
    Dim rsPl, bAdmin, sFlash, sNav, sQ, sLogoUrl

    bAdmin = IsAdmin()
    sNav   = LCase(Nz(sActive, ""))
    sFlash = FlashRender()
    sQ     = Nz(Request.QueryString("q"), "")
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="theme-color" content="#070B14">
<title><%=sTitle%> &middot; <%=APP_NAME%></title>
<link rel="icon" href="<%=AppRoot()%><%=IMG_LOGO%>" type="image/svg+xml">
<link rel="stylesheet" href="<%=AppRoot()%>css/style.css">
<%=sExtraHead%>
</head>
<body class="app">
<a class="skip-link" href="#content">Ir al contenido</a>
<div class="shell">

  <!-- ============================ BARRA LATERAL ============================ -->
  <aside class="sidebar" id="sidebar">
    <div class="brand">
      <a class="brand-link" href="<%=AppRoot()%>home.asp" title="<%=APP_NAME%>">
        <span class="brand-mark"><%=Ico("music", 22)%></span>
        <span class="brand-text">
          <span class="brand-name"><%=APP_NAME%></span>
          <span class="brand-tag">M&uacute;sica en l&iacute;nea</span>
        </span>
      </a>
      <button type="button" class="icon-btn sidebar-close" data-sidebar-close aria-label="Cerrar menu"><%=Ico("close", 18)%></button>
    </div>

    <nav class="nav" aria-label="Navegacion principal">
      <a class="nav-item <%=IIf(sNav="home","is-active","")%>" href="<%=AppRoot()%>home.asp">
        <%=Ico("home", 20)%><span>Inicio</span>
      </a>
      <a class="nav-item <%=IIf(sNav="search","is-active","")%>" href="<%=AppRoot()%>search.asp">
        <%=Ico("search", 20)%><span>Buscar</span>
      </a>
      <a class="nav-item <%=IIf(sNav="library","is-active","")%>" href="<%=AppRoot()%>library.asp">
        <%=Ico("library", 20)%><span>Tu biblioteca</span>
      </a>
      <a class="nav-item <%=IIf(sNav="albums","is-active","")%>" href="<%=AppRoot()%>albums.asp">
        <%=Ico("album", 20)%><span>Albumes</span>
      </a>
      <a class="nav-item <%=IIf(sNav="artists","is-active","")%>" href="<%=AppRoot()%>artists.asp">
        <%=Ico("artist", 20)%><span>Artistas</span>
      </a>
      <a class="nav-item <%=IIf(sNav="songs","is-active","")%>" href="<%=AppRoot()%>songs.asp">
        <%=Ico("music", 20)%><span>Canciones</span>
      </a>
    </nav>

    <div class="sidebar-sep"></div>

    <div class="sidebar-head">
      <span class="sidebar-title"><%=Ico("playlist", 16)%> Tus playlists</span>
      <a class="icon-btn" href="<%=AppRoot()%>playlist.asp?action=new" title="Crear playlist"><%=Ico("plus", 18)%></a>
    </div>

    <div class="sidebar-scroll">
      <ul class="side-list">
<%
    If bAdmin Then
        Set rsPl = DbSelect("SELECT TOP 30 ID_PLAYLIST, TITLE, PHOTO, TYPE FROM TBL_PLAYLIST ORDER BY ID_PLAYLIST DESC", Empty)
    Else
        Set rsPl = DbSelect("SELECT TOP 30 ID_PLAYLIST, TITLE, PHOTO, TYPE FROM TBL_PLAYLIST " & _
                            "WHERE ID_USER = ? OR TYPE = ? ORDER BY TYPE, ID_PLAYLIST DESC", _
                            Array(CurrentUserId(), PLAYLIST_PUBLIC))
    End If
    If rsPl.EOF Then
%>
        <li class="side-empty">Aun no tienes playlists. Crea la primera con el boton <strong>+</strong>.</li>
<%
    Else
        Do While Not rsPl.EOF
%>
        <li>
          <a class="side-item" href="<%=AppRoot()%>playlist.asp?id=<%=rsPl("ID_PLAYLIST")%>" title="<%=HA(rsPl("TITLE"))%>">
            <img class="side-thumb" src="<%=AssetUrl(rsPl("PHOTO"), IMG_FALLBACK_COVER)%>" alt="" loading="lazy">
            <span class="side-item-body">
              <span class="side-item-title"><%=H(rsPl("TITLE"))%></span>
              <span class="side-item-sub"><%=IIf(ToInt(rsPl("TYPE"))=PLAYLIST_PRIVATE, "Privada", "Publica")%></span>
            </span>
          </a>
        </li>
<%
            rsPl.MoveNext
        Loop
    End If
    rsPl.Close
%>
      </ul>

<% If bAdmin Then %>
      <div class="sidebar-sep"></div>
      <div class="sidebar-head">
        <span class="sidebar-title"><%=Ico("shield", 16)%> Administracion</span>
      </div>
      <ul class="side-list">
        <li><a class="side-item side-item-sm" href="<%=AppRoot()%>admin/default.asp"><%=Ico("dashboard", 16)%><span>Panel</span></a></li>
        <li><a class="side-item side-item-sm" href="<%=AppRoot()%>admin/list.asp?e=users"><%=Ico("users", 16)%><span>Usuarios</span></a></li>
        <li><a class="side-item side-item-sm" href="<%=AppRoot()%>admin/list.asp?e=songs"><%=Ico("music", 16)%><span>Canciones</span></a></li>
        <li><a class="side-item side-item-sm" href="<%=AppRoot()%>admin/list.asp?e=albums"><%=Ico("album", 16)%><span>Albumes</span></a></li>
        <li><a class="side-item side-item-sm" href="<%=AppRoot()%>admin/list.asp?e=artists"><%=Ico("artist", 16)%><span>Artistas</span></a></li>
        <li><a class="side-item side-item-sm" href="<%=AppRoot()%>admin/list.asp?e=playlists"><%=Ico("playlist", 16)%><span>Playlists</span></a></li>
      </ul>
<% End If %>
    </div>
  </aside>
  <div class="sidebar-backdrop" data-sidebar-close></div>

  <!-- ============================== CONTENIDO ============================== -->
  <div class="main-wrap">
    <header class="topbar">
      <button type="button" class="icon-btn sidebar-open" data-sidebar-open aria-label="Abrir menu"><%=Ico("queue", 20)%></button>
      <div class="hist-nav">
        <button type="button" class="icon-btn" onclick="history.back()" aria-label="Atras"><%=Ico("chevron-left", 18)%></button>
        <button type="button" class="icon-btn" onclick="history.forward()" aria-label="Adelante"><%=Ico("chevron-right", 18)%></button>
      </div>

      <form class="topbar-search" action="<%=AppRoot()%>search.asp" method="get" role="search">
        <span class="search-ico"><%=Ico("search", 18)%></span>
        <input type="search" name="q" id="globalSearch" value="<%=HA(sQ)%>" placeholder="Que quieres escuchar?" autocomplete="off" aria-label="Buscar">
      </form>

      <div class="topbar-right">
        <a class="btn btn-ghost btn-sm topbar-new" href="<%=AppRoot()%>playlist.asp?action=new"><%=Ico("plus", 16)%><span>Crear playlist</span></a>
        <div class="usermenu" data-usermenu>
          <button type="button" class="usermenu-btn" data-usermenu-toggle aria-haspopup="true" aria-expanded="false">
            <img class="avatar" src="<%=CurrentUserPhotoUrl()%>" alt="">
            <span class="usermenu-name"><%=H(CurrentUserName())%></span>
            <%=Ico("chevron-down", 14)%>
          </button>
          <div class="usermenu-panel" data-usermenu-panel role="menu">
            <div class="usermenu-head">
              <img class="avatar avatar-lg" src="<%=CurrentUserPhotoUrl()%>" alt="">
              <div>
                <div class="usermenu-title"><%=H(CurrentUserName())%></div>
                <div class="usermenu-sub"><%=H(CurrentUserEmail())%></div>
                <span class="badge <%=IIf(bAdmin,"badge-admin","badge-user")%>"><%=H(CurrentRoleLabel())%></span>
              </div>
            </div>
            <a class="usermenu-item" href="<%=AppRoot()%>profile.asp" role="menuitem"><%=Ico("user", 16)%> Mi perfil</a>
            <a class="usermenu-item" href="<%=AppRoot()%>library.asp" role="menuitem"><%=Ico("library", 16)%> Mis playlists</a>
<% If bAdmin Then %>
            <a class="usermenu-item" href="<%=AppRoot()%>admin/default.asp" role="menuitem"><%=Ico("shield", 16)%> Panel de administracion</a>
<% End If %>
            <div class="usermenu-sep"></div>
            <a class="usermenu-item usermenu-danger" href="<%=AppRoot()%>logout.asp" role="menuitem"><%=Ico("logout", 16)%> Cerrar sesion</a>
          </div>
        </div>
      </div>
    </header>

    <main class="content" id="content">
<% If Len(sFlash) > 0 Then Response.Write sFlash End If %>
<%
End Sub

' ============================================================================
'  PageFooter - cierra el contenido y dibuja la barra del reproductor
' ============================================================================
Sub PageFooter()
%>
    </main>
  </div>
</div>

<!-- ============================ REPRODUCTOR ============================ -->
<div class="playerbar" id="playerbar">
  <div class="pb-left">
    <div class="pb-cover">
      <img id="pbCover" src="<%=AppRoot()%><%=IMG_FALLBACK_COVER%>" alt="">
      <span class="pb-cover-empty"><%=Ico("music", 22)%></span>
    </div>
    <div class="pb-meta">
      <a class="pb-title" id="pbTitle" href="#">Sin reproduccion</a>
      <a class="pb-artist" id="pbArtist" href="#">Selecciona una cancion</a>
    </div>
    <button type="button" class="icon-btn pb-like" id="pbLike" title="Me gusta" aria-pressed="false"><%=Ico("heart", 18)%></button>
  </div>

  <div class="pb-center">
    <div class="pb-controls">
      <button type="button" class="icon-btn" id="pbShuffle" title="Aleatorio" aria-pressed="false"><%=Ico("shuffle", 18)%></button>
      <button type="button" class="icon-btn" id="pbPrev" title="Anterior"><%=Ico("prev", 20)%></button>
      <button type="button" class="pb-play" id="pbPlay" title="Reproducir"><%=Ico("play", 22)%></button>
      <button type="button" class="icon-btn" id="pbNext" title="Siguiente"><%=Ico("next", 20)%></button>
      <button type="button" class="icon-btn" id="pbRepeat" title="Repetir" aria-pressed="false"><%=Ico("repeat", 18)%></button>
    </div>
    <div class="pb-progress">
      <span class="pb-time" id="pbCur">0:00</span>
      <input type="range" class="seek" id="pbSeek" min="0" max="100" value="0" step="1" aria-label="Posicion">
      <span class="pb-time" id="pbDur">0:00</span>
    </div>
  </div>

  <div class="pb-right">
    <button type="button" class="icon-btn" id="pbQueueBtn" title="Cola de reproduccion"><%=Ico("queue", 18)%></button>
    <button type="button" class="icon-btn" id="pbMute" title="Silenciar"><%=Ico("volume", 18)%></button>
    <input type="range" class="vol" id="pbVol" min="0" max="100" value="80" step="1" aria-label="Volumen">
  </div>
</div>

<!-- ============================== COLA ============================== -->
<aside class="queuepanel" id="queuepanel" aria-hidden="true">
  <div class="queuepanel-head">
    <span>Cola de reproduccion</span>
    <button type="button" class="icon-btn" id="pbQueueClose" aria-label="Cerrar"><%=Ico("close", 18)%></button>
  </div>
  <div class="queuepanel-body" id="queueList"><p class="muted">La cola esta vacia.</p></div>
</aside>

<audio id="pmAudio" preload="metadata"></audio>

<script>
window.PM = window.PM || {};
PM.appRoot = <%=J(AppRoot())%>;
PM.assets  = <%=J(AppRoot() & ASSET_FOLDER & "/")%>;
PM.user    = { id: <%=CurrentUserId()%>, name: <%=J(CurrentUserName())%>, admin: <%=IIf(IsAdmin(), "true", "false")%> };
</script>
<script src="<%=AppRoot()%>js/app.js"></script>
</body>
</html>
<%
End Sub

' ============================================================================
'  PageHeaderAuth - layout sencillo (sin sidebar ni reproductor) para
'  las pantallas de login / registro
' ============================================================================
Sub PageHeaderAuth(sTitle, sExtraHead)
    Dim sFlash
    sFlash = FlashRender()
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="theme-color" content="#070B14">
<title><%=sTitle%> &middot; <%=APP_NAME%></title>
<link rel="icon" href="<%=AppRoot()%><%=IMG_LOGO%>" type="image/svg+xml">
<link rel="stylesheet" href="<%=AppRoot()%>css/style.css">
<%=sExtraHead%>
</head>
<body class="auth">
<main class="auth-wrap">
  <section class="auth-hero">
    <div class="auth-hero-inner">
      <a class="brand-link" href="<%=AppRoot()%>login.asp" title="<%=APP_NAME%>">
        <span class="brand-mark"><%=Ico("music", 26)%></span>
        <span class="brand-text">
          <span class="brand-name"><%=APP_NAME%></span>
          <span class="brand-tag">M&uacute;sica en l&iacute;nea</span>
        </span>
      </a>
      <h1 class="auth-hero-title"><%=APP_TAGLINE%></h1>
      <p class="auth-hero-sub">Escucha tus canciones favoritas, arma tus playlists y descubre nuevos albumes y artistas. Todo con la identidad azul de PoliMusic.</p>
      <ul class="auth-hero-list">
        <li><%=Ico("music", 18)%> Reproductor web integrado</li>
        <li><%=Ico("playlist", 18)%> Playlists publicas y privadas</li>
        <li><%=Ico("album", 18)%> Catalogo de albumes y artistas</li>
        <li><%=Ico("shield", 18)%> Administracion de contenido</li>
      </ul>
    </div>
  </section>
  <section class="auth-panel">
    <div class="auth-card">
<%=sFlash%>
<%
End Sub

Sub PageFooterAuth()
%>
    </div>
    <p class="auth-foot">&copy; <%=Year(Date())%> <%=APP_NAME%> &middot; Aplicacion web ASP clasico &middot; Escuela Politecnica Nacional</p>
  </section>
</main>
</body>
</html>
<%
End Sub
%>
