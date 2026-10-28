<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' ============================================================================
'  home.asp - Pantalla principal de PoliMusic
' ============================================================================
Dim rsTop, rsAlbums, rsArtists, rsMyPl, rsRecent, nSongs, nAlbums, nArtists, nPl, sUserFirst

RequireLogin

nSongs   = DbCount("SELECT COUNT(*) FROM TBL_SONG", Empty)
nAlbums  = DbCount("SELECT COUNT(*) FROM TBL_ALBUM", Empty)
nArtists = DbCount("SELECT COUNT(*) FROM TBL_ARTIST", Empty)
nPl      = DbCount("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_USER = ? OR TYPE = ?", Array(CurrentUserId(), PLAYLIST_PUBLIC))

' Se muestra el nombre completo del usuario, sin cortarlo por el punto
sUserFirst = CurrentUserName()
If InStr(sUserFirst, " ") > 0 Then sUserFirst = Split(sUserFirst, " ")(0)
If Len(sUserFirst) > 24 Then sUserFirst = Left(sUserFirst, 24)

PageHeader "Inicio", "home", ""
%>

<!-- ================================ HERO ================================ -->
<section class="hero">
  <div class="hero-inner">
    <span class="hero-kicker"><%=Ico("speaker", 15)%> Escuela Polit&eacute;cnica Nacional &middot; M&uacute;sica en l&iacute;nea</span>
    <h1>Hola, <%=H(sUserFirst)%>. Bienvenido a <%=APP_NAME%>.</h1>
    <p>Reproduce tus canciones, explora el cat&aacute;logo de &aacute;lbumes y artistas y organiza tu m&uacute;sica en playlists p&uacute;blicas o privadas.</p>
    <div class="hero-actions">
      <button type="button" class="btn btn-primary btn-lg" data-play-list="#homeTop"><%=Ico("play", 20)%> Reproducir lo m&aacute;s escuchado</button>
      <a class="btn btn-outline btn-lg" href="<%=AppRoot()%>search.asp"><%=Ico("search", 19)%> Explorar cat&aacute;logo</a>
    </div>
    <div class="hero-stats">
      <div><div class="hero-stat-value"><%=nSongs%></div><div class="hero-stat-label">Canciones</div></div>
      <div><div class="hero-stat-value"><%=nAlbums%></div><div class="hero-stat-label">&Aacute;lbumes</div></div>
      <div><div class="hero-stat-value"><%=nArtists%></div><div class="hero-stat-label">Artistas</div></div>
      <div><div class="hero-stat-value"><%=nPl%></div><div class="hero-stat-label">Playlists</div></div>
    </div>
  </div>
</section>

<!-- ========================= LO MAS ESCUCHADO ========================= -->
<section class="section">
  <div class="section-head">
    <h2 class="section-title">Lo m&aacute;s escuchado</h2>
    <a class="section-link" href="<%=AppRoot()%>songs.asp">Ver todo</a>
  </div>
<%
Set rsTop = DbSelect(SongSelectSql() & "ORDER BY ISNULL(S.PLAYS,0) DESC, S.SONG_NAME ASC", Empty)
Dim oTop : Set oTop = Server.CreateObject("Scripting.Dictionary")
oTop("showAlbum") = True
oTop("showPlays") = True
oTop("showAdd")   = True
oTop("containerId") = "homeTop"
RenderTracklist rsTop, oTop
rsTop.Close
%>
</section>

<!-- ============================ TUS PLAYLISTS ============================ -->
<section class="section">
  <div class="section-head">
    <h2 class="section-title"><%=IIf(IsAdmin(), "Playlists del sistema", "Tus playlists")%></h2>
    <a class="section-link" href="<%=AppRoot()%>library.asp">Ver biblioteca</a>
  </div>
<%
If IsAdmin() Then
    Set rsMyPl = DbSelect("SELECT TOP 6 ID_PLAYLIST, TITLE, PHOTO, TYPE, ID_USER FROM TBL_PLAYLIST ORDER BY ID_PLAYLIST DESC", Empty)
Else
    Set rsMyPl = DbSelect("SELECT TOP 6 ID_PLAYLIST, TITLE, PHOTO, TYPE, ID_USER FROM TBL_PLAYLIST WHERE ID_USER = ? ORDER BY ID_PLAYLIST DESC", Array(CurrentUserId()))
End If
If rsMyPl.EOF Then
    RenderEmpty "playlist", "A&uacute;n no tienes playlists", "Crea tu primera playlist y empieza a organizar tus canciones favoritas.", "playlist.asp?action=new", "Crear playlist"
Else
    RenderPlaylistCards rsMyPl, False
End If
rsMyPl.Close
%>
</section>

<!-- ============================== ALBUMES ============================== -->
<section class="section">
  <div class="section-head">
    <h2 class="section-title">&Aacute;lbumes destacados</h2>
    <a class="section-link" href="<%=AppRoot()%>albums.asp">&Aacute;lbumes</a>
  </div>
<%
Set rsAlbums = DbSelect("SELECT TOP 6 ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM ORDER BY ID_ALBUM DESC", Empty)
RenderAlbumCards rsAlbums
rsAlbums.Close
%>
</section>

<!-- ============================== ARTISTAS ============================== -->
<section class="section">
  <div class="section-head">
    <h2 class="section-title">Artistas</h2>
    <a class="section-link" href="<%=AppRoot()%>artists.asp">Todos los artistas</a>
  </div>
<%
Set rsArtists = DbSelect("SELECT TOP 6 ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST ORDER BY ID_ARTIST DESC", Empty)
RenderArtistCards rsArtists
rsArtists.Close
%>
</section>

<!-- ====================== PLAYLISTS PUBLICAS ====================== -->
<section class="section">
  <div class="section-head">
    <h2 class="section-title">Playlists p&uacute;blicas</h2>
    <a class="section-link" href="<%=AppRoot()%>library.asp?tab=public">Explorar</a>
  </div>
<%
Set rsRecent = DbSelect("SELECT TOP 6 P.ID_PLAYLIST, P.TITLE, P.PHOTO, P.TYPE, P.ID_USER, U.USERNAME AS OWNER_NAME " & _
                        "FROM TBL_PLAYLIST P LEFT JOIN TBL_USER U ON U.ID_USER = P.ID_USER " & _
                        "WHERE P.TYPE = ? ORDER BY P.ID_PLAYLIST DESC", Array(PLAYLIST_PUBLIC))
RenderPlaylistCards rsRecent, True
rsRecent.Close
%>
</section>

<%
PageFooter
%>
