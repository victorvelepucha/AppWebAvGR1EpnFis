<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' ============================================================================
'  search.asp - Busqueda en todo el catalogo (canciones, albumes, artistas,
'  playlists). Los usuarios normales solo ven playlists publicas o propias.
' ============================================================================
Dim sQ, sTab, sLike, rs, nTotal, page, pageSize, offset, o, nCount

RequireLogin

sQ    = Nz(Request.QueryString("q"), "")
sTab  = LCase(Nz(Request.QueryString("tab"), "all"))
sLike = "%" & sQ & "%"

page     = ReqPage()
pageSize = ROWS_PER_PAGE
offset   = (page - 1) * pageSize

PageHeader IIf(Len(sQ) > 0, "Buscar: " & H(sQ), "Buscar"), "search", ""
%>

<div class="page-head">
  <div>
    <h1 class="page-title">Buscar</h1>
    <p class="page-sub">
<% If Len(sQ) > 0 Then %>
      Resultados para &laquo;<strong><%=H(sQ)%></strong>&raquo;
<% Else %>
      Escribe en el buscador para encontrar canciones, &aacute;lbumes, artistas y playlists.
<% End If %>
    </p>
  </div>
</div>

<% If Len(sQ) = 0 Then %>
  <div class="empty">
    <div class="empty-ico"><%=Ico("search", 28)%></div>
    <div class="empty-title">Empieza a explorar</div>
    <p class="empty-text">Prueba con el nombre de una canci&oacute;n, un &aacute;lbum, un artista o una playlist.</p>
    <div class="hstack" style="justify-content:center">
      <a class="chip" href="<%=AppRoot()%>search.asp?q=Adventure">Adventure</a>
      <a class="chip" href="<%=AppRoot()%>search.asp?q=Rock">Rock</a>
      <a class="chip" href="<%=AppRoot()%>search.asp?q=Juan">Juan</a>
      <a class="chip" href="<%=AppRoot()%>search.asp?q=Album">Album</a>
    </div>
  </div>
<% Else %>

<nav class="tabs">
  <a class="tab <%=IIf(sTab="all","is-active","")%>" href="<%=AppRoot()%>search.asp?q=<%=Server.URLEncode(sQ)%>&amp;tab=all">Todo</a>
  <a class="tab <%=IIf(sTab="songs","is-active","")%>" href="<%=AppRoot()%>search.asp?q=<%=Server.URLEncode(sQ)%>&amp;tab=songs">Canciones</a>
  <a class="tab <%=IIf(sTab="albums","is-active","")%>" href="<%=AppRoot()%>search.asp?q=<%=Server.URLEncode(sQ)%>&amp;tab=albums">&Aacute;lbumes</a>
  <a class="tab <%=IIf(sTab="artists","is-active","")%>" href="<%=AppRoot()%>search.asp?q=<%=Server.URLEncode(sQ)%>&amp;tab=artists">Artistas</a>
  <a class="tab <%=IIf(sTab="playlists","is-active","")%>" href="<%=AppRoot()%>search.asp?q=<%=Server.URLEncode(sQ)%>&amp;tab=playlists">Playlists</a>
</nav>

<%
' ============================== CANCIONES ==============================
If sTab = "all" Or sTab = "songs" Then
    nTotal = DbCount("SELECT COUNT(*) FROM TBL_SONG S WHERE S.SONG_NAME LIKE ?", Array(sLike))
    If nTotal > 0 Then
%>
<section class="section">
  <div class="section-head">
    <h2 class="section-title">Canciones <span class="muted small">(<%=nTotal%>)</span></h2>
    <a class="section-link" href="<%=AppRoot()%>search.asp?q=<%=Server.URLEncode(sQ)%>&amp;tab=songs">Ver todas</a>
  </div>
<%
        If sTab = "all" Then
            ' En la vista "todo" se limita a 6 filas por bloque
            Set rs = DbSelect("SELECT TOP 6 * FROM (" & SongSelectSql() & "WHERE S.SONG_NAME LIKE ?) AS t ORDER BY t.SONG_NAME", Array(sLike))
        Else
            Set rs = DbSelect(SongSelectSql() & "WHERE S.SONG_NAME LIKE ? ORDER BY S.SONG_NAME " & _
                              "OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY", Array(sLike))
        End If
        Set o = Server.CreateObject("Scripting.Dictionary")
        o("showAlbum") = True
        o("showPlays") = True
        o("showAdd")   = True
        o("containerId") = "searchSongs"
        RenderTracklist rs, o
        rs.Close
        If sTab = "songs" Then
            RenderPager nTotal, page, pageSize, "search.asp", Array("q", sQ, "tab", "songs")
        End If
%>
</section>
<%
    ElseIf sTab = "songs" Then
        Response.Write "<div class=""alert alert-info""><span class=""alert-ico"">i</span><span>No se encontraron canciones.</span></div>"
    End If
End If

' =============================== ALBUMES ===============================
If sTab = "all" Or sTab = "albums" Then
    nTotal = DbCount("SELECT COUNT(*) FROM TBL_ALBUM WHERE TITLE_ALBUM LIKE ?", Array(sLike))
    If nTotal > 0 Then
%>
<section class="section">
  <div class="section-head">
    <h2 class="section-title">&Aacute;lbumes <span class="muted small">(<%=nTotal%>)</span></h2>
  </div>
<%
        If sTab = "all" Then
            Set rs = DbSelect("SELECT TOP 6 ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE TITLE_ALBUM LIKE ? ORDER BY TITLE_ALBUM", Array(sLike))
        Else
            Set rs = DbSelect("SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE TITLE_ALBUM LIKE ? " & _
                              "ORDER BY TITLE_ALBUM OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY", Array(sLike))
        End If
        RenderAlbumCards rs
        rs.Close
        If sTab = "albums" Then RenderPager nTotal, page, pageSize, "search.asp", Array("q", sQ, "tab", "albums")
%>
</section>
<%
    ElseIf sTab = "albums" Then
        Response.Write "<div class=""alert alert-info""><span class=""alert-ico"">i</span><span>No se encontraron &aacute;lbumes.</span></div>"
    End If
End If

' ============================== ARTISTAS ==============================
If sTab = "all" Or sTab = "artists" Then
    nTotal = DbCount("SELECT COUNT(*) FROM TBL_ARTIST WHERE ARTIST_NAME LIKE ? OR ARTIST_LASTNAME LIKE ?", Array(sLike, sLike))
    If nTotal > 0 Then
%>
<section class="section">
  <div class="section-head">
    <h2 class="section-title">Artistas <span class="muted small">(<%=nTotal%>)</span></h2>
  </div>
<%
        If sTab = "all" Then
            Set rs = DbSelect("SELECT TOP 6 ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST " & _
                              "WHERE ARTIST_NAME LIKE ? OR ARTIST_LASTNAME LIKE ? ORDER BY ARTIST_NAME", Array(sLike, sLike))
        Else
            Set rs = DbSelect("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST " & _
                              "WHERE ARTIST_NAME LIKE ? OR ARTIST_LASTNAME LIKE ? ORDER BY ARTIST_NAME " & _
                              "OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY", Array(sLike, sLike))
        End If
        RenderArtistCards rs
        rs.Close
        If sTab = "artists" Then RenderPager nTotal, page, pageSize, "search.asp", Array("q", sQ, "tab", "artists")
%>
</section>
<%
    ElseIf sTab = "artists" Then
        Response.Write "<div class=""alert alert-info""><span class=""alert-ico"">i</span><span>No se encontraron artistas.</span></div>"
    End If
End If

' ============================== PLAYLISTS ==============================
If sTab = "all" Or sTab = "playlists" Then
    Dim sWhere
    If IsAdmin() Then
        sWhere = "WHERE P.TITLE LIKE ?"
    Else
        sWhere = "WHERE P.TITLE LIKE ? AND (P.TYPE = " & PLAYLIST_PUBLIC & " OR P.ID_USER = " & CurrentUserId() & ")"
    End If
    nTotal = DbCount("SELECT COUNT(*) FROM TBL_PLAYLIST P " & sWhere, Array(sLike))
    If nTotal > 0 Then
%>
<section class="section">
  <div class="section-head">
    <h2 class="section-title">Playlists <span class="muted small">(<%=nTotal%>)</span></h2>
  </div>
<%
        If sTab = "all" Then
            Set rs = DbSelect("SELECT TOP 6 P.ID_PLAYLIST, P.TITLE, P.PHOTO, P.TYPE, P.ID_USER, U.USERNAME AS OWNER_NAME " & _
                              "FROM TBL_PLAYLIST P LEFT JOIN TBL_USER U ON U.ID_USER = P.ID_USER " & sWhere & " ORDER BY P.TITLE", Array(sLike))
        Else
            Set rs = DbSelect("SELECT P.ID_PLAYLIST, P.TITLE, P.PHOTO, P.TYPE, P.ID_USER, U.USERNAME AS OWNER_NAME " & _
                              "FROM TBL_PLAYLIST P LEFT JOIN TBL_USER U ON U.ID_USER = P.ID_USER " & sWhere & " ORDER BY P.TITLE " & _
                              "OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY", Array(sLike))
        End If
        RenderPlaylistCards rs, True
        rs.Close
        If sTab = "playlists" Then RenderPager nTotal, page, pageSize, "search.asp", Array("q", sQ, "tab", "playlists")
%>
</section>
<%
    ElseIf sTab = "playlists" Then
        Response.Write "<div class=""alert alert-info""><span class=""alert-ico"">i</span><span>No se encontraron playlists.</span></div>"
    End If
End If
%>

<% End If %>

<%
PageFooter
%>
