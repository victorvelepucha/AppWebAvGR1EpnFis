<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' ============================================================================
'  artist.asp - Detalle de un artista: canciones y albumes relacionados
' ============================================================================
Dim idArtist, rs, o, nSongs, nAlbums, sName, sImg

RequireLogin

idArtist = ReqInt("id")
If idArtist <= 0 Then
    Response.Redirect AppRoot() & "artists.asp"
    Response.End
End If

Set rs = DbSelect("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array(idArtist))
If rs.EOF Then
    rs.Close
    Response.Status = "404 Not Found"
    PageHeader "Artista no encontrado", "artists", ""
    RenderEmpty "artist", "El artista no existe", "No encontramos el artista solicitado en el cat&aacute;logo.", "artists.asp", "Ver todos los artistas"
    PageFooter
    Response.End
End If

sName = ArtistFullName(rs("ARTIST_NAME"), rs("ARTIST_LASTNAME"))
sImg  = AssetUrl(rs("ARTIST_IMAGE"), IMG_FALLBACK_AVATAR)
rs.Close

nSongs  = DbCount("SELECT COUNT(*) FROM TBL_PLAYER WHERE ID_ARTIST = ?", Array(idArtist))
nAlbums = DbCount("SELECT COUNT(DISTINCT T.ID_ALBUM) FROM TBL_TRACKLIST T INNER JOIN TBL_PLAYER P ON P.ID_SONG = T.ID_SONG WHERE P.ID_ARTIST = ?", Array(idArtist))

PageHeader H(sName), "artists", ""
%>

<div class="detail-head">
  <div class="detail-cover detail-cover-round"><img src="<%=HA(sImg)%>" alt="<%=HA(sName)%>"></div>
  <div class="detail-info">
    <div class="detail-kind">Artista</div>
    <h1 class="detail-title"><%=H(sName)%></h1>
    <div class="detail-meta">
      <span><%=Plural(nSongs, "canci&oacute;n", "canciones")%></span>
      <span class="dot"></span><span><%=Plural(nAlbums, "&aacute;lbum", "&aacute;lbumes")%></span>
    </div>
    <div class="detail-actions">
      <button type="button" class="btn btn-primary btn-lg" data-play-list="#artistTracks"><%=Ico("play", 20)%> Reproducir</button>
      <button type="button" class="btn btn-outline" data-play-shuffle="#artistTracks"><%=Ico("shuffle", 18)%> Aleatorio</button>
<% If IsAdmin() Then %>
      <a class="btn btn-secondary" href="<%=AppRoot()%>admin/edit.asp?e=artists&amp;id=<%=idArtist%>"><%=Ico("edit", 17)%> Editar</a>
      <form method="post" action="<%=AppRoot()%>admin/action.asp" class="inline-form" data-confirm="&iquest;Eliminar este artista? Se quitar&aacute; de las canciones asociadas.">
        <input type="hidden" name="entity" value="artists">
        <input type="hidden" name="action" value="delete">
        <input type="hidden" name="id" value="<%=idArtist%>">
        <input type="hidden" name="returnTo" value="artists.asp">
        <button type="submit" class="btn btn-danger"><%=Ico("trash", 17)%> Eliminar</button>
      </form>
<% End If %>
    </div>
  </div>
</div>

<section class="section">
  <div class="section-head">
    <h2 class="section-title">Canciones</h2>
    <span class="muted small"><%=nSongs%> en el cat&aacute;logo</span>
  </div>
<%
Set rs = DbSelect(SongSelectSql() & "WHERE EXISTS (SELECT 1 FROM TBL_PLAYER P WHERE P.ID_SONG = S.ID_SONG AND P.ID_ARTIST = ?) " & _
                  "ORDER BY ISNULL(S.PLAYS,0) DESC, S.SONG_NAME", Array(idArtist))
Set o = Server.CreateObject("Scripting.Dictionary")
o("showAlbum")   = True
o("showPlays")   = True
o("showAdd")     = True
o("containerId") = "artistTracks"
RenderTracklist rs, o
rs.Close
%>
</section>

<section class="section">
  <div class="section-head">
    <h2 class="section-title">&Aacute;lbumes del artista</h2>
  </div>
<%
Set rs = DbSelect("SELECT DISTINCT A.ID_ALBUM, A.TITLE_ALBUM, A.RELEASE_YEAR, A.ALBUM_IMAGE_PATH " & _
                  "FROM TBL_ALBUM A INNER JOIN TBL_TRACKLIST T ON T.ID_ALBUM = A.ID_ALBUM " & _
                  "INNER JOIN TBL_PLAYER P ON P.ID_SONG = T.ID_SONG " & _
                  "WHERE P.ID_ARTIST = ? ORDER BY A.TITLE_ALBUM", Array(idArtist))
RenderAlbumCards rs
rs.Close
%>
</section>

<%
PageFooter
%>
