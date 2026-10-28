<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' ============================================================================
'  album.asp - Detalle de un album con su lista de canciones
' ============================================================================
Dim idAlbum, rs, rsArt, o, nTracks, nPlays, sArtists, bAutoPlay, sCover, sTitle, sYear

RequireLogin

idAlbum = ReqInt("id")
If idAlbum <= 0 Then
    Response.Redirect AppRoot() & "albums.asp"
    Response.End
End If

Set rs = DbSelect("SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(idAlbum))
If rs.EOF Then
    rs.Close
    Response.Status = "404 Not Found"
    PageHeader "Album no encontrado", "albums", ""
    RenderEmpty "album", "El &aacute;lbum no existe", "No encontramos el &aacute;lbum solicitado en el cat&aacute;logo.", "albums.asp", "Ver todos los &aacute;lbumes"
    PageFooter
    Response.End
End If

sTitle = CStr(rs("TITLE_ALBUM").Value)
sYear  = Nz(rs("RELEASE_YEAR").Value, "")
sCover = AssetUrl(rs("ALBUM_IMAGE_PATH").Value, IMG_FALLBACK_COVER)
rs.Close

nTracks  = DbCount("SELECT COUNT(*) FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array(idAlbum))
nPlays   = DbCount("SELECT ISNULL(SUM(S.PLAYS),0) FROM TBL_TRACKLIST T INNER JOIN TBL_SONG S ON S.ID_SONG = T.ID_SONG WHERE T.ID_ALBUM = ?", Array(idAlbum))
bAutoPlay = (Req("play") = "1")

' Artistas presentes en el album
sArtists = ""
Set rsArt = DbSelect("SELECT DISTINCT A.ID_ARTIST, A.ARTIST_NAME, A.ARTIST_LASTNAME FROM TBL_ARTIST A " & _
                     "INNER JOIN TBL_PLAYER P ON P.ID_ARTIST = A.ID_ARTIST " & _
                     "INNER JOIN TBL_TRACKLIST T ON T.ID_SONG = P.ID_SONG " & _
                     "WHERE T.ID_ALBUM = ? ORDER BY A.ARTIST_NAME", Array(idAlbum))
Do While Not rsArt.EOF
    If Len(sArtists) > 0 Then sArtists = sArtists & ", "
    sArtists = sArtists & ArtistFullName(rsArt("ARTIST_NAME"), rsArt("ARTIST_LASTNAME"))
    rsArt.MoveNext
Loop
rsArt.Close
If Len(sArtists) = 0 Then sArtists = "Artista desconocido"

PageHeader H(sTitle), "albums", ""
%>

<div class="detail-head">
  <div class="detail-cover"><img src="<%=HA(sCover)%>" alt="<%=HA(sTitle)%>"></div>
  <div class="detail-info">
    <div class="detail-kind">&Aacute;lbum</div>
    <h1 class="detail-title"><%=H(sTitle)%></h1>
    <div class="detail-meta">
      <span><%=H(sArtists)%></span>
      <span class="dot"></span><span><%=IIf(Len(CStr(sYear))>0, H(sYear), "Sin a&ntilde;o")%></span>
      <span class="dot"></span><span><%=Plural(nTracks, "canci&oacute;n", "canciones")%></span>
      <span class="dot"></span><span><%=FmtNumber(nPlays)%> reproducciones</span>
    </div>
    <div class="detail-actions">
      <button type="button" class="btn btn-primary btn-lg" data-play-list="#albumTracks"><%=Ico("play", 20)%> Reproducir</button>
      <button type="button" class="btn btn-outline" data-play-shuffle="#albumTracks"><%=Ico("shuffle", 18)%> Aleatorio</button>
<% If IsAdmin() Then %>
      <a class="btn btn-secondary" href="<%=AppRoot()%>admin/edit.asp?e=albums&amp;id=<%=idAlbum%>"><%=Ico("edit", 17)%> Editar</a>
      <form method="post" action="<%=AppRoot()%>admin/action.asp" class="inline-form" data-confirm="&iquest;Eliminar este &aacute;lbum? Tambi&eacute;n se quitar&aacute;n sus canciones del listado.">
        <input type="hidden" name="entity" value="albums">
        <input type="hidden" name="action" value="delete">
        <input type="hidden" name="id" value="<%=idAlbum%>">
        <input type="hidden" name="returnTo" value="albums.asp">
        <button type="submit" class="btn btn-danger"><%=Ico("trash", 17)%> Eliminar</button>
      </form>
<% End If %>
    </div>
  </div>
</div>

<section class="section">
  <div class="section-head">
    <h2 class="section-title">Canciones del &aacute;lbum</h2>
    <span class="muted small"><%=Plural(nTracks, "pista", "pistas")%></span>
  </div>
<%
Set rs = DbSelect(SongSelectSql() & "WHERE EXISTS (SELECT 1 FROM TBL_TRACKLIST T WHERE T.ID_SONG = S.ID_SONG AND T.ID_ALBUM = ?) " & _
                  "ORDER BY S.SONG_NAME", Array(idAlbum))
Set o = Server.CreateObject("Scripting.Dictionary")
o("showAlbum")   = False
o("showPlays")   = True
o("showAdd")     = True
o("containerId") = "albumTracks"
RenderTracklist rs, o
rs.Close
%>
</section>

<% If bAutoPlay Then %>
<script>window.PM = window.PM || {}; PM.autoPlay = "#albumTracks";</script>
<% End If %>

<%
PageFooter
%>
