<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' ============================================================================
'  song.asp - Detalle de una cancion
' ============================================================================
Dim idSong, rs, sName, sPath, sArtist, sAlbum, sCover, nPlays, idArtist, idAlbum, sAudio, o

RequireLogin

idSong = ReqInt("id")
If idSong <= 0 Then
    Response.Redirect AppRoot() & "songs.asp"
    Response.End
End If

Set rs = DbSelect(SongSelectSql() & "WHERE S.ID_SONG = ?", Array(idSong))
If rs.EOF Then
    rs.Close
    Response.Status = "404 Not Found"
    PageHeader "Cancion no encontrada", "songs", ""
    RenderEmpty "music", "La canci&oacute;n no existe", "No encontramos la canci&oacute;n solicitada en el cat&aacute;logo.", "songs.asp", "Ver todas las canciones"
    PageFooter
    Response.End
End If

sName    = CStr(rs("SONG_NAME").Value)
sPath    = CStr(Nz(rs("SONG_PATH").Value, ""))
sArtist  = ArtistFullName(rs("ARTIST_NAME").Value, rs("ARTIST_LASTNAME").Value)
sAlbum   = CStr(Nz(rs("ALBUM_TITLE").Value, ""))
sCover   = AssetUrl(rs("ALBUM_IMAGE").Value, IMG_FALLBACK_COVER)
nPlays   = ToInt(rs("PLAYS").Value)
idArtist = ToInt(rs("ID_ARTIST").Value)
idAlbum  = ToInt(rs("ID_ALBUM").Value)
sAudio   = SongAudioUrl(sPath)
rs.Close

PageHeader H(sName), "songs", ""
%>

<div class="detail-head">
  <div class="detail-cover"><img src="<%=HA(sCover)%>" alt="<%=HA(sName)%>"></div>
  <div class="detail-info">
    <div class="detail-kind">Canci&oacute;n</div>
    <h1 class="detail-title"><%=H(sName)%></h1>
    <div class="detail-meta">
<% If idArtist > 0 Then %>
      <a href="<%=AppRoot()%>artist.asp?id=<%=idArtist%>"><%=H(sArtist)%></a>
      <span class="dot"></span>
<% End If %>
<% If idAlbum > 0 Then %>
      <a href="<%=AppRoot()%>album.asp?id=<%=idAlbum%>"><%=H(sAlbum)%></a>
      <span class="dot"></span>
<% End If %>
      <span><%=FmtNumber(nPlays)%> reproducciones</span>
    </div>
    <div class="detail-actions">
<% If Len(sAudio) > 0 Then %>
      <div class="tracklist" data-tracklist id="singleTrack" style="display:none">
        <div class="tl-row" data-song-id="<%=idSong%>" data-song-title="<%=HA(sName)%>" data-song-artist="<%=HA(sArtist)%>"
             data-song-cover="<%=HA(sCover)%>" data-song-src="<%=HA(sAudio)%>" data-song-album="<%=HA(sAlbum)%>"
             data-song-href="<%=AppRoot()%>song.asp?id=<%=idSong%>"></div>
      </div>
      <button type="button" class="btn btn-primary btn-lg" data-play-list="#singleTrack"><%=Ico("play", 20)%> Reproducir</button>
<% Else %>
      <span class="badge badge-warn">Archivo de audio no disponible</span>
<% End If %>
      <button type="button" class="btn btn-outline" data-add-to-playlist data-song-id="<%=idSong%>"><%=Ico("plus", 18)%> Agregar a playlist</button>
<% If IsAdmin() Then %>
      <a class="btn btn-secondary" href="<%=AppRoot()%>admin/edit.asp?e=songs&amp;id=<%=idSong%>"><%=Ico("edit", 17)%> Editar</a>
      <form method="post" action="<%=AppRoot()%>admin/action.asp" class="inline-form" data-confirm="&iquest;Eliminar esta canci&oacute;n? Se quitar&aacute; de &aacute;lbumes y playlists.">
        <input type="hidden" name="entity" value="songs">
        <input type="hidden" name="action" value="delete">
        <input type="hidden" name="id" value="<%=idSong%>">
        <input type="hidden" name="returnTo" value="songs.asp">
        <button type="submit" class="btn btn-danger"><%=Ico("trash", 17)%> Eliminar</button>
      </form>
<% End If %>
    </div>
  </div>
</div>

<section class="section">
  <div class="section-head">
    <h2 class="section-title">Aparece en</h2>
  </div>

  <div class="grid grid-2">
    <div class="panel">
      <div class="panel-head"><span class="panel-title"><%=Ico("album", 17)%> &Aacute;lbumes</span></div>
      <div class="panel-body panel-body-flush">
<%
Set rs = DbSelect("SELECT A.ID_ALBUM, A.TITLE_ALBUM, A.RELEASE_YEAR, A.ALBUM_IMAGE_PATH FROM TBL_ALBUM A " & _
                  "INNER JOIN TBL_TRACKLIST T ON T.ID_ALBUM = A.ID_ALBUM WHERE T.ID_SONG = ? ORDER BY A.TITLE_ALBUM", Array(idSong))
If rs.EOF Then
    Response.Write "<p class=""muted"" style=""padding:16px 20px"">Esta canci&oacute;n no pertenece a ning&uacute;n &aacute;lbum.</p>"
Else
    Response.Write "<div class=""mini-list"" style=""max-height:none"">"
    Do While Not rs.EOF
        Response.Write "<a class=""mini-item"" href=""" & AppRoot() & "album.asp?id=" & rs("ID_ALBUM") & """>" & _
            "<img src=""" & HA(AssetUrl(rs("ALBUM_IMAGE_PATH"), IMG_FALLBACK_COVER)) & """ alt="""" loading=""lazy"">" & _
            "<span class=""mini-item-body""><span class=""mini-item-title"">" & H(rs("TITLE_ALBUM")) & "</span>" & _
            "<span class=""mini-item-sub"">" & H(Nz(rs("RELEASE_YEAR"), "Sin anio")) & "</span></span>" & _
            Ico("chevron-right", 16) & "</a>"
        rs.MoveNext
    Loop
    Response.Write "</div>"
End If
rs.Close
%>
      </div>
    </div>

    <div class="panel">
      <div class="panel-head"><span class="panel-title"><%=Ico("artist", 17)%> Artistas</span></div>
      <div class="panel-body panel-body-flush">
<%
Set rs = DbSelect("SELECT A.ID_ARTIST, A.ARTIST_NAME, A.ARTIST_LASTNAME, A.ARTIST_IMAGE FROM TBL_ARTIST A " & _
                  "INNER JOIN TBL_PLAYER P ON P.ID_ARTIST = A.ID_ARTIST WHERE P.ID_SONG = ? ORDER BY A.ARTIST_NAME", Array(idSong))
If rs.EOF Then
    Response.Write "<p class=""muted"" style=""padding:16px 20px"">Esta canci&oacute;n no tiene artistas asociados.</p>"
Else
    Response.Write "<div class=""mini-list"" style=""max-height:none"">"
    Do While Not rs.EOF
        Response.Write "<a class=""mini-item"" href=""" & AppRoot() & "artist.asp?id=" & rs("ID_ARTIST") & """>" & _
            "<img src=""" & HA(AssetUrl(rs("ARTIST_IMAGE"), IMG_FALLBACK_AVATAR)) & """ alt="""" loading=""lazy"" style=""border-radius:50%"">" & _
            "<span class=""mini-item-body""><span class=""mini-item-title"">" & H(ArtistFullName(rs("ARTIST_NAME"), rs("ARTIST_LASTNAME"))) & "</span>" & _
            "<span class=""mini-item-sub"">Artista</span></span>" & _
            Ico("chevron-right", 16) & "</a>"
        rs.MoveNext
    Loop
    Response.Write "</div>"
End If
rs.Close
%>
      </div>
    </div>
  </div>
</section>

<section class="section">
  <div class="section-head">
    <h2 class="section-title">Playlists que la incluyen</h2>
  </div>
<%
If IsAdmin() Then
    Set rs = DbSelect("SELECT P.ID_PLAYLIST, P.TITLE, P.PHOTO, P.TYPE, P.ID_USER, U.USERNAME AS OWNER_NAME " & _
                      "FROM TBL_PLAYLIST P INNER JOIN TBL_PLAYLIST_SONG PS ON PS.ID_PLAYLIST = P.ID_PLAYLIST " & _
                      "LEFT JOIN TBL_USER U ON U.ID_USER = P.ID_USER WHERE PS.ID_SONG = ? ORDER BY P.TITLE", Array(idSong))
Else
    Set rs = DbSelect("SELECT P.ID_PLAYLIST, P.TITLE, P.PHOTO, P.TYPE, P.ID_USER, U.USERNAME AS OWNER_NAME " & _
                      "FROM TBL_PLAYLIST P INNER JOIN TBL_PLAYLIST_SONG PS ON PS.ID_PLAYLIST = P.ID_PLAYLIST " & _
                      "LEFT JOIN TBL_USER U ON U.ID_USER = P.ID_USER WHERE PS.ID_SONG = ? AND (P.TYPE = ? OR P.ID_USER = ?) " & _
                      "ORDER BY P.TITLE", Array(idSong, PLAYLIST_PUBLIC, CurrentUserId()))
End If
RenderPlaylistCards rs, True
rs.Close
%>
</section>

<%
PageFooter
%>
