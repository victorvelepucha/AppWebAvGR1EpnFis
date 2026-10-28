<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' ============================================================================
'  playlist.asp - Gestion de playlists
'  - El usuario NORMAL puede consultar, crear, actualizar y eliminar SUS
'    propias playlists, y agregar / quitar canciones.
'  - El ADMINISTRADOR puede gestionar cualquier playlist.
'  - Las playlists PUBLICAS son visibles para todos los usuarios.
' ============================================================================
Dim sAction, idPl, rs, bCanManage, sTitle, sType, sPhoto, sErr, sReturnTo
Dim nSongs, o, sMode

RequireLogin

sAction = LCase(Req("action"))
idPl    = ReqInt("id")

' ---------------------------------------------------------------- MUTACIONES
If IsPost() Then
    Select Case sAction

        ' ---------------------------------------------------------- crear
        Case "create"
            sTitle = Req("title")
            sType  = Req("type")
            sPhoto = Req("photo")
            If Len(sPhoto) = 0 Then sPhoto = "~/playlistImages/default.jpg"

            If Len(sTitle) = 0 Or Len(sTitle) > 50 Then
                FlashSet "error", "El titulo de la playlist es obligatorio (maximo 50 caracteres)."
                Response.Redirect AppRoot() & "playlist.asp?action=new"
                Response.End
            End If
            If Len(sType) = 0 Then sType = CStr(PLAYLIST_PUBLIC)

            On Error Resume Next
            idPl = ToInt(DbExecuteScalar("SET NOCOUNT ON; INSERT INTO TBL_PLAYLIST (ID_USER, TITLE, CREATION_DATE, TYPE, PHOTO) " & _
                      "VALUES (?, ?, GETDATE(), CONVERT(int, ?), ?); SELECT CAST(SCOPE_IDENTITY() AS int) AS NEWID;", _
                      Array(CurrentUserId(), sTitle, sType, sPhoto)))
            If Err.Number <> 0 Then
                FlashSet "error", "No se pudo crear la playlist: " & Err.Description
                Err.Clear
                Response.Redirect AppRoot() & "playlist.asp?action=new"
                Response.End
            End If
            On Error GoTo 0

            If idPl <= 0 Then
                idPl = ToInt(DbScalar("SELECT MAX(ID_PLAYLIST) FROM TBL_PLAYLIST WHERE ID_USER = ?", Array(CurrentUserId())))
            End If
            FlashSet "ok", "Playlist """ & sTitle & """ creada correctamente."
            Response.Redirect AppRoot() & "playlist.asp?id=" & idPl
            Response.End

        ' -------------------------------------------------------- actualizar
        Case "update"
            If Not CanManagePlaylist(idPl) Then
                FlashSet "error", "No tienes permiso para modificar esta playlist."
                Response.Redirect AppRoot() & "playlist.asp?id=" & idPl
                Response.End
            End If

            sTitle = Req("title")
            sType  = Req("type")
            sPhoto = Req("photo")
            If Len(sPhoto) = 0 Then sPhoto = "~/playlistImages/default.jpg"

            If Len(sTitle) = 0 Or Len(sTitle) > 50 Then
                FlashSet "error", "El titulo de la playlist es obligatorio (maximo 50 caracteres)."
                Response.Redirect AppRoot() & "playlist.asp?id=" & idPl & "&action=edit"
                Response.End
            End If

            DbExecute "UPDATE TBL_PLAYLIST SET TITLE = ?, TYPE = CONVERT(int, ?), PHOTO = ? WHERE ID_PLAYLIST = ?", _
                      Array(sTitle, sType, sPhoto, idPl)
            FlashSet "ok", "Playlist actualizada correctamente."
            Response.Redirect AppRoot() & "playlist.asp?id=" & idPl
            Response.End

        ' ----------------------------------------------------------- eliminar
        Case "delete"
            If Not CanManagePlaylist(idPl) Then
                FlashSet "error", "No tienes permiso para eliminar esta playlist."
                Response.Redirect AppRoot() & "library.asp"
                Response.End
            End If
            DbExecute "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?", Array(idPl)
            DbExecute "DELETE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(idPl)
            FlashSet "ok", "La playlist fue eliminada."
            Response.Redirect AppRoot() & "library.asp"
            Response.End

        ' -------------------------------------------------- agregar cancion
        Case "addsong"
            Dim idSong : idSong = ReqInt("songId")
            If Not CanManagePlaylist(idPl) Then
                FlashSet "error", "No tienes permiso para modificar esta playlist."
                Response.Redirect AppRoot() & "playlist.asp?id=" & idPl
                Response.End
            End If
            If idSong > 0 Then
                If DbExists("SELECT 1 FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ? AND ID_SONG = ?", Array(idPl, idSong)) Then
                    FlashSet "warn", "La cancion ya estaba en la playlist."
                Else
                    DbExecute "INSERT INTO TBL_PLAYLIST_SONG (ID_PLAYLIST, ID_SONG) VALUES (?, ?)", Array(idPl, idSong)
                    FlashSet "ok", "Cancion agregada a la playlist."
                End If
            End If
            sReturnTo = Req("returnTo")
            If Len(sReturnTo) = 0 Then sReturnTo = "playlist.asp?id=" & idPl
            Response.Redirect AppRoot() & sReturnTo
            Response.End

        ' ----------------------------------------------------- quitar cancion
        Case "removesong"
            Dim idSong2 : idSong2 = ReqInt("songId")
            If Not CanManagePlaylist(idPl) Then
                FlashSet "error", "No tienes permiso para modificar esta playlist."
                Response.Redirect AppRoot() & "playlist.asp?id=" & idPl
                Response.End
            End If
            DbExecute "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ? AND ID_SONG = ?", Array(idPl, idSong2)
            FlashSet "ok", "Cancion quitada de la playlist."
            Response.Redirect AppRoot() & "playlist.asp?id=" & idPl
            Response.End

    End Select
End If

' ============================ VISTA: CREAR / EDITAR ============================
sMode = ""
If sAction = "new" Then
    sMode = "create"
ElseIf sAction = "edit" And idPl > 0 Then
    If Not CanManagePlaylist(idPl) Then
        FlashSet "error", "No tienes permiso para editar esta playlist."
        Response.Redirect AppRoot() & "playlist.asp?id=" & idPl
        Response.End
    End If
    sMode = "edit"
End If

If Len(sMode) > 0 Then
    If sMode = "edit" Then
        Set rs = DbSelect("SELECT ID_PLAYLIST, TITLE, TYPE, PHOTO, CREATION_DATE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(idPl))
        If rs.EOF Then
            rs.Close
            Response.Redirect AppRoot() & "library.asp"
            Response.End
        End If
        sTitle = CStr(rs("TITLE").Value)
        sType  = CStr(ToInt(rs("TYPE").Value))
        sPhoto = CStr(Nz(rs("PHOTO").Value, "~/playlistImages/default.jpg"))
        rs.Close
    Else
        sTitle = ""
        sType  = CStr(PLAYLIST_PUBLIC)
        sPhoto = "~/playlistImages/default.jpg"
    End If

    PageHeader IIf(sMode = "create", "Nueva playlist", "Editar playlist"), "library", ""
%>

<div class="page-head">
  <div>
    <h1 class="page-title"><%=IIf(sMode = "create", "Nueva playlist", "Editar playlist")%></h1>
    <p class="page-sub">Las playlists <strong>p&uacute;blicas</strong> las puede ver todo el mundo; las <strong>privadas</strong> solo t&uacute;.</p>
  </div>
  <a class="btn btn-ghost" href="<%=AppRoot()%>library.asp"><%=Ico("chevron-left", 18)%> Volver</a>
</div>

<form method="post" action="<%=AppRoot()%>playlist.asp" class="panel">
  <input type="hidden" name="action" value="<%=IIf(sMode = "create", "create", "update")%>">
<% If sMode = "edit" Then %>
  <input type="hidden" name="id" value="<%=idPl%>">
<% End If %>
  <div class="panel-head"><span class="panel-title"><%=Ico("playlist", 17)%> Datos de la playlist</span></div>
  <div class="panel-body">
    <div class="form-grid">
      <div class="field field-full">
        <label class="label" for="title">T&iacute;tulo <span class="req">*</span></label>
        <input class="input" type="text" id="title" name="title" maxlength="50" value="<%=HA(sTitle)%>"
               placeholder="Mi playlist favorita" required autofocus>
        <span class="hint">M&aacute;ximo 50 caracteres.</span>
      </div>

<% RenderFilePicker "photo", "playlistImages", sPhoto, "Imagen de la playlist", "image" %>

      <div class="field field-full">
        <span class="label">Visibilidad <span class="req">*</span></span>
        <div class="radio-cards">
          <label class="radio-card<%=IIf(sType = CStr(PLAYLIST_PUBLIC), " is-checked", "")%>">
            <input type="radio" name="type" value="<%=PLAYLIST_PUBLIC%>" <%=IIf(sType = CStr(PLAYLIST_PUBLIC), "checked", "")%>>
            <span><span class="radio-card-title"><%=Ico("globe", 15)%> P&uacute;blica</span>
            <span class="radio-card-sub">Cualquier usuario puede verla y reproducirla.</span></span>
          </label>
          <label class="radio-card<%=IIf(sType = CStr(PLAYLIST_PRIVATE), " is-checked", "")%>">
            <input type="radio" name="type" value="<%=PLAYLIST_PRIVATE%>" <%=IIf(sType = CStr(PLAYLIST_PRIVATE), "checked", "")%>>
            <span><span class="radio-card-title"><%=Ico("lock", 15)%> Privada</span>
            <span class="radio-card-sub">Solo t&uacute; puedes verla.</span></span>
          </label>
        </div>
      </div>
    </div>
  </div>
  <div class="panel-body" style="border-top:1px solid var(--line)">
    <div class="form-actions" style="margin:0;padding:0;border:none">
      <button type="submit" class="btn btn-primary btn-lg"><%=Ico("check", 18)%> <%=IIf(sMode = "create", "Crear playlist", "Guardar cambios")%></button>
      <a class="btn btn-ghost btn-lg" href="<%=IIf(sMode = "edit", AppRoot() & "playlist.asp?id=" & idPl, AppRoot() & "library.asp")%>">Cancelar</a>
    </div>
  </div>
</form>

<%
    PageFooter
    Response.End
End If

' ================================ VISTA: DETALLE ================================
If idPl <= 0 Then
    Response.Redirect AppRoot() & "library.asp"
    Response.End
End If

Set rs = DbSelect("SELECT P.ID_PLAYLIST, P.ID_USER, P.TITLE, P.CREATION_DATE, P.TYPE, P.PHOTO, U.USERNAME AS OWNER_NAME " & _
                  "FROM TBL_PLAYLIST P LEFT JOIN TBL_USER U ON U.ID_USER = P.ID_USER WHERE P.ID_PLAYLIST = ?", Array(idPl))
If rs.EOF Then
    rs.Close
    Response.Status = "404 Not Found"
    PageHeader "Playlist no encontrada", "library", ""
    RenderEmpty "playlist", "La playlist no existe", "No encontramos la playlist solicitada.", "library.asp", "Ir a la biblioteca"
    PageFooter
    Response.End
End If

sTitle      = CStr(rs("TITLE").Value)
sType       = CStr(ToInt(rs("TYPE").Value))
sPhoto      = CStr(Nz(rs("PHOTO").Value, "~/playlistImages/default.jpg"))
bCanManage  = CanManagePlaylist(idPl)
Dim sOwner  : sOwner = CStr(Nz(rs("OWNER_NAME").Value, "desconocido"))
Dim dCreated: dCreated = rs("CREATION_DATE").Value
Dim idOwner : idOwner = ToInt(rs("ID_USER").Value)
rs.Close

' Privada y ajena -> no se muestra
If ToInt(sType) = PLAYLIST_PRIVATE And Not bCanManage Then
    Response.Status = "403 Forbidden"
    PageHeader "Playlist privada", "library", ""
    RenderEmpty "lock", "Esta playlist es privada", "Solo el propietario puede ver el contenido de esta playlist.", "library.asp", "Volver a la biblioteca"
    PageFooter
    Response.End
End If

nSongs = DbCount("SELECT COUNT(*) FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?", Array(idPl))

PageHeader H(sTitle), "library", ""
%>

<div class="detail-head">
  <div class="detail-cover"><img src="<%=HA(AssetUrl(sPhoto, IMG_FALLBACK_COVER))%>" alt="<%=HA(sTitle)%>"></div>
  <div class="detail-info">
    <div class="detail-kind"><%=IIf(ToInt(sType) = PLAYLIST_PRIVATE, "Playlist privada", "Playlist p&uacute;blica")%></div>
    <h1 class="detail-title"><%=H(sTitle)%></h1>
    <div class="detail-meta">
      <span><%=Ico("user", 15)%> <%=H(sOwner)%></span>
      <span class="dot"></span><span><%=Plural(nSongs, "canci&oacute;n", "canciones")%></span>
      <span class="dot"></span><span>Creada el <%=H(FmtDate(dCreated))%></span>
    </div>
    <div class="detail-actions">
<% If nSongs > 0 Then %>
      <button type="button" class="btn btn-primary btn-lg" data-play-list="#playlistTracks"><%=Ico("play", 20)%> Reproducir</button>
      <button type="button" class="btn btn-outline" data-play-shuffle="#playlistTracks"><%=Ico("shuffle", 18)%> Aleatorio</button>
<% End If %>
<% If bCanManage Then %>
      <button type="button" class="btn btn-secondary" data-modal-open="modalAddSongs"><%=Ico("plus", 17)%> Agregar canciones</button>
      <a class="btn btn-secondary" href="<%=AppRoot()%>playlist.asp?id=<%=idPl%>&amp;action=edit"><%=Ico("edit", 17)%> Editar</a>
      <form method="post" action="<%=AppRoot()%>playlist.asp" class="inline-form" data-confirm="&iquest;Eliminar la playlist &quot;<%=HA(sTitle)%>&quot;? Esta accion no se puede deshacer.">
        <input type="hidden" name="action" value="delete">
        <input type="hidden" name="id" value="<%=idPl%>">
        <button type="submit" class="btn btn-danger"><%=Ico("trash", 17)%> Eliminar</button>
      </form>
<% End If %>
    </div>
  </div>
</div>

<section class="section">
  <div class="section-head">
    <h2 class="section-title">Canciones</h2>
    <span class="muted small"><%=Plural(nSongs, "pista", "pistas")%></span>
  </div>
<%
Set rs = DbSelect(SongSelectSql() & "WHERE EXISTS (SELECT 1 FROM TBL_PLAYLIST_SONG PS WHERE PS.ID_SONG = S.ID_SONG AND PS.ID_PLAYLIST = ?) " & _
                  "ORDER BY S.SONG_NAME", Array(idPl))
Set o = Server.CreateObject("Scripting.Dictionary")
o("showAlbum")   = True
o("showPlays")   = True
o("showAdd")     = False
o("playlistId")  = IIf(bCanManage, idPl, 0)
o("containerId") = "playlistTracks"
o("emptyText")   = "Esta playlist todavia no tiene canciones. Usa el boton 'Agregar canciones'."
RenderTracklist rs, o
rs.Close
%>
</section>

<% If bCanManage Then %>
<!-- ====================== MODAL: AGREGAR CANCIONES ====================== -->
<div class="modal" id="modalAddSongs">
  <div class="modal-box modal-box-lg">
    <div class="modal-head">
      <h3 class="modal-title">Agregar canciones a &laquo;<%=H(sTitle)%>&raquo;</h3>
      <button type="button" class="icon-btn" data-modal-close aria-label="Cerrar"><%=Ico("close", 18)%></button>
    </div>
    <div class="modal-body">
      <input class="input mb-2" type="search" id="addSongsFilter" placeholder="Filtrar canciones..." autocomplete="off">
      <div class="mini-list" id="addSongsList">
<%
Set rs = DbSelect(SongSelectSql() & "ORDER BY S.SONG_NAME", Empty)
Do While Not rs.EOF
    Dim bIn : bIn = DbExists("SELECT 1 FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ? AND ID_SONG = ?", Array(idPl, rs("ID_SONG")))
%>
        <form method="post" action="<%=AppRoot()%>playlist.asp" class="mini-form" data-name="<%=HA(LCase(rs("SONG_NAME")))%>">
          <input type="hidden" name="action" value="addsong">
          <input type="hidden" name="id" value="<%=idPl%>">
          <input type="hidden" name="songId" value="<%=rs("ID_SONG")%>">
          <button type="submit" class="mini-item<%=IIf(bIn, " is-added", "")%>" <%=IIf(bIn, "disabled", "")%>>
            <img src="<%=HA(SongCoverFromRow(rs))%>" alt="" loading="lazy">
            <span class="mini-item-body">
              <span class="mini-item-title"><%=H(rs("SONG_NAME"))%></span>
              <span class="mini-item-sub"><%=H(ArtistFullName(rs("ARTIST_NAME"), rs("ARTIST_LASTNAME")))%> &middot; <%=H(Nz(rs("ALBUM_TITLE"), "Sin album"))%></span>
            </span>
<% If bIn Then %>
            <span class="badge badge-ok">Ya esta</span>
<% Else %>
            <span class="ico-add">+</span>
<% End If %>
          </button>
        </form>
<%
    rs.MoveNext
Loop
rs.Close
%>
      </div>
    </div>
    <div class="modal-foot">
      <button type="button" class="btn btn-secondary" data-modal-close>Cerrar</button>
    </div>
  </div>
</div>
<% End If %>

<%
PageFooter
%>
