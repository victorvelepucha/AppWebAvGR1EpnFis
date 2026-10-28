<%
' ============================================================================
'  PoliMusic - Componentes de interfaz reutilizables
'  (listas de canciones, tarjetas, paginacion, selector de archivos)
' ============================================================================

' ============================================================================
'  SongSelectSql - consulta estandar de canciones con su album y artista
'  principal. Se reutiliza en todas las pantallas para que el render sea
'  homogeneo (mismas columnas siempre).
' ============================================================================
Function SongSelectSql()
    SongSelectSql = _
      "SELECT S.ID_SONG, S.SONG_NAME, S.SONG_PATH, ISNULL(S.PLAYS,0) AS PLAYS, " & _
      "  ISNULL(al.ID_ALBUM,0) AS ID_ALBUM, ISNULL(al.TITLE_ALBUM,'') AS ALBUM_TITLE, " & _
      "  ISNULL(al.ALBUM_IMAGE_PATH,'') AS ALBUM_IMAGE, ISNULL(al.RELEASE_YEAR,0) AS RELEASE_YEAR, " & _
      "  ISNULL(ar.ID_ARTIST,0) AS ID_ARTIST, ISNULL(ar.ARTIST_NAME,'') AS ARTIST_NAME, " & _
      "  ISNULL(ar.ARTIST_LASTNAME,'') AS ARTIST_LASTNAME, ISNULL(ar.ARTIST_IMAGE,'') AS ARTIST_IMAGE " & _
      "FROM TBL_SONG S " & _
      "OUTER APPLY (SELECT TOP 1 A.ID_ALBUM, A.TITLE_ALBUM, A.ALBUM_IMAGE_PATH, A.RELEASE_YEAR " & _
      "             FROM TBL_TRACKLIST T INNER JOIN TBL_ALBUM A ON A.ID_ALBUM = T.ID_ALBUM " & _
      "             WHERE T.ID_SONG = S.ID_SONG ORDER BY A.ID_ALBUM) al " & _
      "OUTER APPLY (SELECT TOP 1 AR.ID_ARTIST, AR.ARTIST_NAME, AR.ARTIST_LASTNAME, AR.ARTIST_IMAGE " & _
      "             FROM TBL_PLAYER P INNER JOIN TBL_ARTIST AR ON AR.ID_ARTIST = P.ID_ARTIST " & _
      "             WHERE P.ID_SONG = S.ID_SONG ORDER BY AR.ID_ARTIST) ar "
End Function

' Valor de una opcion dentro de un Dictionary de opciones
Function Opt(opts, key, def)
    If IsObject(opts) Then
        If opts.Exists(key) Then
            Opt = opts(key)
            Exit Function
        End If
    End If
    Opt = def
End Function

' URL del archivo de audio; cadena vacia si el recurso no existe en disco
Function SongAudioUrl(dbPath)
    Dim p
    p = Trim(Nz(dbPath, ""))
    p = Replace(p, "\", "/")
    If Len(p) = 0 Then
        SongAudioUrl = ""
        Exit Function
    End If
    If Left(p, 2) = "~/" Then p = Mid(p, 3)
    If Left(p, 1) = "/" Then p = Mid(p, 2)
    If Left(p, Len(ASSET_FOLDER) + 1) <> ASSET_FOLDER & "/" Then p = ASSET_FOLDER & "/" & p
    If Not AssetFileExists(p) Then
        SongAudioUrl = ""
    Else
        SongAudioUrl = AppRoot() & p
    End If
End Function

' Caratula de una cancion: usa la del album y si no existe, la de respaldo
Function SongCoverFromRow(rs)
    Dim v
    v = AssetUrl(rs("ALBUM_IMAGE"), IMG_FALLBACK_COVER)
    SongCoverFromRow = v
End Function

' ============================================================================
'  RenderTracklist - pinta una lista de canciones interactiva.
'  rs   : RecordSet con las columnas de SongSelectSql
'  opts : Dictionary de opciones
'           showAlbum, showPlays, showIndex, startIndex, showAdd,
'           playlistId, emptyText, containerId
' ============================================================================
Sub RenderTracklist(rs, opts)
    Dim bAlbum, bPlays, bIndex, bAdd, idPl, sEmpty, sCid, n, sArtist, sAudio, bAvail
    bAlbum = CBool(Opt(opts, "showAlbum", True))
    bPlays = CBool(Opt(opts, "showPlays", True))
    bIndex = CBool(Opt(opts, "showIndex", True))
    bAdd   = CBool(Opt(opts, "showAdd", False))
    idPl   = ToInt(Opt(opts, "playlistId", 0))
    sEmpty = Opt(opts, "emptyText", "No hay canciones para mostrar.")
    sCid   = Opt(opts, "containerId", "tracklist")
    n      = ToInt(Opt(opts, "startIndex", 1))

    If rs.EOF Then
        Response.Write "<div class=""empty""><div class=""empty-ico"">" & Ico("music", 28) & _
                       "</div><div class=""empty-title"">Sin resultados</div>" & _
                       "<p class=""empty-text"">" & H(sEmpty) & "</p></div>"
        Exit Sub
    End If

    Response.Write "<div class=""tracklist"" id=""" & H(sCid) & """ data-tracklist>"
    Response.Write "<div class=""tl-head"">" & _
                   "<div class=""tl-num"">#</div><div></div>" & _
                   "<div>Titulo</div>" & _
                   "<div>" & IIf(bAlbum, "Album", "Artista") & "</div>" & _
                   "<div class=""tl-num"">" & IIf(bPlays, "Reproducciones", "") & "</div>" & _
                   "<div class=""tl-actions""></div></div>"

    Do While Not rs.EOF
        sArtist = ArtistFullName(rs("ARTIST_NAME"), rs("ARTIST_LASTNAME"))
        sAudio  = SongAudioUrl(rs("SONG_PATH"))
        bAvail  = (Len(sAudio) > 0)

        Response.Write "<div class=""tl-row" & IIf(bAvail, "", " is-unavailable") & """ " & _
             "data-song-id=""" & rs("ID_SONG") & """ " & _
             "data-song-title=""" & HA(rs("SONG_NAME")) & """ " & _
             "data-song-artist=""" & HA(sArtist) & """ " & _
             "data-song-cover=""" & HA(SongCoverFromRow(rs)) & """ " & _
             "data-song-album=""" & HA(rs("ALBUM_TITLE")) & """ " & _
             "data-song-src=""" & HA(sAudio) & """ " & _
             "data-song-href=""" & AppRoot() & "song.asp?id=" & rs("ID_SONG") & """>"

        ' --- numero / boton de reproduccion
        Response.Write "<div class=""tl-num""><span class=""tl-num-txt"">" & n & "</span>"
        If bAvail Then
            Response.Write "<button type=""button"" class=""icon-btn tl-num-btn"" data-play-song aria-label=""Reproducir " & _
                           HA(rs("SONG_NAME")) & """>" & Ico("play", 15) & "</button>"
        End If
        Response.Write "</div>"

        ' --- caratula
        Response.Write "<img class=""tl-cover"" src=""" & HA(SongCoverFromRow(rs)) & """ alt=""" & HA(rs("SONG_NAME")) & """ loading=""lazy"">"

        ' --- titulo + artista
        Response.Write "<div class=""tl-main""><a class=""tl-name"" href=""" & AppRoot() & "song.asp?id=" & rs("ID_SONG") & """>" & _
                       H(rs("SONG_NAME")) & "</a>"
        If Not bAvail Then
            Response.Write "<span class=""tl-artist""><span class=""badge badge-warn"">Archivo no disponible</span></span>"
        ElseIf ToInt(rs("ID_ARTIST")) > 0 Then
            Response.Write "<a class=""tl-artist"" href=""" & AppRoot() & "artist.asp?id=" & rs("ID_ARTIST") & """>" & H(sArtist) & "</a>"
        Else
            Response.Write "<span class=""tl-artist"">" & H(sArtist) & "</span>"
        End If
        Response.Write "</div>"

        ' --- album (o artista si se pidio lo contrario)
        If bAlbum Then
            If ToInt(rs("ID_ALBUM")) > 0 Then
                Response.Write "<a class=""tl-album"" href=""" & AppRoot() & "album.asp?id=" & rs("ID_ALBUM") & """>" & H(rs("ALBUM_TITLE")) & "</a>"
            Else
                Response.Write "<span class=""tl-album"">-</span>"
            End If
        Else
            Response.Write "<span class=""tl-album"">" & H(sArtist) & "</span>"
        End If

        ' --- reproducciones
        Response.Write "<div class=""tl-plays"">" & IIf(bPlays, FmtNumber(rs("PLAYS")), "") & "</div>"

        ' --- acciones
        Response.Write "<div class=""tl-actions"">"
        If bAdd Then
            Response.Write "<button type=""button"" class=""icon-btn"" data-add-to-playlist data-song-id=""" & rs("ID_SONG") & _
                           """ title=""Agregar a una playlist"">" & Ico("plus", 16) & "</button>"
        End If
        If idPl > 0 Then
            Response.Write "<form method=""post"" action=""" & AppRoot() & "playlist.asp"" class=""inline-form"" " & _
                           "data-confirm=""&iquest;Quitar esta cancion de la playlist?"">" & _
                           "<input type=""hidden"" name=""action"" value=""removesong"">" & _
                           "<input type=""hidden"" name=""id"" value=""" & idPl & """>" & _
                           "<input type=""hidden"" name=""songId"" value=""" & rs("ID_SONG") & """>" & _
                           "<button type=""submit"" class=""icon-btn"" title=""Quitar de la playlist"">" & Ico("trash", 16) & "</button></form>"
        End If
        If bAvail Then
            Response.Write "<button type=""button"" class=""icon-btn"" data-play-song title=""Reproducir"">" & Ico("play", 16) & "</button>"
        End If
        Response.Write "</div>"

        Response.Write "</div>"
        n = n + 1
        rs.MoveNext
    Loop

    Response.Write "</div>"
End Sub

' ============================================================================
'  Botones "Reproducir" de una lista (selector CSS del contenedor)
' ============================================================================
Sub PlayAllButtons(selector, sLabel)
    Response.Write "<button type=""button"" class=""btn btn-primary btn-lg"" data-play-list=""#" & H(selector) & """>" & _
                   Ico("play", 20) & "<span>" & H(IIf(Len(sLabel) > 0, sLabel, "Reproducir")) & "</span></button>"
End Sub

' ============================================================================
'  RenderAlbumCards - tarjetas de albumes
' ============================================================================
Sub RenderAlbumCards(rs)
    Dim cnt
    If rs.EOF Then
        Response.Write "<p class=""muted"">Todavia no hay albumes registrados.</p>"
        Exit Sub
    End If
    Response.Write "<div class=""grid grid-cards"">"
    Do While Not rs.EOF
        cnt = DbCount("SELECT COUNT(*) FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array(rs("ID_ALBUM")))
        Response.Write "<article class=""card"">" & _
            "<a class=""card-link"" href=""" & AppRoot() & "album.asp?id=" & rs("ID_ALBUM") & """ aria-label=""" & HA(rs("TITLE_ALBUM")) & """></a>" & _
            "<div class=""card-media""><img src=""" & HA(AssetUrl(rs("ALBUM_IMAGE_PATH"), IMG_FALLBACK_COVER)) & """ alt=""" & HA(rs("TITLE_ALBUM")) & """ loading=""lazy"">" & _
            "<a class=""play-fab card-play"" href=""" & AppRoot() & "album.asp?id=" & rs("ID_ALBUM") & "&amp;play=1"" title=""Reproducir album"">" & Ico("play", 20) & "</a></div>" & _
            "<h3 class=""card-title"">" & H(rs("TITLE_ALBUM")) & "</h3>" & _
            "<p class=""card-sub"">" & H(Nz(rs("RELEASE_YEAR"), "-")) & " &middot; " & cnt & " canciones</p>" & _
            "</article>"
        rs.MoveNext
    Loop
    Response.Write "</div>"
End Sub

' ============================================================================
'  RenderArtistCards - tarjetas circulares de artistas
' ============================================================================
Sub RenderArtistCards(rs)
    Dim cnt, sName
    If rs.EOF Then
        Response.Write "<p class=""muted"">Todavia no hay artistas registrados.</p>"
        Exit Sub
    End If
    Response.Write "<div class=""grid grid-cards"">"
    Do While Not rs.EOF
        cnt = DbCount("SELECT COUNT(*) FROM TBL_PLAYER WHERE ID_ARTIST = ?", Array(rs("ID_ARTIST")))
        sName = ArtistFullName(rs("ARTIST_NAME"), rs("ARTIST_LASTNAME"))
        Response.Write "<article class=""card card-round"">" & _
            "<a class=""card-link"" href=""" & AppRoot() & "artist.asp?id=" & rs("ID_ARTIST") & """ aria-label=""" & HA(sName) & """></a>" & _
            "<div class=""card-media""><img src=""" & HA(AssetUrl(rs("ARTIST_IMAGE"), IMG_FALLBACK_AVATAR)) & """ alt=""" & HA(sName) & """ loading=""lazy""></div>" & _
            "<h3 class=""card-title center"">" & H(sName) & "</h3>" & _
            "<p class=""card-sub center"">Artista &middot; " & Plural(cnt, "canci&oacute;n", "canciones") & "</p>" & _
            "</article>"
        rs.MoveNext
    Loop
    Response.Write "</div>"
End Sub

' ============================================================================
'  RenderPlaylistCards - tarjetas de playlists
'  bMine = True muestra tambien el propietario (la consulta debe incluir la
'  columna OWNER_NAME). OJO: VBScript evalua las dos ramas de IIf, por lo que
'  la lectura de columnas opcionales se hace con estructuras If normales.
' ============================================================================
Sub RenderPlaylistCards(rs, bMine)
    Dim cnt, sType, sOwner, sSub
    If rs.EOF Then
        Response.Write "<p class=""muted"">No hay playlists para mostrar.</p>"
        Exit Sub
    End If
    Response.Write "<div class=""grid grid-cards"">"
    Do While Not rs.EOF
        cnt = DbCount("SELECT COUNT(*) FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?", Array(rs("ID_PLAYLIST")))
        sType = IIf(ToInt(rs("TYPE")) = PLAYLIST_PRIVATE, "Privada", "Publica")

        sOwner = ""
        If bMine Then
            On Error Resume Next
            sOwner = CStr(Nz(rs("OWNER_NAME").Value, ""))
            If Err.Number <> 0 Then
                Err.Clear
                sOwner = ""
            End If
            On Error GoTo 0
        End If

        sSub = H(sType) & " &middot; " & Plural(cnt, "canci&oacute;n", "canciones")
        If Len(sOwner) > 0 Then sSub = sSub & " &middot; " & H(sOwner)

        Response.Write "<article class=""card"">" & _
            "<a class=""card-link"" href=""" & AppRoot() & "playlist.asp?id=" & rs("ID_PLAYLIST") & """ aria-label=""" & HA(rs("TITLE")) & """></a>" & _
            "<div class=""card-media""><img src=""" & HA(AssetUrl(rs("PHOTO"), IMG_FALLBACK_COVER)) & """ alt=""" & HA(rs("TITLE")) & """ loading=""lazy"">" & _
            "<a class=""play-fab card-play"" href=""" & AppRoot() & "playlist.asp?id=" & rs("ID_PLAYLIST") & "&amp;play=1"" title=""Reproducir playlist"">" & Ico("play", 20) & "</a></div>" & _
            "<h3 class=""card-title"">" & H(rs("TITLE")) & "</h3>" & _
            "<p class=""card-sub"">" & sSub & "</p>" & _
            "</article>"
        rs.MoveNext
    Loop
    Response.Write "</div>"
End Sub

' ============================================================================
'  RenderPager - paginacion numerica
'  baseUrl : pagina destino, p.ej. "songs.asp"
'  extra   : Array de pares clave/valor que se conservan en los enlaces
' ============================================================================
Sub RenderPager(total, page, pageSize, baseUrl, extra)
    Dim pages, i, fromPg, toPg, arr, k
    total = ToInt(total) : page = ToInt(page) : pageSize = ToInt(pageSize)
    If pageSize < 1 Then pageSize = ROWS_PER_PAGE
    If page < 1 Then page = 1
    If total <= pageSize Then Exit Sub
    pages = Int((total + pageSize - 1) / pageSize)

    arr = Array()
    If IsArray(extra) Then arr = extra

    Response.Write "<nav class=""pager"" aria-label=""Paginacion"">"

    ' Anterior
    If page > 1 Then
        Response.Write PagerLink(baseUrl, PageParams(arr, page - 1), "Anterior", "")
    Else
        Response.Write "<span class=""pager-item is-disabled"">Anterior</span>"
    End If

    fromPg = page - 2 : If fromPg < 1 Then fromPg = 1
    toPg = fromPg + 4 : If toPg > pages Then toPg = pages
    fromPg = toPg - 4 : If fromPg < 1 Then fromPg = 1

    If fromPg > 1 Then
        Response.Write PagerLink(baseUrl, PageParams(arr, 1), "1", "")
        If fromPg > 2 Then Response.Write "<span class=""pager-item is-disabled"">...</span>"
    End If

    For i = fromPg To toPg
        Response.Write PagerLink(baseUrl, PageParams(arr, i), CStr(i), IIf(i = page, "is-active", ""))
    Next

    If toPg < pages Then
        If toPg < pages - 1 Then Response.Write "<span class=""pager-item is-disabled"">...</span>"
        Response.Write PagerLink(baseUrl, PageParams(arr, pages), CStr(pages), "")
    End If

    ' Siguiente
    If page < pages Then
        Response.Write PagerLink(baseUrl, PageParams(arr, page + 1), "Siguiente", "")
    Else
        Response.Write "<span class=""pager-item is-disabled"">Siguiente</span>"
    End If

    Response.Write "<span class=""pager-info"">Pagina " & page & " de " & pages & _
                   " &middot; " & total & " registros</span>"
    Response.Write "</nav>"
End Sub

' Construye la cadena de parametros conservando los extras y fijando la pagina
Function PageParams(extra, page)
    Dim i, s
    s = "page=" & page
    If IsArray(extra) Then
        For i = 0 To UBound(extra) Step 2
            If Len(CStr(extra(i + 1))) > 0 Then
                s = s & "&" & CStr(extra(i)) & "=" & Server.URLEncode(CStr(extra(i + 1)))
            End If
        Next
    End If
    PageParams = s
End Function

Function PagerLink(baseUrl, query, label, cls)
    PagerLink = "<a class=""pager-item " & cls & """ href=""" & AppRoot() & H(baseUrl) & "?" & _
                Replace(H(query), "'", "&#39;") & """>" & H(label) & "</a>"
End Function

' ============================================================================
'  ListAssetFiles - enumera los archivos disponibles en una carpeta de assets
' ============================================================================
Function ListAssetFiles(subFolder)
    Dim fso, fld, f, arr, i, n
    arr = Array()
    On Error Resume Next
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(Server.MapPath("/" & ASSET_FOLDER & "/" & subFolder)) Then
        Err.Clear
        On Error GoTo 0
        ListAssetFiles = arr
        Exit Function
    End If
    Set fld = fso.GetFolder(Server.MapPath("/" & ASSET_FOLDER & "/" & subFolder))
    n = fld.Files.Count
    If Err.Number <> 0 Then Err.Clear
    If n > 0 Then
        ReDim arr(n - 1)
        i = 0
        For Each f In fld.Files
            arr(i) = f.Name
            i = i + 1
        Next
    End If
    On Error GoTo 0
    ListAssetFiles = arr
End Function

' ============================================================================
'  RenderFilePicker - campo de texto + selector visual de archivos existentes
'  name       : nombre del campo del formulario
'  subFolder  : subcarpeta dentro de assets (songFiles, albumImages, ...)
'  value      : valor actual en formato "~/subFolder/archivo.ext"
'  sLabel     : etiqueta del campo
'  sKind      : "audio" | "image"  (para el tipo de vista previa)
' ============================================================================
Sub RenderFilePicker(name, subFolder, value, sLabel, sKind)
    Dim arr, i, files, sVal, bAudio
    arr = ListAssetFiles(subFolder)
    bAudio = (LCase(sKind) = "audio")
    sVal = Nz(value, "")
    files = ""
    For i = 0 To UBound(arr)
        files = files & "<button type=""button"" class=""pick-item"" data-pick=""" & HA("~/" & subFolder & "/" & arr(i)) & """>" & _
                IIf(bAudio, Ico("music", 18), "<img src=""" & HA(AppRoot() & ASSET_FOLDER & "/" & subFolder & "/" & arr(i)) & """ alt="""" loading=""lazy"">") & _
                "<span>" & H(arr(i)) & "</span></button>"
    Next
%>
<div class="field field-full">
  <label class="label" for="<%=H(name)%>"><%=H(sLabel)%></label>
  <div class="hstack-tight">
    <input class="input" type="text" id="<%=H(name)%>" name="<%=H(name)%>" value="<%=HA(sVal)%>"
           placeholder="<%=IIf(bAudio,"~/songFiles/archivo.mp3","~/albumImages/portada.png")%>" data-file-input="<%=H(name)%>">
    <button type="button" class="btn btn-secondary btn-sm" data-modal-open="modalPick_<%=H(name)%>"><%=Ico("upload", 16)%> Elegir</button>
  </div>
  <span class="hint">Ruta relativa al directorio de recursos. Ejemplo: <code><%=IIf(bAudio,"~/songFiles/cancion.mp3","~/albumImages/portada.png")%></code></span>
</div>

<div class="modal" id="modalPick_<%=H(name)%>">
  <div class="modal-box modal-box-lg">
    <div class="modal-head">
      <h3 class="modal-title">Seleccionar archivo &middot; <%=H(subFolder)%></h3>
      <button type="button" class="icon-btn" data-modal-close aria-label="Cerrar"><%=Ico("close", 18)%></button>
    </div>
    <div class="modal-body">
<% If Len(files) = 0 Then %>
      <p class="muted">No hay archivos en la carpeta <code>assets/<%=H(subFolder)%></code>.</p>
<% Else %>
      <div class="pick-grid"><%=files%></div>
<% End If %>
    </div>
    <div class="modal-foot">
      <button type="button" class="btn btn-secondary" data-modal-close>Cerrar</button>
    </div>
  </div>
</div>
<%
End Sub

' ============================================================================
'  RenderStat - tarjeta de indicador para el panel de administracion
'  NOTA: los textos se insertan tal cual (HTML ya escapado en el origen).
' ============================================================================
Sub RenderStat(iconName, sValue, sLabel)
    Response.Write "<div class=""stat""><div class=""stat-ico"">" & Ico(iconName, 20) & "</div>" & _
                   "<div class=""stat-value"">" & sValue & "</div>" & _
                   "<div class=""stat-label"">" & sLabel & "</div></div>"
End Sub

' ============================================================================
'  RenderEmpty - bloque de estado vacio
'  NOTA: los textos se insertan tal cual (HTML ya escapado en el origen).
' ============================================================================
Sub RenderEmpty(iconName, sTitle, sText, sBtnHref, sBtnLabel)
    Response.Write "<div class=""empty""><div class=""empty-ico"">" & Ico(iconName, 28) & "</div>" & _
                   "<div class=""empty-title"">" & sTitle & "</div>" & _
                   "<p class=""empty-text"">" & sText & "</p>"
    If Len(sBtnHref) > 0 Then
        Response.Write "<a class=""btn btn-primary"" href=""" & AppRoot() & H(sBtnHref) & """>" & Ico("plus", 17) & " " & sBtnLabel & "</a>"
    End If
    Response.Write "</div>"
End Sub
%>
