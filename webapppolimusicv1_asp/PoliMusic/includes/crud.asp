<%
' ============================================================================
'  PoliMusic - Motor CRUD generico del panel de administracion
'  Define los metadatos de cada entidad (tabla, campos, validaciones y
'  dependencias) para que las paginas admin/*.asp sean genericas.
' ============================================================================

' --------------------------------------------------------------------------
'  Constructor de campos de formulario
' --------------------------------------------------------------------------
Function MkField(nm, lb, tp, rq, mx, ht)
    Dim d
    Set d = Server.CreateObject("Scripting.Dictionary")
    d("name")      = nm
    d("label")     = lb
    d("type")      = tp        ' text | number | date | select | image | audio | password
    d("required")  = rq
    d("maxlength") = mx
    d("hint")      = ht
    d("catalog")   = ""        ' "1" = TBL_CATALOG parent 1, "2" = parent 2
    d("options")   = Array()   ' Array("valor","etiqueta", ...)
    d("folder")    = ""        ' subcarpeta de assets (image / audio)
    d("readonly")  = False
    Set MkField = d
End Function

' Anade un campo al array de campos de una entidad
Function AddFld(arr, f)
    Dim n
    n = UBound(arr) + 1
    ReDim Preserve arr(n)
    Set arr(n) = f
    AddFld = arr
End Function

' --------------------------------------------------------------------------
'  Utilidades de nombre
' --------------------------------------------------------------------------
Function ValidEntity(e)
    e = LCase(Nz(e, ""))
    ValidEntity = (e = "users" Or e = "songs" Or e = "albums" Or e = "artists" Or e = "playlists")
End Function

Function EntityTitle(e)
    Select Case LCase(e)
        Case "users"     : EntityTitle = "Usuarios"
        Case "songs"     : EntityTitle = "Canciones"
        Case "albums"    : EntityTitle = "Albumes"
        Case "artists"   : EntityTitle = "Artistas"
        Case "playlists" : EntityTitle = "Playlists"
        Case Else        : EntityTitle = "Registros"
    End Select
End Function

Function EntitySingular(e)
    Select Case LCase(e)
        Case "users"     : EntitySingular = "usuario"
        Case "songs"     : EntitySingular = "cancion"
        Case "albums"    : EntitySingular = "album"
        Case "artists"   : EntitySingular = "artista"
        Case "playlists" : EntitySingular = "playlist"
        Case Else        : EntitySingular = "registro"
    End Select
End Function

Function EntityIcon(e)
    Select Case LCase(e)
        Case "users"     : EntityIcon = "users"
        Case "songs"     : EntityIcon = "music"
        Case "albums"    : EntityIcon = "album"
        Case "artists"   : EntityIcon = "artist"
        Case "playlists" : EntityIcon = "playlist"
        Case Else        : EntityIcon = "dashboard"
    End Select
End Function

' ============================================================================
'  GetEntity - definicion completa de una entidad
'  Claves: name, title, icon, table, pk, order, searchCols, fields
' ============================================================================
Function GetEntity(e)
    Dim d, f, flds

    Set d = Server.CreateObject("Scripting.Dictionary")
    d("name")       = LCase(e)
    d("title")      = EntityTitle(e)
    d("singular")   = EntitySingular(e)
    d("icon")       = EntityIcon(e)
    d("searchCols") = ""
    d("order")      = ""
    flds = Array()

    Select Case LCase(e)

    ' ============================================================== USUARIOS
    Case "users"
        d("table") = "TBL_USER"
        d("pk")    = "ID_USER"
        d("order") = "USERNAME"
        d("searchCols") = "USERNAME,EMAIL"

        flds = AddFld(flds, MkField("USERNAME", "Nombre de usuario", "text", True, 50, "Entre 3 y 50 caracteres, sin espacios."))
        flds = AddFld(flds, MkField("PASSWORD", "Contrasena", "password", False, 50, "Al editar, dejalo vacio para conservar la contrasena actual."))
        flds = AddFld(flds, MkField("EMAIL", "Correo electronico", "text", True, 50, ""))
        flds = AddFld(flds, MkField("BIRTHDAY", "Fecha de nacimiento", "date", False, 0, ""))
        Set f = MkField("USER_PHOTO", "Foto de perfil", "image", False, 255, "")
        f("folder") = "userImages"
        flds = AddFld(flds, f)
        Set f = MkField("USER_TYPE", "Rol del usuario", "select", True, 0, "Los roles se toman del catalogo (TBL_CATALOG).")
        f("catalog") = "1"
        flds = AddFld(flds, f)

    ' ============================================================= CANCIONES
    Case "songs"
        d("table") = "TBL_SONG"
        d("pk")    = "ID_SONG"
        d("order") = "SONG_NAME"
        d("searchCols") = "SONG_NAME,SONG_PATH"

        flds = AddFld(flds, MkField("SONG_NAME", "Nombre de la cancion", "text", True, 50, ""))
        Set f = MkField("SONG_PATH", "Archivo de audio", "audio", True, 255, "Ruta del archivo dentro de los recursos del sitio.")
        f("folder") = "songFiles"
        flds = AddFld(flds, f)
        flds = AddFld(flds, MkField("PLAYS", "Reproducciones", "number", False, 0, "Contador de veces que se ha reproducido."))

    ' =============================================================== ALBUMES
    Case "albums"
        d("table") = "TBL_ALBUM"
        d("pk")    = "ID_ALBUM"
        d("order") = "TITLE_ALBUM"
        d("searchCols") = "TITLE_ALBUM"

        flds = AddFld(flds, MkField("TITLE_ALBUM", "Titulo del album", "text", True, 50, ""))
        flds = AddFld(flds, MkField("RELEASE_YEAR", "Anio de lanzamiento", "number", False, 0, "Por ejemplo: 1976"))
        Set f = MkField("ALBUM_IMAGE_PATH", "Imagen de portada", "image", False, 255, "")
        f("folder") = "albumImages"
        flds = AddFld(flds, f)

    ' ============================================================== ARTISTAS
    Case "artists"
        d("table") = "TBL_ARTIST"
        d("pk")    = "ID_ARTIST"
        d("order") = "ARTIST_NAME"
        d("searchCols") = "ARTIST_NAME,ARTIST_LASTNAME"

        flds = AddFld(flds, MkField("ARTIST_NAME", "Nombre", "text", True, 100, ""))
        flds = AddFld(flds, MkField("ARTIST_LASTNAME", "Apellido", "text", False, 100, ""))
        Set f = MkField("ARTIST_IMAGE", "Fotografia", "image", False, 255, "")
        f("folder") = "artistImages"
        flds = AddFld(flds, f)

    ' ============================================================= PLAYLISTS
    Case "playlists"
        d("table") = "TBL_PLAYLIST"
        d("pk")    = "ID_PLAYLIST"
        d("order") = "ID_PLAYLIST DESC"
        d("searchCols") = "TITLE"

        flds = AddFld(flds, MkField("TITLE", "Titulo", "text", True, 50, ""))
        Set f = MkField("ID_USER", "Propietario", "select", True, 0, "")
        f("options") = OwnerOptions()
        flds = AddFld(flds, f)
        Set f = MkField("TYPE", "Visibilidad", "select", True, 0, "Los tipos se toman del catalogo (TBL_CATALOG).")
        f("catalog") = "2"
        flds = AddFld(flds, f)
        flds = AddFld(flds, MkField("CREATION_DATE", "Fecha de creacion", "date", True, 0, ""))
        Set f = MkField("PHOTO", "Imagen de la playlist", "image", False, 255, "")
        f("folder") = "playlistImages"
        flds = AddFld(flds, f)

    End Select

    d("fields") = flds
    Set GetEntity = d
End Function

' Devuelve la definicion de un campo por nombre de columna (Nothing si no existe)
Function GetField(d, colName)
    Dim arr, i, f
    Set GetField = Nothing
    arr = d("fields")
    For i = 0 To UBound(arr)
        Set f = arr(i)
        If UCase(f("name")) = UCase(colName) Then
            Set GetField = f
            Exit Function
        End If
    Next
End Function

' Lista de usuarios para el desplegable de propietario de playlist
Function OwnerOptions()
    Dim rs, arr, n
    arr = Array()
    Set rs = DbSelect("SELECT ID_USER, USERNAME FROM TBL_USER ORDER BY USERNAME", Empty)
    n = 0
    Do While Not rs.EOF
        ReDim Preserve arr(n * 2 + 1)
        arr(n * 2)     = CStr(rs("ID_USER").Value)
        arr(n * 2 + 1) = CStr(rs("USERNAME").Value)
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    If n = 0 Then arr = Array()
    OwnerOptions = arr
End Function

' ============================================================================
'  Clausulas de busqueda
' ============================================================================
Function SearchWhere(d)
    Dim cols, i, s
    If Len(Nz(d("searchCols"), "")) = 0 Then
        SearchWhere = ""
        Exit Function
    End If
    cols = Split(d("searchCols"), ",")
    s = ""
    For i = 0 To UBound(cols)
        If Len(s) > 0 Then s = s & " OR "
        s = s & Trim(cols(i)) & " LIKE ?"
    Next
    SearchWhere = "(" & s & ")"
End Function

Function SearchParams(d, sLike)
    Dim cols, arr, i
    If Len(Nz(d("searchCols"), "")) = 0 Then
        SearchParams = Array()
        Exit Function
    End If
    cols = Split(d("searchCols"), ",")
    ReDim arr(UBound(cols))
    For i = 0 To UBound(cols)
        arr(i) = sLike
    Next
    SearchParams = arr
End Function

Function EntityRowCount(d, sLike)
    If Len(Nz(sLike, "")) = 0 Then
        EntityRowCount = DbCount("SELECT COUNT(*) FROM " & d("table"), Empty)
    Else
        EntityRowCount = DbCount("SELECT COUNT(*) FROM " & d("table") & " WHERE " & SearchWhere(d), SearchParams(d, sLike))
    End If
End Function

' ============================================================================
'  Definicion de columnas para los listados del panel de administracion
' ============================================================================
Function MkCol(lb, ex, tp, al)
    Dim d
    Set d = Server.CreateObject("Scripting.Dictionary")
    d("label") = lb
    d("expr")  = ex
    d("type")  = tp     ' number | text | mono | date | image | round | user | catalog1 | catalog2
    d("align") = al
    Set MkCol = d
End Function

Function AddCol(arr, c)
    Dim n
    n = UBound(arr) + 1
    ReDim Preserve arr(n)
    Set arr(n) = c
    AddCol = arr
End Function

Function ListColumns(e)
    Dim arr
    arr = Array()
    Select Case LCase(e)
    Case "users"
        arr = AddCol(arr, MkCol("#", "ID_USER", "number", "right"))
        arr = AddCol(arr, MkCol("Usuario", "USERNAME", "user", ""))
        arr = AddCol(arr, MkCol("Correo electronico", "EMAIL", "text", ""))
        arr = AddCol(arr, MkCol("Rol", "USER_TYPE", "catalog1", ""))
        arr = AddCol(arr, MkCol("Fecha de nacimiento", "BIRTHDAY", "date", ""))
    Case "songs"
        arr = AddCol(arr, MkCol("#", "ID_SONG", "number", "right"))
        arr = AddCol(arr, MkCol("Cancion", "SONG_NAME", "song", ""))
        arr = AddCol(arr, MkCol("Archivo", "SONG_PATH", "mono", ""))
        arr = AddCol(arr, MkCol("Reproducciones", "PLAYS", "number", "right"))
    Case "albums"
        arr = AddCol(arr, MkCol("#", "ID_ALBUM", "number", "right"))
        arr = AddCol(arr, MkCol("Portada", "ALBUM_IMAGE_PATH", "image", ""))
        arr = AddCol(arr, MkCol("Titulo del album", "TITLE_ALBUM", "text", ""))
        arr = AddCol(arr, MkCol("Anio", "RELEASE_YEAR", "number", "right"))
    Case "artists"
        arr = AddCol(arr, MkCol("#", "ID_ARTIST", "number", "right"))
        arr = AddCol(arr, MkCol("Fotografia", "ARTIST_IMAGE", "round", ""))
        arr = AddCol(arr, MkCol("Nombre", "ARTIST_NAME", "text", ""))
        arr = AddCol(arr, MkCol("Apellido", "ARTIST_LASTNAME", "text", ""))
    Case "playlists"
        arr = AddCol(arr, MkCol("#", "ID_PLAYLIST", "number", "right"))
        arr = AddCol(arr, MkCol("Imagen", "PHOTO", "image", ""))
        arr = AddCol(arr, MkCol("Titulo", "TITLE", "text", ""))
        arr = AddCol(arr, MkCol("Propietario", "OWNER_NAME", "text", ""))
        arr = AddCol(arr, MkCol("Visibilidad", "TYPE", "catalog2", ""))
        arr = AddCol(arr, MkCol("Canciones", "SONG_COUNT", "number", "right"))
        arr = AddCol(arr, MkCol("Creacion", "CREATION_DATE", "date", ""))
    End Select
    ListColumns = arr
End Function

' Consulta base del listado (sin WHERE)
Function EntityListSql(e)
    Select Case LCase(e)
    Case "users"
        EntityListSql = "SELECT ID_USER, USERNAME, EMAIL, USER_TYPE, BIRTHDAY, USER_PHOTO FROM TBL_USER "
    Case "songs"
        EntityListSql = "SELECT S.ID_SONG, S.SONG_NAME, S.SONG_PATH, ISNULL(S.PLAYS,0) AS PLAYS FROM TBL_SONG S "
    Case "albums"
        EntityListSql = "SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM "
    Case "artists"
        EntityListSql = "SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST "
    Case "playlists"
        EntityListSql = "SELECT P.ID_PLAYLIST, P.TITLE, P.TYPE, P.CREATION_DATE, P.PHOTO, P.ID_USER, " & _
                        "  U.USERNAME AS OWNER_NAME, " & _
                        "  (SELECT COUNT(*) FROM TBL_PLAYLIST_SONG PS WHERE PS.ID_PLAYLIST = P.ID_PLAYLIST) AS SONG_COUNT " & _
                        "FROM TBL_PLAYLIST P LEFT JOIN TBL_USER U ON U.ID_USER = P.ID_USER "
    Case Else
        EntityListSql = ""
    End Select
End Function

' ============================================================================
'  RenderEntityCell - pinta una celda segun el tipo de columna
' ============================================================================
Sub RenderEntityCell(rs, col, e)
    Dim v, id, sImg, bRound
    v = rs(col("expr")).Value
    bRound = (col("type") = "round")

    Select Case col("type")
        Case "number"
            Response.Write "<td class=""tbl-num"">" & FmtNumber(v) & "</td>"
        Case "date"
            Response.Write "<td>" & H(FmtDate(v)) & "</td>"
        Case "mono"
            Response.Write "<td><code class=""small"">" & H(Truncate(Nz(v, "-"), 46)) & "</code></td>"
        Case "image", "round"
            sImg = AssetUrl(v, IIf(bRound, IMG_FALLBACK_AVATAR, IMG_FALLBACK_COVER))
            Response.Write "<td><img class=""tbl-thumb"" src=""" & HA(sImg) & """ alt="""" loading=""lazy""" & _
                           IIf(bRound, " style=""border-radius:50%""", "") & "></td>"
        Case "catalog1"
            Response.Write "<td><span class=""badge " & IIf(ToInt(v) = USER_TYPE_ADMIN, "badge-admin", "badge-user") & """>" & _
                           H(CatalogLabel("1", CStr(ToInt(v)))) & "</span></td>"
        Case "catalog2"
            Response.Write "<td><span class=""badge " & IIf(ToInt(v) = PLAYLIST_PRIVATE, "badge-warn", "badge-ok") & """>" & _
                           H(CatalogLabel("2", CStr(ToInt(v)))) & "</span></td>"
        Case "user"
            sImg = AssetUrl(rs("USER_PHOTO").Value, IMG_FALLBACK_AVATAR)
            Response.Write "<td><div class=""tbl-cell-main""><img class=""tbl-thumb"" src=""" & HA(sImg) & _
                           """ alt="""" style=""border-radius:50%"" loading=""lazy""><div><div class=""strong"">" & _
                           H(v) & "</div><div class=""small muted"">" & H(rs("EMAIL").Value) & "</div></div></div></td>"
        Case "song"
            sImg = SongCoverUrl(rs("ID_SONG").Value)
            Response.Write "<td><div class=""tbl-cell-main""><img class=""tbl-thumb"" src=""" & HA(sImg) & _
                           """ alt="""" loading=""lazy""><div class=""strong"">" & H(v) & "</div></div></td>"
        Case Else
            Response.Write "<td>" & H(Nz(v, "-")) & "</td>"
    End Select
End Sub

' ============================================================================
'  EntityListParams - parametros del WHERE del listado
' ============================================================================
Function EntityListParams(e, sLike)
    Dim d, arr
    Set d = GetEntity(e)
    If Len(Nz(sLike, "")) = 0 Then
        EntityListParams = Array()
    Else
        arr = SearchParams(d, sLike)
        EntityListParams = arr
    End If
End Function

' ============================================================================
'  RenderFieldControl - pinta el control de formulario de un campo
' ============================================================================
Sub RenderFieldControl(f, sValue)
    Dim sName, sType, rs, i, opts, sSel, sReq, sMax, sInput, sCss

    sName = f("name")
    sType = LCase(f("type"))
    sReq  = IIf(f("required"), " <span class=""req"">*</span>", "")
    sMax  = ""
    If ToInt(f("maxlength")) > 0 Then sMax = " maxlength=""" & f("maxlength") & """"

    Select Case sType

    Case "image", "audio"
        RenderFilePicker sName, f("folder"), sValue, f("label"), IIf(sType = "audio", "audio", "image")

    Case "select"
        Response.Write "<div class=""field""><label class=""label"" for=""" & H(sName) & """>" & f("label") & sReq & "</label>" & _
                       "<select class=""select"" id=""" & H(sName) & """ name=""" & H(sName) & """ required>" & _
                       "<option value="""">-- Seleccione --</option>"
        opts = f("options")
        If Len(Nz(f("catalog"), "")) > 0 Then
            Set rs = DbSelect("SELECT COD_CATALOG, VALUE FROM TBL_CATALOG WHERE COD_CATALOG_PARENT = ? ORDER BY COD_CATALOG", Array(f("catalog")))
            Do While Not rs.EOF
                sSel = IIf(CStr(sValue) = CStr(rs("COD_CATALOG").Value), " selected", "")
                Response.Write "<option value=""" & H(rs("COD_CATALOG").Value) & """" & sSel & ">" & H(rs("VALUE").Value) & "</option>"
                rs.MoveNext
            Loop
            rs.Close
        ElseIf IsArray(opts) Then
            For i = 0 To UBound(opts) Step 2
                sSel = IIf(CStr(sValue) = CStr(opts(i)), " selected", "")
                Response.Write "<option value=""" & H(opts(i)) & """" & sSel & ">" & H(opts(i + 1)) & "</option>"
            Next
        End If
        Response.Write "</select>"
        If Len(Nz(f("hint"), "")) > 0 Then Response.Write "<span class=""hint"">" & H(f("hint")) & "</span>"
        Response.Write "</div>"

    Case "date"
        Response.Write "<div class=""field""><label class=""label"" for=""" & H(sName) & """>" & f("label") & sReq & "</label>" & _
                       "<input class=""input"" type=""date"" id=""" & H(sName) & """ name=""" & H(sName) & """ value=""" & _
                       HA(sValue) & """" & IIf(f("required"), " required", "") & ">"
        If Len(Nz(f("hint"), "")) > 0 Then Response.Write "<span class=""hint"">" & H(f("hint")) & "</span>"
        Response.Write "</div>"

    Case Else
        sCss = "input"
        Select Case sType
            Case "password" : sInput = "password"
            Case "number"   : sInput = "number"
            Case Else       : sInput = "text"
        End Select
        Response.Write "<div class=""field""><label class=""label"" for=""" & H(sName) & """>" & f("label") & sReq & "</label>" & _
                       "<input class=""" & sCss & """ type=""" & sInput & """ id=""" & H(sName) & """ name=""" & H(sName) & _
                       """ value=""" & HA(sValue) & """" & sMax & _
                       IIf(f("required"), " required", "") & _
                       IIf(sType = "password", " autocomplete=""new-password""", "") & ">"
        If Len(Nz(f("hint"), "")) > 0 Then Response.Write "<span class=""hint"">" & H(f("hint")) & "</span>"
        Response.Write "</div>"
    End Select
End Sub

' ============================================================================
'  EntityDelete - borra un registro resolviendo antes sus dependencias
'  Devuelve "" si todo fue bien, o el mensaje de error.
' ============================================================================
Function EntityDelete(e, id)
    Dim d, n

    EntityDelete = ""
    If id <= 0 Then
        EntityDelete = "Identificador invalido."
        Exit Function
    End If

    Set d = GetEntity(e)

    Select Case LCase(e)

    Case "songs"
        DbExecute "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array(id)
        DbExecute "DELETE FROM TBL_PLAYER WHERE ID_SONG = ?", Array(id)
        DbExecute "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = ?", Array(id)
        DbExecute "DELETE FROM TBL_SONG WHERE ID_SONG = ?", Array(id)

    Case "albums"
        DbExecute "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array(id)
        DbExecute "DELETE FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(id)

    Case "artists"
        DbExecute "DELETE FROM TBL_PLAYER WHERE ID_ARTIST = ?", Array(id)
        DbExecute "DELETE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array(id)

    Case "playlists"
        DbExecute "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?", Array(id)
        DbExecute "DELETE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(id)

    Case "users"
        If id = CurrentUserId() Then
            EntityDelete = "No puedes eliminar la cuenta con la que iniciaste sesion."
            Exit Function
        End If
        DbExecute "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST IN (SELECT ID_PLAYLIST FROM TBL_PLAYLIST WHERE ID_USER = ?)", Array(id)
        DbExecute "DELETE FROM TBL_PLAYLIST WHERE ID_USER = ?", Array(id)
        DbExecute "DELETE FROM TBL_USER WHERE ID_USER = ?", Array(id)

    End Select

    ' Comprobar que realmente se borro
    n = DbCount("SELECT COUNT(*) FROM " & d("table") & " WHERE " & d("pk") & " = ?", Array(id))
    If n > 0 Then EntityDelete = "No se pudo eliminar el registro."

End Function
%>
