<%
' ============================================================================
'  PoliMusic - Funciones utilitarias
'  Codificacion HTML/JSON, resolucion de rutas de recursos, formateo, etc.
' ============================================================================

Dim gAssetCache   ' Dictionary con el resultado de las comprobaciones de archivo

' ---------------------------------------------------------------- Nulos / tipos
Function Nz(v, fallback)
    If IsNull(v) Or IsEmpty(v) Then
        Nz = fallback
    ElseIf VarType(v) = vbString And Len(Trim(v)) = 0 Then
        Nz = fallback
    Else
        Nz = v
    End If
End Function

Function IIf(cond, a, b)
    If cond Then IIf = a Else IIf = b
End Function

' Convierte a entero de forma segura. Devuelve 0 si no es numerico.
Function ToInt(v)
    Dim s
    s = Trim(Nz(v, ""))
    If Len(s) = 0 Then
        ToInt = 0
    ElseIf IsNumeric(s) Then
        On Error Resume Next
        ToInt = CLng(s)
        If Err.Number <> 0 Then
            Err.Clear
            ToInt = 0
        End If
        On Error GoTo 0
    Else
        ToInt = 0
    End If
End Function

' Lee un parametro de la peticion (formulario o querystring) como texto
Function Req(name)
    Dim v
    v = Request.Form(name)
    If Len(v) = 0 Then v = Request.QueryString(name)
    Req = Trim(v)
End Function

Function ReqInt(name)
    ReqInt = ToInt(Req(name))
End Function

' Numero de pagina solicitado (>= 1)
Function ReqPage()
    Dim p
    p = ToInt(Request.QueryString("page"))
    If p < 1 Then p = 1
    ReqPage = p
End Function

Function IsPost()
    IsPost = (UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST")
End Function

' --------------------------------------------------------------- Validaciones
' Nombre de usuario: letras, numeros, punto, guion y guion bajo (sin espacios)
Function IsValidUsername(s)
    Dim i, c
    IsValidUsername = False
    s = Trim(Nz(s, ""))
    If Len(s) < 3 Then Exit Function
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If Not ((c >= "a" And c <= "z") Or (c >= "A" And c <= "Z") Or _
                (c >= "0" And c <= "9") Or c = "." Or c = "_" Or c = "-") Then
            Exit Function
        End If
    Next
    IsValidUsername = True
End Function

' Correo electronico con formato basico usuario@dominio.tld
Function IsValidEmail(s)
    Dim p, q, dom
    IsValidEmail = False
    s = Trim(Nz(s, ""))
    If Len(s) < 6 Or Len(s) > 50 Then Exit Function
    If InStr(s, " ") > 0 Then Exit Function
    p = InStr(s, "@")
    If p < 2 Then Exit Function
    If InStr(p + 1, s, "@") > 0 Then Exit Function
    dom = Mid(s, p + 1)
    q = InStrRev(dom, ".")
    If q < 2 Or q >= Len(dom) Then Exit Function
    IsValidEmail = True
End Function

' Fecha valida en formato AAAA-MM-DD
Function IsValidDate(s)
    Dim y, m, d, parts
    IsValidDate = False
    s = Trim(Nz(s, ""))
    If Len(s) <> 10 Then Exit Function
    parts = Split(s, "-")
    If UBound(parts) <> 2 Then Exit Function
    If Not (IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2))) Then Exit Function
    y = CInt(parts(0)) : m = CInt(parts(1)) : d = CInt(parts(2))
    If y < 1900 Or y > 2200 Then Exit Function
    If m < 1 Or m > 12 Then Exit Function
    If d < 1 Or d > 31 Then Exit Function
    On Error Resume Next
    Dim dt
    dt = DateSerial(y, m, d)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    IsValidDate = (Year(dt) = y And Month(dt) = m And Day(dt) = d)
End Function

' ------------------------------------------------------------------- Seguridad
' Codifica una cadena para insertarla en HTML (evita XSS)
Function H(v)
    Dim s
    s = Nz(v, "")
    If Len(s) = 0 Then
        H = ""
        Exit Function
    End If
    s = CStr(s)
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&#39;")
    H = s
End Function

' Codifica para usar dentro de un atributo (incluye salto de linea)
Function HA(v)
    HA = Replace(H(v), vbCrLf, " ")
End Function

' Codifica una cadena como literal JSON
Function J(v)
    Dim s, i, c, code, out
    s = Nz(v, "")
    If Len(s) = 0 Then
        J = """"""
        Exit Function
    End If
    s = CStr(s)
    out = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        code = AscW(c)
        If code < 0 Then code = code + 65536
        Select Case c
            Case """" : out = out & "\"""
            Case "\"  : out = out & "\\"
            Case "/"  : out = out & "\/"
            Case vbCr : out = out & "\r"
            Case vbLf : out = out & "\n"
            Case vbTab: out = out & "\t"
            Case Else
                If code < 32 Or code = 60 Or code = 62 Or code = 38 Then
                    out = out & "\u" & Right("0000" & Hex(code), 4)
                Else
                    out = out & c
                End If
        End Select
    Next
    J = """" & out & """"
End Function

' --------------------------------------------------------------- Recursos (assets)
' Resuelve un path almacenado en la base de datos (prefijo ~/) a una URL real.
'   ~/songFiles/bensound-adventure.mp3  ->  /PoliMusic/assets/songFiles/bensound-adventure.mp3
Function AssetUrl(dbPath, fallback)
    Dim s, p
    s = Trim(Nz(dbPath, ""))
    s = Replace(s, "\", "/")

    If Len(s) = 0 Then
        AssetUrl = AppRoot() & Nz(fallback, IMG_FALLBACK_COVER)
        Exit Function
    End If

    ' Rutas absolutas o externas se respetan tal cual
    If Left(s, 7) = "http://" Or Left(s, 8) = "https://" Or Left(s, 2) = "//" Then
        AssetUrl = s
        Exit Function
    End If

    If Left(s, 2) = "~/" Then
        s = Mid(s, 3)
    ElseIf Left(s, 1) = "~" Then
        s = Mid(s, 2)
    End If
    If Left(s, 1) = "/" Then s = Mid(s, 2)

    ' Evitar rutas relativas al directorio actual
    If Left(s, 3) = "../" Then s = Mid(s, 4)
    Do While Left(s, 3) = "../"
        s = Mid(s, 4)
    Loop

    ' Si ya viene con la carpeta de assets, no duplicarla
    If LCase(Left(s, Len(ASSET_FOLDER) + 1)) = LCase(ASSET_FOLDER) & "/" Then
        p = s
    Else
        p = ASSET_FOLDER & "/" & s
    End If

    ' Si el archivo no existe se usa el recurso alternativo
    If Not AssetFileExists(p) Then
        If Len(Nz(fallback, "")) > 0 And AssetFileExists(fallback) Then
            p = fallback
        End If
    End If

    AssetUrl = AppRoot() & p
End Function

' Comprueba si un recurso relativo a la raiz de la aplicacion existe en disco.
' El resultado se memoriza por peticion para no golpear el disco repetidamente.
Function AssetFileExists(relPath)
    Dim fso, phys, key, res
    If Not IsObject(gAssetCache) Then
        Set gAssetCache = Server.CreateObject("Scripting.Dictionary")
        gAssetCache.CompareMode = 1   ' texto, sin distinguir mayusculas
    End If
    key = Replace(Nz(relPath, ""), "\", "/")
    If Len(key) = 0 Then
        AssetFileExists = False
        Exit Function
    End If
    If gAssetCache.Exists(key) Then
        AssetFileExists = gAssetCache(key)
        Exit Function
    End If
    res = False
    On Error Resume Next
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    phys = Server.MapPath("/" & key)
    If Err.Number = 0 Then res = fso.FileExists(phys)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    gAssetCache(key) = res
    AssetFileExists = res
End Function

' URL de la imagen de portada de una cancion (usa la caratula de su album si existe)
Function SongCoverUrl(idSong)
    Dim v
    v = DbScalar("SELECT TOP 1 A.ALBUM_IMAGE_PATH FROM TBL_TRACKLIST T " & _
                 "INNER JOIN TBL_ALBUM A ON A.ID_ALBUM = T.ID_ALBUM " & _
                 "WHERE T.ID_SONG = ? ORDER BY A.ID_ALBUM", Array(idSong))
    SongCoverUrl = AssetUrl(v, IMG_FALLBACK_COVER)
End Function

' Nombre del artista principal de una cancion
Function SongMainArtist(idSong)
    Dim v
    v = DbScalar("SELECT TOP 1 A.ARTIST_NAME + CASE WHEN ISNULL(A.ARTIST_LASTNAME,'') = '' THEN '' " & _
                 "ELSE ' ' + A.ARTIST_LASTNAME END FROM TBL_PLAYER P " & _
                 "INNER JOIN TBL_ARTIST A ON A.ID_ARTIST = P.ID_ARTIST " & _
                 "WHERE P.ID_SONG = ? ORDER BY A.ID_ARTIST", Array(idSong))
    SongMainArtist = Nz(v, "Artista desconocido")
End Function

' --------------------------------------------------------------- Formateo
Function FmtDate(v)
    If IsNull(v) Or IsEmpty(v) Then
        FmtDate = "-"
        Exit Function
    End If
    On Error Resume Next
    If IsDate(v) Then
        FmtDate = FormatDateTime(CDate(v), 2)
    Else
        FmtDate = "-"
    End If
    If Err.Number <> 0 Then
        Err.Clear
        FmtDate = "-"
    End If
    On Error GoTo 0
End Function

Function FmtDateInput(v)
    If IsNull(v) Or IsEmpty(v) Then
        FmtDateInput = ""
        Exit Function
    End If
    On Error Resume Next
    If IsDate(v) Then
        FmtDateInput = Year(CDate(v)) & "-" & Right("0" & Month(CDate(v)), 2) & "-" & Right("0" & Day(CDate(v)), 2)
    Else
        FmtDateInput = ""
    End If
    On Error GoTo 0
End Function

Function FmtNumber(v)
    Dim n
    n = Nz(v, 0)
    If Not IsNumeric(n) Then n = 0
    FmtNumber = FormatNumber(n, 0, -1, 0, -1)
End Function

Function FmtDuration(seconds)
    Dim s
    s = ToInt(seconds)
    If s < 0 Then s = 0
    FmtDuration = Int(s / 60) & ":" & Right("0" & (s Mod 60), 2)
End Function

Function Truncate(v, n)
    Dim s
    s = Nz(v, "")
    If Len(s) > n Then
        Truncate = Left(s, n - 1) & "..."
    Else
        Truncate = s
    End If
End Function

' Concuerda el numero con el sustantivo: Plural(1,"cancion","canciones")
' Devuelve HTML ya listo, por lo que admite entidades (&oacute;, ...).
Function Plural(n, sSing, sPlur)
    n = ToInt(n)
    If n = 1 Then
        Plural = n & " " & sSing
    Else
        Plural = n & " " & sPlur
    End If
End Function

' Nombre completo del artista
Function ArtistFullName(nombre, apellido)
    Dim s
    s = Trim(Nz(nombre, ""))
    If Len(Trim(Nz(apellido, ""))) > 0 Then s = s & " " & Trim(apellido)
    ArtistFullName = s
End Function

' Iniciales para los avatares generados
Function Initials(v)
    Dim parts, s, out, i
    s = Trim(Nz(v, "?"))
    If Len(s) = 0 Then s = "?"
    parts = Split(s, " ")
    out = ""
    For i = 0 To UBound(parts)
        If Len(Trim(parts(i))) > 0 And Len(out) < 2 Then
            out = out & UCase(Left(Trim(parts(i)), 1))
        End If
    Next
    If Len(out) = 0 Then out = UCase(Left(s, 1))
    Initials = out
End Function

' --------------------------------------------------------------- Navegacion
Function QsUrl(baseUrl, params)
    Dim i, s
    s = ""
    If IsArray(params) Then
        For i = 0 To UBound(params) Step 2
            If Len(s) > 0 Then s = s & "&amp;"
            s = s & Server.URLEncode(params(i)) & "=" & Server.URLEncode(CStr(params(i + 1)))
        Next
    End If
    If Len(s) > 0 Then
        If InStr(baseUrl, "?") > 0 Then
            QsUrl = baseUrl & "&amp;" & s
        Else
            QsUrl = baseUrl & "?" & s
        End If
    Else
        QsUrl = baseUrl
    End If
End Function

' Devuelve la etiqueta de un valor del catalogo (TBL_CATALOG)
Function CatalogLabel(codParent, codValue)
    Dim v
    v = DbScalar("SELECT TOP 1 VALUE FROM TBL_CATALOG WHERE COD_CATALOG_PARENT = ? AND COD_CATALOG = ?", _
                 Array(codParent, codValue))
    CatalogLabel = Nz(v, "-")
End Function

' --------------------------------------------------------------- Mensajes flash
Sub FlashSet(tipo, mensaje)
    Session("PM_FLASH_TYPE") = tipo
    Session("PM_FLASH_MSG") = mensaje
End Sub

' Escribe (y consume) el mensaje flash. Devuelve "" si no hay nada.
Function FlashRender()
    Dim t, m, cls, ico
    t = Nz(Session("PM_FLASH_TYPE"), "")
    m = Nz(Session("PM_FLASH_MSG"), "")
    Session("PM_FLASH_TYPE") = ""
    Session("PM_FLASH_MSG") = ""
    If Len(m) = 0 Then
        FlashRender = ""
        Exit Function
    End If
    cls = "alert-info"
    ico = "i"
    If t = "ok" Then
        cls = "alert-ok" : ico = "OK"
    ElseIf t = "error" Then
        cls = "alert-error" : ico = "!"
    ElseIf t = "warn" Then
        cls = "alert-warn" : ico = "!"
    End If
    FlashRender = "<div class=""alert " & cls & """ role=""status""><span class=""alert-ico"">" & ico & _
                  "</span><span>" & H(m) & "</span></div>"
End Function
%>
