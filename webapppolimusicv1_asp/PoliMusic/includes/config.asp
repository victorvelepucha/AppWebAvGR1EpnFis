<%
' ============================================================================
'  PoliMusic - Configuracion general de la aplicacion
'  Tecnologia : ASP clasico (VBScript) + SQL Server
'  Autor      : Equipo PoliMusic
' ============================================================================

' ---------------------------------------------------------------- Base de datos
Const DB_SERVER      = "localhost"
Const DB_NAME        = "BDD_PoliMusic"
Const DB_USER        = "usr_polimusic"
Const DB_PASSWORD    = "Politecnica1"
Const DB_ENCRYPT     = "true"
Const DB_TRUSTCERT   = "true"

' --------------------------------------------------------------- Aplicacion
Const APP_NAME       = "PoliMusic"
Const APP_TAGLINE    = "Tu musica. En azul."
Const APP_VERSION    = "1.0.0"

' Filas por pagina en los listados
Const ROWS_PER_PAGE  = 12

' ------------------------------------------------------- Tipos de usuario (TBL_CATALOG -> '1')
Const USER_TYPE_NORMAL = 0
Const USER_TYPE_ADMIN  = 1

' ------------------------------------------------------ Tipos de playlist (TBL_CATALOG -> '2')
Const PLAYLIST_PUBLIC  = 0
Const PLAYLIST_PRIVATE = 1

' --------------------------------------------------------- Extensiones validas
Const ALLOWED_AUDIO_EXT = ".mp3|.wav|.ogg|.m4a"
Const ALLOWED_IMAGE_EXT = ".jpg|.jpeg|.png|.gif|.webp|.svg"

' ------------------------------------------------------------------- Rutas
' Los paths almacenados en la base de datos usan el prefijo "~/"
'   ~/songFiles/bensound-adventure.mp3   ->   <APP_ROOT>assets/songFiles/bensound-adventure.mp3
Const ASSET_FOLDER       = "assets"
Const IMG_FALLBACK_COVER = "assets/img/fallback-cover.svg"
Const IMG_FALLBACK_AVATAR= "assets/img/fallback-avatar.svg"
Const IMG_LOGO           = "assets/img/logo.svg"

' Carpeta de subida de archivos (relativa a la raiz de la aplicacion)
Const UPLOAD_FOLDER      = "assets/uploads"

' ============================================================================
'  APP_ROOT : ruta virtual de la raiz de la aplicacion, por ejemplo "/PoliMusic/"
'  Se calcula a partir de SCRIPT_NAME descontando las subcarpetas conocidas,
'  de modo que la aplicacion funcione igual en la raiz de un sitio o en una
'  subcarpeta / aplicacion de IIS.
' ============================================================================
Function AppRoot()
    Dim sPath, sName
    sPath = Request.ServerVariables("SCRIPT_NAME")
    If IsEmpty(sPath) Or IsNull(sPath) Then sPath = "/default.asp"
    sPath = Replace(sPath, "\", "/")
    sPath = Replace(sPath, "/admin/", "/")
    sPath = Replace(sPath, "/api/", "/")
    sPath = Replace(sPath, "/includes/", "/")
    sName = Mid(sPath, InStrRev(sPath, "/") + 1)
    sPath = Left(sPath, InStrRev(sPath, "/"))
    If Len(sPath) = 0 Then sPath = "/"
    AppRoot = sPath
End Function

' Devuelve la URL absoluta del sitio (esquema + host + raiz de la aplicacion)
Function SiteRoot()
    Dim sHost, sPort, sProto
    sProto = "http"
    If LCase(Request.ServerVariables("HTTPS")) = "on" Then sProto = "https"
    sHost = Request.ServerVariables("SERVER_NAME")
    sPort = Request.ServerVariables("SERVER_PORT")
    If sPort <> "" And sPort <> "80" And sPort <> "443" Then sHost = sHost & ":" & sPort
    SiteRoot = sProto & "://" & sHost & AppRoot()
End Function

' Cadena de conexion OLE DB para SQL Server
Function ConnectionString(sProvider)
    ConnectionString = "Provider=" & sProvider & ";" & _
                       "Server=" & DB_SERVER & ";" & _
                       "Database=" & DB_NAME & ";" & _
                       "User ID=" & DB_USER & ";" & _
                       "Password=" & DB_PASSWORD & ";" & _
                       "Encrypt=" & DB_ENCRYPT & ";" & _
                       "TrustServerCertificate=" & DB_TRUSTCERT & ";" & _
                       "Persist Security Info=False;" & _
                       "MARS Connection=False;"
End Function
%>
