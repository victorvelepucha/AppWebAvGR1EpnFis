<%
' ============================================================================
'  PoliMusic - Autenticacion y autorizacion
'  Dos roles: 0 = usuario normal (epn.fis)  |  1 = administrador (userAdmin)
' ============================================================================

' ============================================================================
'  Login - valida credenciales contra TBL_USER.
'  Acepta contrasenas almacenadas como SHA-1 (formato nuevo) y tambien las
'  contrasenas en texto plano heredadas de la carga inicial de la base.
'  Devuelve "" si el acceso es correcto, o el mensaje de error.
' ============================================================================
Function Login(sUser, sPass)
    Dim sql, rs, u
    Login = ""

    If Len(Trim(sUser)) = 0 Or Len(sPass) = 0 Then
        Login = "Debe ingresar usuario y contrasena."
        Exit Function
    End If

    sql = "SELECT TOP 1 ID_USER, USERNAME, EMAIL, USER_PHOTO, USER_TYPE, BIRTHDAY " & _
          "FROM TBL_USER " & _
          "WHERE USERNAME = ? AND ( " & _
          "      UPPER(PASSWORD) = UPPER(CONVERT(varchar(40), HASHBYTES('SHA1', CONVERT(varchar(200), ?)), 2)) " & _
          "      OR PASSWORD = ? " & _
          "      )"
    Set rs = DbSelect(sql, Array(Trim(sUser), sPass, sPass))

    If rs.EOF Then
        rs.Close
        Login = "Usuario o contrasena incorrectos."
        Exit Function
    End If

    Session("PM_USER_ID")    = CLng(rs("ID_USER").Value)
    Session("PM_USERNAME")   = CStr(rs("USERNAME").Value)
    Session("PM_EMAIL")      = CStr(Nz(rs("EMAIL").Value, ""))
    Session("PM_USER_TYPE")  = ToInt(rs("USER_TYPE").Value)
    Session("PM_USER_PHOTO") = CStr(Nz(rs("USER_PHOTO").Value, "~/userImages/default.jpg"))
    Session("PM_BIRTHDAY")   = FmtDateInput(rs("BIRTHDAY").Value)
    Session("PM_LOGIN_TIME") = Now()
    rs.Close
End Function

' Cierra la sesion actual
Sub Logout()
    Session.Contents.RemoveAll()
    Session.Abandon()
End Sub

' ------------------------------------------------------------------ Estado
Function IsLoggedIn()
    IsLoggedIn = (ToInt(Session("PM_USER_ID")) > 0)
End Function

Function IsAdmin()
    IsAdmin = (IsLoggedIn() And ToInt(Session("PM_USER_TYPE")) = USER_TYPE_ADMIN)
End Function

Function CurrentUserId()
    CurrentUserId = ToInt(Session("PM_USER_ID"))
End Function

Function CurrentUserName()
    CurrentUserName = Nz(Session("PM_USERNAME"), "")
End Function

Function CurrentUserEmail()
    CurrentUserEmail = Nz(Session("PM_EMAIL"), "")
End Function

Function CurrentUserType()
    CurrentUserType = ToInt(Session("PM_USER_TYPE"))
End Function

Function CurrentUserPhotoUrl()
    CurrentUserPhotoUrl = AssetUrl(Session("PM_USER_PHOTO"), IMG_FALLBACK_AVATAR)
End Function

' Etiqueta del rol segun TBL_CATALOG
Function CurrentRoleLabel()
    If IsAdmin() Then
        CurrentRoleLabel = CatalogLabel("1", "1")
    Else
        CurrentRoleLabel = CatalogLabel("1", "0")
    End If
End Function

' Refresca los datos de sesion del usuario (tras editar su perfil)
Sub RefreshSession()
    Dim rs
    If Not IsLoggedIn() Then Exit Sub
    Set rs = DbSelect("SELECT ID_USER, USERNAME, EMAIL, USER_PHOTO, USER_TYPE, BIRTHDAY FROM TBL_USER WHERE ID_USER = ?", _
                      Array(CurrentUserId()))
    If Not rs.EOF Then
        Session("PM_USERNAME")   = CStr(rs("USERNAME").Value)
        Session("PM_EMAIL")      = CStr(Nz(rs("EMAIL").Value, ""))
        Session("PM_USER_TYPE")  = ToInt(rs("USER_TYPE").Value)
        Session("PM_USER_PHOTO") = CStr(Nz(rs("USER_PHOTO").Value, "~/userImages/default.jpg"))
        Session("PM_BIRTHDAY")   = FmtDateInput(rs("BIRTHDAY").Value)
    End If
    rs.Close
End Sub

' ------------------------------------------------------- Guardias de acceso
' Exige sesion iniciada; si no, redirige al login conservando el destino.
Sub RequireLogin()
    Dim nextUrl
    If IsLoggedIn() Then Exit Sub
    nextUrl = Request.ServerVariables("URL")
    If Request.ServerVariables("QUERY_STRING") <> "" Then
        nextUrl = nextUrl & "?" & Request.ServerVariables("QUERY_STRING")
    End If
    Response.Redirect AppRoot() & "login.asp?next=" & Server.URLEncode(nextUrl)
    Response.End
End Sub

' Exige rol administrador (403 si el usuario es normal)
Sub RequireAdmin()
    RequireLogin()
    If IsAdmin() Then Exit Sub
    Response.Status = "403 Forbidden"
    Response.Write "<!DOCTYPE html><html lang=""es""><head><meta charset=""utf-8"">" & _
                   "<title>Acceso denegado</title><link rel=""stylesheet"" href=""" & AppRoot() & "css/style.css""></head>" & _
                   "<body class=""page-error""><div class=""error-box"">" & _
                   "<div class=""error-code"">403</div><h1>Acceso restringido</h1>" & _
                   "<p>Solo el usuario administrador puede acceder a esta seccion.</p>" & _
                   "<a class=""btn btn-primary"" href=""" & AppRoot() & "home.asp"">Volver al inicio</a>" & _
                   "</div></body></html>"
    Response.End
End Sub

' Indica si el usuario actual puede gestionar una playlist concreta
Function CanManagePlaylist(idPlaylist)
    Dim idOwner
    CanManagePlaylist = False
    If Not IsLoggedIn() Then Exit Function
    If IsAdmin() Then
        CanManagePlaylist = True
        Exit Function
    End If
    idOwner = ToInt(DbScalar("SELECT ID_USER FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(idPlaylist)))
    CanManagePlaylist = (idOwner > 0 And idOwner = CurrentUserId())
End Function
%>
