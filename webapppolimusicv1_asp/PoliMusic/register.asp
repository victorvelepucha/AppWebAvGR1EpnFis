<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' ============================================================================
'  register.asp - Registro / enrolamiento de usuarios
'  Crea siempre usuarios con rol NORMAL (USER_TYPE = 0).
'  La contrasena se almacena como SHA-1 (mismo formato que usa la base).
' ============================================================================
Dim sUser, sEmail, sPass, sPass2, sBirth, sPhoto, sErr, i
Dim bOk, vPhoto, sMsg

sUser = "" : sEmail = "" : sBirth = "" : sErr = ""
sPhoto = "~/userImages/default.jpg"
bOk = False

If IsLoggedIn() Then
    Response.Redirect AppRoot() & "home.asp"
    Response.End
End If

If IsPost() Then
    sUser  = Req("username")
    sEmail = Req("email")
    sPass  = Request.Form("password")
    sPass2 = Request.Form("password2")
    sBirth = Req("birthday")
    sPhoto = Req("userPhoto")
    If Len(sPhoto) = 0 Then sPhoto = "~/userImages/default.jpg"

    ' ------------------------------------------------------------ validaciones
    If Len(sUser) < 3 Or Len(sUser) > 50 Then
        sErr = "El usuario debe tener entre 3 y 50 caracteres."
    ElseIf Not IsValidUsername(sUser) Then
        sErr = "El usuario solo puede contener letras, numeros, punto, guion y guion bajo."
    ElseIf DbExists("SELECT 1 FROM TBL_USER WHERE USERNAME = ?", Array(sUser)) Then
        sErr = "Ese nombre de usuario ya esta registrado."
    ElseIf Len(sEmail) = 0 Or Len(sEmail) > 50 Then
        sErr = "El correo electronico es obligatorio (maximo 50 caracteres)."
    ElseIf Not IsValidEmail(sEmail) Then
        sErr = "El formato del correo electronico no es valido."
    ElseIf DbExists("SELECT 1 FROM TBL_USER WHERE EMAIL = ?", Array(sEmail)) Then
        sErr = "Ese correo electronico ya esta registrado."
    ElseIf Len(sPass) < 6 Then
        sErr = "La contrasena debe tener al menos 6 caracteres."
    ElseIf Len(sPass) > 50 Then
        sErr = "La contrasena no puede superar los 50 caracteres."
    ElseIf sPass <> sPass2 Then
        sErr = "Las contrasenas no coinciden."
    ElseIf Len(sBirth) > 0 And Not IsValidDate(sBirth) Then
        sErr = "La fecha de nacimiento no es valida (formato AAAA-MM-DD)."
    ElseIf Len(sBirth) > 0 And CDate(sBirth) > Date() Then
        sErr = "La fecha de nacimiento no puede ser futura."
    End If

    If Len(sErr) = 0 Then
        ' --------------------------------------------------------- alta del usuario
        On Error Resume Next
        DbExecute "INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE) " & _
                  "VALUES (?, UPPER(CONVERT(varchar(40), HASHBYTES('SHA1', CONVERT(varchar(200), ?)), 2)), ?, " & _
                  "        TRY_CONVERT(datetime, NULLIF(?, ''), 120), ?, ?)", _
                  Array(sUser, sPass, sEmail, sBirth, sPhoto, CStr(USER_TYPE_NORMAL))
        If Err.Number <> 0 Then
            sErr = "No se pudo crear la cuenta: " & Err.Description
            Err.Clear
        Else
            bOk = True
        End If
        On Error GoTo 0
    End If

    ' Tras crear la cuenta se inicia sesion automaticamente
    If bOk Then
        sErr = Login(sUser, sPass)
        If Len(sErr) = 0 Then
            FlashSet "ok", "Bienvenido a PoliMusic, " & sUser & ". Tu cuenta fue creada correctamente."
            Response.Redirect AppRoot() & "home.asp"
            Response.End
        End If
    End If
End If

PageHeaderAuth "Crear cuenta", ""
%>
      <h1 class="auth-title">Crear tu cuenta</h1>
      <p class="auth-sub">Reg&iacute;strate para guardar tus playlists y escuchar m&uacute;sica.</p>

<% If Len(sErr) > 0 Then %>
      <div class="alert alert-error" role="alert"><span class="alert-ico">!</span><span><%=H(sErr)%></span></div>
<% End If %>

      <form class="auth-form" method="post" action="<%=AppRoot()%>register.asp" novalidate>
        <div class="field">
          <label class="label" for="username">Usuario <span class="req">*</span></label>
          <input class="input" type="text" id="username" name="username" value="<%=HA(sUser)%>"
                 placeholder="epn.fis" maxlength="50" autocomplete="username" required autofocus>
          <span class="hint">Entre 3 y 50 caracteres, sin espacios.</span>
        </div>
        <div class="field">
          <label class="label" for="email">Correo electr&oacute;nico <span class="req">*</span></label>
          <input class="input" type="email" id="email" name="email" value="<%=HA(sEmail)%>"
                 placeholder="usuario@epn.edu.ec" maxlength="50" autocomplete="email" required>
        </div>
        <div class="field">
          <label class="label" for="password">Contrase&ntilde;a <span class="req">*</span></label>
          <input class="input" type="password" id="password" name="password" maxlength="50"
                 placeholder="Minimo 6 caracteres" autocomplete="new-password" required>
        </div>
        <div class="field">
          <label class="label" for="password2">Repetir contrase&ntilde;a <span class="req">*</span></label>
          <input class="input" type="password" id="password2" name="password2" maxlength="50"
                 placeholder="Repite la contrase&ntilde;a" autocomplete="new-password" required>
        </div>
        <div class="field">
          <label class="label" for="birthday">Fecha de nacimiento</label>
          <input class="input" type="date" id="birthday" name="birthday" value="<%=HA(sBirth)%>">
        </div>

<% RenderFilePicker "userPhoto", "userImages", sPhoto, "Foto de perfil", "image" %>

        <button type="submit" class="btn btn-primary btn-lg btn-block"><%=Ico("user", 18)%> Crear cuenta</button>
      </form>

      <p class="auth-alt">&iquest;Ya tienes cuenta? <a href="<%=AppRoot()%>login.asp">Inicia sesi&oacute;n</a></p>
<%
PageFooterAuth
%>
