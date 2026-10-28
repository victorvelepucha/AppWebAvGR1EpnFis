<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' ============================================================================
'  profile.asp - Perfil del usuario autenticado
'  Permite actualizar correo, fecha de nacimiento, foto y contrasena.
'  El rol (USER_TYPE) solo puede cambiarlo un administrador desde el panel.
' ============================================================================
Dim sErr, sOk, sEmail, sBirth, sPhoto, sPass1, sPass2, sPassAct, rs, nPl, nFav, sCreated

RequireLogin

sErr = ""
sOk  = ""

If IsPost() Then
    Dim sForm : sForm = LCase(Req("form"))

    If sForm = "password" Then
        sPassAct = Request.Form("passActual")
        sPass1   = Request.Form("passNueva")
        sPass2   = Request.Form("passNueva2")

        If Len(sPassAct) = 0 Then
            sErr = "Debes indicar tu contrasena actual."
        ElseIf Not DbExists("SELECT 1 FROM TBL_USER WHERE ID_USER = ? AND (UPPER(PASSWORD) = UPPER(CONVERT(varchar(40), HASHBYTES('SHA1', CONVERT(varchar(200), ?)), 2)) OR PASSWORD = ?)", _
                            Array(CurrentUserId(), sPassAct, sPassAct)) Then
            sErr = "La contrasena actual no es correcta."
        ElseIf Len(sPass1) < 6 Then
            sErr = "La nueva contrasena debe tener al menos 6 caracteres."
        ElseIf Len(sPass1) > 50 Then
            sErr = "La nueva contrasena no puede superar los 50 caracteres."
        ElseIf sPass1 <> sPass2 Then
            sErr = "Las nuevas contrasenas no coinciden."
        Else
            DbExecute "UPDATE TBL_USER SET PASSWORD = UPPER(CONVERT(varchar(40), HASHBYTES('SHA1', CONVERT(varchar(200), ?)), 2)) WHERE ID_USER = ?", _
                      Array(sPass1, CurrentUserId())
            FlashSet "ok", "Tu contrasena fue actualizada."
            Response.Redirect AppRoot() & "profile.asp"
            Response.End
        End If

    Else
        sEmail = Req("email")
        sBirth = Req("birthday")
        sPhoto = Req("userPhoto")
        If Len(sPhoto) = 0 Then sPhoto = "~/userImages/default.jpg"

        If Len(sEmail) = 0 Or Len(sEmail) > 50 Then
            sErr = "El correo electronico es obligatorio (maximo 50 caracteres)."
        ElseIf Not IsValidEmail(sEmail) Then
            sErr = "El formato del correo electronico no es valido."
        ElseIf DbExists("SELECT 1 FROM TBL_USER WHERE EMAIL = ? AND ID_USER <> ?", Array(sEmail, CurrentUserId())) Then
            sErr = "Ese correo ya esta siendo usado por otra cuenta."
        ElseIf Len(sBirth) > 0 And Not IsValidDate(sBirth) Then
            sErr = "La fecha de nacimiento no es valida (formato AAAA-MM-DD)."
        ElseIf Len(sBirth) > 0 And CDate(sBirth) > Date() Then
            sErr = "La fecha de nacimiento no puede ser futura."
        Else
            DbExecute "UPDATE TBL_USER SET EMAIL = ?, BIRTHDAY = TRY_CONVERT(datetime, NULLIF(?, ''), 120), USER_PHOTO = ? WHERE ID_USER = ?", _
                      Array(sEmail, sBirth, sPhoto, CurrentUserId())
            RefreshSession
            FlashSet "ok", "Tus datos fueron actualizados correctamente."
            Response.Redirect AppRoot() & "profile.asp"
            Response.End
        End If
    End If
End If

' ------------------------------------------------------------ datos actuales
Set rs = DbSelect("SELECT ID_USER, USERNAME, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE FROM TBL_USER WHERE ID_USER = ?", Array(CurrentUserId()))
If rs.EOF Then
    rs.Close
    Logout
    Response.Redirect AppRoot() & "login.asp"
    Response.End
End If
sEmail = CStr(Nz(rs("EMAIL").Value, ""))
sBirth = FmtDateInput(rs("BIRTHDAY").Value)
sPhoto = CStr(Nz(rs("USER_PHOTO").Value, "~/userImages/default.jpg"))
sCreated = ""
rs.Close

nPl  = DbCount("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_USER = ?", Array(CurrentUserId()))
nFav = 0

PageHeader "Mi perfil", "", ""
%>

<div class="page-head">
  <div>
    <h1 class="page-title">Mi perfil</h1>
    <p class="page-sub">Administra tu informaci&oacute;n personal y tu contrase&ntilde;a.</p>
  </div>
</div>

<% If Len(sErr) > 0 Then %>
<div class="alert alert-error" role="alert"><span class="alert-ico">!</span><span><%=H(sErr)%></span></div>
<% End If %>

<div class="grid grid-2">
  <section class="panel">
    <div class="panel-head"><span class="panel-title"><%=Ico("user", 17)%> Datos personales</span></div>
    <div class="panel-body">
      <div class="hstack mb-3">
        <img class="avatar avatar-lg" src="<%=HA(AssetUrl(sPhoto, IMG_FALLBACK_AVATAR))%>" alt="">
        <div>
          <div class="strong"><%=H(CurrentUserName())%></div>
          <span class="badge <%=IIf(IsAdmin(),"badge-admin","badge-user")%>"><%=H(CurrentRoleLabel())%></span>
        </div>
      </div>

      <form method="post" action="<%=AppRoot()%>profile.asp">
        <input type="hidden" name="form" value="datos">
        <div class="form-grid">
          <div class="field">
            <label class="label" for="username">Usuario</label>
            <input class="input" type="text" id="username" value="<%=HA(CurrentUserName())%>" readonly>
            <span class="hint">El nombre de usuario no se puede cambiar.</span>
          </div>
          <div class="field">
            <label class="label" for="email">Correo electr&oacute;nico <span class="req">*</span></label>
            <input class="input" type="email" id="email" name="email" maxlength="50" value="<%=HA(sEmail)%>" required>
          </div>
          <div class="field">
            <label class="label" for="birthday">Fecha de nacimiento</label>
            <input class="input" type="date" id="birthday" name="birthday" value="<%=HA(sBirth)%>">
          </div>
          <div class="field">
            <label class="label">Playlists creadas</label>
            <input class="input" type="text" value="<%=nPl%>" readonly>
          </div>
<% RenderFilePicker "userPhoto", "userImages", sPhoto, "Foto de perfil", "image" %>
        </div>
        <div class="form-actions">
          <button type="submit" class="btn btn-primary"><%=Ico("check", 17)%> Guardar cambios</button>
          <a class="btn btn-ghost" href="<%=AppRoot()%>home.asp">Cancelar</a>
        </div>
      </form>
    </div>
  </section>

  <section class="panel">
    <div class="panel-head"><span class="panel-title"><%=Ico("lock", 17)%> Cambiar contrase&ntilde;a</span></div>
    <div class="panel-body">
      <form method="post" action="<%=AppRoot()%>profile.asp">
        <input type="hidden" name="form" value="password">
        <div class="form-grid">
          <div class="field field-full">
            <label class="label" for="passActual">Contrase&ntilde;a actual <span class="req">*</span></label>
            <input class="input" type="password" id="passActual" name="passActual" maxlength="50" autocomplete="current-password" required>
          </div>
          <div class="field">
            <label class="label" for="passNueva">Nueva contrase&ntilde;a <span class="req">*</span></label>
            <input class="input" type="password" id="passNueva" name="passNueva" maxlength="50" autocomplete="new-password" required>
          </div>
          <div class="field">
            <label class="label" for="passNueva2">Repetir nueva <span class="req">*</span></label>
            <input class="input" type="password" id="passNueva2" name="passNueva2" maxlength="50" autocomplete="new-password" required>
          </div>
        </div>
        <div class="form-actions">
          <button type="submit" class="btn btn-primary"><%=Ico("shield", 17)%> Actualizar contrase&ntilde;a</button>
        </div>
      </form>
      <div class="alert alert-info mt-3" style="margin-bottom:0">
        <span class="alert-ico">i</span>
        <span>Las contrase&ntilde;as se guardan cifradas con <strong>SHA-1</strong> en la base de datos. El rol de la cuenta solo puede modificarlo un administrador.</span>
      </div>
    </div>
  </section>
</div>

<%
PageFooter
%>
