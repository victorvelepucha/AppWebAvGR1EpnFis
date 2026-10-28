<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' ============================================================================
'  login.asp - Inicio de sesion
' ============================================================================
Dim sUser, sPass, sError, sNext, sRemember

sError = ""
sNext  = Nz(Request.QueryString("next"), "")
sUser  = ""

' --- Validacion del destino posterior al login (evita redirecciones externas)
If Len(sNext) > 0 Then
    If InStr(sNext, "//") > 0 Or InStr(sNext, ":") > 0 Then
        sNext = ""
    ElseIf Left(sNext, 1) <> "/" Then
        sNext = ""
    End If
End If

' Si ya hay sesion iniciada no tiene sentido mostrar el formulario
If IsLoggedIn() Then
    Response.Redirect AppRoot() & "home.asp"
    Response.End
End If

If IsPost() Then
    sUser = Req("username")
    sPass = Request.Form("password")
    sError = Login(sUser, sPass)
    If Len(sError) = 0 Then
        If Len(sNext) > 0 Then
            Response.Redirect sNext
        Else
            Response.Redirect AppRoot() & "home.asp"
        End If
        Response.End
    End If
End If

PageHeaderAuth "Iniciar sesi&oacute;n", ""
%>
      <h1 class="auth-title">Iniciar sesi&oacute;n</h1>
      <p class="auth-sub">Ingresa con tu cuenta para escuchar tu m&uacute;sica en PoliMusic.</p>

<% If Len(sError) > 0 Then %>
      <div class="alert alert-error" role="alert">
        <span class="alert-ico">!</span><span><%=H(sError)%></span>
      </div>
<% End If %>

      <form class="auth-form" method="post" action="<%=AppRoot()%>login.asp<%=IIf(Len(sNext)>0, "?next=" & Server.URLEncode(sNext), "")%>" novalidate>
        <div class="field">
          <label class="label" for="username">Usuario <span class="req">*</span></label>
          <input class="input" type="text" id="username" name="username" value="<%=HA(sUser)%>"
                 placeholder="userAdmin" autocomplete="username" required autofocus>
        </div>
        <div class="field">
          <label class="label" for="password">Contrase&ntilde;a <span class="req">*</span></label>
          <input class="input" type="password" id="password" name="password"
                 placeholder="Tu contrase&ntilde;a" autocomplete="current-password" required>
        </div>
        <button type="submit" class="btn btn-primary btn-lg btn-block"><%=Ico("logout", 18)%> Entrar</button>
      </form>

      <p class="auth-alt">&iquest;No tienes cuenta? <a href="<%=AppRoot()%>register.asp">Reg&iacute;strate</a></p>

      <div class="auth-demo">
        <b>Cuentas de prueba</b><br>
        Administrador: <code>userAdmin</code> / <code>12345678</code><br>
        Usuario normal: <code>epn.fis</code> / <code>12345678</code>
      </div>
<%
PageFooterAuth
%>
