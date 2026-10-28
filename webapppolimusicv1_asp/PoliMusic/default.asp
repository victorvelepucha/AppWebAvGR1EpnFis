<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' ============================================================================
'  default.asp - Punto de entrada de PoliMusic
'  Redirige al inicio si hay sesion activa, o al login en caso contrario.
' ============================================================================
Response.Redirect AppRoot() & IIf(IsLoggedIn(), "home.asp", "login.asp")
Response.End
%>
