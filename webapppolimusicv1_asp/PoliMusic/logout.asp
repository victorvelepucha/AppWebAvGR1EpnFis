<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' ============================================================================
'  logout.asp - Cierre de sesion
' ============================================================================
Logout
Response.Redirect AppRoot() & "login.asp"
Response.End
%>
