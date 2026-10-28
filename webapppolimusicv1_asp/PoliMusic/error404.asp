<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' ============================================================================
'  error404.asp - Pagina de error 404
' ============================================================================
Response.Status = "404 Not Found"
PageHeader "P&aacute;gina no encontrada", "", ""
%>

<div class="page-error" style="min-height:60vh">
  <div class="error-box">
    <div class="error-code">404</div>
    <h1>No encontramos esta p&aacute;gina</h1>
    <p>La direcci&oacute;n que solicitaste no existe o fue movida. Revisa el enlace o vuelve al inicio.</p>
    <div class="hstack" style="justify-content:center">
      <a class="btn btn-primary btn-lg" href="<%=AppRoot()%>home.asp"><%=Ico("home", 18)%> Ir al inicio</a>
      <a class="btn btn-outline btn-lg" href="<%=AppRoot()%>search.asp"><%=Ico("search", 18)%> Buscar m&uacute;sica</a>
    </div>
  </div>
</div>

<%
PageFooter
%>
