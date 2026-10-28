<!--#include file="config.asp" -->
<!--#include file="db.asp" -->
<!--#include file="util.asp" -->
<!--#include file="icons.asp" -->
<!--#include file="auth.asp" -->
<!--#include file="ui.asp" -->
<!--#include file="crud.asp" -->
<!--#include file="layout.asp" -->
<%
' ============================================================================
'  PoliMusic - Arranque comun a todas las paginas
'  (config + acceso a datos + utilidades + iconos + seguridad + UI + layout)
'  IMPORTANTE: la salida es UTF-8 y los archivos .asp se mantienen en ASCII
'  puro; los acentos se escriben con entidades HTML (&aacute;, &ntilde;, ...).
' ============================================================================

Response.Buffer    = True
Response.CodePage  = 65001
Response.Charset   = "utf-8"

On Error Resume Next
Session.CodePage = 65001
If Err.Number <> 0 Then Err.Clear
On Error GoTo 0

' Evita el cacheo de paginas dinamicas (el contenido depende de la sesion)
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "X-Content-Type-Options", "nosniff"
Response.AddHeader "X-Frame-Options", "SAMEORIGIN"
Response.AddHeader "Referrer-Policy", "strict-origin-when-cross-origin"
%>
