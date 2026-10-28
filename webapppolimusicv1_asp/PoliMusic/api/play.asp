<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' ============================================================================
'  api/play.asp - Incrementa el contador de reproducciones de una cancion.
'  Respuesta JSON.  Se invoca desde el reproductor al iniciar cada pista.
' ============================================================================
Dim idSong, nPlays

Response.ContentType = "application/json"
Response.CacheControl = "no-store"

If Not IsLoggedIn() Then
    Response.Status = "401 Unauthorized"
    Response.Write "{""ok"":false,""error"":" & J("Sesion no iniciada") & "}"
    Response.End
End If

idSong = ToInt(Request.QueryString("id"))
If idSong <= 0 Then
    Response.Status = "400 Bad Request"
    Response.Write "{""ok"":false,""error"":" & J("Identificador de cancion invalido") & "}"
    Response.End
End If

If Not DbExists("SELECT 1 FROM TBL_SONG WHERE ID_SONG = ?", Array(idSong)) Then
    Response.Status = "404 Not Found"
    Response.Write "{""ok"":false,""error"":" & J("La cancion no existe") & "}"
    Response.End
End If

DbExecute "UPDATE TBL_SONG SET PLAYS = ISNULL(PLAYS, 0) + 1 WHERE ID_SONG = ?", Array(idSong)

nPlays = DbCount("SELECT ISNULL(PLAYS,0) FROM TBL_SONG WHERE ID_SONG = ?", Array(idSong))

Response.Write "{""ok"":true,""id"":" & idSong & ",""plays"":" & nPlays & "}"
Response.End
%>
