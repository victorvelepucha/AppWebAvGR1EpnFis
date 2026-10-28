<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' ============================================================================
'  api/playlist_song.asp - Agrega una cancion a una playlist (POST, JSON).
'  Solo el propietario de la playlist (o un administrador) puede hacerlo.
' ============================================================================
Dim idPl, idSong

Response.ContentType = "application/json"
Response.CacheControl = "no-store"

If Not IsLoggedIn() Then
    Response.Status = "401 Unauthorized"
    Response.Write "{""ok"":false,""error"":" & J("Sesion no iniciada") & "}"
    Response.End
End If

If Not IsPost() Then
    Response.Status = "405 Method Not Allowed"
    Response.Write "{""ok"":false,""error"":" & J("Se requiere una peticion POST") & "}"
    Response.End
End If

idPl   = ToInt(Request.Form("playlistId"))
idSong = ToInt(Request.Form("songId"))

If idPl <= 0 Or idSong <= 0 Then
    Response.Status = "400 Bad Request"
    Response.Write "{""ok"":false,""error"":" & J("Parametros incompletos") & "}"
    Response.End
End If

If Not CanManagePlaylist(idPl) Then
    Response.Status = "403 Forbidden"
    Response.Write "{""ok"":false,""error"":" & J("No tienes permiso para modificar esta playlist") & "}"
    Response.End
End If

If Not DbExists("SELECT 1 FROM TBL_SONG WHERE ID_SONG = ?", Array(idSong)) Then
    Response.Status = "404 Not Found"
    Response.Write "{""ok"":false,""error"":" & J("La cancion no existe") & "}"
    Response.End
End If

If DbExists("SELECT 1 FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ? AND ID_SONG = ?", Array(idPl, idSong)) Then
    Response.Write "{""ok"":true,""already"":true,""message"":" & J("La cancion ya estaba en la playlist") & "}"
    Response.End
End If

On Error Resume Next
DbExecute "INSERT INTO TBL_PLAYLIST_SONG (ID_PLAYLIST, ID_SONG) VALUES (?, ?)", Array(idPl, idSong)
If Err.Number <> 0 Then
    Response.Status = "500 Internal Server Error"
    Response.Write "{""ok"":false,""error"":" & J("No se pudo agregar la cancion: " & Err.Description) & "}"
    Err.Clear
    Response.End
End If
On Error GoTo 0

Response.Write "{""ok"":true,""already"":false,""playlistId"":" & idPl & ",""songId"":" & idSong & "}"
Response.End
%>
