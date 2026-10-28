<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' ============================================================================
'  api/playlists.asp - Devuelve en JSON las playlists que el usuario puede
'  gestionar, indicando si ya contienen una cancion concreta.
'  Se usa en el modal "Agregar a playlist".
' ============================================================================
Dim idSong, rs, items, cnt, bHas

Response.ContentType = "application/json"
Response.CacheControl = "no-store"

If Not IsLoggedIn() Then
    Response.Status = "401 Unauthorized"
    Response.Write "{""ok"":false,""error"":" & J("Sesion no iniciada") & "}"
    Response.End
End If

idSong = ToInt(Request.QueryString("forSong"))

Set rs = DbSelect("SELECT P.ID_PLAYLIST, P.TITLE, P.PHOTO, P.TYPE FROM TBL_PLAYLIST P " & _
                  "WHERE P.ID_USER = ? ORDER BY P.ID_PLAYLIST DESC", Array(CurrentUserId()))

items = ""
cnt = 0
Do While Not rs.EOF
    bHas = False
    If idSong > 0 Then
        bHas = DbExists("SELECT 1 FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ? AND ID_SONG = ?", Array(rs("ID_PLAYLIST"), idSong))
    End If
    If cnt > 0 Then items = items & ","
    items = items & "{""id"":" & rs("ID_PLAYLIST") & _
            ",""title"":" & J(rs("TITLE")) & _
            ",""photo"":" & J(AssetUrl(rs("PHOTO"), IMG_FALLBACK_COVER)) & _
            ",""type"":" & J(IIf(ToInt(rs("TYPE")) = PLAYLIST_PRIVATE, "Privada", "Publica")) & _
            ",""songs"":" & DbCount("SELECT COUNT(*) FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?", Array(rs("ID_PLAYLIST"))) & _
            ",""has"":" & IIf(bHas, "true", "false") & "}"
    cnt = cnt + 1
    rs.MoveNext
Loop
rs.Close

Response.Write "{""ok"":true,""count"":" & cnt & ",""items"":[" & items & "]}"
Response.End
%>
