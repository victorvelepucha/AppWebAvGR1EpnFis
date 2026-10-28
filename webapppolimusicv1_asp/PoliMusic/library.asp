<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' ============================================================================
'  library.asp - Biblioteca de playlists del usuario
'  Pestanas: Mis playlists / Playlists publicas
' ============================================================================
Dim sTab, rs, nMine, nPublic, sWhere, nTotal, page, pageSize, offset

RequireLogin

sTab     = LCase(Nz(Request.QueryString("tab"), "mine"))
If sTab <> "public" Then sTab = "mine"
page     = ReqPage()
pageSize = ROWS_PER_PAGE
offset   = (page - 1) * pageSize

nMine   = DbCount("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_USER = ?", Array(CurrentUserId()))
nPublic = DbCount("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE TYPE = ?", Array(PLAYLIST_PUBLIC))

PageHeader "Tu biblioteca", "library", ""
%>

<div class="page-head">
  <div>
    <h1 class="page-title">Tu biblioteca</h1>
    <p class="page-sub">
<% If IsAdmin() Then %>
      Como administrador puedes gestionar todas las playlists del sistema.
<% Else %>
      Consulta, crea, edita y elimina tus playlists. Tambi&eacute;n puedes explorar las playlists p&uacute;blicas de otros usuarios.
<% End If %>
    </p>
  </div>
  <a class="btn btn-primary" href="<%=AppRoot()%>playlist.asp?action=new"><%=Ico("plus", 17)%> Nueva playlist</a>
</div>

<nav class="tabs">
  <a class="tab <%=IIf(sTab="mine","is-active","")%>" href="<%=AppRoot()%>library.asp?tab=mine">
    <%=IIf(IsAdmin(), "Todas las playlists", "Mis playlists")%> <span class="muted">(<%=IIf(IsAdmin(), DbCount("SELECT COUNT(*) FROM TBL_PLAYLIST", Empty), nMine)%>)</span>
  </a>
  <a class="tab <%=IIf(sTab="public","is-active","")%>" href="<%=AppRoot()%>library.asp?tab=public">
    Playlists p&uacute;blicas <span class="muted">(<%=nPublic%>)</span>
  </a>
</nav>

<%
If sTab = "mine" Then
    If IsAdmin() Then
        sWhere = ""
        nTotal = DbCount("SELECT COUNT(*) FROM TBL_PLAYLIST", Empty)
    Else
        sWhere = "WHERE P.ID_USER = " & CurrentUserId()
        nTotal = nMine
    End If
    If nTotal = 0 Then
        RenderEmpty "playlist", "A&uacute;n no tienes playlists", "Crea tu primera playlist para organizar tus canciones favoritas.", "playlist.asp?action=new", "Crear mi primera playlist"
    Else
        Set rs = DbSelect("SELECT P.ID_PLAYLIST, P.TITLE, P.PHOTO, P.TYPE, P.ID_USER, U.USERNAME AS OWNER_NAME " & _
                          "FROM TBL_PLAYLIST P LEFT JOIN TBL_USER U ON U.ID_USER = P.ID_USER " & sWhere & _
                          " ORDER BY P.ID_PLAYLIST DESC OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY", Empty)
        RenderPlaylistCards rs, IIf(IsAdmin(), True, False)
        rs.Close
        RenderPager nTotal, page, pageSize, "library.asp", Array("tab", "mine")
    End If
Else
    nTotal = nPublic
    If nTotal = 0 Then
        RenderEmpty "globe", "No hay playlists p&uacute;blicas", "Todav&iacute;a ning&uacute;n usuario ha compartido una playlist p&uacute;blica.", "", ""
    Else
        Set rs = DbSelect("SELECT P.ID_PLAYLIST, P.TITLE, P.PHOTO, P.TYPE, P.ID_USER, U.USERNAME AS OWNER_NAME " & _
                          "FROM TBL_PLAYLIST P LEFT JOIN TBL_USER U ON U.ID_USER = P.ID_USER " & _
                          "WHERE P.TYPE = ? ORDER BY P.ID_PLAYLIST DESC OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY", _
                          Array(PLAYLIST_PUBLIC))
        RenderPlaylistCards rs, True
        rs.Close
        RenderPager nTotal, page, pageSize, "library.asp", Array("tab", "public")
    End If
End If

PageFooter
%>
