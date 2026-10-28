<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' ============================================================================
'  songs.asp - Listado completo de canciones con orden y paginacion
' ============================================================================
Dim rs, o, nTotal, page, pageSize, offset, sSort, sOrder, sQ, sLike, sWhere

RequireLogin

page     = ReqPage()
pageSize = ROWS_PER_PAGE
offset   = (page - 1) * pageSize
sSort    = LCase(Nz(Request.QueryString("sort"), "plays"))
sQ       = Nz(Request.QueryString("q"), "")
sLike    = "%" & sQ & "%"

If sSort = "name" Then
    sOrder = "S.SONG_NAME ASC"
ElseIf sSort = "recent" Then
    sOrder = "S.ID_SONG DESC"
Else
    sOrder = "ISNULL(S.PLAYS,0) DESC, S.SONG_NAME ASC"
End If

If Len(sQ) > 0 Then
    sWhere = "WHERE S.SONG_NAME LIKE ? "
    nTotal = DbCount("SELECT COUNT(*) FROM TBL_SONG S " & sWhere, Array(sLike))
Else
    sWhere = ""
    nTotal = DbCount("SELECT COUNT(*) FROM TBL_SONG", Empty)
End If

PageHeader "Canciones", "songs", ""
%>

<div class="page-head">
  <div>
    <h1 class="page-title">Canciones</h1>
    <p class="page-sub"><%=nTotal%> canciones disponibles en el cat&aacute;logo.</p>
  </div>
  <div class="hstack">
    <a class="chip <%=IIf(sSort="plays","is-active","")%>" href="<%=AppRoot()%>songs.asp?sort=plays">M&aacute;s escuchadas</a>
    <a class="chip <%=IIf(sSort="name","is-active","")%>" href="<%=AppRoot()%>songs.asp?sort=name">A - Z</a>
    <a class="chip <%=IIf(sSort="recent","is-active","")%>" href="<%=AppRoot()%>songs.asp?sort=recent">Recientes</a>
    <button type="button" class="btn btn-primary btn-sm" data-play-list="#allSongs"><%=Ico("play", 16)%> Reproducir todo</button>
  </div>
</div>

<%
If Len(sQ) > 0 Then
    Set rs = DbSelect(SongSelectSql() & sWhere & "ORDER BY " & sOrder & _
                      " OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY", Array(sLike))
Else
    Set rs = DbSelect(SongSelectSql() & "ORDER BY " & sOrder & _
                      " OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY", Empty)
End If

Set o = Server.CreateObject("Scripting.Dictionary")
o("showAlbum")   = True
o("showPlays")   = True
o("showAdd")     = True
o("startIndex")  = offset + 1
o("containerId") = "allSongs"
RenderTracklist rs, o
rs.Close

RenderPager nTotal, page, pageSize, "songs.asp", Array("sort", sSort, "q", sQ)
PageFooter
%>
