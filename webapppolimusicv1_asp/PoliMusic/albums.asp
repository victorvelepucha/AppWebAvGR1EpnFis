<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' ============================================================================
'  albums.asp - Listado completo de albumes
' ============================================================================
Dim rs, nTotal, page, pageSize, offset, sSort, sOrder, sLetter

RequireLogin

page     = ReqPage()
pageSize = ROWS_PER_PAGE
offset   = (page - 1) * pageSize
sSort    = LCase(Nz(Request.QueryString("sort"), "title"))
sLetter  = Nz(Request.QueryString("letter"), "")

If sSort = "year" Then
    sOrder = "RELEASE_YEAR DESC, TITLE_ALBUM"
ElseIf sSort = "recent" Then
    sOrder = "ID_ALBUM DESC"
Else
    sOrder = "TITLE_ALBUM"
End If

nTotal = DbCount("SELECT COUNT(*) FROM TBL_ALBUM", Empty)

PageHeader "&Aacute;lbumes", "albums", ""
%>

<div class="page-head">
  <div>
    <h1 class="page-title">&Aacute;lbumes</h1>
    <p class="page-sub"><%=nTotal%> &aacute;lbumes en el cat&aacute;logo de <%=APP_NAME%>.</p>
  </div>
  <div class="hstack">
    <a class="chip <%=IIf(sSort="title","is-active","")%>" href="<%=AppRoot()%>albums.asp?sort=title">A - Z</a>
    <a class="chip <%=IIf(sSort="year","is-active","")%>" href="<%=AppRoot()%>albums.asp?sort=year">Por a&ntilde;o</a>
    <a class="chip <%=IIf(sSort="recent","is-active","")%>" href="<%=AppRoot()%>albums.asp?sort=recent">Recientes</a>
  </div>
</div>

<%
Set rs = DbSelect("SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM ORDER BY " & sOrder & _
                  " OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY", Empty)
RenderAlbumCards rs
rs.Close
RenderPager nTotal, page, pageSize, "albums.asp", Array("sort", sSort)
PageFooter
%>
