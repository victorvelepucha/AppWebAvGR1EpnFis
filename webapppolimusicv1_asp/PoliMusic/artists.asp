<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' ============================================================================
'  artists.asp - Listado completo de artistas
' ============================================================================
Dim rs, nTotal, page, pageSize, offset, sSort, sOrder

RequireLogin

page     = ReqPage()
pageSize = ROWS_PER_PAGE
offset   = (page - 1) * pageSize
sSort    = LCase(Nz(Request.QueryString("sort"), "name"))

If sSort = "songs" Then
    sOrder = "(SELECT COUNT(*) FROM TBL_PLAYER P WHERE P.ID_ARTIST = A.ID_ARTIST) DESC, A.ARTIST_NAME"
ElseIf sSort = "recent" Then
    sOrder = "A.ID_ARTIST DESC"
Else
    sOrder = "A.ARTIST_NAME, A.ARTIST_LASTNAME"
End If

nTotal = DbCount("SELECT COUNT(*) FROM TBL_ARTIST", Empty)

PageHeader "Artistas", "artists", ""
%>

<div class="page-head">
  <div>
    <h1 class="page-title">Artistas</h1>
    <p class="page-sub"><%=nTotal%> artistas registrados en <%=APP_NAME%>.</p>
  </div>
  <div class="hstack">
    <a class="chip <%=IIf(sSort="name","is-active","")%>" href="<%=AppRoot()%>artists.asp?sort=name">A - Z</a>
    <a class="chip <%=IIf(sSort="songs","is-active","")%>" href="<%=AppRoot()%>artists.asp?sort=songs">Con m&aacute;s canciones</a>
    <a class="chip <%=IIf(sSort="recent","is-active","")%>" href="<%=AppRoot()%>artists.asp?sort=recent">Recientes</a>
  </div>
</div>

<%
Set rs = DbSelect("SELECT A.ID_ARTIST, A.ARTIST_NAME, A.ARTIST_LASTNAME, A.ARTIST_IMAGE FROM TBL_ARTIST A " & _
                  "ORDER BY " & sOrder & " OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY", Empty)
RenderArtistCards rs
rs.Close
RenderPager nTotal, page, pageSize, "artists.asp", Array("sort", sSort)
PageFooter
%>
