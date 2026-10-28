<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' ============================================================================
'  admin/list.asp - Listado generico de una entidad con busqueda y paginacion
'  Parametros: e (entidad), q (busqueda), page, sort
' ============================================================================
Dim e, d, sQ, sLike, page, pageSize, offset, nTotal, rs, cols, i, col, sOrder, sSort

RequireAdmin

e = LCase(Req("e"))
If Not ValidEntity(e) Then
    Response.Redirect AppRoot() & "admin/default.asp"
    Response.End
End If

Set d       = GetEntity(e)
sQ          = Req("q")
sLike       = "%" & sQ & "%"
page        = ReqPage()
pageSize    = ROWS_PER_PAGE
offset      = (page - 1) * pageSize
sSort       = LCase(Nz(Request.QueryString("sort"), ""))

nTotal = EntityRowCount(d, sQ)

' Orden: por nombre ascendente o por clave descendente
If sSort = "recent" Then
    sOrder = d("pk") & " DESC"
ElseIf Len(Nz(d("order"), "")) > 0 Then
    sOrder = d("order")
Else
    sOrder = d("pk")
End If

PageHeader "Administrar " & EntityTitle(e), "admin", ""
%>

<div class="page-head">
  <div>
    <h1 class="page-title"><%=Ico(d("icon"), 26)%> Administrar <%=EntityTitle(e)%></h1>
    <p class="page-sub"><%=nTotal%> registros en la tabla <code><%=H(d("table"))%></code>.</p>
  </div>
  <div class="hstack">
    <a class="btn btn-ghost" href="<%=AppRoot()%>admin/default.asp"><%=Ico("chevron-left", 17)%> Panel</a>
    <a class="btn btn-primary" href="<%=AppRoot()%>admin/edit.asp?e=<%=e%>"><%=Ico("plus", 17)%> Nuevo <%=H(d("singular"))%></a>
  </div>
</div>

<div class="panel">
  <div class="panel-head">
    <form class="hstack" method="get" action="<%=AppRoot()%>admin/list.asp" style="flex:1">
      <input type="hidden" name="e" value="<%=H(e)%>">
      <div class="topbar-search" style="max-width:360px">
        <span class="search-ico"><%=Ico("search", 17)%></span>
        <input type="search" name="q" value="<%=HA(sQ)%>" placeholder="Buscar en <%=HA(EntityTitle(e))%>...">
      </div>
      <button type="submit" class="btn btn-secondary btn-sm">Buscar</button>
<% If Len(sQ) > 0 Then %>
      <a class="btn btn-ghost btn-sm" href="<%=AppRoot()%>admin/list.asp?e=<%=e%>">Limpiar</a>
<% End If %>
    </form>
    <div class="hstack-tight">
      <a class="chip <%=IIf(sSort="","is-active","")%>" href="<%=AppRoot()%>admin/list.asp?e=<%=e%><%=IIf(Len(sQ)>0, "&amp;q=" & Server.URLEncode(sQ), "")%>">Orden por defecto</a>
      <a class="chip <%=IIf(sSort="recent","is-active","")%>" href="<%=AppRoot()%>admin/list.asp?e=<%=e%>&amp;sort=recent<%=IIf(Len(sQ)>0, "&amp;q=" & Server.URLEncode(sQ), "")%>">Recientes</a>
    </div>
  </div>

  <div class="panel-body panel-body-flush">
<%
If nTotal = 0 Then
    Response.Write "<div style=""padding:22px"">"
    RenderEmpty d("icon"), "Sin registros", IIf(Len(sQ) > 0, "No se encontraron coincidencias para la busqueda.", "Todavia no hay registros en esta tabla."), "admin/edit.asp?e=" & e, "Crear " & d("singular")
    Response.Write "</div>"
Else
    If Len(sQ) > 0 Then
        Set rs = DbSelect(EntityListSql(e) & "WHERE " & SearchWhere(d) & " ORDER BY " & sOrder & _
                          " OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY", SearchParams(d, sLike))
    Else
        Set rs = DbSelect(EntityListSql(e) & "ORDER BY " & sOrder & _
                          " OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY", Empty)
    End If

    cols = ListColumns(e)
%>
    <div class="table-wrap">
      <table class="tbl">
        <thead>
          <tr>
<%
    For i = 0 To UBound(cols)
        Set col = cols(i)
        Response.Write "<th" & IIf(col("align") = "right", " class=""tbl-num""", "") & ">" & H(col("label")) & "</th>"
    Next
%>
            <th style="text-align:right">Acciones</th>
          </tr>
        </thead>
        <tbody>
<%
    Do While Not rs.EOF
        Response.Write "<tr>"
        For i = 0 To UBound(cols)
            Set col = cols(i)
            RenderEntityCell rs, col, e
        Next
        Response.Write "<td><div class=""tbl-actions"">" & _
            "<a class=""btn btn-secondary btn-xs"" href=""" & AppRoot() & "admin/edit.asp?e=" & e & "&amp;id=" & rs(d("pk")) & """>" & Ico("edit", 14) & " Editar</a>"
        If e = "songs" Then
            Response.Write "<a class=""btn btn-ghost btn-xs"" href=""" & AppRoot() & "song.asp?id=" & rs(d("pk")) & """ title=""Ver"">" & Ico("eye", 14) & "</a>"
        ElseIf e = "albums" Then
            Response.Write "<a class=""btn btn-ghost btn-xs"" href=""" & AppRoot() & "album.asp?id=" & rs(d("pk")) & """ title=""Ver"">" & Ico("eye", 14) & "</a>"
        ElseIf e = "artists" Then
            Response.Write "<a class=""btn btn-ghost btn-xs"" href=""" & AppRoot() & "artist.asp?id=" & rs(d("pk")) & """ title=""Ver"">" & Ico("eye", 14) & "</a>"
        ElseIf e = "playlists" Then
            Response.Write "<a class=""btn btn-ghost btn-xs"" href=""" & AppRoot() & "playlist.asp?id=" & rs(d("pk")) & """ title=""Ver"">" & Ico("eye", 14) & "</a>"
        End If
        Response.Write "<form method=""post"" action=""" & AppRoot() & "admin/action.asp"" class=""inline-form"" " & _
            "data-confirm=""&iquest;Eliminar este registro? Esta accion no se puede deshacer."">" & _
            "<input type=""hidden"" name=""entity"" value=""" & e & """>" & _
            "<input type=""hidden"" name=""action"" value=""delete"">" & _
            "<input type=""hidden"" name=""id"" value=""" & rs(d("pk")) & """>" & _
            "<input type=""hidden"" name=""returnTo"" value=""admin/list.asp?e=" & e & """>" & _
            "<button type=""submit"" class=""btn btn-danger btn-xs"">" & Ico("trash", 14) & "</button></form>"
        Response.Write "</div></td></tr>"
        rs.MoveNext
    Loop
    rs.Close
%>
        </tbody>
      </table>
    </div>
<%
    Response.Write "<div style=""padding:0 20px 20px"">"
    RenderPager nTotal, page, pageSize, "admin/list.asp", Array("e", e, "q", sQ, "sort", sSort)
    Response.Write "</div>"
End If
%>
  </div>
</div>

<div class="alert alert-info mt-3">
  <span class="alert-ico">i</span>
  <span>Las operaciones de creaci&oacute;n, actualizaci&oacute;n y eliminaci&oacute;n est&aacute;n disponibles <strong>&uacute;nicamente</strong> para usuarios con rol <strong>Administrador</strong>. Al eliminar un registro se limpian autom&aacute;ticamente sus relaciones (listas de canciones, artistas y playlists).</span>
</div>

<%
PageFooter
%>
