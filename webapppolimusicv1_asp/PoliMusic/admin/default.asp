<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' ============================================================================
'  admin/default.asp - Panel de administracion (solo rol ADMIN)
' ============================================================================
Dim nUsers, nAdmins, nSongs, nAlbums, nArtists, nPl, nPub, nPriv, nPlays, rs

RequireAdmin

nUsers  = DbCount("SELECT COUNT(*) FROM TBL_USER", Empty)
nAdmins = DbCount("SELECT COUNT(*) FROM TBL_USER WHERE USER_TYPE = ?", Array(USER_TYPE_ADMIN))
nSongs  = DbCount("SELECT COUNT(*) FROM TBL_SONG", Empty)
nAlbums = DbCount("SELECT COUNT(*) FROM TBL_ALBUM", Empty)
nArtists= DbCount("SELECT COUNT(*) FROM TBL_ARTIST", Empty)
nPl     = DbCount("SELECT COUNT(*) FROM TBL_PLAYLIST", Empty)
nPub    = DbCount("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE TYPE = ?", Array(PLAYLIST_PUBLIC))
nPriv   = DbCount("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE TYPE = ?", Array(PLAYLIST_PRIVATE))
nPlays  = DbCount("SELECT ISNULL(SUM(PLAYS),0) FROM TBL_SONG", Empty)

PageHeader "Panel de administracion", "admin", ""
%>

<div class="page-head">
  <div>
    <h1 class="page-title">Panel de administraci&oacute;n</h1>
    <p class="page-sub">Gesti&oacute;n completa de usuarios, canciones, &aacute;lbumes, artistas y playlists de <%=APP_NAME%>.</p>
  </div>
  <span class="badge badge-admin"><%=Ico("shield", 14)%> <%=H(CurrentRoleLabel())%></span>
</div>

<section class="section">
  <div class="stat-grid">
<%
RenderStat "users", nUsers, "Usuarios"
RenderStat "shield", nAdmins, "Administradores"
RenderStat "music", nSongs, "Canciones"
RenderStat "album", nAlbums, "&Aacute;lbumes"
RenderStat "artist", nArtists, "Artistas"
RenderStat "playlist", nPl, "Playlists"
RenderStat "trend", FmtNumber(nPlays), "Reproducciones totales"
%>
  </div>
</section>

<section class="section">
  <div class="section-head"><h2 class="section-title">Operaciones CRUD</h2></div>
  <div class="grid grid-wide">
<%
Dim ent, i, arrEnt
arrEnt = Array("users", "songs", "albums", "artists", "playlists")
For i = 0 To UBound(arrEnt)
    ent = arrEnt(i)
%>
    <article class="card">
      <div class="hstack mb-2">
        <span class="stat-ico" style="margin:0"><%=Ico(EntityIcon(ent), 20)%></span>
        <div class="grow">
          <h3 class="card-title" style="margin:0"><%=H(EntityTitle(ent))%></h3>
          <p class="card-sub" style="margin:0"><%=EntityRowCount(GetEntity(ent), "")%> registros</p>
        </div>
      </div>
      <div class="hstack-tight">
        <a class="btn btn-secondary btn-sm" href="<%=AppRoot()%>admin/list.asp?e=<%=ent%>"><%=Ico("eye", 15)%> Ver listado</a>
        <a class="btn btn-primary btn-sm" href="<%=AppRoot()%>admin/edit.asp?e=<%=ent%>"><%=Ico("plus", 15)%> Nuevo</a>
      </div>
    </article>
<%
Next
%>
  </div>
</section>

<div class="grid grid-2">
  <section class="panel">
    <div class="panel-head">
      <span class="panel-title"><%=Ico("chart", 17)%> Canciones m&aacute;s reproducidas</span>
      <a class="section-link" href="<%=AppRoot()%>admin/list.asp?e=songs&amp;sort=plays">Ver todas</a>
    </div>
    <div class="panel-body panel-body-flush">
<%
Set rs = DbSelect("SELECT TOP 5 ID_SONG, SONG_NAME, PLAYS FROM TBL_SONG ORDER BY ISNULL(PLAYS,0) DESC", Empty)
If rs.EOF Then
    Response.Write "<p class=""muted"" style=""padding:16px 20px"">No hay canciones registradas.</p>"
Else
    Response.Write "<div class=""mini-list"" style=""max-height:none"">"
    Do While Not rs.EOF
        Response.Write "<a class=""mini-item"" href=""" & AppRoot() & "song.asp?id=" & rs("ID_SONG") & """>" & _
            "<span class=""mini-item-body""><span class=""mini-item-title"">" & H(rs("SONG_NAME")) & "</span>" & _
            "<span class=""mini-item-sub"">" & FmtNumber(rs("PLAYS")) & " reproducciones</span></span>" & _
            Ico("trend", 16) & "</a>"
        rs.MoveNext
    Loop
    Response.Write "</div>"
End If
rs.Close
%>
    </div>
  </section>

  <section class="panel">
    <div class="panel-head">
      <span class="panel-title"><%=Ico("users", 17)%> Usuarios recientes</span>
      <a class="section-link" href="<%=AppRoot()%>admin/list.asp?e=users">Ver todos</a>
    </div>
    <div class="panel-body panel-body-flush">
<%
Set rs = DbSelect("SELECT TOP 5 ID_USER, USERNAME, EMAIL, USER_PHOTO, USER_TYPE FROM TBL_USER ORDER BY ID_USER DESC", Empty)
If rs.EOF Then
    Response.Write "<p class=""muted"" style=""padding:16px 20px"">No hay usuarios registrados.</p>"
Else
    Response.Write "<div class=""mini-list"" style=""max-height:none"">"
    Do While Not rs.EOF
        Response.Write "<a class=""mini-item"" href=""" & AppRoot() & "admin/edit.asp?e=users&amp;id=" & rs("ID_USER") & """>" & _
            "<img src=""" & HA(AssetUrl(rs("USER_PHOTO"), IMG_FALLBACK_AVATAR)) & """ alt="""" style=""border-radius:50%"" loading=""lazy"">" & _
            "<span class=""mini-item-body""><span class=""mini-item-title"">" & H(rs("USERNAME")) & "</span>" & _
            "<span class=""mini-item-sub"">" & H(rs("EMAIL")) & "</span></span>" & _
            "<span class=""badge " & IIf(ToInt(rs("USER_TYPE")) = USER_TYPE_ADMIN, "badge-admin", "badge-user") & """>" & _
            H(CatalogLabel("1", CStr(ToInt(rs("USER_TYPE"))))) & "</span></a>"
        rs.MoveNext
    Loop
    Response.Write "</div>"
End If
rs.Close
%>
    </div>
  </section>
</div>

<section class="section">
  <div class="panel">
    <div class="panel-head"><span class="panel-title"><%=Ico("info", 17)%> Estado del sistema</span></div>
    <div class="panel-body">
      <div class="form-grid">
        <div class="field">
          <span class="label">Base de datos</span>
          <span><%=H(DB_NAME)%> en <%=H(DB_SERVER)%></span>
        </div>
        <div class="field">
          <span class="label">Proveedor OLE DB</span>
          <span><%=H(Nz(Application("PM_DB_PROVIDER"), "no detectado"))%></span>
        </div>
        <div class="field">
          <span class="label">Versi&oacute;n de la aplicaci&oacute;n</span>
          <span><%=APP_VERSION%></span>
        </div>
        <div class="field">
          <span class="label">Playlists p&uacute;blicas / privadas</span>
          <span><%=nPub%> / <%=nPriv%></span>
        </div>
      </div>
    </div>
  </div>
</section>

<%
PageFooter
%>
