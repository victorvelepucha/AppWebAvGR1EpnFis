<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' ============================================================================
'  admin/edit.asp - Formulario generico de alta / edicion
'  Parametros: e (entidad), id (0 o vacio = crear)
' ============================================================================
Dim e, d, id, rs, flds, i, f, sVal, sValores, sAction, bNew, sExtra, bPassword

RequireAdmin

e = LCase(Req("e"))
If Not ValidEntity(e) Then
    Response.Redirect AppRoot() & "admin/default.asp"
    Response.End
End If

Set d  = GetEntity(e)
id    = ReqInt("id")
bNew  = (id <= 0)

sValores = ""   ' Dictionary con los valores actuales
Set rs = Nothing

If Not bNew Then
    ' Se leen todas las columnas de la tabla para poder editar cualquier campo
    Set rs = DbSelect("SELECT * FROM " & d("table") & " WHERE " & d("pk") & " = ?", Array(id))
    If rs.EOF Then
        rs.Close
        FlashSet "error", "El registro solicitado no existe."
        Response.Redirect AppRoot() & "admin/list.asp?e=" & e
        Response.End
    End If
End If

' Mensaje de error de validacion reenviado por action.asp
sExtra = Nz(Session("PM_FORM_ERROR"), "")
Session("PM_FORM_ERROR") = ""

' Valores enviados previamente (para no perder lo escrito si hubo un error)
Dim oPrev
Set oPrev = Nothing
If IsObject(Session("PM_FORM_VALUES")) Then
    Set oPrev = Session("PM_FORM_VALUES")
    Set Session("PM_FORM_VALUES") = Nothing
End If

PageHeader IIf(bNew, "Nuevo " & d("singular"), "Editar " & d("singular")), "admin", ""
%>

<div class="page-head">
  <div>
    <h1 class="page-title"><%=Ico(d("icon"), 26)%> <%=IIf(bNew, "Nuevo " & d("singular"), "Editar " & d("singular"))%></h1>
    <p class="page-sub">
<% If bNew Then %>
      Completa los datos para crear un registro en <code><%=H(d("table"))%></code>.
<% Else %>
      Modificando el registro <strong>#<%=id%></strong> de <code><%=H(d("table"))%></code>.
<% End If %>
    </p>
  </div>
  <a class="btn btn-ghost" href="<%=AppRoot()%>admin/list.asp?e=<%=e%>"><%=Ico("chevron-left", 17)%> Volver al listado</a>
</div>

<% If Len(sExtra) > 0 Then %>
<div class="alert alert-error" role="alert"><span class="alert-ico">!</span><span><%=H(sExtra)%></span></div>
<% End If %>

<form method="post" action="<%=AppRoot()%>admin/action.asp" class="panel" autocomplete="off">
  <input type="hidden" name="entity" value="<%=H(e)%>">
  <input type="hidden" name="action" value="save">
<% If Not bNew Then %>
  <input type="hidden" name="id" value="<%=id%>">
<% End If %>

  <div class="panel-head">
    <span class="panel-title"><%=Ico("edit", 17)%> Datos del registro</span>
    <span class="badge"><%=H(d("table"))%></span>
  </div>

  <div class="panel-body">
    <div class="form-grid">
<%
flds = d("fields")
For i = 0 To UBound(flds)
    Set f = flds(i)
    sVal = ""

    If Not bNew Then
        On Error Resume Next
        sVal = rs(f("name")).Value
        If Err.Number <> 0 Then
            Err.Clear
            sVal = ""
        End If
        On Error GoTo 0
    End If

    ' Valores previos (tras un error de validacion) tienen prioridad
    If Not (oPrev Is Nothing) Then
        If oPrev.Exists(f("name")) Then
            If oPrev(f("name")) <> "__KEEP__" Then sVal = oPrev(f("name"))
        End If
    End If

    ' La contrasena nunca se muestra en el formulario
    If LCase(f("name")) = "password" Then sVal = ""

    ' Las fechas se formatean para el control <input type="date">
    If f("type") = "date" Then sVal = FmtDateInput(sVal)

    ' Fechas con valor por defecto al crear
    If bNew And LCase(f("name")) = "creation_date" And Len(CStr(sVal)) = 0 Then
        sVal = FmtDateInput(Now())
    End If

    RenderFieldControl f, sVal
Next
%>
    </div>
  </div>

  <div class="panel-body" style="border-top:1px solid var(--line)">
    <div class="form-actions" style="margin:0;padding:0;border:none">
      <button type="submit" class="btn btn-primary btn-lg"><%=Ico("check", 18)%> <%=IIf(bNew, "Crear registro", "Guardar cambios")%></button>
      <a class="btn btn-ghost btn-lg" href="<%=AppRoot()%>admin/list.asp?e=<%=e%>">Cancelar</a>
<% If Not bNew Then %>
      <span class="grow"></span>
      <button type="submit" class="btn btn-danger" form="frmDelete"><%=Ico("trash", 17)%> Eliminar</button>
<% End If %>
    </div>
  </div>
</form>

<% If Not bNew Then %>
<form id="frmDelete" method="post" action="<%=AppRoot()%>admin/action.asp"
      data-confirm="&iquest;Eliminar este registro definitivamente? Se quitar&aacute;n tambi&eacute;n sus relaciones.">
  <input type="hidden" name="entity" value="<%=H(e)%>">
  <input type="hidden" name="action" value="delete">
  <input type="hidden" name="id" value="<%=id%>">
  <input type="hidden" name="returnTo" value="admin/list.asp?e=<%=e%>">
</form>
<% End If %>

<%
If Not bNew Then rs.Close
PageFooter
%>
