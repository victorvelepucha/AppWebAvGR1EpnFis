<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' ============================================================================
'  admin/action.asp - Procesa las altas, modificaciones y bajas del panel
'  Recibe siempre por POST: entity, action (save|delete) e id
' ============================================================================
Dim e, d, sAction, id, flds, i, f, sVal, sErr, sReturnTo, sSql, sCols, sVals, sSet, sExpr
Dim prm, nPrm, bNew, sName, sWhere

RequireAdmin

e       = LCase(Req("entity"))
sAction = LCase(Req("action"))
id      = ReqInt("id")
sReturnTo = Req("returnTo")

If Not ValidEntity(e) Then
    FlashSet "error", "Entidad no valida."
    Response.Redirect AppRoot() & "admin/default.asp"
    Response.End
End If

Set d = GetEntity(e)
If Len(sReturnTo) = 0 Then sReturnTo = "admin/list.asp?e=" & e

' ============================================================ ELIMINAR
If sAction = "delete" Then
    sErr = EntityDelete(e, id)
    If Len(sErr) > 0 Then
        FlashSet "error", sErr
    Else
        FlashSet "ok", "El registro de " & d("singular") & " fue eliminado correctamente."
    End If
    Response.Redirect AppRoot() & sReturnTo
    Response.End
End If

If sAction <> "save" Then
    FlashSet "error", "Accion no reconocida."
    Response.Redirect AppRoot() & sReturnTo
    Response.End
End If

' ============================================================== VALIDAR
bNew = (id <= 0)
flds = d("fields")
sErr = ""

Dim oVals
Set oVals = Server.CreateObject("Scripting.Dictionary")

For i = 0 To UBound(flds)
    Set f = flds(i)
    sName = f("name")
    sVal  = Req(sName)

    ' La contrasena vacia al editar significa "no cambiar"
    If LCase(sName) = "password" And Not bNew And Len(sVal) = 0 Then
        oVals(sName) = "__KEEP__"
    Else
        oVals(sName) = sVal
    End If

    ' ---- obligatorios
    If f("required") And Len(sVal) = 0 And LCase(sName) <> "password" Then
        sErr = "El campo """ & f("label") & """ es obligatorio."
    End If

    ' ---- longitud maxima
    If Len(sErr) = 0 And ToInt(f("maxlength")) > 0 And Len(sVal) > ToInt(f("maxlength")) Then
        sErr = "El campo """ & f("label") & """ supera la longitud maxima de " & f("maxlength") & " caracteres."
    End If

    ' ---- tipos
    If Len(sErr) = 0 Then
        Select Case LCase(f("type"))
            Case "number"
                If Len(sVal) > 0 And Not IsNumeric(sVal) Then
                    sErr = "El campo """ & f("label") & """ debe ser numerico."
                End If
            Case "date"
                If Len(sVal) > 0 And Not IsValidDate(sVal) Then
                    sErr = "El campo """ & f("label") & """ no tiene una fecha valida (AAAA-MM-DD)."
                End If
            Case "select"
                If Len(sVal) > 0 And Not IsNumeric(sVal) Then
                    sErr = "El campo """ & f("label") & """ tiene un valor no valido."
                End If
        End Select
    End If
Next

' ---------------------------------------------------- reglas por entidad
If Len(sErr) = 0 And e = "users" Then
    Dim sUser : sUser = Nz(oVals("USERNAME"), "")
    Dim sMail : sMail = Nz(oVals("EMAIL"), "")
    If Not IsValidUsername(sUser) Then
        sErr = "El nombre de usuario solo admite letras, numeros, punto, guion y guion bajo (3 a 50 caracteres)."
    ElseIf Not IsValidEmail(sMail) Then
        sErr = "El correo electronico no tiene un formato valido."
    ElseIf bNew And Len(Nz(oVals("PASSWORD"), "")) = 0 Then
        sErr = "Debe indicar una contrasena para el nuevo usuario."
    ElseIf Len(Nz(oVals("PASSWORD"), "")) > 0 And Len(oVals("PASSWORD")) < 6 Then
        sErr = "La contrasena debe tener al menos 6 caracteres."
    End If
    If Len(sErr) = 0 Then
        If bNew Then
            If DbExists("SELECT 1 FROM TBL_USER WHERE USERNAME = ?", Array(sUser)) Then sErr = "Ya existe un usuario con ese nombre."
            If Len(sErr) = 0 Then
                If DbExists("SELECT 1 FROM TBL_USER WHERE EMAIL = ?", Array(sMail)) Then sErr = "Ya existe un usuario con ese correo."
            End If
        Else
            If DbExists("SELECT 1 FROM TBL_USER WHERE USERNAME = ? AND ID_USER <> ?", Array(sUser, id)) Then sErr = "Ya existe otro usuario con ese nombre."
            If Len(sErr) = 0 Then
                If DbExists("SELECT 1 FROM TBL_USER WHERE EMAIL = ? AND ID_USER <> ?", Array(sMail, id)) Then sErr = "Ya existe otro usuario con ese correo."
            End If
        End If
    End If
    ' No permitir quitarse a si mismo el rol de administrador
    If Len(sErr) = 0 And Not bNew And id = CurrentUserId() And ToInt(oVals("USER_TYPE")) <> USER_TYPE_ADMIN Then
        sErr = "No puedes quitarte a ti mismo el rol de administrador."
    End If
End If

If Len(sErr) = 0 And e = "songs" Then
    If Len(Nz(oVals("SONG_NAME"), "")) = 0 Then sErr = "El nombre de la cancion es obligatorio."
    If Len(sErr) = 0 Then
        If bNew Then
            If DbExists("SELECT 1 FROM TBL_SONG WHERE SONG_NAME = ?", Array(oVals("SONG_NAME"))) Then sErr = "Ya existe una cancion con ese nombre."
        Else
            If DbExists("SELECT 1 FROM TBL_SONG WHERE SONG_NAME = ? AND ID_SONG <> ?", Array(oVals("SONG_NAME"), id)) Then sErr = "Ya existe otra cancion con ese nombre."
        End If
    End If
End If

If Len(sErr) = 0 And e = "albums" Then
    If Len(Nz(oVals("TITLE_ALBUM"), "")) = 0 Then sErr = "El titulo del album es obligatorio."
End If

If Len(sErr) = 0 And e = "artists" Then
    If Len(Nz(oVals("ARTIST_NAME"), "")) = 0 Then sErr = "El nombre del artista es obligatorio."
End If

If Len(sErr) = 0 And e = "playlists" Then
    If Len(Nz(oVals("TITLE"), "")) = 0 Then sErr = "El titulo de la playlist es obligatorio."
    If Len(sErr) = 0 And ToInt(oVals("ID_USER")) <= 0 Then sErr = "Debe seleccionar un propietario para la playlist."
End If

' --------------------------------------------------- error de validacion
If Len(sErr) > 0 Then
    Session("PM_FORM_ERROR") = sErr
    Set Session("PM_FORM_VALUES") = oVals
    Response.Redirect AppRoot() & "admin/edit.asp?e=" & e & IIf(bNew, "", "&id=" & id)
    Response.End
End If

' ============================================================== GUARDAR
sCols = "" : sVals = "" : sSet = "" : prm = Array() : nPrm = 0

For i = 0 To UBound(flds)
    Set f = flds(i)
    sName = f("name")
    sVal  = oVals(sName)

    If sVal <> "__KEEP__" Then
        ' Expresion SQL que convierte el parametro al tipo de la columna
        Select Case LCase(f("type"))
            Case "password"
                sExpr = "UPPER(CONVERT(varchar(40), HASHBYTES('SHA1', CONVERT(varchar(200), ?)), 2))"
            Case "date"
                sExpr = "TRY_CONVERT(datetime, NULLIF(?, ''), 120)"
            Case "number", "select"
                sExpr = "CONVERT(int, NULLIF(?, ''))"
            Case Else
                sExpr = "?"
        End Select

        sCols = sCols & ", " & sName
        sVals = sVals & ", " & sExpr
        sSet  = sSet & ", " & sName & " = " & sExpr

        ReDim Preserve prm(nPrm)
        prm(nPrm) = sVal
        nPrm = nPrm + 1
    End If
Next

If Len(sCols) = 0 Then
    FlashSet "warn", "No habia cambios que guardar."
    Response.Redirect AppRoot() & sReturnTo
    Response.End
End If

If bNew Then
    sSql = "SET NOCOUNT ON; INSERT INTO " & d("table") & " (" & Mid(sCols, 3) & ") VALUES (" & Mid(sVals, 3) & ")"
Else
    ReDim Preserve prm(nPrm)
    prm(nPrm) = CStr(id)
    sSql = "UPDATE " & d("table") & " SET " & Mid(sSet, 3) & " WHERE " & d("pk") & " = ?"
End If

On Error Resume Next
DbExecute sSql, prm
If Err.Number <> 0 Then
    sErr = "No se pudo guardar el registro: " & Err.Description
    Err.Clear
    On Error GoTo 0
    Session("PM_FORM_ERROR") = sErr
    Set Session("PM_FORM_VALUES") = oVals
    Response.Redirect AppRoot() & "admin/edit.asp?e=" & e & IIf(bNew, "", "&id=" & id)
    Response.End
End If
On Error GoTo 0

If bNew Then
    FlashSet "ok", "El registro de " & d("singular") & " fue creado correctamente."
Else
    FlashSet "ok", "Los cambios fueron guardados correctamente."
End If
Response.Redirect AppRoot() & sReturnTo
Response.End
%>
