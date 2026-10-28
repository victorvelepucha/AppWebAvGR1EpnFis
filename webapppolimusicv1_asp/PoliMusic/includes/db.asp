<%
' ============================================================================
'  PoliMusic - Capa de acceso a datos (ADODB / OLE DB)
'  Todas las consultas usan parametros -> proteccion contra inyeccion SQL.
' ============================================================================

' Tipos de datos ADO
Const adCmdText        = 1
Const adParamInput     = 1
Const adVarWChar       = 202
Const adUseClient      = 3
Const adLockReadOnly   = 1
Const adOpenStatic     = 3
Const adInteger        = 3

' ============================================================================
'  GetConnection
'  Abre una conexion a SQL Server probando varios proveedores OLE DB, ya que
'  la disponibilidad depende de lo que este instalado en el servidor IIS.
'  El proveedor que funciona se memoriza en Application para no reintentar.
' ============================================================================
Function GetConnection()
    Dim cn, arrProviders, i, sProv

    arrProviders = Array("MSOLEDBSQL", "MSOLEDBSQL19", "SQLNCLI11", "SQLNCLI10", "SQLOLEDB")

    ' 1) Intentar con el proveedor ya memorizado
    sProv = Application("PM_DB_PROVIDER")
    If Len(sProv) > 0 Then
        Set cn = TryConnect(sProv)
        If Not cn Is Nothing Then
            Set GetConnection = cn
            Exit Function
        End If
        Application("PM_DB_PROVIDER") = ""
    End If

    ' 2) Recorrer la lista de proveedores conocidos
    For i = 0 To UBound(arrProviders)
        Set cn = TryConnect(arrProviders(i))
        If Not cn Is Nothing Then
            Application("PM_DB_PROVIDER") = arrProviders(i)
            Set GetConnection = cn
            Exit Function
        End If
    Next

    ' 3) Ninguno funciono
    Err.Raise vbObjectError + 9001, "PoliMusic.DB", _
        "No se pudo establecer conexion con SQL Server (" & DB_SERVER & "/" & DB_NAME & "). " & _
        "Verifique que el servicio este activo, que las credenciales sean correctas y que " & _
        "exista un proveedor OLE DB para SQL Server instalado."
End Function

' Intenta abrir una conexion con un proveedor concreto. Devuelve Nothing si falla.
Function TryConnect(sProvider)
    Dim cn
    On Error Resume Next
    Set cn = Server.CreateObject("ADODB.Connection")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Set TryConnect = Nothing
        Exit Function
    End If
    cn.ConnectionTimeout = 15
    cn.CommandTimeout = 60
    cn.Open ConnectionString(sProvider)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        cn.Close
        Set cn = Nothing
        Set TryConnect = Nothing
        Exit Function
    End If
    On Error GoTo 0
    Set TryConnect = cn
End Function

' ============================================================================
'  BuildCommand - crea un ADODB.Command con los parametros indicados.
'  prm puede ser Empty (sin parametros) o un Array de valores.
'  Todos los parametros se envian como nvarchar; las conversiones explicitas
'  (CONVERT/NULLIF) se escriben en el SQL cuando se necesitan otros tipos.
' ============================================================================
Function BuildCommand(cn, sql, prm)
    Dim cmd, i, v, n
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = adCmdText
    cmd.CommandText = sql
    cmd.CommandTimeout = 60

    If IsArray(prm) Then
        For i = 0 To UBound(prm)
            v = prm(i)
            If IsNull(v) Or IsEmpty(v) Then
                n = 4000
                cmd.Parameters.Append cmd.CreateParameter("p" & i, adVarWChar, adParamInput, n, Null)
            Else
                v = CStr(v)
                n = Len(v)
                If n < 1 Then n = 1
                cmd.Parameters.Append cmd.CreateParameter("p" & i, adVarWChar, adParamInput, n, v)
            End If
        Next
    End If

    Set BuildCommand = cmd
End Function

' ============================================================================
'  DbSelect - ejecuta un SELECT y devuelve un RecordSet desconectado.
'  El llamador debe cerrarlo con rs.Close.
' ============================================================================
Function DbSelect(sql, prm)
    Dim cn, cmd, rs
    Set cn = GetConnection()
    Set cmd = BuildCommand(cn, sql, prm)
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic, adLockReadOnly
    Set rs.ActiveConnection = Nothing   ' desconectar el cursor
    cn.Close
    Set DbSelect = rs
End Function

' ============================================================================
'  DbExecute - ejecuta INSERT / UPDATE / DELETE. Devuelve filas afectadas.
' ============================================================================
Function DbExecute(sql, prm)
    Dim cn, cmd, n
    Set cn = GetConnection()
    Set cmd = BuildCommand(cn, sql, prm)
    cmd.Execute n, , 129   ' adExecuteNoRecords
    cn.Close
    DbExecute = n
End Function

' Igual que DbExecute pero devuelve el valor de una columna con OUTPUT / SCOPE_IDENTITY
Function DbExecuteScalar(sql, prm)
    Dim cn, cmd, rs, v
    Set cn = GetConnection()
    Set cmd = BuildCommand(cn, sql, prm)
    Set rs = cmd.Execute
    v = Null
    If Not rs.EOF Then v = rs(0).Value
    rs.Close
    cn.Close
    DbExecuteScalar = v
End Function

' Devuelve el primer valor de la primera fila (o Null)
Function DbScalar(sql, prm)
    Dim rs, v
    Set rs = DbSelect(sql, prm)
    v = Null
    If Not rs.EOF Then v = rs(0).Value
    rs.Close
    DbScalar = v
End Function

' Devuelve True si la consulta devuelve al menos una fila
Function DbExists(sql, prm)
    Dim rs, b
    Set rs = DbSelect(sql, prm)
    b = Not rs.EOF
    rs.Close
    DbExists = b
End Function

' Ejecuta una sentencia que devuelve un unico entero (COUNT, MAX, ...)
Function DbCount(sql, prm)
    Dim v
    v = DbScalar(sql, prm)
    If IsNull(v) Or IsEmpty(v) Then
        DbCount = 0
    Else
        DbCount = CLng(v)
    End If
End Function
%>
