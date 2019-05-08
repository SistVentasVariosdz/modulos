Attribute VB_Name = "Mod_Menu"
Option Explicit
Public sQuery As String

Public Sub EjecutaOpcion1(ByVal sNameOpcion As String, perfil As String, empresa As String)  'CurrentNodeKey As String)
On Error GoTo EjecutaOpcion
    Dim tDllName As String
    Dim sopcion As String
    Dim lValDev As Long
    Dim rutexe As String
    Dim nomfor As String
    Dim nivel As Integer
    Dim tipo As String
    Dim icono As String
    Dim cod_padre As String
    Dim des_opcion As String
    On Error GoTo EjecutaOpcion
    Get_Datos_form sNameOpcion, rutexe, nomfor, nivel, tipo, icono, cod_padre, des_opcion

    sopcion = tipo 'GetSubString(CurrentNodeKey, 6)

    tDllName = rutexe ' Trim(GetSubString(CurrentNodeKey, 3))

     If sopcion = "C" Or sopcion = "P" Or sopcion = "M" Then
        If sDllName <> tDllName Then
          If Not oFormObjDLL Is Nothing Then
            Set oFormObjDLL = Nothing
          End If

          If Not objFormDLL Is Nothing Then
            Set objFormDLL = Nothing
          End If
          sDllName = tDllName
          Set objFormDLL = CreateObject(sDllName & ".clsForm")
        End If

        Set oFormObjDLL = objFormDLL.GetForm(nomfor) 'Trim(GetSubString(CurrentNodeKey, 4)))
        If Not (oFormObjDLL Is Nothing) Then
            objFormDLL.Cod_Empresa = empresa
            objFormDLL.UserName = vusu
            objFormDLL.Cod_Perfil = perfil
            objFormDLL.Rutas = App.Path
            objFormDLL.ConnectEmpresa = cCONNECT
            objFormDLL.ConnectSeguridad = cSEGURIDAD
            objFormDLL.Language = iLanguage
    On Error GoTo EjecutaOpcion
            If sopcion = "M" Then
                oFormObjDLL.Show 1
            Else
                oFormObjDLL.Show 1
            End If

            Set oFormObjDLL = Nothing
        End If
    Else
    End If
     Exit Sub
EjecutaOpcion:
    ErrorHandler Err, "EjecutaOpcion"
    Set oFormObjDLL = Nothing
    'Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function Get_Datos_form1(ByVal sopcion As String, ByRef rutexe As String, ByRef nomfor As String, ByRef nivel As Integer, ByRef tipo As String, ByRef icono As String, ByRef cod_padre As String, ByRef des_opcion As String)
    Dim iCount As Integer
    Dim mRs As ADODB.Recordset

    sQuery = "SELECT isnull(RUTEXE,''),isnull(nomfor,''),isnull(nivel,0),isnull(tipo,''),isnull(icono,''),isnull(cod_padre,''),isnull(des_opcion,'') FROM SEG_OPCIONES  WHERE COD_OPCION='" & sopcion & "'"
    Set mRs = New ADODB.Recordset
    mRs.ActiveConnection = cSEGURIDAD
    mRs.CursorType = adOpenStatic
    mRs.Open sQuery
    iCount = mRs.RecordCount
    If iCount > 0 Then
       rutexe = mRs(0)
       nomfor = mRs(1)
       nivel = mRs(2)
       tipo = mRs(3)
       icono = mRs(4)
       cod_padre = mRs(5)
       des_opcion = mRs(6)
    End If
    Set mRs = Nothing
End Function


