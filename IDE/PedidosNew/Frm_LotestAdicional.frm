VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form Frm_LotestAdicional 
   Caption         =   "Listado de Precios por Exportaciones"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3120
      TabIndex        =   1
      Top             =   4350
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"Frm_LotestAdicional.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4215
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   7435
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      SelectionStyle  =   1
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "Frm_LotestAdicional.frx":008D
      Column(2)       =   "Frm_LotestAdicional.frx":0155
      FormatStylesCount=   6
      FormatStyle(1)  =   "Frm_LotestAdicional.frx":01F9
      FormatStyle(2)  =   "Frm_LotestAdicional.frx":0331
      FormatStyle(3)  =   "Frm_LotestAdicional.frx":03E1
      FormatStyle(4)  =   "Frm_LotestAdicional.frx":0495
      FormatStyle(5)  =   "Frm_LotestAdicional.frx":056D
      FormatStyle(6)  =   "Frm_LotestAdicional.frx":0625
      ImageCount      =   0
      PrinterProperties=   "Frm_LotestAdicional.frx":0705
   End
End
Attribute VB_Name = "Frm_LotestAdicional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public sCod_Cliente   As String

Public sCod_PurOrd    As String

Public sCod_LotPurOrd As String

Public sCod_EstCli    As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, _
                                   ByVal ActionType As Integer, _
                                   ByVal ActionName As String)

    Select Case ActionName

        Case "REVISAR"

            If GridEX1.RowCount = 0 Then Exit Sub
            Revisar

        Case "SALIR"
            Unload Me
    End Select

End Sub

Private Sub Revisar()

    On Error GoTo ERR_REVISAR

    Dim strSql As String

    If GridEX1.RowCount > 0 Then
        GridEX1.Row = 1

        For i = 1 To GridEX1.RowCount

            If GridEX1.value(GridEX1.Columns("Sel").Index) <> "0" Then
                strSql = "EXEC TG_MAN_lotest_adicional '" & sCod_Cliente & "','" & sCod_PurOrd & "','" & sCod_LotPurOrd & "','" & sCod_EstCli & "','S','" & vusu & "'"
                ExecuteCommandSQL cCONNECT, strSql
            End If

            GridEX1.MoveNext
        Next

        GridEX1.Row = 1
    End If
    
    Cargar
    
    Exit Sub

ERR_REVISAR:
    MsgBox Err.Description
End Sub

Public Sub Cargar()

    On Error GoTo ERR_CARGAR

    Dim strSql As String

    strSql = "EXEC TG_MUESTRA_LOTEST_ADICIONAL '" & sCod_Cliente & "','" & sCod_PurOrd & "','" & sCod_LotPurOrd & "','" & sCod_EstCli & "'"
    
    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSql, cCONNECT)
    
    GridEX1.Columns("SEL").ColumnType = jgexCheckBox
    GridEX1.Columns("SEL").Visible = True
    GridEX1.Columns("SEL").EditType = jgexEditCheckBox
    GridEX1.Columns("SEl").Width = 500
    
    Exit Sub

ERR_CARGAR:
    MsgBox Err.Description
End Sub

Private Sub GridEX1_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)

    If GridEX1.col = GridEX1.Columns("SEL").Index Then
        GridEX1.AllowEdit = True
    Else
        GridEX1.AllowEdit = False
    End If

End Sub
