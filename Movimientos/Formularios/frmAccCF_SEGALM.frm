VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAccCF_SEGALM 
   Caption         =   "Acceso por Almacen (CF)"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   2430
      Left            =   3780
      TabIndex        =   5
      Top             =   1650
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   4286
      Custom          =   $"frmAccCF_SEGALM.frx":0000
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   600
      ControlHeigth   =   600
      ControlSeparator=   0
   End
   Begin VB.ComboBox cboAlmacen 
      Height          =   315
      Left            =   1425
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   270
      Width           =   1695
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   6915
      TabIndex        =   1
      Top             =   5130
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   900
      Custom          =   "0~0~SALIR~True~True~&Salir~0~0~1~~0~False~False~&Salir~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX gexLgconAcc 
      Height          =   4080
      Left            =   4530
      TabIndex        =   0
      Top             =   795
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   7197
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmAccCF_SEGALM.frx":00C5
      Column(2)       =   "frmAccCF_SEGALM.frx":018D
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAccCF_SEGALM.frx":0231
      FormatStyle(2)  =   "frmAccCF_SEGALM.frx":0369
      FormatStyle(3)  =   "frmAccCF_SEGALM.frx":0419
      FormatStyle(4)  =   "frmAccCF_SEGALM.frx":04CD
      FormatStyle(5)  =   "frmAccCF_SEGALM.frx":05A5
      FormatStyle(6)  =   "frmAccCF_SEGALM.frx":065D
      ImageCount      =   0
      PrinterProperties=   "frmAccCF_SEGALM.frx":073D
   End
   Begin GridEX20.GridEX gexLgsinAcc 
      Height          =   4080
      Left            =   45
      TabIndex        =   4
      Top             =   795
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   7197
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmAccCF_SEGALM.frx":0915
      Column(2)       =   "frmAccCF_SEGALM.frx":09DD
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAccCF_SEGALM.frx":0A81
      FormatStyle(2)  =   "frmAccCF_SEGALM.frx":0BB9
      FormatStyle(3)  =   "frmAccCF_SEGALM.frx":0C69
      FormatStyle(4)  =   "frmAccCF_SEGALM.frx":0D1D
      FormatStyle(5)  =   "frmAccCF_SEGALM.frx":0DF5
      FormatStyle(6)  =   "frmAccCF_SEGALM.frx":0EAD
      ImageCount      =   0
      PrinterProperties=   "frmAccCF_SEGALM.frx":0F8D
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   2670
      Top             =   5205
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Almacen"
      Height          =   255
      Left            =   300
      TabIndex        =   3
      Top             =   315
      Width           =   1095
   End
End
Attribute VB_Name = "frmAccCF_SEGALM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Strsql As String, sCod_Almacen As String

Private Sub FillAlmacen()
On Error GoTo Fin
Dim rstAux As ADODB.Recordset
    Strsql = "SELECT Cod_Almacen, Nom_Almacen FROM CF_ALMACEN"
    
    cboAlmacen.Clear
    
    Set rstAux = CargarRecordSetDesconectado(Strsql, cConnect)
    With rstAux
    If .RecordCount > 0 Then .MoveFirst
    Do Until .EOF
        cboAlmacen.AddItem !Cod_Almacen & " " & !Nom_Almacen
        .MoveNext
    Loop
    .Close
    End With
    Set rstAux = Nothing
    If cboAlmacen.ListCount > 0 Then cboAlmacen.ListIndex = 0
Exit Sub
Fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Cargar Almacenes LG"
End Sub

Private Sub MostrarAccAlm()
On Error GoTo Fin

sCod_Almacen = Left(cboAlmacen, 2)

Screen.MousePointer = 11

'Usuario con acceso restringido al Almacen especificado
Strsql = "SELECT Cod_Usuario, Nom_Usuario FROM Seguridad.dbo.SEG_USUARIOS " & _
         "WHERE  Cod_Usuario NOT IN (SELECT Cod_Usuario FROM CF_SEGALM " & _
         "WHERE  Cod_Almacen = '" & sCod_Almacen & "') ORDER BY Cod_Usuario"
Set gexLgsinAcc.ADORecordset = CargarRecordSetDesconectado(Strsql, cConnect)
gexLgsinAcc.Columns("Cod_Usuario").Width = 1200
gexLgsinAcc.Columns("Nom_Usuario").Width = 3000
gexLgsinAcc.Columns("Cod_Usuario").Caption = "Usuario"
gexLgsinAcc.Columns("Nom_Usuario").Caption = "Nombre"

'Usuario con acceso restringido al Almacen especificado
Strsql = "SELECT a.Cod_Usuario, b.Nom_Usuario FROM CF_SEGALM a, " & _
         "Seguridad.dbo.SEG_USUARIOS b WHERE  a.Cod_Almacen = '" & sCod_Almacen & _
         "' AND a.Cod_Usuario = b.Cod_Usuario Order By a.Cod_Usuario"
Set gexLgconAcc.ADORecordset = CargarRecordSetDesconectado(Strsql, cConnect)
gexLgconAcc.Columns("Cod_Usuario").Width = 1200
gexLgconAcc.Columns("Nom_Usuario").Width = 3000
gexLgconAcc.Columns("Cod_Usuario").Caption = "Usuario"
gexLgconAcc.Columns("Nom_Usuario").Caption = "Nombre"

Screen.MousePointer = 0
Exit Sub
Fin:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical + vbOKOnly, "Mostrar Access por Almacen"
End Sub

Private Sub cboAlmacen_Click()
    MostrarAccAlm
End Sub

Private Sub Form_Load()
    FillAlmacen
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Unload Me
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim iRow As Long
    Screen.MousePointer = 11
    Select Case ActionName
    Case "ADDALL"
        gexLgsinAcc.Visible = False
        For iRow = 1 To gexLgsinAcc.RowCount
            gexLgsinAcc.Row = iRow
            If Not ActAccAlm("I", gexLgsinAcc.Value(gexLgsinAcc _
            .Columns("Cod_Usuario").Index)) Then Exit For
        Next iRow
    Case "ADDONE"
        If gexLgsinAcc.RowCount = 0 Then Exit Sub
        ActAccAlm "I", gexLgsinAcc.Value(gexLgsinAcc.Columns("Cod_Usuario").Index)
    Case "DROPONE"
        If gexLgconAcc.RowCount = 0 Then Exit Sub
        If ActAccAlm("D", gexLgconAcc.Value(gexLgsinAcc _
        .Columns("Cod_Usuario").Index)) Then
            MostrarAccAlm
        End If
    Case "DROPALL"
        gexLgconAcc.Visible = False
        For iRow = 1 To gexLgconAcc.RowCount
            gexLgconAcc.Row = iRow
            If Not ActAccAlm("D", gexLgconAcc.Value(gexLgsinAcc _
            .Columns("Cod_Usuario").Index)) Then Exit For
        Next iRow
    End Select
    MostrarAccAlm
    gexLgsinAcc.Visible = True
    gexLgconAcc.Visible = True
    Screen.MousePointer = 0
End Sub

Private Function ActAccAlm(Accion As String, Cod_Usuario As String) As Boolean
On Error GoTo Fin
    ActAccAlm = False
    Strsql = "EXEC UP_MAN_CF_SEGALM '" & Accion & "', '" & Cod_Usuario & _
             "', '" & sCod_Almacen & "'"
    ExecuteSQL cConnect, Strsql
    ActAccAlm = True
Exit Function
Fin:
End Function
