VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmShowGuiasxFact 
   Caption         =   "Autorización de Pago de Documentos"
   ClientHeight    =   7230
   ClientLeft      =   1095
   ClientTop       =   1380
   ClientWidth     =   11550
   Icon            =   "frmShowGuiasxFact.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   11550
   WindowState     =   2  'Maximized
   Begin GridEX20.GridEX GridEX2 
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3625
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ScrollToolTipColumn=   ""
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmShowGuiasxFact.frx":030A
      Column(2)       =   "frmShowGuiasxFact.frx":03D2
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmShowGuiasxFact.frx":0476
      FormatStyle(2)  =   "frmShowGuiasxFact.frx":05AE
      FormatStyle(3)  =   "frmShowGuiasxFact.frx":065E
      FormatStyle(4)  =   "frmShowGuiasxFact.frx":0712
      FormatStyle(5)  =   "frmShowGuiasxFact.frx":07EA
      FormatStyle(6)  =   "frmShowGuiasxFact.frx":08A2
      ImageCount      =   0
      PrinterProperties=   "frmShowGuiasxFact.frx":0982
   End
   Begin GridEX20.GridEX GridEX3 
      Height          =   2055
      Left            =   2880
      TabIndex        =   7
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   3625
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ScrollToolTipColumn=   ""
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmShowGuiasxFact.frx":0B5A
      Column(2)       =   "frmShowGuiasxFact.frx":0C22
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmShowGuiasxFact.frx":0CC6
      FormatStyle(2)  =   "frmShowGuiasxFact.frx":0DFE
      FormatStyle(3)  =   "frmShowGuiasxFact.frx":0EAE
      FormatStyle(4)  =   "frmShowGuiasxFact.frx":0F62
      FormatStyle(5)  =   "frmShowGuiasxFact.frx":103A
      FormatStyle(6)  =   "frmShowGuiasxFact.frx":10F2
      ImageCount      =   0
      PrinterProperties=   "frmShowGuiasxFact.frx":11D2
   End
   Begin VB.Frame FraBuscar 
      Caption         =   "Argumentos de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   11520
      Begin VB.ComboBox Cbo_Almacen 
         Height          =   315
         Left            =   1920
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox optTodos 
         Caption         =   "&Todos"
         Height          =   255
         Left            =   4080
         TabIndex        =   8
         Top             =   270
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpFecEmiFin 
         Height          =   315
         Left            =   3990
         TabIndex        =   2
         Top             =   650
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   65929217
         CurrentDate     =   37543
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   675
         Left            =   8280
         TabIndex        =   3
         Top             =   200
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   1191
         Custom          =   $"frmShowGuiasxFact.frx":13AA
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1000
         ControlHeigth   =   650
         ControlSeparator=   40
      End
      Begin MSComCtl2.DTPicker dtpFecEmiIni 
         Height          =   315
         Left            =   1950
         TabIndex        =   1
         Top             =   650
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   65929217
         CurrentDate     =   37543
      End
      Begin VB.Label Label1 
         Caption         =   "Rango Fecha de Emisión:"
         Height          =   360
         Left            =   90
         TabIndex        =   10
         Top             =   627
         Width           =   2355
      End
      Begin VB.Label Label2 
         Caption         =   "Almacen"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   1095
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5820
      Left            =   60
      TabIndex        =   4
      Top             =   1065
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   10266
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmShowGuiasxFact.frx":1493
      Column(2)       =   "frmShowGuiasxFact.frx":155B
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmShowGuiasxFact.frx":15FF
      FormatStyle(2)  =   "frmShowGuiasxFact.frx":1737
      FormatStyle(3)  =   "frmShowGuiasxFact.frx":17E7
      FormatStyle(4)  =   "frmShowGuiasxFact.frx":189B
      FormatStyle(5)  =   "frmShowGuiasxFact.frx":1973
      FormatStyle(6)  =   "frmShowGuiasxFact.frx":1A2B
      FormatStyle(7)  =   "frmShowGuiasxFact.frx":1B0B
      FormatStyle(8)  =   "frmShowGuiasxFact.frx":1BB7
      ImageCount      =   0
      PrinterProperties=   "frmShowGuiasxFact.frx":1C67
   End
   Begin VB.Label lbMotivo 
      AutoSize        =   -1  'True
      Caption         =   "Motivo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6240
      TabIndex        =   14
      Top             =   6960
      Width           =   690
   End
   Begin VB.Label lbDes_Motivo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   7080
      TabIndex        =   13
      Top             =   6960
      Width           =   45
   End
   Begin VB.Label lbDescripcion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1440
      TabIndex        =   12
      Top             =   6960
      Width           =   45
   End
   Begin VB.Label Descripcion 
      AutoSize        =   -1  'True
      Caption         =   "Descripcion :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   6960
      Width           =   1170
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   6435
      Top             =   4905
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmShowGuiasxFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iRowAnterior As Long
Dim iColAnterior As Long
Dim bClickColSelec As Boolean
Dim bCargaGRid As Boolean
Dim bPuedeAutorizar  As Boolean
Dim sTipoDocAutorizar As String
Dim Doc As String
Public sSQL As String



Private Sub Form_Load()
  dtpFecEmiIni.Value = Date:  dtpFecEmiIni.Value = ""
  dtpFecEmiFin.Value = Date:  dtpFecEmiFin.Value = ""
  
'  FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name) & "/SALIR"
  iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))
  
  If InStr(FunctButt1.FunctionsUser, "AUTORIZARPAGO") <> 0 Then
      bPuedeAutorizar = True
  End If
  Set GridEX2.ADORecordset = CargarRecordSetDesconectado("select Cod_CondVent, Des_CondVent from lg_condvent", cCONNECT)
  FillAlmacen
  GridEX2.ColumnAutoResize = True
'  GridEX2.ClearFields
'  GridEX2.Rebind
  
  'GridEX2 will act as the drop down list
  'for column 'SupplierID' in GridEX1
  GridEX2.ActAsDropDown = True
  GridEX2.BoundColumnIndex = 1
  GridEX2.ReplaceColumnIndex = 2
  GridEX2.Columns("Cod_CondVent").Visible = False
End Sub

Private Sub DtFecVencimiento_Change()
  GridEX1.ClearFields
  dtpFecEmiIni.Value = ""
  dtpFecEmiFin.Value = ""
End Sub

Private Sub dtpFecEmiIni_Change()
  GridEX1.ClearFields
  dtpFecEmiFin.Value = dtpFecEmiIni
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Msg As Variant
    Select Case ActionName
    Case "BUSCAR"
        Buscar
    Case "AUTORIZARPAGO"
        If GridEX1.RowCount = 0 Then Exit Sub
        Msg = MsgBox("¿Esta seguro de autorizar pago?", vbYesNo)
        If Msg = vbNo Then Exit Sub
        If Left(Cbo_Almacen, 2) = "20" Or Left(Cbo_Almacen, 2) = "30" Or Left(Cbo_Almacen, 2) = "40" Or Left(Cbo_Almacen, 2) = "41" Or Left(Cbo_Almacen, 2) = "90" Then
            Autorizar_Hilos
        ElseIf Left(Cbo_Almacen, 2) = "03" Then
            Autorizar_Hilos_ALM_03
        Else
            Autorizar_Mat_SubP
        End If
    Case "SALIR"
        Unload Me
    End Select
End Sub

Private Sub GridEX1_AfterColEdit(ByVal ColIndex As Integer)
    On Error GoTo Error_Handler
    
    Dim oGroup As GridEX20.JSGroup
    Select Case ColIndex
      Case Is = GridEX1.Columns("Sel").Index
            If Left(Cbo_Almacen, 2) = "20" Or Left(Cbo_Almacen, 2) = "30" Or Left(Cbo_Almacen, 2) = "40" Or Left(Cbo_Almacen, 2) = "41" Or Left(Cbo_Almacen, 2) = "90" Then
                Hilados
            ElseIf Left(Cbo_Almacen, 2) = "03" Then
                Hilados_ALM_03
            Else
                Mat_Prima_SubP
            End If
      Case Is = GridEX1.Columns("Precio Kg").Index
            GridEX1.Value(GridEX1.Columns("Monto Despacho").Index) = GridEX1.Value(GridEX1.Columns("Precio Kg").Index) * GridEX1.Value(GridEX1.Columns("Cantidad Despachada").Index)
            GridEX1.Value(GridEX1.Columns("sel").Index) = 0
      Case Is = GridEX1.Columns("Ser_Docum").Index
            GridEX1.Value(GridEX1.Columns("Fac_Cli").Index) = GridEX1.Value(GridEX1.Columns("Ser_Docum").Index) & GridEX1.Value(GridEX1.Columns("Num_Docum_Ventas").Index) & " " & GridEX1.Value(GridEX1.Columns("Cliente").Index)
            GridEX1.Groups.Clear
            Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Fac_Cli").Index, jgexSortAscending)
            GridEX1.Value(GridEX1.Columns("sel").Index) = 0
      Case Is = GridEX1.Columns("Num_Docum_Ventas").Index
            GridEX1.Value(GridEX1.Columns("Fac_Cli").Index) = GridEX1.Value(GridEX1.Columns("Ser_Docum").Index) & GridEX1.Value(GridEX1.Columns("Num_Docum_Ventas").Index) & " " & GridEX1.Value(GridEX1.Columns("Cliente").Index)
            GridEX1.Groups.Clear
            Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Fac_Cli").Index, jgexSortAscending)
            GridEX1.Value(GridEX1.Columns("sel").Index) = 0
      Case Is = GridEX1.Columns("Gastos Financieros").Index
            Cambio_Importe "Gastos Financieros"
            GridEX1.Value(GridEX1.Columns("sel").Index) = 0
      Case Is = GridEX1.Columns("Otros").Index
            Cambio_Importe "Otros"
            GridEX1.Value(GridEX1.Columns("sel").Index) = 0
      End Select
    Exit Sub
    Resume
    
Error_Handler:
      errores err.Number
      If ColIndex = GridEX1.Columns("Sel").Index Then
         GridEX1.Value(GridEX1.Columns("sel").Index) = 0
      End If
End Sub

Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
    Select Case ColIndex
        Case Is = GridEX1.Columns("Fecha").Index
           Cancel = False
        Case Is = GridEX1.Columns("Ser_Docum").Index
           Cancel = False
        Case Is = GridEX1.Columns("Num_Docum_Ventas").Index
           Cancel = False
        Case Is = GridEX1.Columns("SEL").Index
           Cancel = False
        Case Is = GridEX1.Columns("Precio Kg").Index
           Cancel = False
        Case Is = GridEX1.Columns("Tipo Pago").Index
           Cancel = False
        Case Is = GridEX1.Columns("Moneda").Index
           Cancel = False
        Case Is = GridEX1.Columns("Gastos Financieros").Index
           Cancel = False
        Case Is = GridEX1.Columns("Otros").Index
           Cancel = False
        Case Is = GridEX1.Columns("Und").Index
           Cancel = False
        Case Else
           Cancel = True
     End Select
End Sub

Private Sub GridEX1_Click()
'On Error Resume Next
    Dim ColIndex As Long
    Dim oRowData As JSRowData
    Dim sGrupo As String
    Dim iRow As Long
    Dim I As Long
    Dim sCaptionGroup As String
    
    bCargaGRid = True
    If GridEX1.RowCount > 0 Then
        ColIndex = GridEX1.Col
        
        If Not GridEX1.IsGroupItem(GridEX1.Row) Then
            If ColIndex = 0 Then Exit Sub
            If UCase(GridEX1.Columns(ColIndex).Key) = "SEL" Then
                bClickColSelec = True
                SendKeys "{ENTER}"
            End If
        Else
            If GridEX1.IsGroupItem(GridEX1.Row) Then
            End If
        End If
    End If
End Sub

Private Sub GridEX1_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
  Dim ocol As JSColumn
  Dim oRow As JSRowData
  Dim vCurrentRow As Variant
  Dim oRowGroup As JSRowData
  Dim sProveedor As String
  
  iColAnterior = LastCol
  iRowAnterior = LastRow
  If GridEX1.Row <> 0 Then
      Set oRow = GridEX1.GetRowData(GridEX1.Row)
  End If
  If GridEX1.RowCount > 0 Then
    On Error Resume Next
    lbDescripcion.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Descripcion").Index)), "", GridEX1.Value(GridEX1.Columns("Descripcion").Index))
    lbMotivo.Visible = False
    If Left(Cbo_Almacen, 2) = "01" Then
      lbMotivo.Visible = True
      lbDes_Motivo.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Motivo").Index)), "", GridEX1.Value(GridEX1.Columns("Motivo").Index))
    End If
  End If
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
    Dim strGroupCaption As String
    
    If RowBuffer.RowType = jgexRowTypeGroupHeader Then
        strGroupCaption = RTrim(RowBuffer.GroupCaption) & " (" & RowBuffer.RecordCount & " Documentos " & "" & ") "
        RowBuffer.GroupCaption = strGroupCaption
    End If
End Sub

Private Sub GridEX2_Click()
     Dim Serie As String, Nro_Factura As String, iPos, I As Integer, lvSw As Boolean

     GridEX1.Redraw = False
     lvSw = True
     Serie = GridEX1.Value(GridEX1.Columns("Ser_Docum").Index)
     Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Docum_Ventas").Index)
      
     GridEX1.MoveFirst
     For I = 0 To GridEX1.RowCount
        If Serie = GridEX1.Value(GridEX1.Columns("Ser_Docum").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Docum_Ventas").Index) Then
          If lvSw Then iPos = GridEX1.Row
          lvSw = False
          GridEX1.Value(GridEX1.Columns("Cod_Tip_Pago").Index) = GridEX2.Value(GridEX2.Columns("Cod_CondVent").Index)
          GridEX1.Value(GridEX1.Columns("Tipo Pago").Index) = GridEX2.Value(GridEX2.Columns("Des_CondVent").Index)
        End If
        GridEX1.MoveNext
      Next I
      GridEX1.Row = iPos
      GridEX1.Redraw = True
      SendKeys "{TAB}"
End Sub

Private Sub GridEX3_Click()
    Dim Serie As String, Nro_Factura As String, iPos, I As Integer, lvSw As Boolean

    GridEX1.Redraw = False
    Serie = GridEX1.Value(GridEX1.Columns("Ser_Docum").Index)
    Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Docum_Ventas").Index)
    lvSw = True
    GridEX1.MoveFirst
    For I = 0 To GridEX1.RowCount
      If Serie = GridEX1.Value(GridEX1.Columns("Ser_Docum").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Docum_Ventas").Index) Then
        If lvSw Then iPos = GridEX1.Row
        lvSw = False
        GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index) = GridEX3.Value(GridEX3.Columns("Cod_Moneda").Index)
        GridEX1.Value(GridEX1.Columns("Moneda").Index) = GridEX3.Value(GridEX3.Columns("Descripcion").Index)
      End If
      GridEX1.MoveNext
    Next I
    GridEX1.Row = iPos
    GridEX1.Redraw = True
    SendKeys "{TAB}"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub



'****************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************

Private Sub FillAlmacen()
    Dim rstAux As ADODB.Recordset
    Dim strSql As String
        
    strSql = "Ventas_Ayuda_Almacenes_Hilos"
             
    Set rstAux = CargarRecordSetDesconectado(strSql, cCONNECT)
    Cbo_Almacen.Clear
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
            Cbo_Almacen.AddItem !Cod_Almacen & " " & !Nom_Almacen
            .MoveNext
        Loop
        .Close
    End With
    If Cbo_Almacen.ListCount > 0 Then Cbo_Almacen.ListIndex = 0
    Set rstAux = Nothing
End Sub

Private Sub Buscar()
    On Error GoTo drDepurar
    
    Dim oGroup As GridEX20.JSGroup
    Dim oFormat As JSFormatStyle
    
    GridEX3.ColumnAutoResize = True
    
    GridEX3.ActAsDropDown = True
    GridEX3.BoundColumnIndex = 1
    GridEX3.ReplaceColumnIndex = 2
    
    GridEX1.ClearFields
    
    GridEX1.DefaultGroupMode = jgexDGMExpanded
    bCargaGRid = False
   
    
    If Left(Cbo_Almacen, 2) = "03" Then
        If optTodos Then
          sSQL = "Hilado_2004..Ventas_Muestra_Documentos_Pendientes_Facturar_Todos '" & dtpFecEmiIni & "','" & dtpFecEmiFin & "'"
        Else
          sSQL = "Hilado_2004..Ventas_Muestra_Documentos_Pendientes_Facturar '" & dtpFecEmiIni & "','" & dtpFecEmiFin & "'"
        End If
        Set GridEX3.ADORecordset = CargarRecordSetDesconectado("select sopcsopc as Cod_Moneda,soptsopc as Descripcion from Hilado_2004..conmoneda", cCONNECT)
        Formateo
    ElseIf Left(Cbo_Almacen, 2) = "20" Or Left(Cbo_Almacen, 2) = "30" Or Left(Cbo_Almacen, 2) = "40" Or Left(Cbo_Almacen, 2) = "41" Or Left(Cbo_Almacen, 2) = "90" Then
        sSQL = "EXEC HIL_VENTAS_MUESTRA_DOCUMENTOS_PENDIENTES_FACTURAR '" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','" & Left(Cbo_Almacen, 2) & "'"
        Set GridEX3.ADORecordset = CargarRecordSetDesconectado("select sopcsopc as Cod_Moneda,soptsopc as Descripcion from Hilado_2004..conmoneda", cCONNECT)
        NewFormateo
    Else
        sSQL = "Ventas_Muestra_Documentos_Pendientes_Facturar_Mat_Prima_SubP '" & Left(Cbo_Almacen, 2) & "','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "'"
        Set GridEX3.ADORecordset = CargarRecordSetDesconectado("select Cod_Moneda as cod_Moneda,Nom_Moneda as Descripcion from tg_moneda", cCONNECT)
        Formateo
    End If
    
   
    
    Exit Sub
    Resume
drDepurar:
      errores err.Number
End Sub

Sub Mat_Prima_SubP()
    Dim sSQL As String

    sSQL = "Ventas_Cambio_Estado_DocAlm_Mat_Prima_SubP '$' , '$' , '$' , '$' , '$' , $ , '$' , $ , $ , $ , $ , '$' "
    sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
                         GridEX1.Value(GridEX1.Columns("Num_Movstk").Index), _
                         GridEX1.Value(GridEX1.Columns("Ser_Docum").Index), _
                         GridEX1.Value(GridEX1.Columns("Num_Docum_Ventas").Index), _
                         GridEX1.Value(GridEX1.Columns("Cod_Tip_Pago").Index), _
                         GridEX1.Value(GridEX1.Columns("precio kg").Index), _
                         GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                         GridEX1.Value(GridEX1.Columns("Gastos Financieros").Index), _
                         GridEX1.Value(GridEX1.Columns("Otros").Index), _
                         IIf(GridEX1.Value(GridEX1.Columns("Sel").Index) = 0, 0, 1), _
                         GridEX1.Value(GridEX1.Columns("Proforma").Index), _
                         GridEX1.Value(GridEX1.Columns("Cod_Item").Index))
    ExecuteCommandSQL cCONNECT, sSQL
End Sub

Sub Hilados()
   Dim sSQL As String

   sSQL = "HIL_VENTAS_CAMBIA_ESTADO_DOCALMACEN '$' , '$' , '$' , '$' , '$' , '$' , $ , '$' ,'$','$',$,$,$,'$','$','$'"
   sSQL = VBsprintf(sSQL, GridEX1.Value(GridEX1.Columns("pedcclic").Index), _
                         GridEX1.Value(GridEX1.Columns("pedcpedc").Index), _
                         GridEX1.Value(GridEX1.Columns("Cod_Doc").Index), _
                         GridEX1.Value(GridEX1.Columns("Ser_Docum").Index), _
                         GridEX1.Value(GridEX1.Columns("Num_Docum_Ventas").Index), _
                         GridEX1.Value(GridEX1.Columns("Cod_Tip_Pago").Index), _
                         GridEX1.Value(GridEX1.Columns("precio kg").Index), _
                         GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                         GridEX1.Value(GridEX1.Columns("cod hilado").Index), _
                         GridEX1.Value(GridEX1.Columns("petchilc").Index), _
                         GridEX1.Value(GridEX1.Columns("Gastos Financieros").Index), _
                         GridEX1.Value(GridEX1.Columns("Otros").Index), _
                         IIf(GridEX1.Value(GridEX1.Columns("Sel").Index) = 0, 0, 1), _
                         GridEX1.Value(GridEX1.Columns("Fecha").Index), _
                         GridEX1.Value(GridEX1.Columns("Und").Index), _
                         Left(Cbo_Almacen, 2))
                         
   'ExecuteCommandSQL cCONNECT, sSQL
End Sub


Private Sub MuestraSubTotales()
    Dim colTemp As JSColumn
    
    GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
    Set colTemp = GridEX1.Columns("Moneda")
    colTemp.AggregateFunction = jgexAggregateNone
    colTemp.TotalRowPrefix = "SUB TOTAL "
    
    GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
    Set colTemp = GridEX1.Columns("Cantidad Despachada")
    colTemp.AggregateFunction = jgexSum
    colTemp.TotalRowPrefix = ""
    
    GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
    Set colTemp = GridEX1.Columns("Monto Despacho")
    colTemp.AggregateFunction = jgexSum
    colTemp.TotalRowPrefix = ""
End Sub

Private Sub SetColores()
    Dim fmtCon As JSFmtCondition
    Dim fmtCond2 As JSFmtCondition
    Dim fmtCond3 As JSFmtCondition
    
    Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("SEL").Index, jgexEqual, -1)
    
    With GridEX1.FmtConditions
            .ApplyGroupCondition = True
            .ShowGroupConditionCount = True
            .GroupConditionCountTitle = "Documento(s) Autorizado(s)"
            Set fmtCon = .GroupCondition
    End With
    fmtCon.SetCondition GridEX1.Columns("SEL").Index, jgexEqual, -1
    fmtCon.FormatStyle.FontBold = True
    fmtCon.FormatStyle.BackColor = &HFFFFC0   '&HC0FFC0    ' &HC0E0FF    ' '&HC0FFFF
End Sub


Private Sub Autorizar_Hilos()
    On Error GoTo errorx
    Dim sSQL As String
    Dim aMess(4), I As Integer
      
    GridEX1.MoveFirst
    
    For I = 0 To GridEX1.RowCount
      If GridEX1.Value(GridEX1.Columns("SEL").Index) Then
        sSQL = "EXEC HIL_VENTAS_CAMBIA_ESTADO_DOCALMACEN '$' , '$' , '$' , '$' , '$' , '$' , $ , '$' ,'$','$',$,$,$,'$','$','$'"
        sSQL = VBsprintf(sSQL, GridEX1.Value(GridEX1.Columns("pedcclic").Index), _
                               GridEX1.Value(GridEX1.Columns("pedcpedc").Index), _
                               GridEX1.Value(GridEX1.Columns("Cod_Doc").Index), _
                               GridEX1.Value(GridEX1.Columns("Ser_Docum").Index), _
                               GridEX1.Value(GridEX1.Columns("Num_Docum_Ventas").Index), _
                               GridEX1.Value(GridEX1.Columns("Cod_Tip_Pago").Index), _
                               GridEX1.Value(GridEX1.Columns("precio kg").Index), _
                               GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                               GridEX1.Value(GridEX1.Columns("cod hilado").Index), _
                               GridEX1.Value(GridEX1.Columns("petchilc").Index), _
                               GridEX1.Value(GridEX1.Columns("Gastos Financieros").Index), _
                               GridEX1.Value(GridEX1.Columns("Otros").Index), _
                               IIf(GridEX1.Value(GridEX1.Columns("Sel").Index) = 0, 0, 1), _
                               GridEX1.Value(GridEX1.Columns("Fecha").Index), _
                               GridEX1.Value(GridEX1.Columns("Und").Index), Left(Cbo_Almacen, 2))
        ExecuteCommandSQL cCONNECT, sSQL
      End If
      GridEX1.MoveNext
    Next I
    ExecuteCommandSQL cCONNECT, "HIL_VENTAS_GENERA_DOCUM_AUTORIZADOS '" & vusu & "','" & Left(Cbo_Almacen, 2) & "'"
    Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
    Buscar
     
    Exit Sub
    Resume
errorx:
        errores err.Number
End Sub

Private Sub Autorizar_Mat_SubP()
    On Error GoTo errorx
    Dim sSQL As String
    Dim aMess(4), I As Integer
      
    GridEX1.MoveFirst
    For I = 0 To GridEX1.RowCount
      If GridEX1.Value(GridEX1.Columns("SEL").Index) Then
    
      sSQL = "Ventas_Cambio_Estado_DocAlm_Mat_Prima_SubP '$' , '$' , '$' , '$' , '$' , $ , '$' , $ , $ , $ , $ , '$' "
      sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
                             GridEX1.Value(GridEX1.Columns("Num_Movstk").Index), _
                             GridEX1.Value(GridEX1.Columns("Ser_Docum").Index), _
                             GridEX1.Value(GridEX1.Columns("Num_Docum_Ventas").Index), _
                             GridEX1.Value(GridEX1.Columns("Cod_Tip_Pago").Index), _
                             GridEX1.Value(GridEX1.Columns("precio kg").Index), _
                             GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                             GridEX1.Value(GridEX1.Columns("Gastos Financieros").Index), _
                             GridEX1.Value(GridEX1.Columns("Otros").Index), _
                             IIf(GridEX1.Value(GridEX1.Columns("Sel").Index) = 0, 0, 1), _
                             GridEX1.Value(GridEX1.Columns("Proforma").Index), _
                             GridEX1.Value(GridEX1.Columns("Cod_Item").Index))
        ExecuteCommandSQL cCONNECT, sSQL
      End If
      GridEX1.MoveNext
    Next I
    ExecuteCommandSQL cCONNECT, "Ventas_Genera_Docum_Autorizados_Mat_Prima_SubP '" & vusu & "','" & Left(Cbo_Almacen, 2) & "'"
    Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
    Buscar
    Exit Sub
    Resume
errorx:
        errores err.Number
End Sub

Sub Cambio_Nro_Factura()
    Dim Serie As String, Nro_Factura As String, iPos, I As Integer, lvSw As Boolean

    GridEX1.Redraw = False
    lvSw = True
    Doc = GridEX1.Value(GridEX1.Columns("Cod_Doc").Index)
    Serie = GridEX1.Value(GridEX1.Columns("Ser_Docum").Index)
    Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Docum_Ventas").Index)
    GridEX1.MoveFirst
    For I = 0 To GridEX1.RowCount
        If Doc = GridEX1.Value(GridEX1.Columns("Cod_Doc").Index) Then
            If lvSw Then iPos = GridEX1.Row
            lvSw = False
            GridEX1.Value(GridEX1.Columns("Ser_Docum").Index) = Serie
            GridEX1.Value(GridEX1.Columns("Nro_Docum_Ventas").Index) = Nro_Factura
        End If
        GridEX1.MoveNext
    Next I
    GridEX1.Row = iPos
    GridEX1.Redraw = True
    SendKeys "{TAB}"
End Sub

Sub Cambio_Importe(Campo As String)
      Dim Fac_Cli As String, Importe As String, iPos, I As Integer, lvSw As Boolean

      GridEX1.Redraw = False
      lvSw = True
      Fac_Cli = GridEX1.Value(GridEX1.Columns("Fac_Cli").Index)
      Importe = GridEX1.Value(GridEX1.Columns(Campo).Index)
      GridEX1.MoveFirst
      For I = 0 To GridEX1.RowCount
        If Fac_Cli = GridEX1.Value(GridEX1.Columns("Fac_Cli").Index) Then
            If lvSw Then iPos = GridEX1.Row
            lvSw = False
            GridEX1.Value(GridEX1.Columns(Campo).Index) = Importe
        End If
        GridEX1.MoveNext
      Next I
      GridEX1.Row = iPos
      GridEX1.Redraw = True
End Sub





Sub Hilados_ALM_03()
   Dim sSQL As String

   sSQL = "Hilado_2004..Ventas_Cambio_Estado_DocAlm '$' , '$' , '$' , '$' , '$' , '$' , $ , '$' ,'$','$',$,$,$,'$','$'"
   sSQL = VBsprintf(sSQL, GridEX1.Value(GridEX1.Columns("pedcclic").Index), _
                         GridEX1.Value(GridEX1.Columns("pedcpedc").Index), _
                         GridEX1.Value(GridEX1.Columns("Cod_Doc").Index), _
                         GridEX1.Value(GridEX1.Columns("Ser_Docum").Index), _
                         GridEX1.Value(GridEX1.Columns("Num_Docum_Ventas").Index), _
                         GridEX1.Value(GridEX1.Columns("Cod_Tip_Pago").Index), _
                         GridEX1.Value(GridEX1.Columns("precio kg").Index), _
                         GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                         GridEX1.Value(GridEX1.Columns("cod hilado").Index), _
                         GridEX1.Value(GridEX1.Columns("petchilc").Index), _
                         GridEX1.Value(GridEX1.Columns("Gastos Financieros").Index), _
                         GridEX1.Value(GridEX1.Columns("Otros").Index), _
                         IIf(GridEX1.Value(GridEX1.Columns("Sel").Index) = 0, 0, 1), _
                         GridEX1.Value(GridEX1.Columns("Fecha").Index), _
                         GridEX1.Value(GridEX1.Columns("Und").Index))
   ExecuteCommandSQL cCONNECT, sSQL
End Sub


Private Sub Autorizar_Hilos_ALM_03()
    On Error GoTo errorx
    Dim sSQL As String
    Dim aMess(4), I As Integer
      
    GridEX1.MoveFirst
    
    For I = 0 To GridEX1.RowCount
      If GridEX1.Value(GridEX1.Columns("SEL").Index) Then
        sSQL = " Hilado_2004..Ventas_Cambio_Estado_DocAlm '$' , '$' , '$' , '$' , '$' , '$' , $ , '$' ,'$','$',$,$,$,'$','$'"
        sSQL = VBsprintf(sSQL, GridEX1.Value(GridEX1.Columns("pedcclic").Index), _
                               GridEX1.Value(GridEX1.Columns("pedcpedc").Index), _
                               GridEX1.Value(GridEX1.Columns("Cod_Doc").Index), _
                               GridEX1.Value(GridEX1.Columns("Ser_Docum").Index), _
                               GridEX1.Value(GridEX1.Columns("Num_Docum_Ventas").Index), _
                               GridEX1.Value(GridEX1.Columns("Cod_Tip_Pago").Index), _
                               GridEX1.Value(GridEX1.Columns("precio kg").Index), _
                               GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                               GridEX1.Value(GridEX1.Columns("cod hilado").Index), _
                               GridEX1.Value(GridEX1.Columns("petchilc").Index), _
                               GridEX1.Value(GridEX1.Columns("Gastos Financieros").Index), _
                               GridEX1.Value(GridEX1.Columns("Otros").Index), _
                               IIf(GridEX1.Value(GridEX1.Columns("Sel").Index) = 0, 0, 1), _
                               GridEX1.Value(GridEX1.Columns("Fecha").Index), _
                               GridEX1.Value(GridEX1.Columns("Und").Index))
        ExecuteCommandSQL cCONNECT, sSQL
      End If
      GridEX1.MoveNext
    Next I
    ExecuteCommandSQL cCONNECT, "HIL_VENTAS_GENERA_DOCUM_AUTORIZADOS '" & vusu & "','" & Left(Cbo_Almacen, 2) & "'"
    Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
    Buscar
     
    Exit Sub
    Resume
errorx:
        errores err.Number
End Sub


Sub NewFormateo()
    On Error GoTo drDepurar
        
    Dim oGroup As GridEX20.JSGroup
    Dim oFormat As JSFormatStyle
     
    GridEX3.Columns("Cod_Moneda").Visible = False
     
    GridEX3.ColumnAutoResize = True
    
    GridEX3.ActAsDropDown = True
    GridEX3.BoundColumnIndex = 1
    GridEX3.ReplaceColumnIndex = 2
    
    GridEX3.Columns("Cod_Moneda").Visible = False
    GridEX1.ClearFields
    
    GridEX1.DefaultGroupMode = jgexDGMExpanded
    bCargaGRid = False
    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)
     
    
    Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Fac_Cli").Index, jgexSortAscending)
    
    MuestraSubTotales
    GridEX1.BackColorRowGroup = &H80000005
    
    GridEX1.ColumnHeaderHeight = 500
    
    GridEX1.Columns("fecha").Width = 975
    GridEX1.Columns("Ser_Docum").Width = 435
    GridEX1.Columns("Num_Docum_Ventas").Width = 965
    GridEX1.Columns("cliente").Width = 0
    GridEX1.Columns("nro_Guia").Width = 1125
    GridEX1.Columns("descripcion").Width = 3690
    GridEX1.Columns("moneda").Width = 870
    GridEX1.Columns("precio kg").Width = 840
    GridEX1.Columns("cantidad despachada").Width = 1050
    GridEX1.Columns("monto despacho").Width = 855
    GridEX1.Columns("SEL").Width = 495
    GridEX1.Columns("Fac_Cli").Width = 0
    GridEX1.Columns("Gastos Financieros").Width = 900
    GridEX1.Columns("otros").Width = 810
    GridEX1.Columns("Tipo Pago").Width = 960
    GridEX1.Columns("Und").Width = 375
    If Left(Cbo_Almacen, 2) = "20" Or Left(Cbo_Almacen, 2) = "30" Or Left(Cbo_Almacen, 2) = "40" Or Left(Cbo_Almacen, 2) = "41" Or Left(Cbo_Almacen, 2) = "90" Then
        'VR
        GridEX1.Columns("OC HILADO").Width = 915
        GridEX1.Columns("Nro Pedido/OC").Width = 0
        'VR
        GridEX1.Columns("cod hilado").Width = 1110
        GridEX1.Columns("Conos").Width = 525
        GridEX1.Columns("Cod_Doc").Visible = False
        GridEX1.Columns("parxmotc").Visible = False
        GridEX1.Columns("pedcclic").Visible = False
        GridEX1.Columns("petchilc").Visible = False
        GridEX1.Columns("pedcpedc").Visible = False
    Else
        GridEX1.Columns("Proforma").Width = 945
        GridEX1.Columns("Cod_Item").Width = 780
        GridEX1.Columns("Num_Movstk").Visible = False
        GridEX1.Columns("Cod_Cliente").Visible = False
        GridEX1.Columns("Und").Visible = False
    End If
    GridEX1.Columns("nro_parte").Visible = False
    GridEX1.Columns("Cod_Tip_Pago").Visible = False
    GridEX1.Columns("Cod_Moneda").Visible = False
    
    GridEX1.Columns("Ser_Docum").Caption = "Serie"
    GridEX1.Columns("Num_Docum_Ventas").Caption = "Nro Factura"
    
    GridEX1.Columns("cantidad despachada").Format = "#######0.00"
    GridEX1.Columns("precio kg").Format = "#######0.0000"
    GridEX1.Columns("monto despacho").Format = "#######0.00"
    
    GridEX1.Columns("SEL").ColumnType = jgexCheckBox
    GridEX1.Columns("SEL").Visible = True
    GridEX1.Columns("SEL").EditType = jgexEditCheckBox
    GridEX1.Columns("SEL").Width = 500
    
    
    With GridEX1.Columns("Tipo Pago")
        .TextAlignment = jgexAlignLeft
        .EditType = jgexEditCombo
        Set .DropDownControl = GridEX2
    End With
    With GridEX1.Columns("moneda")
        .TextAlignment = jgexAlignLeft
        .EditType = jgexEditCombo
        Set .DropDownControl = GridEX3
    End With
    
    SetColores
    GridEX1.DefaultGroupMode = jgexDGMCollapsed
    
    If dtpFecEmiIni.Value <> "" Then
        GridEX1.DefaultGroupMode = jgexDGMExpanded
    End If
    
    GridEX1.ContinuousScroll = True
    Exit Sub
    Resume
drDepurar:
      errores err.Number
    
End Sub

Sub Formateo()
    On Error GoTo drDepurar
    
    Dim oGroup As GridEX20.JSGroup
    Dim oFormat As JSFormatStyle
     
    GridEX3.ColumnAutoResize = True
    
    GridEX3.ActAsDropDown = True
    GridEX3.BoundColumnIndex = 1
    GridEX3.ReplaceColumnIndex = 2
    
    GridEX3.Columns("Cod_Moneda").Visible = False
    GridEX1.ClearFields
    
    GridEX1.DefaultGroupMode = jgexDGMExpanded
    bCargaGRid = False
    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)

     
    Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Fac_Cli").Index, jgexSortAscending)
    
    MuestraSubTotales
    GridEX1.BackColorRowGroup = &H80000005
    
    GridEX1.ColumnHeaderHeight = 500
    
    GridEX1.Columns("fecha").Width = 975
    GridEX1.Columns("Ser_Docum").Width = 435
    GridEX1.Columns("Num_Docum_Ventas").Width = 965
    GridEX1.Columns("cliente").Width = 0
    GridEX1.Columns("nro_Guia").Width = 1125
    GridEX1.Columns("descripcion").Width = 3690
    GridEX1.Columns("moneda").Width = 870
    GridEX1.Columns("precio kg").Width = 840
    GridEX1.Columns("cantidad despachada").Width = 1050
    GridEX1.Columns("monto despacho").Width = 855
    GridEX1.Columns("SEL").Width = 495
    GridEX1.Columns("Fac_Cli").Width = 0
    GridEX1.Columns("Gastos Financieros").Width = 900
    GridEX1.Columns("otros").Width = 810
    GridEX1.Columns("Tipo Pago").Width = 960
    GridEX1.Columns("Und").Width = 375
    If Left(Cbo_Almacen, 2) = "03" Then
        GridEX1.Columns("Nro Pedido/OC").Width = 915
        GridEX1.Columns("cod hilado").Width = 1110
        GridEX1.Columns("Conos").Width = 525
        GridEX1.Columns("Cod_Doc").Visible = False
        GridEX1.Columns("parxmotc").Visible = False
        GridEX1.Columns("pedcclic").Visible = False
        GridEX1.Columns("petchilc").Visible = False
        GridEX1.Columns("pedcpedc").Visible = False
    Else
        GridEX1.Columns("Proforma").Width = 945
        GridEX1.Columns("Cod_Item").Width = 780
        GridEX1.Columns("Num_Movstk").Visible = False
        GridEX1.Columns("Cod_Cliente").Visible = False
        GridEX1.Columns("Und").Visible = False
    End If
    GridEX1.Columns("nro_parte").Visible = False
    GridEX1.Columns("Cod_Tip_Pago").Visible = False
    GridEX1.Columns("Cod_Moneda").Visible = False
    
    GridEX1.Columns("Ser_Docum").Caption = "Serie"
    GridEX1.Columns("Num_Docum_Ventas").Caption = "Nro Factura"
    
    GridEX1.Columns("cantidad despachada").Format = "#######0.00"
    GridEX1.Columns("precio kg").Format = "#######0.0000"
    GridEX1.Columns("monto despacho").Format = "#######0.00"
    
    GridEX1.Columns("SEL").ColumnType = jgexCheckBox
    GridEX1.Columns("SEL").Visible = True
    GridEX1.Columns("SEL").EditType = jgexEditCheckBox
    GridEX1.Columns("SEL").Width = 500
    
    
    With GridEX1.Columns("Tipo Pago")
        .TextAlignment = jgexAlignLeft
        .EditType = jgexEditCombo
        Set .DropDownControl = GridEX2
    End With
    With GridEX1.Columns("moneda")
        .TextAlignment = jgexAlignLeft
        .EditType = jgexEditCombo
        Set .DropDownControl = GridEX3
    End With
    
    SetColores
    GridEX1.DefaultGroupMode = jgexDGMCollapsed
    
    If dtpFecEmiIni.Value <> "" Then
        GridEX1.DefaultGroupMode = jgexDGMExpanded
    End If
    
    GridEX1.ContinuousScroll = True
    Exit Sub
    Resume
drDepurar:
      errores err.Number
    
End Sub
