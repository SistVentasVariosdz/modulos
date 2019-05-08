VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmShowGuiasxFact_SaldosTelaTenida 
   Caption         =   "Autorización de Pago de Documentos Saldos - Tela Cruda / Teñida"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
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
      Height          =   1125
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11520
      Begin VB.CheckBox optTodos 
         Caption         =   "&Todos"
         Height          =   255
         Left            =   4080
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox Cbo_Almacen 
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dtpFecEmiIni 
         Height          =   315
         Left            =   1950
         TabIndex        =   5
         Top             =   675
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   101253121
         CurrentDate     =   37543
      End
      Begin MSComCtl2.DTPicker dtpFecEmiFin 
         Height          =   315
         Left            =   3990
         TabIndex        =   6
         Top             =   675
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   101253121
         CurrentDate     =   37543
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   525
         Left            =   8040
         TabIndex        =   7
         Top             =   360
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   926
         Custom          =   $"frmShowGuiasxFact_SaldosTelaTenida.frx":0000
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1000
         ControlHeigth   =   500
         ControlSeparator=   40
      End
      Begin VB.Label Label1 
         Caption         =   "Rango Fecha de Emisión:"
         Height          =   360
         Left            =   90
         TabIndex        =   9
         Top             =   705
         Width           =   2355
      End
      Begin VB.Label Label2 
         Caption         =   "Almacen"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin GridEX20.GridEX GridEX2 
      Height          =   2055
      Left            =   60
      TabIndex        =   0
      Top             =   4320
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
      Column(1)       =   "frmShowGuiasxFact_SaldosTelaTenida.frx":00E9
      Column(2)       =   "frmShowGuiasxFact_SaldosTelaTenida.frx":01B1
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmShowGuiasxFact_SaldosTelaTenida.frx":0255
      FormatStyle(2)  =   "frmShowGuiasxFact_SaldosTelaTenida.frx":038D
      FormatStyle(3)  =   "frmShowGuiasxFact_SaldosTelaTenida.frx":043D
      FormatStyle(4)  =   "frmShowGuiasxFact_SaldosTelaTenida.frx":04F1
      FormatStyle(5)  =   "frmShowGuiasxFact_SaldosTelaTenida.frx":05C9
      FormatStyle(6)  =   "frmShowGuiasxFact_SaldosTelaTenida.frx":0681
      ImageCount      =   0
      PrinterProperties=   "frmShowGuiasxFact_SaldosTelaTenida.frx":0761
   End
   Begin GridEX20.GridEX GridEX3 
      Height          =   2055
      Left            =   2820
      TabIndex        =   1
      Top             =   4320
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
      Column(1)       =   "frmShowGuiasxFact_SaldosTelaTenida.frx":0939
      Column(2)       =   "frmShowGuiasxFact_SaldosTelaTenida.frx":0A01
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmShowGuiasxFact_SaldosTelaTenida.frx":0AA5
      FormatStyle(2)  =   "frmShowGuiasxFact_SaldosTelaTenida.frx":0BDD
      FormatStyle(3)  =   "frmShowGuiasxFact_SaldosTelaTenida.frx":0C8D
      FormatStyle(4)  =   "frmShowGuiasxFact_SaldosTelaTenida.frx":0D41
      FormatStyle(5)  =   "frmShowGuiasxFact_SaldosTelaTenida.frx":0E19
      FormatStyle(6)  =   "frmShowGuiasxFact_SaldosTelaTenida.frx":0ED1
      ImageCount      =   0
      PrinterProperties=   "frmShowGuiasxFact_SaldosTelaTenida.frx":0FB1
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5340
      Left            =   0
      TabIndex        =   10
      Top             =   1185
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   9419
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmShowGuiasxFact_SaldosTelaTenida.frx":1189
      Column(2)       =   "frmShowGuiasxFact_SaldosTelaTenida.frx":1251
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmShowGuiasxFact_SaldosTelaTenida.frx":12F5
      FormatStyle(2)  =   "frmShowGuiasxFact_SaldosTelaTenida.frx":142D
      FormatStyle(3)  =   "frmShowGuiasxFact_SaldosTelaTenida.frx":14DD
      FormatStyle(4)  =   "frmShowGuiasxFact_SaldosTelaTenida.frx":1591
      FormatStyle(5)  =   "frmShowGuiasxFact_SaldosTelaTenida.frx":1669
      FormatStyle(6)  =   "frmShowGuiasxFact_SaldosTelaTenida.frx":1721
      FormatStyle(7)  =   "frmShowGuiasxFact_SaldosTelaTenida.frx":1801
      FormatStyle(8)  =   "frmShowGuiasxFact_SaldosTelaTenida.frx":18AD
      ImageCount      =   0
      PrinterProperties=   "frmShowGuiasxFact_SaldosTelaTenida.frx":195D
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   6375
      Top             =   4905
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tela :"
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
      Left            =   60
      TabIndex        =   24
      Top             =   6630
      Width           =   510
   End
   Begin VB.Label lbDesTela 
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
      Left            =   660
      TabIndex        =   23
      Top             =   6630
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Comb :"
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
      Left            =   5340
      TabIndex        =   22
      Top             =   6660
      Width           =   630
   End
   Begin VB.Label lbComb 
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
      Left            =   6060
      TabIndex        =   21
      Top             =   6660
      Width           =   45
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Nro Rollos :"
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
      Left            =   7500
      TabIndex        =   20
      Top             =   6660
      Width           =   1050
   End
   Begin VB.Label lbRollos 
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
      Left            =   8700
      TabIndex        =   19
      Top             =   6660
      Width           =   45
   End
   Begin VB.Label lbCalidad 
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
      Left            =   7260
      TabIndex        =   18
      Top             =   6630
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Calidad :"
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
      Left            =   6300
      TabIndex        =   17
      Top             =   6660
      Width           =   795
   End
   Begin VB.Label lbDes_Color 
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
      Left            =   9660
      TabIndex        =   16
      Top             =   6660
      Width           =   45
   End
   Begin VB.Label lbCod_Color 
      AutoSize        =   -1  'True
      Caption         =   "Color :"
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
      Left            =   8940
      TabIndex        =   15
      Top             =   6660
      Width           =   570
   End
   Begin VB.Label lbObservacion 
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
      Left            =   1380
      TabIndex        =   14
      Top             =   6960
      Width           =   45
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Observacion :"
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
      Left            =   60
      TabIndex        =   13
      Top             =   6960
      Width           =   1245
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Guia :"
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
      Left            =   8940
      TabIndex        =   12
      Top             =   6960
      Width           =   510
   End
   Begin VB.Label lbGuia 
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
      Left            =   9660
      TabIndex        =   11
      Top             =   6960
      Width           =   45
   End
End
Attribute VB_Name = "frmShowGuiasxFact_SaldosTelaTenida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim iRowAnterior As Long
'Dim iColAnterior As Long
'Dim bClickColSelec As Boolean
'Dim bCargaGRid As Boolean
'Dim bPuedeAutorizar  As Boolean
'Dim sTipoDocAutorizar As String
'Dim Doc As String
'Dim strSQL As String
'Public codigo As String
'Public Descripcion As String
'Public tipoadd As String
'Dim sCod_TipoFact  As String
'
Dim sSer_Factura_Orig As String
Dim sNum_Factura_Orig As String

Private Sub Form_Load()
  dtpFecEmiIni.Value = Date - 30
  dtpFecEmiFin.Value = Date
  FillAlmacen
  iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))
  If InStr(FunctButt1.FunctionsUser, "AUTORIZARPAGO") <> 0 Then
      bPuedeAutorizar = True
  End If
  Set GridEX2.ADORecordset = CargarRecordSetDesconectado("select Cod_CondVent,Des_CondVent as Descripcion from lg_condvent", cCONNECT)
  GridEX2.ColumnAutoResize = True
  GridEX2.ActAsDropDown = True
  GridEX2.BoundColumnIndex = 1
  GridEX2.ReplaceColumnIndex = 2
  GridEX2.Columns("Cod_CondVent").Visible = False

  Set GridEX3.ADORecordset = CargarRecordSetDesconectado("select Cod_Moneda as cod_Moneda,Nom_Moneda as Descripcion from tg_moneda", cCONNECT)
  GridEX3.ColumnAutoResize = True
  GridEX3.ActAsDropDown = True
  GridEX3.BoundColumnIndex = 1
  GridEX3.ReplaceColumnIndex = 2
  GridEX3.Columns("Cod_Moneda").Visible = False
End Sub
'
Private Sub BUSCAR()
On Error GoTo drDepurar
Dim sSQL As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle

Dim TipoPresentacion As String
TipoPresentacion = DevuelveCampo("SELECT Tip_Presentacion FROM dbo.Lg_Almacen WHERE Cod_Almacen ='" & Left(Cbo_Almacen, 3) & "'", cCONNECT)

If TipoPresentacion = "T" Then
    sSQL = "Ventas_Muestra_Documentos_Pendientes_Facturar_Saldos_Tela_Tenida '" & Left(Cbo_Almacen, 3) & "','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','" & IIf(optTodos, "*", "") & "'"
ElseIf TipoPresentacion = "C" Then
    sSQL = "Ventas_Muestra_Documentos_Pendientes_Facturar_Saldos_Tela_Cruda '" & Left(Cbo_Almacen, 3) & "','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','" & IIf(optTodos, "*", "") & "'"
End If

GridEX1.ClearFields

GridEX1.DefaultGroupMode = jgexDGMExpanded
bCargaGRid = False
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)

Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Fac_Cli").Index, jgexSortAscending)

MuestraSubTotales
GridEX1.BackColorRowGroup = &H80000005

GridEX1.ColumnHeaderHeight = 500

GridEX1.Columns("SEL").ColumnType = jgexCheckBox
GridEX1.Columns("SEL").Visible = True
GridEX1.Columns("SEL").EditType = jgexEditCheckBox
GridEX1.Columns("SEL").Width = 500

GridEX1.Columns("Fecha").Width = 1005
GridEX1.Columns("Ser_Factura").Width = 400
GridEX1.Columns("Num_Factura").Width = 1100
GridEX1.Columns("Cod_Cliente").Visible = False
GridEX1.Columns("nom_cliente").Width = 500
GridEX1.Columns("nro_Guia").Width = 1260

GridEX1.Columns("OC").Width = 800
GridEX1.Columns("Ser_Parte_Salida").Visible = False
GridEX1.Columns("Numero_Parte_Salida").Visible = False
GridEX1.Columns("Cod_Tela").Width = 870
GridEX1.Columns("Pre_Unitario").Width = 550
GridEX1.Columns("MontoDespacho").Width = 845
GridEX1.Columns("Moneda").Width = 400
GridEX1.Columns("Cod_Moneda").Visible = False
GridEX1.Columns("Sel").Width = 390
GridEX1.Columns("Fac_Cli").Width = 1110
GridEX1.Columns("Gastos_Financieros").Width = 585
GridEX1.Columns("Otros").Width = 615
GridEX1.Columns("Kgs_Movimiento").Width = 1000
GridEX1.Columns("Kgs_a_Facturar").Width = 1000
GridEX1.Columns("Observaciones").Width = 1500
GridEX1.Columns("num_movstk").Width = 750
GridEX1.Columns("Num_Secuencia").Visible = False
GridEX1.Columns("Cod_CondVent").Width = 825
GridEX1.Columns("Condicion_Venta").Width = 705
GridEX1.Columns("COD_condvent").Visible = False
GridEX1.Columns("Fecha").Caption = "Fecha"
GridEX1.Columns("Ser_Factura").Caption = "Ser/Fact"
GridEX1.Columns("Num_Factura").Caption = "N/Fact"
GridEX1.Columns("nom_cliente").Caption = "Cliente"
GridEX1.Columns("nro_Guia").Caption = "NroGuia"
GridEX1.Columns("Pre_Unitario").Caption = "Precio Unitario"
GridEX1.Columns("Moneda").Caption = "Moneda"
GridEX1.Columns("Cod_Moneda").Caption = "Moneda"
GridEX1.Columns("Sel").Caption = "Sel"
GridEX1.Columns("Fac_Cli").Caption = "Fac_Cli"
GridEX1.Columns("Otros").Caption = "Otros"
GridEX1.Columns("Observaciones").Caption = "Observaciones"
GridEX1.Columns("num_movstk").Caption = "Nro.Movstk"
GridEX1.Columns("Num_Secuencia").Caption = "Secuencia"
GridEX1.Columns("Cod_CondVent").Caption = "Cond.Vent"
GridEX1.Columns("Condicion_Venta").Caption = "Condicion.Venta"

With GridEX1.Columns("Condicion_Venta")
  .TextAlignment = jgexAlignLeft
  .EditType = jgexEditCombo
  Set .DropDownControl = GridEX2
End With

With GridEX1.Columns("moneda")
  .TextAlignment = jgexAlignLeft
  .EditType = jgexEditCombo
  Set .DropDownControl = GridEX3
End With

With GridEX1.Columns("Fecha")
  .EditType = jgexEditCalendarDropDown
End With
SetColores
GridEX1.DefaultGroupMode = jgexDGMCollapsed
If dtpFecEmiIni.Value <> "" Then
    GridEX1.DefaultGroupMode = jgexDGMExpanded
End If
If GridEX1.RowCount > 0 Then
    GridEX1.Row = 1
End If
GridEX1.ContinuousScroll = True
Exit Sub
Resume
drDepurar:
  errores err.Number
End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
'End Sub
'
Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Msg As Variant
    Select Case ActionName
    Case "BUSCAR"
      BUSCAR
    Case "AUTORIZARPAGO"
        If GridEX1.RowCount = 0 Then Exit Sub
        Msg = MsgBox("¿Esta seguro de autorizar pago?", vbYesNo)
        If Msg = vbNo Then Exit Sub
        Autorizar
    Case "SALIR"
        Unload Me
    End Select
End Sub

Private Sub GridEX1_AfterColEdit(ByVal ColIndex As Integer)
Dim TipoPresentacion As String
TipoPresentacion = DevuelveCampo("SELECT Tip_Presentacion FROM dbo.Lg_Almacen WHERE Cod_Almacen ='" & Left(Cbo_Almacen, 3) & "'", cCONNECT)

If TipoPresentacion = "T" Then
    AfterColEdit_Tenido (ColIndex)
ElseIf TipoPresentacion = "C" Then
    AfterColEdit_Crudo (ColIndex)
End If

End Sub
Sub AfterColEdit_Tenido(ByVal ColIndex As Integer)
Dim sSQL As String
On Error GoTo Error_Handler

Dim oGroup As GridEX20.JSGroup


Select Case ColIndex
  Case Is = GridEX1.Columns("Sel").Index

      sSQL = "Ventas_Cambio_Estado_DocAlm_Saldos_Tela_Tenida '$','$','$','$','$','$','$','$','$','$','$','$'"

      sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
                       GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
                       GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Fecha").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
                       GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                       GridEX1.Value(GridEX1.Columns("Gastos_Financieros").Index), _
                       GridEX1.Value(GridEX1.Columns("Otros").Index), _
                       GridEX1.Value(GridEX1.Columns("Und").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Secuencia").Index))


    ExecuteCommandSQL cCONNECT, sSQL
    SeleccionarOtrosRegTenido GridEX1.Value(GridEX1.Columns("Sel").Index)
  Case Is = GridEX1.Columns("Pre_Unitario").Index
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
    GridEX1.Value(GridEX1.Columns("MontoDespacho").Index) = GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index) * GridEX1.Value(GridEX1.Columns("Kgs_a_Facturar").Index)
  Case Is = GridEX1.Columns("Kgs_a_Facturar").Index
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
    GridEX1.Value(GridEX1.Columns("MontoDespacho").Index) = GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index) * GridEX1.Value(GridEX1.Columns("Kgs_a_Facturar").Index)
  Case Is = GridEX1.Columns("Ser_Factura").Index
    GridEX1.Value(GridEX1.Columns("Fac_Cli").Index) = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) & "-" & RPad(GridEX1.Value(GridEX1.Columns("Num_Factura").Index), 13, " ") & "  " & GridEX1.Value(GridEX1.Columns("Nom_Cliente").Index)
    GridEX1.Groups.Clear
    Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Fac_Cli").Index, jgexSortAscending)
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
    
  Case Is = GridEX1.Columns("Gastos_Financieros").Index
    Cambio_Importe "Gastos_Financieros"
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  Case Is = GridEX1.Columns("Otros").Index
    Cambio_Importe "Otros"
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  Case Is = GridEX1.Columns("Fecha").Index
    Cambio_Fecha GridEX1.Value(GridEX1.Columns("Fecha").Index)
  End Select
Exit Sub
Resume
Error_Handler:
  errores err.Number
  If ColIndex = GridEX1.Columns("Sel").Index Then
     GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  End If
End Sub '
Sub AfterColEdit_Crudo(ByVal ColIndex As Integer)

Dim sSQL As String
On Error GoTo Error_Handler

Dim oGroup As GridEX20.JSGroup


Select Case ColIndex
  Case Is = GridEX1.Columns("Sel").Index

      sSQL = "Ventas_Cambio_Estado_DocAlm_Saldos_Tela_Cruda '$','$','$','$','$','$','$','$','$','$','$','$'"

      sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
                       GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
                       GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Fecha").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
                       GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                       GridEX1.Value(GridEX1.Columns("Gastos_Financieros").Index), _
                       GridEX1.Value(GridEX1.Columns("Otros").Index), _
                       GridEX1.Value(GridEX1.Columns("Und").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Secuencia").Index))


    ExecuteCommandSQL cCONNECT, sSQL
    SeleccionarOtrosRegCrudo GridEX1.Value(GridEX1.Columns("Sel").Index)
  Case Is = GridEX1.Columns("Pre_Unitario").Index
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
    GridEX1.Value(GridEX1.Columns("MontoDespacho").Index) = GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index) * GridEX1.Value(GridEX1.Columns("Kgs_a_Facturar").Index)
  Case Is = GridEX1.Columns("Kgs_a_Facturar").Index
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
    GridEX1.Value(GridEX1.Columns("MontoDespacho").Index) = GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index) * GridEX1.Value(GridEX1.Columns("Kgs_a_Facturar").Index)
  Case Is = GridEX1.Columns("Ser_Factura").Index
    GridEX1.Value(GridEX1.Columns("Fac_Cli").Index) = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) & "-" & RPad(GridEX1.Value(GridEX1.Columns("Num_Factura").Index), 13, " ") & "  " & GridEX1.Value(GridEX1.Columns("Nom_Cliente").Index)
    GridEX1.Groups.Clear
    Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Fac_Cli").Index, jgexSortAscending)
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
    
  Case Is = GridEX1.Columns("Gastos_Financieros").Index
    Cambio_Importe "Gastos_Financieros"
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  Case Is = GridEX1.Columns("Otros").Index
    Cambio_Importe "Otros"
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  Case Is = GridEX1.Columns("Fecha").Index
    Cambio_Fecha GridEX1.Value(GridEX1.Columns("Fecha").Index)
  End Select
Exit Sub
Resume
Error_Handler:
  errores err.Number
  If ColIndex = GridEX1.Columns("Sel").Index Then
     GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  End If
End Sub
'
'
Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)

  Select Case ColIndex
    Case Is = GridEX1.Columns("Ser_Factura").Index
        sSer_Factura_Orig = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
        sNum_Factura_Orig = RPad(GridEX1.Value(GridEX1.Columns("Num_Factura").Index), 13, " ")
        Cancel = False
    Case Is = GridEX1.Columns("Num_Factura").Index
        sSer_Factura_Orig = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
        sNum_Factura_Orig = RPad(GridEX1.Value(GridEX1.Columns("Num_Factura").Index), 13, " ")
        Cancel = False
    Case Is = GridEX1.Columns("SEL").Index
      Cancel = False
    Case Is = GridEX1.Columns("Pre_Unitario").Index
      Cancel = False
    Case Is = GridEX1.Columns("Condicion_Venta").Index
      Cancel = False
    Case Is = GridEX1.Columns("Moneda").Index
      Cancel = False
   Case Is = GridEX1.Columns("Gastos_Financieros").Index
      Cancel = False
   Case Is = GridEX1.Columns("Otros").Index
      Cancel = False
   
   Case Is = GridEX1.Columns("Fecha").Index
      Cancel = False
   Case Else
      Cancel = True
    End Select
End Sub
'
Private Sub GridEX1_Click()

'On Error Resume Next
    Dim ColIndex As Long
    Dim oRowData As JSRowData
    Dim SGRUPO As String
    Dim iRow As Long
    Dim i As Long
    Dim sCaptionGroup As String

    bCargaGRid = True

        If GridEX1.RowCount > 0 Then
        ColIndex = GridEX1.Col

        If Not GridEX1.IsGroupItem(GridEX1.Row) Then
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
      lbDesTela.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Tela").Index)), "", GridEX1.Value(GridEX1.Columns("Tela").Index))
      lbComb.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Comb").Index)), "", GridEX1.Value(GridEX1.Columns("Comb").Index))
      lbCalidad.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Calidad").Index)), "", GridEX1.Value(GridEX1.Columns("Calidad").Index))
      lbRollos.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Numero_Rollos").Index)), "", GridEX1.Value(GridEX1.Columns("Numero_Rollos").Index))
      If lbCod_Color.Visible Then lbDes_Color.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Color").Index)), "", GridEX1.Value(GridEX1.Columns("Color").Index))
      lbGuia.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("nro_Guia").Index)), "", GridEX1.Value(GridEX1.Columns("nro_Guia").Index))
      lbObservacion.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Observaciones").Index)), "", GridEX1.Value(GridEX1.Columns("Observaciones").Index))
    End If
End Sub
'
Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)

Dim strGroupCaption As String

If RowBuffer.RowType = jgexRowTypeGroupHeader Then
    strGroupCaption = RTrim(RowBuffer.GroupCaption) & " (" & RowBuffer.RecordCount & " Documentos " & "" & ") "
    RowBuffer.GroupCaption = strGroupCaption
End If

End Sub
'
Private Sub MuestraSubTotales()
Dim colTemp As JSColumn

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Nom_Cliente")
colTemp.AggregateFunction = jgexAggregateNone
colTemp.TotalRowPrefix = "SUB TOTAL "

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Kgs_a_Facturar")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("MontoDespacho")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

End Sub
'
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
'
'
Private Sub Autorizar()
On Error GoTo errorx

Dim sSQL As String
Dim aMess(4), i As Integer
Dim TipoPresentacion As String

GridEX1.MoveFirst

TipoPresentacion = DevuelveCampo("SELECT Tip_Presentacion FROM dbo.Lg_Almacen WHERE Cod_Almacen ='" & Left(Cbo_Almacen, 3) & "'", cCONNECT)

For i = 1 To GridEX1.RowCount
    If GridEX1.Value(GridEX1.Columns("Sel").Index) = "" Then
        If TipoPresentacion = "T" Then
            sSQL = "Ventas_Cambio_Estado_DocAlm_Saldos_Tela_Tenida '$','$','$','$','$','$','$','$','$','$','$','$'"
            sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
            GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
            GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
            GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
            GridEX1.Value(GridEX1.Columns("Fecha").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
            GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
            GridEX1.Value(GridEX1.Columns("Gastos_Financieros").Index), _
            GridEX1.Value(GridEX1.Columns("Otros").Index), _
            GridEX1.Value(GridEX1.Columns("Und").Index), _
            GridEX1.Value(GridEX1.Columns("Num_Secuencia").Index))
        ElseIf TipoPresentacion = "C" Then
            sSQL = "Ventas_Cambio_Estado_DocAlm_Saldos_Tela_Cruda '$','$','$','$','$','$','$','$','$','$','$','$'"
            sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
            GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
            GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
            GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
            GridEX1.Value(GridEX1.Columns("Fecha").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
            GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
            GridEX1.Value(GridEX1.Columns("Gastos_Financieros").Index), _
            GridEX1.Value(GridEX1.Columns("Otros").Index), _
            GridEX1.Value(GridEX1.Columns("Und").Index), _
            GridEX1.Value(GridEX1.Columns("Num_Secuencia").Index))
        End If
        ExecuteCommandSQL cCONNECT, sSQL
    End If
    GridEX1.MoveNext
Next i

If TipoPresentacion = "T" Then
    ExecuteCommandSQL cCONNECT, "Ventas_Genera_Docum_Autorizados_Tela_Saldos_Tenida '" & vusu & "','" & Left(Cbo_Almacen, 2) & "'"
ElseIf TipoPresentacion = "C" Then
    ExecuteCommandSQL cCONNECT, "Ventas_Genera_Docum_Autorizados_Tela_Saldos_Cruda '" & vusu & "','" & Left(Cbo_Almacen, 2) & "'"
End If
Mensaje kMESSAGE_INF_PROCESS_SATISFACTO

BUSCAR

Exit Sub
Resume
errorx:
    errores err.Number
End Sub
'
'
Sub Cambio_Importe(Campo As String)

Dim Fac_Cli As String, Importe As String, iPos, i As Integer, lvSW As Boolean

  GridEX1.Redraw = False

  lvSW = True

  Fac_Cli = GridEX1.Value(GridEX1.Columns("Fac_Cli").Index)
  Importe = GridEX1.Value(GridEX1.Columns(Campo).Index)

  GridEX1.MoveFirst
  For i = 0 To GridEX1.RowCount
    If Fac_Cli = GridEX1.Value(GridEX1.Columns("Fac_Cli").Index) Then
      If lvSW Then iPos = GridEX1.Row
      lvSW = False
      GridEX1.Value(GridEX1.Columns(Campo).Index) = Importe
    End If
    GridEX1.MoveNext
  Next i

  GridEX1.Row = iPos

  GridEX1.Redraw = True

End Sub
'
'Private Sub GridEX2_Click()
'
'Dim Serie As String, Nro_Factura As String, iPos, I As Integer, lvSw As Boolean
'
'  GridEX1.Redraw = False
'
'  lvSw = True
'
'  Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
'  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)
'
'
'  GridEX1.MoveFirst
'  For I = 0 To GridEX1.RowCount
'    If Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
'      If lvSw Then iPos = GridEX1.Row
'      lvSw = False
'      GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index) = GridEX2.Value(GridEX2.Columns("Cod_CondVent").Index)
'      GridEX1.Value(GridEX1.Columns("Condicion_Venta").Index) = GridEX2.Value(GridEX2.Columns("Descripcion").Index)
'    End If
'    GridEX1.MoveNext
'  Next I
'
'  GridEX1.Row = iPos
'
'  GridEX1.Redraw = True
'
'  SendKeys "{TAB}"
'
'End Sub
'
'Private Sub GridEX3_Click()
'
'Dim Serie As String, Nro_Factura As String, iPos, I As Integer, lvSw As Boolean
'
'  GridEX1.Redraw = False
'
'  Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
'  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)
'  lvSw = True
'  GridEX1.MoveFirst
'  For I = 0 To GridEX1.RowCount
'    If Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
'      If lvSw Then iPos = GridEX1.Row
'      lvSw = False
'      GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index) = GridEX3.Value(GridEX3.Columns("Cod_Moneda").Index)
'      GridEX1.Value(GridEX1.Columns("Moneda").Index) = GridEX3.Value(GridEX3.Columns("Descripcion").Index)
'    End If
'    GridEX1.MoveNext
'  Next I
'
'  GridEX1.Row = iPos
'
'  GridEX1.Redraw = True
'
'  SendKeys "{TAB}"
'
'End Sub
'
'
Private Sub FillAlmacen()

Dim rstAux As ADODB.Recordset
Dim strSQL As String

strSQL = "Ventas_Ayuda_Almacenes_Saldos"

Set rstAux = CargarRecordSetDesconectado(strSQL, cCONNECT)
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

Private Sub SeleccionarOtrosRegTenido(Valor As Variant)
Dim Serie As String, Nro_Factura As String, iPos, i As Integer, lvSW As Boolean
Dim sSQL As String
  
  GridEX1.Redraw = False
  lvSW = True
  Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)

  GridEX1.MoveFirst
  For i = 0 To GridEX1.RowCount
    If Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
      If lvSW Then iPos = GridEX1.Row
        lvSW = False
        GridEX1.Value(GridEX1.Columns("Sel").Index) = Valor
      sSQL = "Ventas_Cambio_Estado_DocAlm_Saldos_Tela_Tenida '$','$','$','$','$','$','$','$','$','$','$','$'"
      sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
                       GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
                       GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Fecha").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
                       GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                       GridEX1.Value(GridEX1.Columns("Gastos_Financieros").Index), _
                       GridEX1.Value(GridEX1.Columns("Otros").Index), _
                       GridEX1.Value(GridEX1.Columns("Und").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Secuencia").Index))
      
      ExecuteCommandSQL cCONNECT, sSQL
    End If
    GridEX1.MoveNext
  Next i
  GridEX1.Row = iPos
  GridEX1.Redraw = True
End Sub
Private Sub SeleccionarOtrosRegCrudo(Valor As Variant)
Dim Serie As String, Nro_Factura As String, iPos, i As Integer, lvSW As Boolean
Dim sSQL As String
  
  GridEX1.Redraw = False
  lvSW = True
  Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)

  GridEX1.MoveFirst
  For i = 0 To GridEX1.RowCount
    If Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
      If lvSW Then iPos = GridEX1.Row
        lvSW = False
        GridEX1.Value(GridEX1.Columns("Sel").Index) = Valor
      sSQL = "Ventas_Cambio_Estado_DocAlm_Saldos_Tela_Cruda '$','$','$','$','$','$','$','$','$','$','$','$'"
      sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
                       GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
                       GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Fecha").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
                       GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                       GridEX1.Value(GridEX1.Columns("Gastos_Financieros").Index), _
                       GridEX1.Value(GridEX1.Columns("Otros").Index), _
                       GridEX1.Value(GridEX1.Columns("Und").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Secuencia").Index))
      
      ExecuteCommandSQL cCONNECT, sSQL
    End If
    GridEX1.MoveNext
  Next i
  GridEX1.Row = iPos
  GridEX1.Redraw = True
End Sub
'
'

'
'
'
'Private Sub txtCod_Emabarque_Venta_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        BuscaModoTransporte 1
'        SendKeys "{TAB}"
'    End If
'End Sub
'
'Public Sub BuscaModoTransporte(Opcion As String)
'Dim rstAux As ADODB.Recordset
'
'    strSQL = "SELECT Cod_Embarque, Des_Embarque FROM TG_TIPEMB WHERE "
'
'    txtCod_Embarque = Trim(txtCod_Embarque)
'    txtDes_Embarque = Trim(txtDes_Embarque)
'
'    Select Case Opcion
'    Case 1: strSQL = strSQL & "Cod_Embarque like '%" & txtCod_Embarque & "%'"
'    Case 2: strSQL = strSQL & "Des_Embarque LIKE '%" & txtDes_Embarque & "%'"
'    End Select
'
'    Set frmBusqGeneral3.oParent = Me
'    frmBusqGeneral3.SQuery = strSQL
'    frmBusqGeneral3.CARGAR_DATOS
'    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
'
'    frmBusqGeneral3.gexLista.Columns("Cod_Embarque").Width = 700
'    frmBusqGeneral3.gexLista.Columns("Des_Embarque").Width = 2000
'
'    frmBusqGeneral3.gexLista.Columns("Cod_Embarque").Caption = "Embarque"
'    frmBusqGeneral3.gexLista.Columns("Des_Embarque").Caption = "Descrip."
'
'    If frmBusqGeneral3.gexLista.RowCount > 1 Then
'        frmBusqGeneral3.Show vbModal
'    Else
'        frmBusqGeneral3.cmdAceptar.Value = True
'    End If
'
'    txtCod_Embarque = ""
'    txtDes_Embarque = ""
'
'    If codigo <> "" Then
'        txtCod_Embarque = codigo
'        txtDes_Embarque = Descripcion
'    End If
'    Unload frmBusqGeneral3
'    Set frmBusqGeneral3 = Nothing
'
'    codigo = ""
'    Descripcion = ""
'End Sub
'
'Public Function CargaValores(ByRef ObjTemp As Object) As Boolean
'    ObjTemp.TxtAbr_Cliente.Text = TxtAbr_Cliente.Text
'    ObjTemp.TxtAbr_Cliente.Tag = TxtAbr_Cliente.Tag
'    ObjTemp.txtDes_Cliente.Text = TxtNom_Cliente.Text
'    'ObjTemp.txtCOD_TEMCLI.Text = gexLista.Value(gexLista.Columns("COD_TEMCLI").Index)
'    'ObjTemp.CARGA_ESTCLI
'End Function
'
'
Private Sub Cambio_Fecha(sFecha As String)
Dim Serie As String, Nro_Factura As String, iPos, i As Integer, lvSW As Boolean
Dim sSQL As String
  GridEX1.Redraw = False

  lvSW = True

  Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)


  GridEX1.MoveFirst
  For i = 0 To GridEX1.RowCount
    If Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
      If lvSW Then iPos = GridEX1.Row
      lvSW = False
        GridEX1.Value(GridEX1.Columns("Fecha").Index) = sFecha
    End If
    GridEX1.MoveNext
  Next i

  GridEX1.Row = iPos

  GridEX1.Redraw = True

End Sub
'
'Public Sub BuscaRef_Embarque(Opcion As String)
'Dim rstAux As ADODB.Recordset
'Dim rsData As ADODB.Recordset
'
'    strSQL = "SELECT Ref_Embarque , Obs_Embarque FROM TG_EMBARQUE WHERE FLG_STATUS in ('T','F') AND COD_TIPANEX = '" & GridEX1.Value(GridEX1.Columns("COD_TIPANEX").Index) & "' AND  COD_ANXO = '" & GridEX1.Value(GridEX1.Columns("COD_ANXO").Index) & "' AND COD_CLIENTE = '" & GridEX1.Value(GridEX1.Columns("COD_CLIENTE").Index) & "' AND "
'
'    txtRef_Embarque = Trim(txtRef_Embarque)
'
'    Select Case Opcion
'    Case 1: strSQL = strSQL & "Ref_Embarque like '%" & txtRef_Embarque & "%'"
'    End Select
'
'    Set frmBusqGeneral3.oParent = Me
'    frmBusqGeneral3.SQuery = strSQL
'    frmBusqGeneral3.CARGAR_DATOS
'    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
'
'    frmBusqGeneral3.gexLista.Columns("Ref_Embarque").Width = 1700
'    frmBusqGeneral3.gexLista.Columns("obs_Embarque").Width = 2000
'
'    frmBusqGeneral3.gexLista.Columns("Ref_Embarque").Caption = "Número Embarque"
'    frmBusqGeneral3.gexLista.Columns("Obs_Embarque").Caption = "Observaciones"
'
'    If frmBusqGeneral3.gexLista.RowCount = 0 Then
'        MsgBox "Embarque no existe", 1
'        Exit Sub
'    End If
'
'    If frmBusqGeneral3.gexLista.RowCount > 1 Then
'        frmBusqGeneral3.Show vbModal
'    Else
'        frmBusqGeneral3.cmdAceptar.Value = True
'    End If
'
'    txtRef_Embarque = ""
'
'
'    If codigo <> "" Then
'        txtRef_Embarque = codigo
'        If txtRef_Embarque <> "" Then
'            strSQL = "TG_Embarques_Muestra '$','$','$','$','$','$','$'"
'            strSQL = VBsprintf(strSQL, "3", 0, txtRef_Embarque, "", "", "", "")
'            Set rsData = GetDataSet(cCONNECT, strSQL)
'            If Not rsData Is Nothing Then
'                Do While Not rsData.EOF
'                    If RTrim(txtCod_Termino_Venta) = "" Then
'                        txtCod_Termino_Venta = FixNulos(rsData("Cod_Termino_venta").Value, vbString)
'                        txtDes_Termino_Venta = FixNulos(rsData("Des_Termino_Venta").Value, vbString)
'                    End If
'                    If RTrim(txtCod_Embarque.Text) = "" Then
'                        txtCod_Embarque.Text = FixNulos(rsData("Cod_Embarque").Value, vbString)
'                        txtDes_Embarque.Text = FixNulos(rsData("Des_Embarque").Value, vbString)
'                    End If
'                    If RTrim(txtNom_Embarque.Text) = "" Then
'                        txtNom_Embarque.Text = FixNulos(rsData("Nom_Embarque").Value, vbString)
'                    End If
'
'                    rsData.MoveNext
'                Loop
'                rsData.Close
'            End If
'            Set rsData = Nothing
'
'        End If
'    End If
'    Unload frmBusqGeneral3
'    Set frmBusqGeneral3 = Nothing
'
'    codigo = ""
'    Descripcion = ""
'End Sub
'
'
'
'Private Sub Cambio_PO_Factura(sPO As String)
'Dim sSQL As String
'On Error GoTo errx
'
'    GridEX1.Value(GridEX1.Columns("Cod_PurOrd_Factura").Index) = sPO
'
'    sSQL = "UP_MAN_TEMP_Ventas_PurOrd_Factura '$','$','$','$','$',$,'$','$','$','$','$','$','$'"
'
'    sSQL = VBsprintf(sSQL, "I", vusu, Left(Cbo_Almacen, 2), _
'            GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
'            GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
'            GridEX1.Value(GridEX1.Columns("Num_Packing").Index), _
'            GridEX1.Value(GridEX1.Columns("Cod_cliente").Index), _
'            GridEX1.Value(GridEX1.Columns("Cod_PurOrd").Index), _
'            GridEX1.Value(GridEX1.Columns("Cod_LotPurOrd").Index), _
'            GridEX1.Value(GridEX1.Columns("Cod_Estcli").Index), _
'            GridEX1.Value(GridEX1.Columns("Cod_ColCli").Index), _
'            GridEX1.Value(GridEX1.Columns("Cod_Talla").Index), _
'            GridEX1.Value(GridEX1.Columns("Cod_PurOrd_Factura").Index))
'
'    ExecuteCommandSQL cCONNECT, sSQL
'
'Exit Sub
'errx:
'    errores err.Number
'
'End Sub
'
'
'

'
'

