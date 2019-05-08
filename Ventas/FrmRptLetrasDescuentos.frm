VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRptLetrasDescuentos 
   Caption         =   "Descargo de Letras en Descuento"
   ClientHeight    =   7710
   ClientLeft      =   615
   ClientTop       =   1215
   ClientWidth     =   14595
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   14595
   Begin VB.CheckBox chkCanceladas 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "Solo Canceladas"
      Height          =   255
      Left            =   11400
      TabIndex        =   7
      Top             =   240
      Width           =   1695
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   570
      Left            =   5520
      TabIndex        =   2
      Top             =   7080
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   1005
      Custom          =   $"FrmRptLetrasDescuentos.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1200
      ControlHeigth   =   550
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   1335
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14535
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ordenado Por"
         Height          =   615
         Left            =   5040
         TabIndex        =   15
         Top             =   600
         Width           =   3735
         Begin VB.OptionButton optDescuento 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&Descuento"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&Garantia"
            Height          =   255
            Left            =   1440
            TabIndex        =   17
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Cobranza"
            Height          =   255
            Left            =   2520
            TabIndex        =   16
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   4575
         Begin VB.OptionButton optGestion 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&Gestion"
            Height          =   255
            Left            =   360
            TabIndex        =   14
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optContabilidad 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&Contabilidad"
            Height          =   255
            Left            =   1560
            TabIndex        =   13
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ordenado Por"
         Height          =   615
         Left            =   8880
         TabIndex        =   8
         Top             =   600
         Width           =   4335
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Cliente"
            Height          =   255
            Left            =   3240
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Fecha Cancelacion"
            Height          =   255
            Left            =   1440
            TabIndex        =   10
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Nro Letra"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin MSComCtl2.DTPicker txtFec_Ini 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61997057
         CurrentDate     =   41012
         MinDate         =   41012
      End
      Begin MSComCtl2.DTPicker txtFec_Fin 
         Height          =   315
         Left            =   3360
         TabIndex        =   1
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61997057
         CurrentDate     =   37543
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Desde :"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Hasta :"
         Height          =   255
         Left            =   2730
         TabIndex        =   5
         Top             =   270
         Width           =   615
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5580
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1440
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   9843
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "FrmRptLetrasDescuentos.frx":00D5
      Column(2)       =   "FrmRptLetrasDescuentos.frx":019D
      FormatStylesCount=   8
      FormatStyle(1)  =   "FrmRptLetrasDescuentos.frx":0241
      FormatStyle(2)  =   "FrmRptLetrasDescuentos.frx":0379
      FormatStyle(3)  =   "FrmRptLetrasDescuentos.frx":0429
      FormatStyle(4)  =   "FrmRptLetrasDescuentos.frx":04DD
      FormatStyle(5)  =   "FrmRptLetrasDescuentos.frx":05B5
      FormatStyle(6)  =   "FrmRptLetrasDescuentos.frx":066D
      FormatStyle(7)  =   "FrmRptLetrasDescuentos.frx":074D
      FormatStyle(8)  =   "FrmRptLetrasDescuentos.frx":07F9
      ImageCount      =   0
      PrinterProperties=   "FrmRptLetrasDescuentos.frx":08A9
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   240
      Top             =   6120
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmRptLetrasDescuentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, strOrden As String, strStatus As String

Private Sub Form_Load()
  txtFec_Ini = Date
  txtFec_Fin = Date
  strOrden = "L"
  strStatus = "D"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "BUSCAR"
  If Format(txtFec_Ini, "mm") <> Format(txtFec_Fin, "mm") Or Format(txtFec_Ini, "yyyy") <> Format(txtFec_Fin, "yyyy") Then
    MsgBox "La fecha de inicio y fin deben ser del mismo mes y año"
    Exit Sub
  End If
  
  If Format(txtFec_Ini, "dd") <> "01" Then
    MsgBox "La fecha de Inicio debe ser el primer dia del mes"
    Exit Sub
  End If
  
'  If Format(txtFec_Fin, "dd") <> "30" Or Format(txtFec_Fin, "dd") <> "31" Or Format(txtFec_Fin, "dd") <> "28" Or Format(txtFec_Fin, "dd") <> "29" Then
'    MsgBox "La fecha de Fin debe ser el último dia del mes"
'    Exit Sub
'  End If

  If DevuelveCampo("select DBO.tg_obtiene_dia_ultimo_ano_mes  ('" & Year(txtFec_Fin) & "','" & Format(Month(txtFec_Fin), "00") & "')", cCONNECT) <> txtFec_Fin Then
    MsgBox "La fecha de Fin debe ser el último dia del mes"
    Exit Sub
  End If
    
  Call CARGA_GRID
Case "IMPRIMIR"
  Call Reporte
Case "SALIR"
  Unload Me
End Select
End Sub

Sub CARGA_GRID()

Dim oGroup As GridEX20.JSGroup

On Error GoTo errCarga

strSQL = "Ventas_Muestra_Descargos_Letras_Descuentos_prueba '" & txtFec_Ini & "','" & txtFec_Fin & "'," & IIf(chkCanceladas, 1, 0) & ",'" & strOrden & "','" & strStatus & "','" & Format(txtFec_Fin, "yyyy") & "','" & Format(txtFec_Fin, "mm") & "'"

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

GridEX1.ColumnHeaderHeight = 500

Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Banco").Index, jgexSortAscending)

GridEX1.Columns("Banco").Visible = False
GridEX1.Columns("Nro_Letra").Width = 975
GridEX1.Columns("Ruc").Width = 1245
GridEX1.Columns("Cliente").Width = 3615
GridEX1.Columns("Fec_Cancel").Width = 1095
GridEX1.Columns("Fec_VenDoc").Width = 1095
GridEX1.Columns("Moneda").Width = 555
GridEX1.Columns("Tipo_Cambio").Width = 765
GridEX1.Columns("Importe_Saldo_An").Width = 960
GridEX1.Columns("Importe_Saldo_An").Format = "###,###.00"
GridEX1.Columns("Pago_Amortizacion").Width = 1050
GridEX1.Columns("Pago_Amortizacion").Format = "###,###.00"
GridEX1.Columns("Saldo_Letra").Width = 1050
GridEX1.Columns("Saldo_Letra").Format = "###,###.00"
GridEX1.Columns("Condicion").Width = 1065
GridEX1.Columns("Num_Letra_Banco").Width = 1500


GridEX1.DefaultGroupMode = jgexDGMExpanded

GridEX1.BackColorRowGroup = &H80000005

MuestraSubTotales

Exit Sub
errCarga:
    ErrorHandler err, "Carga Grid"
End Sub

Private Sub MuestraSubTotales()

Dim colTemp As JSColumn

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Fec_VenDoc")
colTemp.AggregateFunction = jgexAggregateNone
colTemp.TotalRowPrefix = "SUB TOTAL "

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Importe_Saldo_An")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Pago_Amortizacion")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Saldo_Letra")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

End Sub


Sub Reporte()
On Error GoTo hand
Dim oo As Object, strSubTitle As String, strTitle As String
Dim strSQL As String
Dim sEmpresa As String
    strSQL = "SELECT DES_EMPRESA FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA='" & vemp & "'"
    sEmpresa = DevuelveCampo(strSQL, cCONNECT)

If GridEX1.RowCount = 0 Then Exit Sub

Set oo = CreateObject("excel.application")

oo.Workbooks.Open vRuta & "\RptLetrasenDescuento.xlt"
oo.Visible = True
oo.displayalerts = False

If Month(txtFec_Ini) = Month(txtFec_Ini) Then
  strSubTitle = " DE " & Format(txtFec_Ini, "MMMM")
Else
  strSubTitle = " DESDE EL " & txtFec_Ini & " HASTA EL " & txtFec_Fin
End If

If strStatus = "D" Then
  strTitle = " DESCUENTO "
ElseIf strStatus = "G" Then
  strTitle = " COBRANZA GARANTIA "
Else
  strTitle = " COBRANZA LIBRE "
End If
oo.Run "reporte", GridEX1.ADORecordset, strTitle & IIf(optGestion, "GESTION", "CONTABILIDAD") & strSubTitle, sEmpresa

Set oo = Nothing

Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub

Private Sub optDescuento_Click()
  strStatus = "D"
End Sub

Private Sub Option1_Click()
  strOrden = "L"
End Sub

Private Sub Option2_Click()
  strOrden = "F"
End Sub

Private Sub Option3_Click()
  strOrden = "C"
End Sub

Private Sub Option4_Click()
  strStatus = "B"
End Sub

Private Sub Option5_Click()
  strStatus = "G"
End Sub

Private Sub txtFec_Ini_Change()
  txtFec_Fin = txtFec_Ini
End Sub
