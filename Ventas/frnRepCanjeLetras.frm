VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frnRepCanjeLetras 
   Caption         =   "Canje de Facturas x Letras por Cobrar"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14580
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   14580
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14535
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   495
         Left            =   13200
         TabIndex        =   2
         Top             =   150
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin MSComCtl2.DTPicker txtFec_Ini 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16318465
         CurrentDate     =   37543
      End
      Begin MSComCtl2.DTPicker txtFec_Fin 
         Height          =   315
         Left            =   3360
         TabIndex        =   4
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16318465
         CurrentDate     =   37543
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta :"
         Height          =   255
         Left            =   2730
         TabIndex        =   6
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Desde :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   270
         Width           =   615
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   570
      Left            =   5520
      TabIndex        =   0
      Top             =   6120
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   1005
      Custom          =   $"frnRepCanjeLetras.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1200
      ControlHeigth   =   550
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5220
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   840
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   9208
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
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
      Column(1)       =   "frnRepCanjeLetras.frx":0090
      Column(2)       =   "frnRepCanjeLetras.frx":0158
      FormatStylesCount=   8
      FormatStyle(1)  =   "frnRepCanjeLetras.frx":01FC
      FormatStyle(2)  =   "frnRepCanjeLetras.frx":0334
      FormatStyle(3)  =   "frnRepCanjeLetras.frx":03E4
      FormatStyle(4)  =   "frnRepCanjeLetras.frx":0498
      FormatStyle(5)  =   "frnRepCanjeLetras.frx":0570
      FormatStyle(6)  =   "frnRepCanjeLetras.frx":0628
      FormatStyle(7)  =   "frnRepCanjeLetras.frx":0708
      FormatStyle(8)  =   "frnRepCanjeLetras.frx":07B4
      ImageCount      =   0
      PrinterProperties=   "frnRepCanjeLetras.frx":0864
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   240
      Top             =   6120
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frnRepCanjeLetras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String

Private Sub Form_Load()
  txtFec_Ini = Date
  txtFec_Fin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
    Call Reporte
Case "SALIR"
    Unload Me
End Select
End Sub

Sub CARGA_GRID()

Dim oGroup As GridEX20.JSGroup

strSQL = "Cn_Ventas_Muestra_CANJE_Letras_x_COBRAR '" & txtFec_Ini & "','" & txtFec_Fin & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

GridEX1.ColumnHeaderHeight = 500


GridEX1.Columns("Letra").Width = 975
GridEX1.Columns("Ruc").Width = 1245
GridEX1.Columns("Cliente").Width = 3615
GridEX1.Columns("Fecha_Vencimiento").Width = 1095
GridEX1.Columns("Moneda").Width = 555
GridEX1.Columns("Tipo_Cambio").Width = 765
GridEX1.Columns("Importe").Width = 960
GridEX1.Columns("Importe").Format = "###,###.00"
GridEX1.Columns("Importe_DocCanjeado").Width = 1050
GridEX1.Columns("Importe_DocCanjeado").Format = "###,###.00"

GridEX1.DefaultGroupMode = jgexDGMExpanded

GridEX1.BackColorRowGroup = &H80000005

MuestraSubTotales

Exit Sub
errCarga:
    ErrorHandler err, "Carga Grid"
End Sub

Private Sub MuestraSubTotales()

Dim colTemp As JSColumn


End Sub


Sub Reporte()
On Error GoTo hand
Dim oo As Object
Dim strSQL As String
Dim sEmpresa As String

    strSQL = "SELECT DES_EMPRESA FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA='" & vemp & "'"
    sEmpresa = DevuelveCampo(strSQL, cCONNECT)

If GridEX1.RowCount = 0 Then Exit Sub

Set oo = CreateObject("excel.application")

oo.Workbooks.Open vRuta & "\RptCanjeLetras.XLT"
oo.Visible = True
oo.DisplayAlerts = False
oo.Run "reporte", GridEX1.ADORecordset, sEmpresa

Set oo = Nothing

Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Call CARGA_GRID
End Sub

Private Sub txtFec_Ini_Change()
  txtFec_Fin = txtFec_Ini
End Sub

