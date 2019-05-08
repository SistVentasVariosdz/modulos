VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRptAnticipos_Canjes 
   Caption         =   "Cancelaciones - Anticipos"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11895
      Begin VB.TextBox Txt_Origen 
         Height          =   300
         Left            =   4440
         TabIndex        =   1
         Top             =   277
         Width           =   615
      End
      Begin VB.TextBox Txt_Descripcion 
         Height          =   300
         Left            =   5160
         TabIndex        =   2
         Top             =   277
         Width           =   1695
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   495
         Left            =   10440
         TabIndex        =   3
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
      Begin MSComCtl2.DTPicker DTPFecha 
         Height          =   255
         Left            =   1440
         TabIndex        =   0
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   64159745
         CurrentDate     =   38590
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Origen"
         Height          =   195
         Left            =   3840
         TabIndex        =   8
         Top             =   330
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   540
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   570
      Left            =   4680
      TabIndex        =   4
      Top             =   6120
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   1005
      Custom          =   $"FrmRptAnticipos_Canjes.frx":0000
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
      Top             =   840
      Width           =   11880
      _ExtentX        =   20955
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
      Column(1)       =   "FrmRptAnticipos_Canjes.frx":0090
      Column(2)       =   "FrmRptAnticipos_Canjes.frx":0158
      FormatStylesCount=   8
      FormatStyle(1)  =   "FrmRptAnticipos_Canjes.frx":01FC
      FormatStyle(2)  =   "FrmRptAnticipos_Canjes.frx":0334
      FormatStyle(3)  =   "FrmRptAnticipos_Canjes.frx":03E4
      FormatStyle(4)  =   "FrmRptAnticipos_Canjes.frx":0498
      FormatStyle(5)  =   "FrmRptAnticipos_Canjes.frx":0570
      FormatStyle(6)  =   "FrmRptAnticipos_Canjes.frx":0628
      FormatStyle(7)  =   "FrmRptAnticipos_Canjes.frx":0708
      FormatStyle(8)  =   "FrmRptAnticipos_Canjes.frx":07B4
      ImageCount      =   0
      PrinterProperties=   "FrmRptAnticipos_Canjes.frx":0864
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   240
      Top             =   6120
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmRptAnticipos_Canjes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim vOpcion As String
Public codigo As String, Descripcion As String

Private Sub Form_Load()
DTPFecha.Value = Date
vOpcion = "T"
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
On Error GoTo errCarga

strSQL = "Cn_Ventas_Emision_Parte_Cancelaciones '" & DTPFecha & "','" & vOpcion & "','" & Txt_Origen & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)



Exit Sub
errCarga:
    ErrorHandler err, "Carga Grid"
End Sub

Sub Reporte()
Dim sempresa As String
On Error GoTo hand
Dim oo As Object

sempresa = DevuelveCampo("SELECT des_empresa FROM seg_empresas WHERE Cod_Empresa ='" & vemp & "'", cSEGURIDAD)

strSQL = "Cn_Ventas_Emision_Parte_Cancelaciones '" & DTPFecha & "','" & vOpcion & "','" & Txt_Origen & "'"

Set oo = CreateObject("excel.application")
oo.Workbooks.Open vRuta & "\RptCancelacion_Facturas.xlt"
oo.Visible = True
oo.DisplayAlerts = False
oo.Run "reporte", DTPFecha.Value, strSQL, cCONNECT, vOpcion, sempresa
Set oo = Nothing

Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Call CARGA_GRID
End Sub

Private Sub Txt_Descripcion_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Origen", "Des_Origen", "cn_origen where ", Txt_Origen, Txt_Descripcion, 2, Me)
End Sub

Private Sub Txt_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Origen", "Des_Origen", "cn_origen where ", Txt_Origen, Txt_Descripcion, 1, Me)
End Sub
