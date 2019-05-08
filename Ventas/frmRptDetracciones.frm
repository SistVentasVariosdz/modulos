VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRptDetracciones 
   Caption         =   "Reporte de Detracciones Ventas"
   ClientHeight    =   7635
   ClientLeft      =   450
   ClientTop       =   795
   ClientWidth     =   11610
   Icon            =   "frmRptDetracciones.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   11610
   Begin VB.Frame FraBuscar 
      Caption         =   "Opciones de Busqueda"
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
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11520
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   675
         Left            =   8280
         TabIndex        =   0
         Top             =   150
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   1191
         Custom          =   $"frmRptDetracciones.frx":030A
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
         Left            =   1200
         TabIndex        =   1
         Top             =   450
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   37543
      End
      Begin MSComCtl2.DTPicker dtpFecEmiFin 
         Height          =   315
         Left            =   4440
         TabIndex        =   2
         Top             =   450
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   37543
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         Height          =   195
         Left            =   3720
         TabIndex        =   6
         Top             =   510
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   510
         Width           =   555
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   6420
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   11324
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
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmRptDetracciones.frx":03E2
      Column(2)       =   "frmRptDetracciones.frx":04AA
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmRptDetracciones.frx":054E
      FormatStyle(2)  =   "frmRptDetracciones.frx":0686
      FormatStyle(3)  =   "frmRptDetracciones.frx":0736
      FormatStyle(4)  =   "frmRptDetracciones.frx":07EA
      FormatStyle(5)  =   "frmRptDetracciones.frx":08C2
      FormatStyle(6)  =   "frmRptDetracciones.frx":097A
      FormatStyle(7)  =   "frmRptDetracciones.frx":0A5A
      FormatStyle(8)  =   "frmRptDetracciones.frx":0B06
      ImageCount      =   0
      PrinterProperties=   "frmRptDetracciones.frx":0BB6
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   10680
      Top             =   5760
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmRptDetracciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public codigo As String, Descripcion As String

Private Sub dtpFecEmiIni_Change()
  dtpFecEmiFin.Value = dtpFecEmiIni.Value
End Sub

Private Sub Form_Load()
    
  iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))
  
  
  dtpFecEmiIni = Date - 1
  dtpFecEmiFin = Date - 1
    
End Sub

Private Sub Buscar()

Dim sSQL As String

sSQL = "Ventas_Muestra_Detracciones '" & dtpFecEmiIni & "','" & dtpFecEmiFin & "'"

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)

'GridEX1.Columns("Nro_Documento").Width = 1305
'GridEX1.Columns("Nro_Documento").Caption = "Nro_Factura"

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Msg As Variant
    Select Case ActionName
    Case "BUSCAR"
      Buscar
    Case "IMPRIMIR"
      If GridEX1.RowCount = 0 Then Exit Sub
      Reporte
    Case "SALIR"
       Unload Me
    End Select
End Sub

Public Sub Reporte()
  
On Error GoTo ErrorImpresion

    VB.Screen.MousePointer = vbHourglass
    
    Dim oo As Object
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\rptDetraccionesVentas.XLT"
    
    oo.Visible = True

    oo.Run "REPORTE", GridEX1.ADORecordset, " FACTURAS AFECTAS A DETRACCION DEL " & dtpFecEmiIni & "  AL " & dtpFecEmiFin

    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
    
    Exit Sub
    Resume
ErrorImpresion:
    Screen.MousePointer = vbNormal
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte" & err.Description, vbCritical, "Impresion"
End Sub

