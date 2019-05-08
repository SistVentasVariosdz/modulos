VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCtaCteCliFacExt 
   Caption         =   "Consulta Cta.Corriente Factoring Exterior"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   11835
   StartUpPosition =   2  'CenterScreen
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
      Height          =   2205
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11745
      Begin VB.TextBox TxtDDolares 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   900
         Width           =   1455
      End
      Begin VB.TextBox txtDes_TipAnex 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2640
         MaxLength       =   11
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   1920
         TabIndex        =   11
         Top             =   720
         Width           =   6135
         Begin MSComCtl2.DTPicker dtpFecEmiIni 
            Height          =   315
            Left            =   1980
            TabIndex        =   12
            Top             =   120
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   55312385
            CurrentDate     =   37543
         End
         Begin MSComCtl2.DTPicker dtpFecEmiFin 
            Height          =   315
            Left            =   4080
            TabIndex        =   13
            Top             =   120
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   55312385
            CurrentDate     =   37543
         End
         Begin VB.Label Label1 
            Caption         =   "Rango Fecha de Emisión:"
            Height          =   240
            Left            =   0
            TabIndex        =   14
            Top             =   120
            Width           =   2235
         End
      End
      Begin VB.OptionButton opTodas 
         Caption         =   "Todas"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton oprCanceladas 
         Caption         =   "Canceladas"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton opPendiente 
         Caption         =   "Pendientes"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   435
         Left            =   7650
         TabIndex        =   7
         Top             =   225
         Width           =   1305
      End
      Begin VB.TextBox TxtDsoles 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Txt_Importe 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox TxtDOtros 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1290
         Width           =   1455
      End
      Begin VB.TextBox txtNum_Ruc 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   690
         MaxLength       =   11
         TabIndex        =   3
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox txtCod_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2130
         MaxLength       =   4
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   360
      End
      Begin VB.TextBox txtDes_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   4785
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total $"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9120
         TabIndex        =   20
         Top             =   1785
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Deuda Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9660
         TabIndex        =   19
         Top             =   180
         Width           =   1305
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Soles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9120
         TabIndex        =   18
         Top             =   555
         Width           =   615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Dolares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9120
         TabIndex        =   17
         Top             =   930
         Width           =   840
      End
      Begin VB.Label Label11 
         Caption         =   "Otra Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   9120
         TabIndex        =   16
         Top             =   1245
         Width           =   900
      End
      Begin VB.Label Label3 
         Caption         =   "Ruc :"
         Height          =   180
         Left            =   240
         TabIndex        =   15
         Tag             =   "Anexo Type"
         Top             =   405
         Width           =   435
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4860
      Left            =   0
      TabIndex        =   21
      Top             =   2280
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   8573
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
      Column(1)       =   "frmCtaCteCliFacExt.frx":0000
      Column(2)       =   "frmCtaCteCliFacExt.frx":00C8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmCtaCteCliFacExt.frx":016C
      FormatStyle(2)  =   "frmCtaCteCliFacExt.frx":02A4
      FormatStyle(3)  =   "frmCtaCteCliFacExt.frx":0354
      FormatStyle(4)  =   "frmCtaCteCliFacExt.frx":0408
      FormatStyle(5)  =   "frmCtaCteCliFacExt.frx":04E0
      FormatStyle(6)  =   "frmCtaCteCliFacExt.frx":0598
      FormatStyle(7)  =   "frmCtaCteCliFacExt.frx":0678
      FormatStyle(8)  =   "frmCtaCteCliFacExt.frx":0724
      ImageCount      =   0
      PrinterProperties=   "frmCtaCteCliFacExt.frx":07D4
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   525
      Left            =   8640
      TabIndex        =   22
      Top             =   7200
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   926
      Custom          =   $"frmCtaCteCliFacExt.frx":09AC
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1500
      ControlHeigth   =   500
      ControlSeparator=   10
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   300
      Top             =   7170
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmCtaCteCliFacExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OP_Opcion As String
Public StrEstus As String
Public strSQL  As String
Public codigo, Descripcion As String, TipoAdd As String, strCod_Anxo As String
Public oGroup As GridEX20.JSGroup
Public oFormat As JSFormatStyle


Private Sub cmdBuscar_Click()
Buscar
End Sub

Private Sub dtpFecEmiFin_Validate(Cancel As Boolean)
If dtpFecEmiIni > dtpFecEmiFin Then
  MsgBox "Fecha Final no puede ser menor a la fecha Inicial", vbInformation, "AVISO"
  dtpFecEmiIni = dtpFecEmiFin
End If
End Sub

Private Sub dtpFecEmiIni_Change()
  GridEX1.ClearFields
  dtpFecEmiFin.Value = Date
End Sub

Private Sub dtpFecEmiIni_Validate(Cancel As Boolean)
If dtpFecEmiIni > dtpFecEmiFin Then
  MsgBox "Fecha Inicial no puede ser mayor a la fecha final", vbInformation, "AVISO"
  dtpFecEmiIni = dtpFecEmiFin
End If
End Sub

Private Sub Form_Load()
  'txtCod_TipAne = "C"
  dtpFecEmiIni.Value = Date
  dtpFecEmiFin.Value = Date
    StrEstus = "P"

End Sub


Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName

Case Is = "VERDETALLE"
  Call GridEX1_DblClick

Case Is = "IMPRIMIR"
  Call Reporte
  
End Select
End Sub

Private Sub GridEX1_Click()
Txt_Importe = Format(GridEX1.Value(GridEX1.Columns("imp_total").Index), "##,##0.00")
TxtDDolares.Text = Format(GridEX1.Value(GridEX1.Columns("SALDO_DOLARES").Index), "##,##0.00")
TxtDsoles = Format(GridEX1.Value(GridEX1.Columns("SALDO_SOLES").Index), "##,##0.00")
TxtDOtros = Format(GridEX1.Value(GridEX1.Columns("SALDO_OTROS").Index), "##,##0.00")
End Sub

Private Sub opPendiente_Click()
StrEstus = "P"

End Sub

Private Sub oprCanceladas_Click()
StrEstus = "C"
End Sub


Private Sub opTodas_Click()
StrEstus = "T"
End Sub

Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 1, Me)
    FunctButt1.SetFocus
  End If
End Sub


Private Sub txtCod_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAne, txtDes_TipAnex, 1, Me)
    
End Sub


Private Sub txtDes_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 2, Me)
End Sub


Public Sub Buscar()

    If (IsNull(dtpFecEmiIni) Or IsNull(dtpFecEmiIni)) Then
        MsgBox "Ingrese un Rango de Fechas", vbInformation, "AVISO"
        Exit Sub
      End If
    
     ' If (dtpFecEmiFin - dtpFecEmiIni) > 60 Then
        'MsgBox "No puede Ingresar un Rango Mayor a 60 Dias", vbInformation, "AVISO"
        'Exit Sub
     ' End If
      
strSQL = "FI_CONSULTA_CTACTE_FACTURAS_CLIENTES_CON_FACTORING_EXTERIOR '" & StrEstus & "','" & dtpFecEmiIni.Value & "','" & dtpFecEmiFin.Value & "','" & txtCod_TipAne.Text & "','" & strCod_Anxo & "'"

GridEX1.ClearFields
GridEX1.DefaultGroupMode = jgexDGMExpanded
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
Configurar

 End Sub



Sub Configurar()

Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Cliente").Index, jgexSortAscending)

GridEX1.BackColorRowGroup = &H80000005

Txt_Importe = Format(GridEX1.Value(GridEX1.Columns("imp_total").Index), "##,##0.00")
TxtDDolares.Text = Format(GridEX1.Value(GridEX1.Columns("SALDO_DOLARES").Index), "##,##0.00")
TxtDsoles = Format(GridEX1.Value(GridEX1.Columns("SALDO_SOLES").Index), "##,##0.00")
TxtDOtros = Format(GridEX1.Value(GridEX1.Columns("SALDO_OTROS").Index), "##,##0.00")



GridEX1.Columns("Fec_Emision").Width = 1125
'GridEX1.Columns("Fec_VenDoc").Width = 1080
GridEX1.Columns("Num_Registro").Width = 1155
GridEX1.Columns("Moneda").Width = 720

If txtCod_TipAne = "" Then
GridEX1.DefaultGroupMode = jgexDGMCollapsed
Else
GridEX1.DefaultGroupMode = jgexDGMExpanded
End If

GridEX1.ContinuousScroll = True

End Sub




Private Sub GridEX1_DblClick()
  If GridEX1.RowCount = 0 Then Exit Sub
  Load frmCtaCteCliFacExtDetalle
  frmCtaCteCliFacExtDetalle.Caption = "Detalle Cliente: " & Trim(GridEX1.Value(GridEX1.Columns("Cliente").Index)) & Space(10) & "Documento : " & GridEX1.Value(GridEX1.Columns("Documento").Index)
  frmCtaCteCliFacExtDetalle.sNumCorre = GridEX1.Value(GridEX1.Columns("num_corre_ventas").Index)
  frmCtaCteCliFacExtDetalle.Buscar
  frmCtaCteCliFacExtDetalle.Show vbModal
End Sub



Sub Reporte()
On Error GoTo ERROR
'Dim ssql As String
Dim oo As Object
Dim Ruta As String, sRutaLogo As String

strSQL = "SELECT Ruta_Logo = ISNULL(Ruta_Logo, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
sRutaLogo = DevuelveCampo(strSQL, cCONNECT)
    
If GridEX1.RowCount = 0 Then Exit Sub

    Ruta = vRuta & "\RptCtaCorriente_ClienteFactoring.XLT"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "Reporte", GridEX1.ADORecordset, sRutaLogo
    Set oo = Nothing
Exit Sub
ERROR:
    errores err.Number
End Sub


