VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMovStocks 
   Caption         =   "Consulta Movimientos Stock"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14220
   Icon            =   "frmMovStocks.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   14220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   540
      Left            =   120
      TabIndex        =   12
      Top             =   6120
      Width           =   1500
   End
   Begin VB.Frame Frame3 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   14055
      Begin GridEX20.GridEX gexDetalle 
         Height          =   4575
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   8070
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         GridLineStyle   =   2
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         ColumnHeaderHeight=   285
         ColumnsCount    =   9
         Column(1)       =   "frmMovStocks.frx":030A
         Column(2)       =   "frmMovStocks.frx":03F6
         Column(3)       =   "frmMovStocks.frx":04CA
         Column(4)       =   "frmMovStocks.frx":059E
         Column(5)       =   "frmMovStocks.frx":066E
         Column(6)       =   "frmMovStocks.frx":073A
         Column(7)       =   "frmMovStocks.frx":0806
         Column(8)       =   "frmMovStocks.frx":08D2
         Column(9)       =   "frmMovStocks.frx":099E
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmMovStocks.frx":0A6E
         FormatStyle(2)  =   "frmMovStocks.frx":0BA6
         FormatStyle(3)  =   "frmMovStocks.frx":0C56
         FormatStyle(4)  =   "frmMovStocks.frx":0D0A
         FormatStyle(5)  =   "frmMovStocks.frx":0DE2
         FormatStyle(6)  =   "frmMovStocks.frx":0E9A
         ImageCount      =   0
         PrinterProperties=   "frmMovStocks.frx":0F7A
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   14055
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   12720
         TabIndex        =   3
         Top             =   360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   900
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fechas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5280
         TabIndex        =   7
         Top             =   120
         Width           =   4575
         Begin MSComCtl2.DTPicker DTPFin 
            Height          =   315
            Left            =   2760
            TabIndex        =   2
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Format          =   104988673
            CurrentDate     =   37460
         End
         Begin MSComCtl2.DTPicker DTPInicio 
            Height          =   315
            Left            =   720
            TabIndex        =   1
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   104988673
            CurrentDate     =   37460
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
            Height          =   255
            Left            =   2280
            TabIndex        =   9
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   300
            Width           =   495
         End
      End
      Begin VB.TextBox txtListado 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.ComboBox CboAlmacen 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   9840
      Top             =   120
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmMovStocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public varCadena_Movs As String

Private Sub cboAlmacen_Click()
If cboAlmacen.ListIndex <> -1 Then
    Load frmSelectipMov
    frmSelectipMov.varCOD_ALMACEN = Trim(Right(cboAlmacen, 3))
    frmSelectipMov.CARGA_MOVIMIENTOS
    Set frmSelectipMov.oParent = Me
    frmSelectipMov.Show 1
    If varCadena_Movs <> "" Then txtListado.Text = varCadena_Movs
End If

End Sub

Private Sub CmdImprimir_Click()
    Call Reporte
End Sub

Private Sub DTPFin_Change()
'    If DTPFin.Value < Me.DTPInicio.Value Then
'        DTPFin.Value = Me.DTPInicio.Value
'    End If
End Sub

Private Sub DTPFin_Click()
'    If DTPFin.Value < Me.DTPInicio.Value Then
'        DTPFin.Value = Me.DTPInicio.Value
'    End If
End Sub

Private Sub DTPInicio_Click()
'    If DTPFin.Value < Me.DTPInicio.Value Then
'        Me.DTPInicio.Value = DTPFin.Value
'    End If
End Sub

Private Sub DTPInicio_Validate(Cancel As Boolean)
'    DTPInicio_Click
End Sub

Private Sub Form_Load()
LlenaCombo cboAlmacen, "Select a.Nom_Almacen+space(100)+ a.Cod_Almacen from lg_almacen a, lg_segalm b  where a.cod_almacen=b.cod_almacen and b.cod_usuario='" & vusu & "' order by 1", cConnect
Me.DTPInicio.Value = Date - 30
Me.DTPFin.Value = Date
'CARGA_GRID
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    If varCadena_Movs = "" Then
        MsgBox "NO ha seleccionado ningún Tipo de Movimiento", vbInformation, "Movimiento de Stocks"
        cboAlmacen.SetFocus
        Exit Sub
    End If
    If Len(varCadena_Movs) > 9 Then
        If Me.DTPFin.Value - Me.DTPInicio.Value > 60 Then
            MsgBox "La diferencia de dias entre las fechas no puede ser mayor a 60. Sirvase verificar", vbInformation, "Mensaje"
            Exit Sub
        End If
    End If
    CARGA_GRID
End Sub

Sub CARGA_GRID()
Dim Rs_Lista As ADODB.Recordset
Dim CN_Lista As ADODB.Connection
Dim StrSql As String

    Set CN_Lista = Nothing
    Set CN_Lista = New ADODB.Connection
    CN_Lista.CommandTimeout = 900
    CN_Lista.ConnectionString = cConnect
    CN_Lista.ConnectionTimeout = 900
    CN_Lista.Open cConnect
    
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    'Esta cadena es para devolver el Codigo de Cliente
    StrSql = "EXEC SM_MUESTRA_MOVIMIENTOS_ALMACEN '" & Trim(Right(cboAlmacen, 3)) & "','" & DTPInicio.Value & "','" & DTPFin.Value & "','" & varCadena_Movs & "','" & vusu & "'"
    
    Rs_Lista.Open StrSql, CN_Lista
    Set gexDetalle.ADORecordset = Rs_Lista

ConfigurarGrid
End Sub

Sub ConfigurarGrid()
    gexDetalle.Columns("cod_Tipmov").Visible = False
    
    gexDetalle.Columns(1).Width = 1000
    gexDetalle.Columns(2).Width = 0
    gexDetalle.Columns(3).Width = 1000
    gexDetalle.Columns(4).Width = 800
    gexDetalle.Columns(5).Width = 800
    gexDetalle.Columns(6).Width = 800
    gexDetalle.Columns(8).Width = 800
    gexDetalle.Columns(9).Width = 800
    gexDetalle.Columns(10).Width = 800
    gexDetalle.Columns(12).Width = 500
    gexDetalle.Columns(7).Width = 3000
End Sub

Public Sub Reporte()
On Error GoTo ErrorImpresion
    Dim oo As Object
    Set oo = CreateObject("excel.application")
    'oo.Workbooks.Open App.Path & "\RptMovStockFecha.xlt"
    oo.Workbooks.Open vRuta & "\RptMovStockFecha.xlt"
    oo.Visible = True
    oo.Run "REPORTE", Me.cboAlmacen.Text, Me.DTPInicio.Value, Me.DTPFin.Value, Me.varCadena_Movs, cConnect, vusu
    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte de Movimientos x Fecha " & err.Description, vbCritical, "Impresion"
End Sub

