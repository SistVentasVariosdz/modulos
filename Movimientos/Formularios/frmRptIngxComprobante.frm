VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRptIngxComprobante 
   Caption         =   "Reporte Ingresos por Comprobante"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   Icon            =   "frmRptIngxComprobante.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1515
      TabIndex        =   5
      Top             =   1890
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "0~0~IMPRIMIR~True~True~&Imprimir~0~0~1~~0~False~False~&Imprimir~~1~0~CANCELAR~True~True~&Cancelar~0~0~2~~0~False~False~&Cancelar~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   1680
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   5250
      Begin VB.ComboBox cboAlmacen 
         Height          =   315
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   225
         Width           =   4050
      End
      Begin MSComCtl2.DTPicker dtpAnoMes 
         Height          =   315
         Left            =   1035
         TabIndex        =   3
         Top             =   675
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy MMM"
         Format          =   23658499
         CurrentDate     =   37817
      End
      Begin VB.ComboBox cboOpcion 
         Height          =   315
         ItemData        =   "frmRptIngxComprobante.frx":27A2
         Left            =   1050
         List            =   "frmRptIngxComprobante.frx":27AF
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1140
         Width           =   4065
      End
      Begin VB.Label Label3 
         Caption         =   "Opciones:"
         Height          =   225
         Left            =   135
         TabIndex        =   7
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "Año / Mes :"
         Height          =   255
         Left            =   105
         TabIndex        =   4
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen"
         Height          =   225
         Left            =   120
         TabIndex        =   2
         Top             =   285
         Width           =   720
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   4275
      Top             =   1890
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmRptIngxComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrSql As String

Private Sub Form_Load()
    LoadAlmacenes
    dtpAnoMes = Date
    cboAlmacen.ListIndex = 0
    cboOpcion.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "IMPRIMIR"
        Reporte
    Case "CANCELAR"
        Unload Me
    End Select
End Sub

Private Sub Reporte()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String, varLogo As Variant, rstAux As ADODB.Recordset

    Screen.MousePointer = 11
    Ruta = vRuta & "\IngXComprobante-Guia.XLT"
    'Ruta = App.Path & "\kardex.xlt"
'    Usu = "Usuario : " & vusu
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    
    StrSql = "SELECT Ruta_Logo FROM SEG_EMPRESAS WHERE cod_EMPRESA ='" & Trim(vemp1) & "'"
    varLogo = DevuelveCampo(StrSql, cSEGURIDAD)
    StrSql = IIf(IsNull(varLogo), "", varLogo)
    
    oo.Run "Reporte", Mid(cboAlmacen, 1, 2), Format(dtpAnoMes, "yyyy"), _
            Format(dtpAnoMes, "mm"), cboOpcion.ListIndex + 1, StrSql, cConnect
            
    Set oo = Nothing
    Screen.MousePointer = 0
Exit Sub
hand:
    Screen.MousePointer = 0
    ErrorHandler Err, "GeneraReportes"
    Set oo = Nothing
End Sub

Private Sub LoadAlmacenes()
Dim rstAux As ADODB.Recordset
    StrSql = "SELECT Cod_Almacen, Nom_Almacen, Tip_Presentacion FROM LG_ALMACEN WHERE Tip_Item = 'T'"
    Set rstAux = CargarRecordSetDesconectado(StrSql, cConnect)
    cboAlmacen.Clear
    With rstAux
    If .RecordCount > 0 Then .MoveFirst
    Do Until .EOF
        cboAlmacen.AddItem !Cod_Almacen & " - " & !Nom_Almacen & Space(100) & !Tip_presentacion
        .MoveNext
    Loop
    .Close
    End With
    Set rstAux = Nothing
End Sub


