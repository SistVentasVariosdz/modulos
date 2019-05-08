VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRptVentasxGrupo 
   Caption         =   "Emision Reporte Ventas por Grupo "
   ClientHeight    =   1755
   ClientLeft      =   2280
   ClientTop       =   2970
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   5355
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1447
      TabIndex        =   3
      Top             =   1170
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmRptVentasxGrupo.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   5295
      Begin VB.ComboBox CboOrigen 
         Height          =   315
         ItemData        =   "FrmRptVentasxGrupo.frx":0090
         Left            =   1560
         List            =   "FrmRptVentasxGrupo.frx":0092
         TabIndex        =   4
         Top             =   600
         Width           =   3465
      End
      Begin MSComCtl2.DTPicker dtpAnoMes 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM yyyy"
         Format          =   16515075
         CurrentDate     =   37887
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Origen : "
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
         TabIndex        =   5
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año - Mes : "
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
         TabIndex        =   1
         Top             =   240
         Width           =   1050
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   120
      Top             =   1080
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmRptVentasxGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String


Private Sub Form_Load()
Call LLENA_COMBO
dtpAnoMes.Value = Date
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "IMPRIMIR"
            Select Case CboOrigen.ListIndex
                Case 4: Call Reporte_VTAmensualDeHilado
                Case Else
                    Call Reporte
            End Select
            
        Case "SALIR"
            Unload Me
    End Select
End Sub


Public Sub Reporte()
Dim oo As Object
Dim strSQL As String
Dim sEmpresa As String
    strSQL = "SELECT DES_EMPRESA FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA='" & vemp & "'"
    sEmpresa = DevuelveCampo(strSQL, cCONNECT)

On Error GoTo ErrorImpresion

    VB.Screen.MousePointer = vbHourglass
    
    strSQL = "Ventas_Emision_Articulos_por_Grupo '" & Format(dtpAnoMes.Value, "YYYY") & "','" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2) & "','" & Mid(CboOrigen.Text, 1, 1) & "'"
    Set oo = CreateObject("excel.application")
    oo.WorkBooks.Open vRuta & "\RptVentasxGrupo.XLT"

    oo.Visible = True
    oo.Run "REPORTE", strSQL, Format(dtpAnoMes.Value, "YYYY") & "-" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2), cCONNECT, CboOrigen.Text, sEmpresa
        
    Screen.MousePointer = vbNormal
    
    Set oo = Nothing
    
    Exit Sub
ErrorImpresion:
    Screen.MousePointer = vbNormal
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte" & err.Description, vbCritical, "Impresion"
End Sub

Sub Reporte_VTAmensualDeHilado()
    Dim oo As Object
    Dim oRsHilado_R As New Recordset
    Dim oRsCliente_R As New Recordset
    Dim oRsCliente_D As New Recordset
    Dim sRutaLogo As String
    Dim sTitulo As String
    
    strSQL = "SELECT Ruta_Logo From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
    sRutaLogo = DevuelveCampo(strSQL, cCONNECT)
    sRutaLogo = CStr(IIf(IsNull(sRutaLogo), "", sRutaLogo))
        
    strSQL = "exec Ventas_Emision_Articulos_por_Grupo '" & Format(dtpAnoMes.Value, "YYYY") & "', '" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2) & "', '',null,null,null,'','001'"
    Set oRsCliente_D = CargarRecordSetDesconectado(strSQL, cCONNECT)
    strSQL = "exec Ventas_Emision_Articulos_por_Grupo '" & Format(dtpAnoMes.Value, "YYYY") & "', '" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2) & "', '',null,null,null,'','CLI'"
    Set oRsCliente_R = CargarRecordSetDesconectado(strSQL, cCONNECT)
    strSQL = "exec Ventas_Emision_Articulos_por_Grupo '" & Format(dtpAnoMes.Value, "YYYY") & "', '" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2) & "', '',null,null,null,'','HIL'"
    Set oRsHilado_R = CargarRecordSetDesconectado(strSQL, cCONNECT)
    
    Set oo = CreateObject("excel.application")
        
    oo.WorkBooks.Open vRuta & "\rptVentaDeHiladoMensual.XLT"
    oo.Visible = True
    oo.DisplayAlerts = False
    
    sTitulo = Format(dtpAnoMes.Value, "MMMM yyyy")

    oo.Run "reporte", sRutaLogo, oRsHilado_R, oRsCliente_R, oRsCliente_D, sTitulo
    Set oo = Nothing
    Exit Sub

Errox:
        MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
'
'
'Sub Reporte_VTAmensualPorCLienteResumen()
'    Dim oo As Object
'    Dim oRs As New Recordset
'    Dim sRutaLogo As String
'    Dim sTitulo As String
'
'    strSQL = "SELECT Ruta_Logo From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
'    sRutaLogo = DevuelveCampo(strSQL, cCONNECT)
'    sRutaLogo = CStr(IIf(IsNull(sRutaLogo), "", sRutaLogo))
'
'    strSQL = "exec Ventas_Emision_Articulos_por_Grupo '" & Format(dtpAnoMes.Value, "YYYY") & "', '" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2) & "', '',null,null,null,'','000'"
'
'
'    Set oRs = CargarRecordSetDesconectado(strSQL, cCONNECT)
'    Set oo = CreateObject("excel.application")
'
'    oo.WorkBooks.Open vRuta & "\rptVentaMensualPorClienteResumen.XLT"
'    oo.Visible = True
'    oo.DisplayAlerts = False
'
'    'Select Case Month(dtpAnoMes)
'    'End Select
'
'    sTitulo = Format(dtpAnoMes.Value, "MMMM yyyy")
'
'    oo.Run "reporte", sRutaLogo, oRs, sTitulo
'    Set oo = Nothing
'    Exit Sub
'
'Errox:
'        MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
'End Sub



Sub LLENA_COMBO()
    CboOrigen.AddItem "N - Nacional"
    CboOrigen.AddItem "E - Extranjero"
    CboOrigen.AddItem "T - Todos"
    CboOrigen.AddItem "G - Transferencia Gratuita"
    CboOrigen.AddItem "Venta de Hilado Mensual"
    
    CboOrigen.ListIndex = -1
End Sub
