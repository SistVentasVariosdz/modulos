VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRptRelacionArticulos_x_Grupos 
   Caption         =   "Relacion de Articulos x Grupo"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Opciones"
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
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4695
      Begin VB.OptionButton OptPrecios 
         Caption         =   "Precios Articulos x Grupo"
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton OptImportes 
         Caption         =   "Articulos x Grupo Resumen"
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   4695
      Begin VB.ComboBox CboOrigen 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   600
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker dtpAnoMes 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM yyyy"
         Format          =   62193667
         CurrentDate     =   37887
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
         TabIndex        =   5
         Top             =   300
         Width           =   1050
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
         TabIndex        =   4
         Top             =   700
         Width           =   750
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   473
      TabIndex        =   0
      Top             =   1920
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   900
      Custom          =   $"FrmRptRelacionArticulos_x_Grupos.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   0
      Top             =   1200
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmRptRelacionArticulos_x_Grupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim sFam As String
Dim scadena As String
Dim smoneda As String


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

   Call Reporte
 


 Case "DIARIO"
    Call Reporte_Diario
Case "SALIR"
    Unload Me
End Select
End Sub
Sub Reporte_Diario()
On Error GoTo err:
Dim oo As Object
Dim sRutaLogo As String

Set oo = CreateObject("excel.application")
sFam = DevuelveCampo("select cod_tipodiarioventas from cn_control", cCONNECT)
scadena = 8
smoneda = "SOL"
    strSQL = "SELECT Ruta_Logo = ISNULL(Ruta_Logo, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
    sRutaLogo = DevuelveCampo(strSQL, cCONNECT)
    
    strSQL = "CN_GENERA_REPORTE_DIARIO_GENERAL '" & Format(dtpAnoMes.Value, "YYYY") & "','" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2) & "','" & sFam & "','" & smoneda & "','" & scadena & "'"
    oo.workbooks.Open vRuta & "\RptDiarioGeneral_Ventas.XLT"
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.run "Reporte", Format(dtpAnoMes.Value, "YYYY"), Right("00" & Format(dtpAnoMes.Value, "MM"), 2), sFam, smoneda, strSQL, cCONNECT, 70, 1, sRutaLogo



Set oo = Nothing
Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub
Public Sub Reporte()
Dim oo As Object
Dim Valor As String

On Error GoTo ErrorImpresion

If MsgBox(" Desea incluir Relacion de Gastos Cobrados en Facturas,NDebito y NCredito ? ", vbInformation + vbYesNo, "AVISO") = vbYes Then
   Valor = "S"
Else
   Valor = "N"
End If


    VB.Screen.MousePointer = vbHourglass
    
    Set oo = CreateObject("excel.application")
    If OptImportes Then
        strSQL = "Ventas_Emision_Articulos_por_Grupo_Resumen_1 '" & Format(dtpAnoMes.Value, "YYYY") & "','" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2) & "','" & Mid(CboOrigen.Text, 1, 1) & "'"
        oo.workbooks.Open vRuta & "\RptArticulos_x_Grupo_Resumen.XLT"
    Else
        strSQL = "Ventas_Emision_Articulos_por_Grupo_Resumen_2 '" & Format(dtpAnoMes.Value, "YYYY") & "','" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2) & "','" & Mid(CboOrigen.Text, 1, 1) & "'"
        oo.workbooks.Open vRuta & "\RptPrecios_Articulos_x_Grupo_Resumen.xlt"
    End If

    oo.Visible = True
    oo.run "REPORTE", Format(dtpAnoMes.Value, "YYYY") & "-" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2), CboOrigen.Text, Valor, Format(dtpAnoMes.Value, "YYYY"), Right("00" & Format(dtpAnoMes.Value, "MM"), 2), strSQL, cCONNECT
        
    Screen.MousePointer = vbNormal
    
    Set oo = Nothing
    
    Exit Sub
ErrorImpresion:
    Screen.MousePointer = vbNormal
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte" & err.Description, vbCritical, "Impresion"
End Sub

Sub LLENA_COMBO()
CboOrigen.AddItem "N - Nacional"
CboOrigen.AddItem "E - Extranjero"
CboOrigen.AddItem "T - Todos"

CboOrigen.ListIndex = -1
End Sub

