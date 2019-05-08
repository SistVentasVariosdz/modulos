VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReporteAnioPeriodo 
   Caption         =   "Reporte Año - Periódo"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   LinkTopic       =   "frmReporteAnioPeriodo"
   ScaleHeight     =   2715
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   135
      TabIndex        =   1
      Top             =   180
      Width           =   4380
      Begin MSComCtl2.DTPicker dtpAnoMes 
         Height          =   330
         Left            =   2010
         TabIndex        =   2
         Top             =   405
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MMM yyyy"
         Format          =   62062595
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
         Height          =   225
         Left            =   705
         TabIndex        =   3
         Top             =   450
         Width           =   1050
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   390
      TabIndex        =   0
      Top             =   1830
      Width           =   3855
      _ExtentX        =   6588
      _ExtentY        =   900
      Custom          =   $"frmReporteAnioPeriodo.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   120
      Top             =   1080
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmReporteAnioPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String

 
Private Sub Form_Load()
FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
dtpAnoMes.Value = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

 

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
    Call Reporte
Case "REGVARIA"
    Call registroVariaciones
Case "SALIR"
    Unload Me
End Select
End Sub

Sub registroVariaciones()
 
   Load frmRegistroVariaciones
   frmRegistroVariaciones.Caption = "Registro de Variaciones " & Right("00" & Format(dtpAnoMes.Value, "MM"), 2) & " /" & Format(dtpAnoMes.Value, "YYYY")
   frmRegistroVariaciones.sAnio = Format(dtpAnoMes.Value, "YYYY")
   frmRegistroVariaciones.sMes = Right("00" & Format(dtpAnoMes.Value, "MM"), 2)
  frmRegistroVariaciones.CargarDatos
  frmRegistroVariaciones.Show vbModal
   Set frmRegistroVariaciones = Nothing

End Sub
Sub Reporte()
Dim strSQL As String
Dim periodo As String
 On Error GoTo Errox
 
Dim oo As Object
Dim Ruta As String
Ruta = ""

Ruta = vRuta & "\RptReporteAnioPeriodo.XLT"
strSQL = "Ventas_Emision_Articulos_por_Grupo_Resumen_1 '" & Format(dtpAnoMes.Value, "YYYY") & "','" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2) & "','R'"
    
    
    
Set oo = CreateObject("excel.application")
oo.Workbooks.Open Ruta
oo.Visible = True
oo.DisplayAlerts = False

periodo = Format(dtpAnoMes.Value, "YYYY") & "-" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2)
oo.Run "REPORTE", strSQL, cCONNECT, Format(dtpAnoMes.Value, "YYYY"), Right("00" & Format(dtpAnoMes.Value, "MM"), 2)
   
Set oo = Nothing
Exit Sub

Errox:
    ErrorHandler Err, "Reporte"
End Sub

Public Sub Reporte2()
Dim oo As Object

On Error GoTo ErrorImpresion

    VB.Screen.MousePointer = vbHourglass
    
    'strSql = "Ventas_Emision_Articulos_por_Grupo_Resumen_1 '" & Format(dtpAnoMes.Value, "YYYY") & "','" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2) & "','" & Mid(CboOrigen.Text, 1, 1) & "'"
    strSQL = "Ventas_Emision_Articulos_por_Grupo_Resumen_1 '" & Format(dtpAnoMes.Value, "YYYY") & "','" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2) & "','R'"
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\RptReporteAnioPeriodo.XLT"

    oo.Visible = True
    'oo.Run "REPORTE", strSql, Format(dtpAnoMes.Value, "YYYY") & "-" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2), cCONNECT, CboOrigen.Text
    oo.Run "REPORTE", strSQL, cCONNECT, Format(dtpAnoMes.Value, "YYYY") & "-" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2)
        
    Screen.MousePointer = vbNormal
    
    Set oo = Nothing
    
    Exit Sub
ErrorImpresion:
    Screen.MousePointer = vbNormal
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte" & Err.Description, vbCritical, "Impresion"
End Sub


