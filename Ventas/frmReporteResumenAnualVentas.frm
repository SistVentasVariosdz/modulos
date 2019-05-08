VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReporteResumenAnualVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen Anual de Ventas"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   3660
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3420
      Begin MSComCtl2.DTPicker dtpAnoMes 
         Height          =   330
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   56492035
         CurrentDate     =   37887
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año : "
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
         TabIndex        =   2
         Top             =   308
         Width           =   525
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmReporteResumenAnualVentas.frx":0000
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
      Top             =   0
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmReporteResumenAnualVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    dtpAnoMes.Value = DateAdd("yyyy", -1, Date)
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
   Call Reporte
Case "SALIR"
    Unload Me
End Select
End Sub

Sub Reporte()
On Error GoTo Errox
Dim oo As Object
Dim strSQL As String, periodo As String, Ruta As String
    
    strSQL = "Ventas_Emision_Resumen_ANUAL '" & Format(dtpAnoMes.Value, "YYYY") & "'"
    
    If MsgBox("Imprimir usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir") = vbYes Then
        Ruta = vRuta & "\RptResumenAnualVentas.XLT"
        Set oo = CreateObject("excel.application")
        oo.Workbooks.Open Ruta
        oo.Visible = True
        oo.DisplayAlerts = False
        
        oo.Run "REPORTE", strSQL, cCONNECT, Format(dtpAnoMes.Value, "YYYY")
    Else
        Ruta = vRuta & "\RptResumenAnualVentas.OTS"
        Set oo = CreateObject("ooBusiness.Calc")
        oo.OfficeTemplateSheet = Ruta
        oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
        oo.MacroLibraryName = "Library1"
        oo.MacroModuleName = "Module1"
        oo.MacroName = "Reporte"
        
        oo.Run strSQL, cCONNECT, Format(dtpAnoMes.Value, "YYYY")
    End If
    Set oo = Nothing
Exit Sub
Errox:
    ErrorHandler err, "Reporte"
End Sub


