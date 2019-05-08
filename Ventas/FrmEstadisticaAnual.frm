VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmEstadisticaAnual 
   Caption         =   "Estadistica Anual"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin MSComCtl2.DTPicker dtpFechaActual 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   840
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyyy"
         Format          =   57999363
         CurrentDate     =   40544
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   1080
         TabIndex        =   3
         Top             =   1440
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"FrmEstadisticaAnual.frx":0000
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label1 
         Caption         =   "ESTADISTICA ANUAL EN DOLARES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label4 
         Caption         =   "Año :"
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   960
         Width           =   585
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   120
      Top             =   0
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmEstadisticaAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
 Case "IMPRI"
  reporteEstadistico
 Case "SALIR"
    Unload Me
End Select
End Sub


Sub reporteEstadistico()
Dim strSQL As String
Dim oo As Object
Dim sRuta_Logo As String
 
Set oo = CreateObject("excel.application")
oo.Workbooks.Open vRuta & "\RptEstadisticaAnual.XLT"
oo.Visible = True
oo.DisplayAlerts = False

Dim rutaLogo As String
rutaLogo = DevuelveCampo("select ruta_logo=isNUll(ruta_logo,'') from seguridad..seg_empresas where cod_empresa='" & vemp & "'", cCONNECT)


strSQL = "gerencial_encuentra_ventas_ultimos_2_anios  '" & dtpFechaActual.Year & "'"
oo.Visible = True
oo.DisplayAlerts = False
oo.Run "reporte", strSQL, cCONNECT, rutaLogo, dtpFechaActual.Year

Set oo = Nothing

Exit Sub
errReporte:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub
