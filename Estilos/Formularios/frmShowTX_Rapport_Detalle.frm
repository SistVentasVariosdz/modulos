VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmShowTX_Rapport_Detalle 
   Caption         =   "Detalle de Combinación  -Rapport"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   7350
      TabIndex        =   2
      Top             =   3255
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "0~0~IMPRIMIR~True~True~&Imprimir~0~0~1~~0~False~False~&Imprimir~~1~0~SALIR~True~True~&Salir~1~0~3~~0~False~False~&Salir~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame2 
      Height          =   3150
      Left            =   60
      TabIndex        =   0
      Top             =   -30
      Width           =   9765
      Begin GridEX20.GridEX GridEX1 
         Height          =   2775
         Left            =   90
         TabIndex        =   1
         Top             =   210
         Width           =   9540
         _ExtentX        =   16828
         _ExtentY        =   4895
         Version         =   "2.0"
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
         Column(1)       =   "frmShowTX_Rapport_Detalle.frx":0000
         Column(2)       =   "frmShowTX_Rapport_Detalle.frx":00C8
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmShowTX_Rapport_Detalle.frx":016C
         FormatStyle(2)  =   "frmShowTX_Rapport_Detalle.frx":02A4
         FormatStyle(3)  =   "frmShowTX_Rapport_Detalle.frx":0354
         FormatStyle(4)  =   "frmShowTX_Rapport_Detalle.frx":0408
         FormatStyle(5)  =   "frmShowTX_Rapport_Detalle.frx":04E0
         FormatStyle(6)  =   "frmShowTX_Rapport_Detalle.frx":0598
         FormatStyle(7)  =   "frmShowTX_Rapport_Detalle.frx":0678
         FormatStyle(8)  =   "frmShowTX_Rapport_Detalle.frx":0724
         ImageCount      =   0
         PrinterProperties=   "frmShowTX_Rapport_Detalle.frx":07D4
      End
   End
End
Attribute VB_Name = "frmShowTX_Rapport_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rapport_number As Integer
Dim StrSql As String
Dim col As Integer
Dim i As Integer
Sub CARGA_GRID()
StrSql = "SG_GeneraReporteRapport " & rapport_number
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(StrSql, cCONNECT)

GridEX1.Columns("rapport_comb").Width = "700"

col = GridEX1.Columns.Count

For i = 2 To col
    GridEX1.Columns(i).Width = "2200"
Next i

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
Set oo = CreateObject("excel.application")
oo.workbooks.Open vRuta & "\ReporteRapport.xlt"
oo.Visible = True
oo.run "Reporte", rapport_number, col, cCONNECT
Screen.MousePointer = vbNormal
oo.Visible = True
Set oo = Nothing
End Sub
