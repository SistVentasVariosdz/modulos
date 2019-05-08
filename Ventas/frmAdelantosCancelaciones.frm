VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAdelantosCancelaciones 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   Icon            =   "frmAdelantosCancelaciones.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3285
      TabIndex        =   2
      Top             =   4110
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmAdelantosCancelaciones.frx":030A
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   3825
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   9075
      Begin GridEX20.GridEX gexDetLetra 
         Height          =   3525
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   6218
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAdelantosCancelaciones.frx":039A
         Column(2)       =   "frmAdelantosCancelaciones.frx":0462
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdelantosCancelaciones.frx":0506
         FormatStyle(2)  =   "frmAdelantosCancelaciones.frx":063E
         FormatStyle(3)  =   "frmAdelantosCancelaciones.frx":06EE
         FormatStyle(4)  =   "frmAdelantosCancelaciones.frx":07A2
         FormatStyle(5)  =   "frmAdelantosCancelaciones.frx":087A
         FormatStyle(6)  =   "frmAdelantosCancelaciones.frx":0932
         ImageCount      =   0
         PrinterProperties=   "frmAdelantosCancelaciones.frx":0A12
      End
   End
End
Attribute VB_Name = "frmAdelantosCancelaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strSQL As String

Public Sub CARGA_GRID()

On Error GoTo hand

Dim RS As ADODB.Recordset

    Set RS = CreateObject("ADODB.Recordset")
    RS.CursorLocation = adUseClient
    RS.Open strSQL, cCONNECT

    Set gexDetLetra.ADORecordset = RS

    Set RS = Nothing

    ConfigurarGrid

Exit Sub
Resume
hand:
ErrorHandler err, "CARGA_GRID"
Set RS = Nothing

End Sub

Sub ConfigurarGrid()
  gexDetLetra.Columns("Nro_Documento").Width = 1500
  gexDetLetra.Columns("Cliente").Width = 2535
  gexDetLetra.Columns("Fecha_Emision").Width = 1230
  gexDetLetra.Columns("Moneda").Width = 765
  gexDetLetra.Columns("Imp_Total").Width = 960
  gexDetLetra.Columns("Imp_Total").Format = "###,###.00"
  gexDetLetra.Columns("Importe_Cancelado").Width = 1500
  gexDetLetra.Columns("Importe_Cancelado").Format = "###,###.00"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

Select Case ActionName
  Case "IMPRIMIR"
    Reporte
  Case "SALIR"
    Unload Me
End Select
End Sub

Public Sub Reporte()
  
On Error GoTo ErrorImpresion

If gexDetLetra.RowCount = 0 Then Exit Sub

VB.Screen.MousePointer = vbHourglass

Dim oo As Object
Set oo = CreateObject("excel.application")

oo.Workbooks.Open vRuta & "\rptAdelantosCancelaciones.xlt"
oo.Visible = True
oo.Run "REPORTE", gexDetLetra.ADORecordset, UCase(Me.Caption)


Screen.MousePointer = vbNormal
oo.Visible = True
Set oo = Nothing

Exit Sub
Resume
ErrorImpresion:
    Screen.MousePointer = vbNormal
    Set oo = Nothing
    Error err.Number
End Sub

