VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmDetLetra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Letra"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   Icon            =   "frmDetLetra.frx":0000
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
      Custom          =   $"frmDetLetra.frx":030A
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
         Column(1)       =   "frmDetLetra.frx":039A
         Column(2)       =   "frmDetLetra.frx":0462
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmDetLetra.frx":0506
         FormatStyle(2)  =   "frmDetLetra.frx":063E
         FormatStyle(3)  =   "frmDetLetra.frx":06EE
         FormatStyle(4)  =   "frmDetLetra.frx":07A2
         FormatStyle(5)  =   "frmDetLetra.frx":087A
         FormatStyle(6)  =   "frmDetLetra.frx":0932
         ImageCount      =   0
         PrinterProperties=   "frmDetLetra.frx":0A12
      End
   End
End
Attribute VB_Name = "frmDetLetra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vNum_Correlativo As String

Public Sub CARGA_GRID()
On Error GoTo hand
Dim RS As ADODB.Recordset

    Set RS = CreateObject("ADODB.Recordset")
    RS.CursorLocation = adUseClient
    RS.Open "EXEC Ventas_Letras_Muestra_Detalle '" & vNum_Correlativo & "'", cCONNECT

    Set gexDetLetra.ADORecordset = RS

    Set RS = Nothing

    ConfigurarGrid

Exit Sub
hand:
ErrorHandler err, "CARGA_GRID"
Set RS = Nothing

End Sub

Sub ConfigurarGrid()
  gexDetLetra.Columns("Nro_Doc").Width = 1470
  gexDetLetra.Columns("Fecha_Emision").Width = 1215
  gexDetLetra.Columns("Fecha_Vencimiento").Width = 1545
  gexDetLetra.Columns("Moneda").Width = 720
  gexDetLetra.Columns("Importe").Width = 930
  gexDetLetra.Columns("Importe").Format = "###,###.00"
  gexDetLetra.Columns("Num_Corre").Width = 1500
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

Select Case ActionName
  Case "IMPRIMIR"
    'Reporte
  Case "SALIR"
    Unload Me
End Select
End Sub


