VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMuestraGeneral 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4530
   ClientLeft      =   2265
   ClientTop       =   2745
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   4275
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   840
      TabIndex        =   1
      Top             =   3840
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmMuestraGeneral.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   6165
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmMuestraGeneral.frx":0090
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmMuestraGeneral.frx":03E2
      Column(2)       =   "frmMuestraGeneral.frx":04AA
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmMuestraGeneral.frx":054E
      FormatStyle(2)  =   "frmMuestraGeneral.frx":0686
      FormatStyle(3)  =   "frmMuestraGeneral.frx":0736
      FormatStyle(4)  =   "frmMuestraGeneral.frx":07EA
      FormatStyle(5)  =   "frmMuestraGeneral.frx":08C2
      FormatStyle(6)  =   "frmMuestraGeneral.frx":097A
      FormatStyle(7)  =   "frmMuestraGeneral.frx":0A5A
      FormatStyle(8)  =   "frmMuestraGeneral.frx":0F12
      ImageCount      =   1
      ImagePicture(1) =   "frmMuestraGeneral.frx":135E
      PrinterProperties=   "frmMuestraGeneral.frx":16B0
   End
End
Attribute VB_Name = "frmMuestraGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strSQL As String, sCod_Cliente As String, sCod_TempCli As String

Public Function BUSCAR() As Boolean
On Error GoTo errores
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
'GridEX1.FrozenColumns = 3

Exit Function
errores:
    errores Err.Number
End Function

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
  If GridEX1.RowCount <> 0 Then Reporte
Case Is = "SALIR"
  Unload Me
End Select
End Sub

Sub Reporte()

Dim oo As Object

On Error GoTo ErrorImpresion

Set oo = CreateObject("excel.application")
oo.workbooks.Open vRuta & "\RptMuestras.XLT"
oo.Visible = True
oo.DisplayAlerts = False
oo.run "reporte", "es_muestra_matriz_muestras_estilos_colores_po '" & sCod_Cliente & "','" & sCod_TempCli & "','" & GridEX1.Value(GridEX1.Columns("COD_PURORD").Index) & "'", cCONNECT, Left(Me.Caption, InStr(1, Me.Caption, "/", vbTextCompare) - 1), Right(Me.Caption, Len(Me.Caption) - InStr(1, Me.Caption, "/", vbTextCompare)), "M1", GridEX1.Value(GridEX1.Columns("COD_PURORD").Index)
Set oo = Nothing
Unload Me
Exit Sub
    
ErrorImpresion:

    Set oo = Nothing
    MsgBox Err.Description, vbCritical, "Impresion"

End Sub
