VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmTransaccionesDetalle_Cheque_Diferido 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4770
   ClientLeft      =   240
   ClientTop       =   1500
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   10125
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2400
      TabIndex        =   1
      Top             =   4080
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   900
      Custom          =   $"frmTransaccionesDetalle_Cheque_Diferido.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   6800
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmTransaccionesDetalle_Cheque_Diferido.frx":012C
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmTransaccionesDetalle_Cheque_Diferido.frx":047E
      Column(2)       =   "frmTransaccionesDetalle_Cheque_Diferido.frx":0546
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmTransaccionesDetalle_Cheque_Diferido.frx":05EA
      FormatStyle(2)  =   "frmTransaccionesDetalle_Cheque_Diferido.frx":0722
      FormatStyle(3)  =   "frmTransaccionesDetalle_Cheque_Diferido.frx":07D2
      FormatStyle(4)  =   "frmTransaccionesDetalle_Cheque_Diferido.frx":0886
      FormatStyle(5)  =   "frmTransaccionesDetalle_Cheque_Diferido.frx":095E
      FormatStyle(6)  =   "frmTransaccionesDetalle_Cheque_Diferido.frx":0A16
      FormatStyle(7)  =   "frmTransaccionesDetalle_Cheque_Diferido.frx":0AF6
      FormatStyle(8)  =   "frmTransaccionesDetalle_Cheque_Diferido.frx":0FAE
      ImageCount      =   1
      ImagePicture(1) =   "frmTransaccionesDetalle_Cheque_Diferido.frx":13FA
      PrinterProperties=   "frmTransaccionesDetalle_Cheque_Diferido.frx":174C
   End
End
Attribute VB_Name = "frmTransaccionesDetalle_Cheque_Diferido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dFecha As Date, intSecuencia As Integer, strCod_Banco As String, strCod_Moneda As String

Public Sub Buscar()

On Error GoTo errores

Dim StrSql As String

StrSql = "Ventas_Man_Cheques_Diferidos_Detalle 'V','" & dFecha & "'," & intSecuencia

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(StrSql, cCONNECT)
GridEX1.Columns("Nro_Cheque").Width = 1530
GridEX1.Columns("Importe").Width = 975
GridEX1.Columns("Ruc").Width = 1740
GridEX1.Columns("Anexo").Width = 3420
GridEX1.Columns("Cod_Banco").Visible = False

Exit Sub
Resume
errores:
    errores Err.Number
End Sub


Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

Dim lvSql As String

On Error GoTo DrpDepurar

Select Case ActionName
Case Is = "ADICIONAR"
  With frmTransaccionesDetalle_Cheque_DiferidoMan
    .strCod_Banco = strCod_Banco
    .strCod_Moneda = strCod_Moneda
    .Caption = "Adiciona Cheque Diferido"
    .strStore = "Ventas_Man_Cheques_Diferidos_Detalle 'I','" & dFecha & "'," & intSecuencia
    .Show 1
    If .lfAceptar Then Buscar
  End With
Case Is = "MODIFICAR"
  If GridEX1.RowCount = 0 Then Exit Sub
  
  Load frmTransaccionesDetalle_Cheque_DiferidoMan
  With frmTransaccionesDetalle_Cheque_DiferidoMan
    .Caption = "Modificar Cheque Diferido"
    .strStore = "Ventas_Man_Cheques_Diferidos_Detalle 'U','" & dFecha & "'," & intSecuencia
    .strCod_Banco = GridEX1.Value(GridEX1.Columns("Cod_Banco").Index)
    .strCod_Moneda = GridEX1.Value(GridEX1.Columns("Moneda").Index)
    .txtNro = GridEX1.Value(GridEX1.Columns("Nro_Cheque").Index)
    .txtNro.Enabled = False
    .txt_Importe.Text = GridEX1.Value(GridEX1.Columns("Importe").Index)
    .Show 1
    If .lfAceptar Then Buscar
  End With
Case Is = "ELIMINAR"
  If GridEX1.RowCount = 0 Then Exit Sub
  
  If GridEX1.RowCount = 0 Then Exit Sub
  If MsgBox("Esta Seguro de Eliminar este Registro", vbYesNo, "ADVERTENCIA") = vbYes Then
     lvSql = "Ventas_Man_Cheques_Diferidos_Detalle 'D','" & dFecha & "'," & intSecuencia & ",'" & GridEX1.Value(GridEX1.Columns("Cod_Banco").Index) & "','" & GridEX1.Value(GridEX1.Columns("Moneda").Index) & "','" & GridEX1.Value(GridEX1.Columns("Nro_Cheque").Index) & "'"
    ExecuteCommandSQL cCONNECT, lvSql
    Buscar
  End If
Case Is = "SALIR"
  Unload Me
End Select

Exit Sub
Resume
DrpDepurar:

errores Err.Number

End Sub

