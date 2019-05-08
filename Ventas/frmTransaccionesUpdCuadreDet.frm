VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmTransaccionesUpdCuadreDet 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3555
   ClientLeft      =   1005
   ClientTop       =   1290
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   9810
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3000
      TabIndex        =   1
      Top             =   2880
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   900
      Custom          =   $"frmTransaccionesUpdCuadreDet.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   4683
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmTransaccionesUpdCuadreDet.frx":00DE
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmTransaccionesUpdCuadreDet.frx":0430
      Column(2)       =   "frmTransaccionesUpdCuadreDet.frx":04F8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmTransaccionesUpdCuadreDet.frx":059C
      FormatStyle(2)  =   "frmTransaccionesUpdCuadreDet.frx":06D4
      FormatStyle(3)  =   "frmTransaccionesUpdCuadreDet.frx":0784
      FormatStyle(4)  =   "frmTransaccionesUpdCuadreDet.frx":0838
      FormatStyle(5)  =   "frmTransaccionesUpdCuadreDet.frx":0910
      FormatStyle(6)  =   "frmTransaccionesUpdCuadreDet.frx":09C8
      FormatStyle(7)  =   "frmTransaccionesUpdCuadreDet.frx":0AA8
      FormatStyle(8)  =   "frmTransaccionesUpdCuadreDet.frx":0F60
      ImageCount      =   1
      ImagePicture(1) =   "frmTransaccionesUpdCuadreDet.frx":13AC
      PrinterProperties=   "frmTransaccionesUpdCuadreDet.frx":16FE
   End
End
Attribute VB_Name = "frmTransaccionesUpdCuadreDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public StrSql As String, strCod_Anexo As String, strCod_TipAnexo, strTipo_Det As String, _
       intNum_Transaccion As Integer, dFecha As Date, strCod_Moneda As String, _
       dTipo_Cambio As Double, StrSql_Man As String, intSecuencia As Integer, strCod_Det As String

Sub Carga_Grid()

On Error GoTo errores

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(StrSql, cCONNECT)

If strTipo_Det = gcAnticipos Then
  GridEX1.Columns("Nro_Anticipo").Width = 1050
  GridEX1.Columns("Cod_Tipanex").Visible = False
  GridEX1.Columns("Cod_Anxo").Visible = False
Else
  GridEX1.Columns("Doc").Width = 435
  GridEX1.Columns("Serie").Width = 495
  GridEX1.Columns("Fecha").Width = 945
  GridEX1.Columns("num_corre").Visible = False
End If

GridEX1.Columns("Secuencia").Visible = False
GridEX1.Columns("Moneda").Width = 720
GridEX1.Columns("Imp_Cancelado").Width = 1245


Exit Sub
Resume
errores:
    errores Err.Number
End Sub
Sub Conf_Anticipo()
  GridEX1.Columns("Secuencia").Visible = False
  GridEX1.Columns("Nro_Anticipo").Width = 1050
  GridEX1.Columns("Moneda").Width = 720
  GridEX1.Columns("Imp_Cancelado").Width = 1245
End Sub

Sub Conf_Docs()
  GridEX1.Columns("Secuencia").Visible = False
  GridEX1.Columns("Doc").Width = 435
  GridEX1.Columns("Serie").Width = 495
  GridEX1.Columns("Fecha").Width = 945
  GridEX1.Columns("Moneda").Width = 720
  GridEX1.Columns("Imp_Cancelado").Width = 1245
End Sub

Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

Dim lvSql As String

On Error GoTo DrpDepurar

Select Case ActionName
Case Is = "MODIFICAR"
  If GridEX1.RowCount = 0 Then Exit Sub
  If strTipo_Det = gcAnticipos Then
    Carga_Mantenimieno StrSql_Man, True, False
  Else
    Carga_Mantenimieno StrSql_Man, False, True
  End If
Case Is = "ELIMINAR"
  If GridEX1.RowCount = 0 Then Exit Sub
  If MsgBox("Esta seguro de Eliminar este Detalle", vbYesNo, "IMPORTANTE") = vbYes Then
    If strTipo_Det = gcAnticipos Then
      lvSql = StrSql_Man & "'D','" & dFecha & "'," & GridEX1.Value(GridEX1.Columns("Secuencia").Index) & ",'" _
              & GridEX1.Value(GridEX1.Columns("Cod_Tipanex").Index) & "','" & GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index) & "'," _
              & GridEX1.Value(GridEX1.Columns("Nro_Anticipo").Index) & "," _
              & GridEX1.Value(GridEX1.Columns("Imp_Cancelado").Index) & "," & dTipo_Cambio
    Else
      lvSql = StrSql_Man & "'D','" & dFecha & "'," & GridEX1.Value(GridEX1.Columns("Secuencia").Index) & ",'" _
              & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "'," & GridEX1.Value(GridEX1.Columns("Imp_Cancelado").Index) & "," & dTipo_Cambio & ",'" & vusu & "'"
    End If
    ExecuteCommandSQL cCONNECT, lvSql
    Carga_Grid
  End If

Case Is = "SALIR"
  Unload Me
End Select

Exit Sub
Resume
DrpDepurar:

errores Err.Number

End Sub

Sub Carga_Mantenimieno(Store As String, dAnticipo As Boolean, dDoc As Boolean)
  With frmTransaccionesUpdCuadreMan
    .strCod_Moneda = strCod_Moneda
    .dFecha = dFecha
    .strStore = Store
    .frAnticipo.Visible = dAnticipo
    .frDocumento.Visible = dDoc
    .StrOption = "U"
    .intSecuencia = intSecuencia
    .strTipo_Det = strCod_Det
    If strTipo_Det = gcAnticipos Then
      .strCod_TipAnexo = GridEX1.Value(GridEX1.Columns("Cod_Tipanex").Index)
      .strCod_Anexo = GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index)
      .txtNro_Anticipo = GridEX1.Value(GridEX1.Columns("Nro_Anticipo").Index)
      .txtNro_Anticipo.Enabled = False
    Else
      .strCod_TipAnexo = strCod_TipAnexo
      .strCod_Anexo = strCod_Anexo
      .txtCod_TipDoc = GridEX1.Value(GridEX1.Columns("Doc").Index)
      .txtSer_Docum = GridEX1.Value(GridEX1.Columns("Serie").Index)
      .txtNum_Docum = GridEX1.Value(GridEX1.Columns("Nro").Index)
      .txtCod_TipDoc.Enabled = False
      .txtDes_TipDoc.Enabled = False
      .txtSer_Docum.Enabled = False
      .txtNum_Docum.Enabled = False
      .strNum_Corre = GridEX1.Value(GridEX1.Columns("Num_Corre").Index)
    End If
    .TxtTipo_Cambio = dTipo_Cambio
    .txtCod_Moneda = GridEX1.Value(GridEX1.Columns("Moneda").Index)
    
    If GridEX1.Value(GridEX1.Columns("Moneda").Index) = strCod_Moneda Then
      .txtImporte = GridEX1.Value(GridEX1.Columns("Imp_Cancelado").Index)
    Else
      If GridEX1.Value(GridEX1.Columns("Moneda").Index) = "SOL" Then
        .txtImporte = GridEX1.Value(GridEX1.Columns("Imp_Cancelado").Index) * dTipo_Cambio
      Else
        .txtImporte = GridEX1.Value(GridEX1.Columns("Imp_Cancelado").Index) / dTipo_Cambio
      End If
    End If
    .Calcula_Importe_Converido
    .Caption = Me.Caption & " ( Modifiacion ) "
    .Show 1
    If .lfAceptar Then Carga_Grid
  End With
End Sub

