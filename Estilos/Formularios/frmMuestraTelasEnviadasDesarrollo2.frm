VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form frmMuestraTelasEnviadasDesarrollo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Telas Enviadas a Desarrollo de Comercial"
   ClientHeight    =   6120
   ClientLeft      =   990
   ClientTop       =   1650
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   10110
   Begin GridEX20.GridEX GridEX1 
      Height          =   5400
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   9525
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmMuestraTelasEnviadasDesarrollo.frx":0000
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmMuestraTelasEnviadasDesarrollo.frx":0352
      Column(2)       =   "frmMuestraTelasEnviadasDesarrollo.frx":041A
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmMuestraTelasEnviadasDesarrollo.frx":04BE
      FormatStyle(2)  =   "frmMuestraTelasEnviadasDesarrollo.frx":05F6
      FormatStyle(3)  =   "frmMuestraTelasEnviadasDesarrollo.frx":06A6
      FormatStyle(4)  =   "frmMuestraTelasEnviadasDesarrollo.frx":075A
      FormatStyle(5)  =   "frmMuestraTelasEnviadasDesarrollo.frx":0832
      FormatStyle(6)  =   "frmMuestraTelasEnviadasDesarrollo.frx":08EA
      FormatStyle(7)  =   "frmMuestraTelasEnviadasDesarrollo.frx":09CA
      FormatStyle(8)  =   "frmMuestraTelasEnviadasDesarrollo.frx":0E82
      ImageCount      =   1
      ImagePicture(1) =   "frmMuestraTelasEnviadasDesarrollo.frx":12CE
      PrinterProperties=   "frmMuestraTelasEnviadasDesarrollo.frx":1620
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3480
      TabIndex        =   1
      Top             =   5520
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmMuestraTelasEnviadasDesarrollo.frx":17F8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmMuestraTelasEnviadasDesarrollo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
carga_grid
End Sub

Sub carga_grid()
Dim strSQL As String
On Error GoTo hand
  VB.Screen.MousePointer = vbHourglass
  strSQL = "TX_DESARROLLO_MUESTRA_TELAS_ENVIADAS_A_DESARROLLO"
  Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
  Configura_Grid
  VB.Screen.MousePointer = 0
Exit Sub
Resume
hand:
    VB.Screen.MousePointer = 0
    ErrorHandler Err, "LOAD"
End Sub
Sub Configura_Grid()
  GridEX1.Columns("Cod_Tela").Width = 930
  GridEX1.Columns("Cod_Tela").Caption = "Cod Tela"
  GridEX1.Columns("Des_Tela").Width = 3105
  GridEX1.Columns("Des_Tela").Caption = "Descripcion"
  GridEX1.Columns("Ancho_Acab").Width = 1260
  GridEX1.Columns("Ancho_Acab").Caption = "Ancho Acab"
  GridEX1.Columns("Gramaje_Acab").Width = 1290
  GridEX1.Columns("Gramaje_Acab").Caption = "Gramaje Acab"
  GridEX1.Columns("Fec_Envio_a_Desarrollo").Width = 1665
  GridEX1.Columns("Fec_Envio_a_Desarrollo").Caption = "Fecha Envio"
  GridEX1.Columns("Fec_Envio_a_Desarrollo").Format = "dd/mm/yyyy"
  GridEX1.Columns("Usuario_Envio_Desarrollo").Width = 1725
  GridEX1.Columns("Usuario_Envio_Desarrollo").Caption = "Usuario Envio"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

On Error GoTo hand

Dim strSQL As String

Select Case ActionName
Case Is = "RECEPCIONAR"
  If GridEX1.RowCount = 0 Then Exit Sub
  If MsgBox("Esta seguro de Recepcionar la Tela " & vbCr & GridEX1.Value(GridEX1.Columns("Cod_Tela").Index) & GridEX1.Value(GridEX1.Columns("Des_Tela").Index), vbYesNo, "ADVERTENCIA") = vbYes Then
    strSQL = "TX_DESARROLLO_CAMBIA_ESTADO_TELA_A_RECIBIDO '" & GridEX1.Value(GridEX1.Columns("Cod_Tela").Index) & "','" & ComputerName & "','" & vusu & "'"
    ExecuteCommandSQL cCONNECT, strSQL
    carga_grid
  End If
Case Is = "SALIR"
  Unload Me
End Select

Exit Sub
Resume
hand:
    VB.Screen.MousePointer = 0
    ErrorHandler Err, "RECEPCION"
End Sub
