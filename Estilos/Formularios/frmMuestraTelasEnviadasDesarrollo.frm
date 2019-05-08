VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form frmMuestraTelasEnviadasDesarrollo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Telas Enviadas a Desarrollo de Comercial"
   ClientHeight    =   7035
   ClientLeft      =   1170
   ClientTop       =   1560
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   10110
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar Tela"
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdRepcepcionar 
      Caption         =   "&Recepcionar Tela"
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10020
      Begin VB.OptionButton optCerrar 
         Caption         =   "Telas a &Cerrar"
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optRecepcion 
         Caption         =   "Telas a &Recepcionar"
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   510
         Left            =   8715
         TabIndex        =   2
         Top             =   195
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   900
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5400
      Left            =   0
      TabIndex        =   3
      Top             =   840
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
      DataMode        =   1
      ColumnHeaderHeight=   285
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmMuestraTelasEnviadasDesarrollo.frx":0000
      FormatStyle(2)  =   "frmMuestraTelasEnviadasDesarrollo.frx":0138
      FormatStyle(3)  =   "frmMuestraTelasEnviadasDesarrollo.frx":01E8
      FormatStyle(4)  =   "frmMuestraTelasEnviadasDesarrollo.frx":029C
      FormatStyle(5)  =   "frmMuestraTelasEnviadasDesarrollo.frx":0374
      FormatStyle(6)  =   "frmMuestraTelasEnviadasDesarrollo.frx":042C
      FormatStyle(7)  =   "frmMuestraTelasEnviadasDesarrollo.frx":050C
      ImageCount      =   0
      PrinterProperties=   "frmMuestraTelasEnviadasDesarrollo.frx":052C
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   1080
      Top             =   6480
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmMuestraTelasEnviadasDesarrollo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
On Error GoTo hand

Dim strSQL As String

  If GridEX1.RowCount = 0 Then Exit Sub
  If MsgBox("Esta seguro de Cerrar la Tela " & vbCr & GridEX1.Value(GridEX1.Columns("Cod_Tela").Index) & GridEX1.Value(GridEX1.Columns("Des_Tela").Index), vbYesNo, "ADVERTENCIA") = vbYes Then
    strSQL = "TX_DESARROLLO_CAMBIA_ESTADO_TELA_A_CERRADO '" & GridEX1.Value(GridEX1.Columns("Cod_Tela").Index) & "','" & ComputerName & "','" & vusu & "'"
    ExecuteCommandSQL cCONNECT, strSQL
    CARGA_GRID
  End If

Exit Sub
Resume
hand:
    VB.Screen.MousePointer = 0
    ErrorHandler Err, "RECEPCION"
End Sub

Private Sub cmdRepcepcionar_Click()
On Error GoTo hand

Dim strSQL As String

  If GridEX1.RowCount = 0 Then Exit Sub
  If MsgBox("Esta seguro de Recepcionar la Tela " & vbCr & GridEX1.Value(GridEX1.Columns("Cod_Tela").Index) & GridEX1.Value(GridEX1.Columns("Des_Tela").Index), vbYesNo, "ADVERTENCIA") = vbYes Then
    strSQL = "TX_DESARROLLO_CAMBIA_ESTADO_TELA_A_RECIBIDO '" & GridEX1.Value(GridEX1.Columns("Cod_Tela").Index) & "','" & ComputerName & "','" & vusu & "'"
    ExecuteCommandSQL cCONNECT, strSQL
    CARGA_GRID
  End If

Exit Sub
Resume
hand:
    VB.Screen.MousePointer = 0
    ErrorHandler Err, "RECEPCION"

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Sub CARGA_GRID()
Dim strSQL As String
On Error GoTo hand
  VB.Screen.MousePointer = vbHourglass
  strSQL = "TX_DESARROLLO_MUESTRA_TELAS_ENVIADAS_A_DESARROLLO '" & IIf(optRecepcion, "E", "T") & "'"
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

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
  CARGA_GRID
End Sub

Private Sub optCerrar_Click()
  cmdRepcepcionar.Visible = False
  cmdCerrar.Visible = True
End Sub

Private Sub optRecepcion_Click()
  cmdRepcepcionar.Visible = True
  cmdCerrar.Visible = False
End Sub
