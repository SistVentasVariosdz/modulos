VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmCosGerenciaDetHilos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6495
   ClientLeft      =   960
   ClientTop       =   2865
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   10125
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   4200
      TabIndex        =   1
      Top             =   5880
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   900
      Custom          =   "0~0~SALIR~Verdadero~Verdadero~&Salir~0~0~1~~0~Falso~Falso~&Salir~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   9975
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmCosGerenciaDetHilos.frx":0000
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmCosGerenciaDetHilos.frx":0352
      Column(2)       =   "frmCosGerenciaDetHilos.frx":041A
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmCosGerenciaDetHilos.frx":04BE
      FormatStyle(2)  =   "frmCosGerenciaDetHilos.frx":05F6
      FormatStyle(3)  =   "frmCosGerenciaDetHilos.frx":06A6
      FormatStyle(4)  =   "frmCosGerenciaDetHilos.frx":075A
      FormatStyle(5)  =   "frmCosGerenciaDetHilos.frx":0832
      FormatStyle(6)  =   "frmCosGerenciaDetHilos.frx":08EA
      FormatStyle(7)  =   "frmCosGerenciaDetHilos.frx":09CA
      FormatStyle(8)  =   "frmCosGerenciaDetHilos.frx":0E82
      ImageCount      =   1
      ImagePicture(1) =   "frmCosGerenciaDetHilos.frx":12CE
      PrinterProperties=   "frmCosGerenciaDetHilos.frx":1620
   End
End
Attribute VB_Name = "frmCosGerenciaDetHilos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strSQL As String, lvCodAlmace As String, lvCodTelaCruda As String

Public Function BUSCAR() As Boolean
On Error GoTo errores
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
'GridEX1.FrozenColumns = 3

Exit Function
errores:
    errores err.Number
End Function

Private Sub Form_Load()

End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case Is = "SALIR"
  Unload Me
End Select
End Sub

