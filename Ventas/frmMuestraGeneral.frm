VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMuestraGeneral 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4530
   ClientLeft      =   285
   ClientTop       =   1275
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   11385
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   4800
      TabIndex        =   1
      Top             =   3840
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
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   6165
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmMuestraGeneral.frx":0000
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmMuestraGeneral.frx":0352
      Column(2)       =   "frmMuestraGeneral.frx":041A
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmMuestraGeneral.frx":04BE
      FormatStyle(2)  =   "frmMuestraGeneral.frx":05F6
      FormatStyle(3)  =   "frmMuestraGeneral.frx":06A6
      FormatStyle(4)  =   "frmMuestraGeneral.frx":075A
      FormatStyle(5)  =   "frmMuestraGeneral.frx":0832
      FormatStyle(6)  =   "frmMuestraGeneral.frx":08EA
      FormatStyle(7)  =   "frmMuestraGeneral.frx":09CA
      FormatStyle(8)  =   "frmMuestraGeneral.frx":0E82
      ImageCount      =   1
      ImagePicture(1) =   "frmMuestraGeneral.frx":12CE
      PrinterProperties=   "frmMuestraGeneral.frx":1620
   End
End
Attribute VB_Name = "frmMuestraGeneral"
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
    errores Err.Number
End Function

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case Is = "SALIR"
  Unload Me
End Select
End Sub

