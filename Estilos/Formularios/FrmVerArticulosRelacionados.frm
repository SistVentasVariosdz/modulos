VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmVerArticulosRelacionados 
   Caption         =   "Artículos Relacionados"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   Begin GridEX20.GridEX GridEX1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6800
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "FrmVerArticulosRelacionados.frx":0000
      FormatStyle(2)  =   "FrmVerArticulosRelacionados.frx":0138
      FormatStyle(3)  =   "FrmVerArticulosRelacionados.frx":01E8
      FormatStyle(4)  =   "FrmVerArticulosRelacionados.frx":029C
      FormatStyle(5)  =   "FrmVerArticulosRelacionados.frx":0374
      FormatStyle(6)  =   "FrmVerArticulosRelacionados.frx":042C
      FormatStyle(7)  =   "FrmVerArticulosRelacionados.frx":050C
      ImageCount      =   0
      PrinterProperties=   "FrmVerArticulosRelacionados.frx":052C
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3157
      TabIndex        =   1
      Top             =   3960
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
End
Attribute VB_Name = "FrmVerArticulosRelacionados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public vCod_Tela As String, vDes_Tela As String

Private Sub Form_Load()
FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "SALIR"
    Unload Me
End Select
End Sub

Sub CARGA_GRID()
On Error GoTo errGrid

strSQL = "SM_CONULTAR_ARTICULOS_RELACIONADOS '" & vCod_Tela & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

GridEX1.Columns("COD_TELA").Width = 1500
GridEX1.Columns("DES_TELA").Width = 5500

Exit Sub
errGrid:
    MsgBox Err.Description, vbCritical, "Grid"
    
End Sub
