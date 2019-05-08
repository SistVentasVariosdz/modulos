VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmVerRevisionesTelas 
   Caption         =   "Revisiones Tela"
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
      FormatStyle(1)  =   "FrmVerRevisionesTelas.frx":0000
      FormatStyle(2)  =   "FrmVerRevisionesTelas.frx":0138
      FormatStyle(3)  =   "FrmVerRevisionesTelas.frx":01E8
      FormatStyle(4)  =   "FrmVerRevisionesTelas.frx":029C
      FormatStyle(5)  =   "FrmVerRevisionesTelas.frx":0374
      FormatStyle(6)  =   "FrmVerRevisionesTelas.frx":042C
      FormatStyle(7)  =   "FrmVerRevisionesTelas.frx":050C
      ImageCount      =   0
      PrinterProperties=   "FrmVerRevisionesTelas.frx":052C
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2520
      TabIndex        =   1
      Top             =   3960
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmVerRevisionesTelas.frx":0704
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmVerRevisionesTelas"
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
Case "AGREGAR"
    Load FrmRevisionTela
    FrmRevisionTela.vCod_Tela = Me.vCod_Tela
    FrmRevisionTela.txtcod_tela = Me.vCod_Tela
    FrmRevisionTela.txtdes_tela.Text = Me.vDes_Tela
    FrmRevisionTela.Show vbModal
    Set FrmRevisionTela = Nothing
    Call CARGA_GRID
Case "SALIR"
    Unload Me
End Select
End Sub

Sub CARGA_GRID()
On Error GoTo errGrid

strSQL = "tx_sm_muestra_tx_telas_revisadas '" & vCod_Tela & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

GridEX1.Columns("Num_Secuencia").Width = 700
GridEX1.Columns("Cod_Usuario").Width = 900
GridEX1.Columns("Fec_Revision").Width = 1100
GridEX1.Columns("Observaciones").Width = 2000

GridEX1.Columns("Num_Secuencia").Caption = "Sec."

Exit Sub
errGrid:
    MsgBox Err.Description, vbCritical, "Grid"
End Sub
