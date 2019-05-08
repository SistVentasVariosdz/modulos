VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmVerSalidasxDua 
   Caption         =   "Salidas por DUA"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   8280
      TabIndex        =   1
      Top             =   4800
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   900
      Custom          =   "0~0~SALUIR~Verdadero~Verdadero~&Salir~0~0~1~~0~Falso~Falso~&Salir~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX DGridLista 
      Height          =   4740
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   8361
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "FrmVerSalidasxDua.frx":0000
      Column(2)       =   "FrmVerSalidasxDua.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "FrmVerSalidasxDua.frx":016C
      FormatStyle(2)  =   "FrmVerSalidasxDua.frx":02A4
      FormatStyle(3)  =   "FrmVerSalidasxDua.frx":0354
      FormatStyle(4)  =   "FrmVerSalidasxDua.frx":0408
      FormatStyle(5)  =   "FrmVerSalidasxDua.frx":04E0
      FormatStyle(6)  =   "FrmVerSalidasxDua.frx":0598
      ImageCount      =   0
      PrinterProperties=   "FrmVerSalidasxDua.frx":0678
   End
End
Attribute VB_Name = "FrmVerSalidasxDua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Ser_OrdComp As String, Cod_OrdComp As String, Sec_OrdComp
Dim strSQL As String


Sub CARGA_GRID()
On Error GoTo errCargaGrid
strSQL = "LG_CONSULTA_SALIDAS_POR_DUA_OC '" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Sec_OrdComp & "'"
Set DGridLista.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)

DGridLista.Columns("NP").Width = 1000
DGridLista.Columns("Cod_Almacen").Width = 900
DGridLista.Columns("Num_MovStk").Width = 1400
DGridLista.Columns("Num_Secuencia").Width = 900
DGridLista.Columns("Cantidad_Asignada").Width = 1300

DGridLista.Columns("Cod_Almacen").Caption = "Almacen"
DGridLista.Columns("Num_MovStk").Caption = "Movimiento"
DGridLista.Columns("Num_Secuencia").Caption = "Secuencia"
DGridLista.Columns("Cantidad_Asignada").Caption = "Can.Asignada"

Exit Sub
errCargaGrid:
    MsgBox err.Description, vbCritical, "Carga Grid"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Unload Me
End Sub
