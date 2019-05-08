VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmListaPacking 
   Caption         =   "Lista de Packing por Venta"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   495
      Left            =   8520
      TabIndex        =   1
      Top             =   3720
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      Custom          =   "0~0~SALIR~Verdadero~Verdadero~&Salir~0~0~1~~0~Falso~Falso~&Salir~"
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX dgLista 
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   5953
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "FrmListaPacking.frx":0000
      Column(2)       =   "FrmListaPacking.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "FrmListaPacking.frx":016C
      FormatStyle(2)  =   "FrmListaPacking.frx":02A4
      FormatStyle(3)  =   "FrmListaPacking.frx":0354
      FormatStyle(4)  =   "FrmListaPacking.frx":0408
      FormatStyle(5)  =   "FrmListaPacking.frx":04E0
      FormatStyle(6)  =   "FrmListaPacking.frx":0598
      ImageCount      =   0
      PrinterProperties=   "FrmListaPacking.frx":0678
   End
End
Attribute VB_Name = "FrmListaPacking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public numcorre As String

Private Sub cmdSalir_Click()
Unload Me
End Sub


Sub cargarGrid()
On Error GoTo dprDepurar

Dim sSql As String

sSql = "VENTAS_Muestra_PackingList  '" & numcorre & "'"

Set dgLista.ADORecordset = CargarRecordSetDesconectado(sSql, cCONNECT)

dgLista.ContinuousScroll = True
dgLista.AllowEdit = False

Exit Sub

dprDepurar:

errores err.Number

End Sub

Private Sub Form_Load()
Dim sSeguridad  As String
sSeguridad = get_botones1(Me, vper, vemp, Me.Name)
FunctButt1.FunctionsUser = sSeguridad
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
             Case "SALIR"
            Unload Me
    End Select
End Sub
