VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmGastosAsociados 
   Caption         =   "MUESTRA EMBARQUE"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   9120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9045
      _ExtentX        =   15954
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
      ColumnsCount    =   2
      Column(1)       =   "frmGastosAsociados.frx":0000
      Column(2)       =   "frmGastosAsociados.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmGastosAsociados.frx":016C
      FormatStyle(2)  =   "frmGastosAsociados.frx":02A4
      FormatStyle(3)  =   "frmGastosAsociados.frx":0354
      FormatStyle(4)  =   "frmGastosAsociados.frx":0408
      FormatStyle(5)  =   "frmGastosAsociados.frx":04E0
      FormatStyle(6)  =   "frmGastosAsociados.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmGastosAsociados.frx":0678
   End
End
Attribute VB_Name = "frmGastosAsociados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public vnum_emb As String
 
 

Public Function BUSCAR() As Boolean
On Error GoTo errores
Dim ssql As String
Dim vBookmark As Variant

ssql = "TG_EMBARQUE_MUESTRA_CN_DOCUM '$'"
ssql = VBsprintf(ssql, vnum_emb)
  
vBookmark = GridEX1.Row
GridEX1.ClearFields

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(ssql, cCONNECT)

GridEX1.Row = vBookmark


GridEX1.ContinuousScroll = True

GridEX1.FrozenColumns = 3
Exit Function

errores:
    errores err.Number
End Function

Private Sub btnSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim seguridad As String
seguridad = get_botones1(Me, vusu, vemp, Me.Name)

End Sub
