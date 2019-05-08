VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form frmBusqPartidas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda de Partidas"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   8460
   StartUpPosition =   2  'CenterScreen
   Begin GridEX20.GridEX GridEX1 
      Height          =   3270
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5768
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorBkg    =   -2147483624
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmBusqPartidas.frx":0000
      Column(2)       =   "frmBusqPartidas.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmBusqPartidas.frx":016C
      FormatStyle(2)  =   "frmBusqPartidas.frx":02A4
      FormatStyle(3)  =   "frmBusqPartidas.frx":0354
      FormatStyle(4)  =   "frmBusqPartidas.frx":0408
      FormatStyle(5)  =   "frmBusqPartidas.frx":04E0
      FormatStyle(6)  =   "frmBusqPartidas.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmBusqPartidas.frx":0678
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2895
      TabIndex        =   0
      Tag             =   "&OK"
      Top             =   3480
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4380
      TabIndex        =   1
      Tag             =   "&Cancel"
      Top             =   3465
      Width           =   1185
   End
End
Attribute VB_Name = "frmBusqPartidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public sQuery As String, sCod_TipOrdTra As String
Dim Rs_Carga As New ADODB.Recordset
Public Codigo As String, Descripcion As String

Sub CARGAR_DATOS()
On Error GoTo Cargar_DatosErr

Rs_Carga.ActiveConnection = cConnect
Rs_Carga.CursorType = adOpenStatic
Rs_Carga.CursorLocation = adUseClient
Rs_Carga.LockType = adLockReadOnly
Rs_Carga.Open sQuery
Set GridEX1.ADORecordset = Rs_Carga
With GridEX1
    
    If sCod_TipOrdTra = "TJ" Then
        .Columns("COD_TIPORDTRA").Width = 510
        .Columns("OT").Width = 480
        .Columns("LOTE").Width = 3120
    Else
        .Columns("COD_TIPORDTRA").Caption = "Tip.Orden"
        .Columns("COD_ORDTRA").Caption = "Nro.Orden"
        .Columns("KGS_PROGR").Caption = "Kgs.Prog."
        .Columns("KGS_CRUDO").Caption = "Kgs.Crudo"
        .Columns("COD_ORDPRO_TEXT").Caption = "Ord.Pro."
        .Columns("FEC_1ER_ENVIO").Caption = "1er.Envío"
        .Columns("GUIAS").Caption = "Guías"
        
        .Columns("COD_TIPORDTRA").Width = 800
        .Columns("COD_ORDTRA").Width = 800
        .Columns("KGS_PROGR").Width = 1000
        .Columns("KGS_CRUDO").Width = 1000
        .Columns("COD_ORDPRO_TEXT").Width = 2000
        .Columns("FEC_1ER_ENVIO").Width = 1300
        .Columns("GUIAS").Width = 1000
    End If
End With
Exit Sub
Cargar_DatosErr:
    Set Rs_Carga = Nothing
    ErrorHandler Err, "Cargar_Datos"
End Sub

Private Sub Form_Load()
Call FormSet(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Rs_Carga = Nothing
End Sub

Private Sub cmdAceptar_Click()
    GridEX1_DblClick
End Sub

Private Sub cmdCancelar_Click()
    With oParent
        .Codigo = ""
        .Descripcion = ""
    End With
Unload Me
End Sub

Private Sub GridEX1_DblClick()
On Error Resume Next
If GridEX1.RowCount > 0 Then
    With oParent
        .Codigo = GridEX1.Value(1)
        .Descripcion = GridEX1.Value(2)
    End With
End If
Unload Me
End Sub

Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GridEX1_DblClick
End Sub
