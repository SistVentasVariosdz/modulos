VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmEstilosNPs 
   Caption         =   "Estilos NPs"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   DrawMode        =   16  'Merge Pen
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   4200
      Width           =   1935
   End
   Begin GridEX20.GridEX DGridLista 
      Height          =   3825
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6747
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmEstilosNPs.frx":0000
      Column(2)       =   "frmEstilosNPs.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmEstilosNPs.frx":016C
      FormatStyle(2)  =   "frmEstilosNPs.frx":02A4
      FormatStyle(3)  =   "frmEstilosNPs.frx":0354
      FormatStyle(4)  =   "frmEstilosNPs.frx":0408
      FormatStyle(5)  =   "frmEstilosNPs.frx":04E0
      FormatStyle(6)  =   "frmEstilosNPs.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmEstilosNPs.frx":0678
   End
End
Attribute VB_Name = "frmEstilosNPs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public codCliente As String
Public codTemporada As String
Public codItem As String


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
CargaLista
End Sub

Public Sub CargaLista()
    Dim StrSQL As String
    Dim xRow As Variant
    
    'Esta cadena es para devolver el Codigo de Cliente
     
    'StrSQL = "EXEC ES_SM_ItemServicios_ClienteTemp '" & Opcion & "','" & DevuelveCampo(StrSQL, cCONNECT) & "','" & txttemporada.Text & "','" & txtcod_item.Text & "','" & txtCodStatus.Text & "', '" & txtCodProveedor2.Text & "' "
    StrSQL = "EXEC ES_MUESTRA_OPS_TEMPORADA_CLIENTE_ITEM '" & codCliente & "','" & codTemporada & "','" & codItem & "'"
    
    xRow = DGridLista.Row
    Set DGridLista.ADORecordset = CargarRecordSetDesconectado(StrSQL, cCONNECT)
    DGridLista.Columns("OP").Width = 1000
    DGridLista.Columns("OP").Caption = "Op"
    
    DGridLista.Columns("COMPONENTE").Width = 2000
    DGridLista.Columns("COMPONENTE").Caption = "Componente"
    
    DGridLista.Columns("ESTILO_VERSION").Width = 1500
    DGridLista.Columns("ESTILO_VERSION").Caption = "Estilo Versión"
    
    
    DGridLista.Row = xRow
    DGridLista.Enabled = True
    'SeteaGrid
End Sub

