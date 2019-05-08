VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form frmHiladosRequeridos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hilados Requeridos"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "frmHiladosRequeridos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   7560
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   8775
      Begin GridEX20.GridEX gexDetalle 
         Height          =   2655
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   8460
         _ExtentX        =   14923
         _ExtentY        =   4683
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         RecordNavigatorString=   "Registro:|de"
         HoldSortSettings=   -1  'True
         GridLineStyle   =   2
         MethodHoldFields=   -1  'True
         ContScroll      =   -1  'True
         Options         =   8
         RecordsetType   =   1
         GroupByBoxInfoText=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ImageCount      =   3
         ImagePicture1   =   "frmHiladosRequeridos.frx":08CA
         ImagePicture2   =   "frmHiladosRequeridos.frx":09DC
         ImagePicture3   =   "frmHiladosRequeridos.frx":0CF6
         RowHeaders      =   -1  'True
         DataMode        =   1
         HeaderFontName  =   "Tahoma"
         FontName        =   "Tahoma"
         GridLines       =   1
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         ColumnsCount    =   7
         Column(1)       =   "frmHiladosRequeridos.frx":1010
         Column(2)       =   "frmHiladosRequeridos.frx":110C
         Column(3)       =   "frmHiladosRequeridos.frx":11DC
         Column(4)       =   "frmHiladosRequeridos.frx":12B0
         Column(5)       =   "frmHiladosRequeridos.frx":137C
         Column(6)       =   "frmHiladosRequeridos.frx":1450
         Column(7)       =   "frmHiladosRequeridos.frx":1520
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmHiladosRequeridos.frx":15F4
         FormatStyle(2)  =   "frmHiladosRequeridos.frx":16D4
         FormatStyle(3)  =   "frmHiladosRequeridos.frx":17FC
         FormatStyle(4)  =   "frmHiladosRequeridos.frx":18AC
         FormatStyle(5)  =   "frmHiladosRequeridos.frx":1960
         FormatStyle(6)  =   "frmHiladosRequeridos.frx":1A38
         ImageCount      =   3
         ImagePicture(1) =   "frmHiladosRequeridos.frx":1AF0
         ImagePicture(2) =   "frmHiladosRequeridos.frx":1C02
         ImagePicture(3) =   "frmHiladosRequeridos.frx":1F1C
         PrinterProperties=   "frmHiladosRequeridos.frx":2236
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin VB.TextBox TxtProveedor 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3840
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   240
         Width           =   4695
      End
      Begin VB.TextBox TxtOC 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Proveedor:"
         Height          =   255
         Left            =   2880
         TabIndex        =   4
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "O.C.:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   280
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmHiladosRequeridos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public varSer_OrdComp As String
Public varCod_OrdComp As String
Public varCod_Proveedor As String

Sub CARGA_GRID()
On Error GoTo hand
Dim Rs As ADODB.Recordset

    Set Rs = New ADODB.Recordset
    Rs.ActiveConnection = cCONNECT
    Rs.CursorLocation = adUseClient
    Rs.CursorType = adOpenStatic
    Rs.Open "exec SM_HILADOS_REQUERIDOS_ENVIADOS_OC '" & varSer_OrdComp & "','" & varCod_OrdComp & "'"
    
    'If Rs.RecordCount Then
    Set gexDetalle.ADORecordset = Rs
    
    Set Rs = Nothing
Exit Sub
hand:
    Set Rs = Nothing
    ErrorHandler Err, "CARGA_GRID"
End Sub

Private Sub CmdImprimir_Click()
Dim oo As Object
Dim Ruta As String
    If gexDetalle.RowCount = 0 Then Exit Sub
        Ruta = vRuta & "\Hilos-Requeridos.xlt"
        'Ruta = "C:\Archivos de Programa\Gestion de pedidos\Hilos-Requeridos.xlt"

    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False

    oo.Run "reporte", cCONNECT, gexDetalle.ADORecordset
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim Strsql As String
    TxtOC.Text = varSer_OrdComp & "-" & varCod_OrdComp
    Strsql = "select des_proveedor from lg_proveedor where cod_proveedor='" & varCod_Proveedor & "'"
    TxtProveedor.Text = varCod_Proveedor & "-" & DevuelveCampo(Strsql, cCONNECT)
    CARGA_GRID
End Sub
