VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmMuestraGuias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes de Compra Relacionadas a la Guia"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   4230
      TabIndex        =   2
      Top             =   3390
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Height          =   3165
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   5415
      Begin GridEX20.GridEX gexList 
         Height          =   2835
         Left            =   135
         TabIndex        =   1
         Top             =   180
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   5001
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "FrmMuestraGuias.frx":0000
         Column(2)       =   "FrmMuestraGuias.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "FrmMuestraGuias.frx":016C
         FormatStyle(2)  =   "FrmMuestraGuias.frx":02A4
         FormatStyle(3)  =   "FrmMuestraGuias.frx":0354
         FormatStyle(4)  =   "FrmMuestraGuias.frx":0408
         FormatStyle(5)  =   "FrmMuestraGuias.frx":04E0
         FormatStyle(6)  =   "FrmMuestraGuias.frx":0598
         ImageCount      =   0
         PrinterProperties=   "FrmMuestraGuias.frx":0678
      End
   End
End
Attribute VB_Name = "FrmMuestraGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vCod_Almacen As String
Public vCod_Proveedor As String
Public vNum_Guia As String

Private Sub Command1_Click()
    Unload Me
End Sub

Sub CARGA_GRID()
Dim StrSql As String
On Error GoTo hand

StrSql = "EXEC UP_SEL_BUSCA_GUIAS '" & vCod_Almacen & "','" & vCod_Proveedor & "','" & vNum_Guia & "'"

Set gexList.ADORecordset = CargarRecordSetDesconectado(StrSql, cConnect)

Exit Sub
hand:
    ErrorHandler Err, "CARGA_GRID"
End Sub
