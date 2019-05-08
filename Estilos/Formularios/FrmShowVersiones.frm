VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmShowVersiones 
   Caption         =   "Seleccionar Versión"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3000
      TabIndex        =   2
      Top             =   3480
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmShowVersiones.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin GridEX20.GridEX GridEX1 
         Height          =   2970
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   5239
         Version         =   "2.0"
         HoldSortSettings=   -1  'True
         DefaultGroupMode=   1
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "FrmShowVersiones.frx":0096
         Column(2)       =   "FrmShowVersiones.frx":015E
         FormatStylesCount=   8
         FormatStyle(1)  =   "FrmShowVersiones.frx":0202
         FormatStyle(2)  =   "FrmShowVersiones.frx":033A
         FormatStyle(3)  =   "FrmShowVersiones.frx":03EA
         FormatStyle(4)  =   "FrmShowVersiones.frx":049E
         FormatStyle(5)  =   "FrmShowVersiones.frx":0576
         FormatStyle(6)  =   "FrmShowVersiones.frx":062E
         FormatStyle(7)  =   "FrmShowVersiones.frx":070E
         FormatStyle(8)  =   "FrmShowVersiones.frx":07BA
         ImageCount      =   0
         PrinterProperties=   "FrmShowVersiones.frx":086A
      End
   End
End
Attribute VB_Name = "FrmShowVersiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public oParent As Object
Public vCod_Cliente As String, vCod_Temcli As String, vCod_EstCli As String
Public vCod_EstPro As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    Call GridEX1_DblClick
Case "CANCELAR"
    Unload Me
End Select
End Sub

Sub CARGA_GRID()
On Error GoTo errCarga_Grid

strSQL = "Es_Muestra_Versiones_Estilo_Propio '" & vCod_EstPro & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

GridEX1.Columns("cod_version").Width = 600
GridEX1.Columns("des_version").Width = 3000
GridEX1.Columns("tip_version").Width = 800
GridEX1.Columns("des_tipo_version").Width = 2000
GridEX1.Columns("fec_creacion").Width = 1350

GridEX1.Columns("cod_version").Caption = "Version"
GridEX1.Columns("des_version").Caption = "Descripcion"
GridEX1.Columns("tip_version").Caption = "Tipo"
GridEX1.Columns("des_tipo_version").Caption = "Des. Tipo"

Exit Sub
errCarga_Grid:
    ErrorHandler Err, "CARGA_GRID"
End Sub

Private Sub GridEX1_DblClick()
On Error GoTo errVersion
If GridEX1.RowCount Then
    strSQL = "Es_Actualiza_Version_Costeo_Estilo '" & vCod_Cliente & "','" & vCod_Temcli & "','" & vCod_EstCli & "','" & vCod_EstPro & "','" & GridEX1.Value(GridEX1.Columns("cod_version").Index) & "'"
    Call ExecuteSQL(cCONNECT, strSQL)
End If
Unload Me
Exit Sub
errVersion:
    ErrorHandler Err, "Version Costeo"
End Sub
