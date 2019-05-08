VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmRptAnexosContables 
   Caption         =   "Reporte de Anexos"
   ClientHeight    =   6720
   ClientLeft      =   135
   ClientTop       =   990
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   11895
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11895
      Begin VB.TextBox txtDes_TipAnex 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   1
         Top             =   240
         Width           =   1545
      End
      Begin VB.TextBox txtCod_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   960
         MaxLength       =   4
         TabIndex        =   0
         Text            =   "C"
         Top             =   240
         Width           =   360
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   495
         Left            =   10440
         TabIndex        =   2
         Top             =   150
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo :"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   255
         Width           =   495
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   570
      Left            =   4680
      TabIndex        =   3
      Top             =   6120
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   1005
      Custom          =   $"FrmRptAnexosContables.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1200
      ControlHeigth   =   550
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5220
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   840
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   9208
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "FrmRptAnexosContables.frx":0090
      Column(2)       =   "FrmRptAnexosContables.frx":0158
      FormatStylesCount=   8
      FormatStyle(1)  =   "FrmRptAnexosContables.frx":01FC
      FormatStyle(2)  =   "FrmRptAnexosContables.frx":0334
      FormatStyle(3)  =   "FrmRptAnexosContables.frx":03E4
      FormatStyle(4)  =   "FrmRptAnexosContables.frx":0498
      FormatStyle(5)  =   "FrmRptAnexosContables.frx":0570
      FormatStyle(6)  =   "FrmRptAnexosContables.frx":0628
      FormatStyle(7)  =   "FrmRptAnexosContables.frx":0708
      FormatStyle(8)  =   "FrmRptAnexosContables.frx":07B4
      ImageCount      =   0
      PrinterProperties=   "FrmRptAnexosContables.frx":0864
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   240
      Top             =   6120
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmRptAnexosContables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public codigo As String, Descripcion As String

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
    Call Reporte
Case "SALIR"
    Unload Me
End Select
End Sub


Sub CARGA_GRID()
On Error GoTo errCarga

strSQL = "Cn_Muestra_Anexos_Contables '" & txtCod_TipAne & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

GridEX1.Columns("Nombre").Width = 4260
GridEX1.Columns("Ruc").Width = 1320
GridEX1.Columns("Direccion").Width = 4545
GridEX1.Columns("Telefono").Width = 1305

Exit Sub
errCarga:
    ErrorHandler Err, "Carga Grid"
End Sub

Sub Reporte()
On Error GoTo hand
Dim oo As Object

If GridEX1.RowCount = 0 Then Exit Sub

Call txtCod_TipAne_KeyPress(13)

Screen.MousePointer = vbHourglass

Set oo = CreateObject("excel.application")
oo.Workbooks.Open vRuta & "\ReporteAnexosContables.xlt"

oo.DisplayAlerts = False
oo.Run "reporte", GridEX1.ADORecordset, UCase("Reporte de " & txtDes_TipAnex)

oo.Visible = True

Set oo = Nothing

Screen.MousePointer = vbDefault

Exit Sub
hand:
    Screen.MousePointer = vbDefault
    ErrorHandler Err, "GeneraReportes"
    Set oo = Nothing
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Call CARGA_GRID
End Sub

Private Sub txtCod_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAne, txtDes_TipAnex, 1, Me)
End Sub

Private Sub txtDes_TipAnex_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAne, txtDes_TipAnex, 2, Me)
End Sub
