VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmShowCtaCteDet 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4530
   ClientLeft      =   825
   ClientTop       =   1740
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   11385
   Begin VB.Frame frFecha 
      Height          =   1695
      Left            =   3840
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   3735
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   1380
         TabIndex        =   3
         Top             =   240
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   63438849
         CurrentDate     =   37543
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   510
         Left            =   600
         TabIndex        =   5
         Top             =   840
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmShowCtaCteDet.frx":0000
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   300
         Width           =   540
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   6165
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmShowCtaCteDet.frx":0096
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmShowCtaCteDet.frx":03E8
      Column(2)       =   "frmShowCtaCteDet.frx":04B0
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmShowCtaCteDet.frx":0554
      FormatStyle(2)  =   "frmShowCtaCteDet.frx":068C
      FormatStyle(3)  =   "frmShowCtaCteDet.frx":073C
      FormatStyle(4)  =   "frmShowCtaCteDet.frx":07F0
      FormatStyle(5)  =   "frmShowCtaCteDet.frx":08C8
      FormatStyle(6)  =   "frmShowCtaCteDet.frx":0980
      FormatStyle(7)  =   "frmShowCtaCteDet.frx":0A60
      FormatStyle(8)  =   "frmShowCtaCteDet.frx":0F18
      ImageCount      =   1
      ImagePicture(1) =   "frmShowCtaCteDet.frx":1364
      PrinterProperties=   "frmShowCtaCteDet.frx":16B6
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3240
      TabIndex        =   1
      Top             =   3720
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   900
      Custom          =   $"frmShowCtaCteDet.frx":188E
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1500
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmShowCtaCteDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strSQL As String

Public Function Buscar() As Boolean
On Error GoTo errores
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
'GridEX1.FrozenColumns = 3

Exit Function
errores:
    errores err.Number
End Function

Private Sub Form_Load()
Dim sSeguridad  As String

  sSeguridad = get_botones1(Me, vper, vemp, Me.Name)

  FunctButt1.FunctionsUser = sSeguridad
  
  If DevuelveCampo("select Flg_Mod_Fecha from Cn_Ventas_Control_Usuario where cod_usuario = '" & vusu & "'", cCONNECT) <> "*" Then
    FunctButt1.ChangeProperty "VISIBLE", 0, "FALSE"
    FunctButt1.ChangeProperty "VISIBLE", 1, "FALSE"
  End If
End Sub


Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

On Error GoTo errores

Dim lvSql As String

Select Case ActionName
Case Is = "MODFECANCEL"
  If GridEX1.RowCount = 0 Then Exit Sub
  FunctButt1.Visible = False
  GridEX1.Enabled = False
  frFecha.Visible = True
  dtpFecha = GridEX1.Value(GridEX1.Columns("Fec_Cobranza").Index)
  dtpFecha.SetFocus
Case "REVIERTE"
  If GridEX1.RowCount = 0 Then Exit Sub
  If MsgBox("ESTA SEGURO DE REVERTIR EL VOUCHER", vbYesNo, "AVISO") = vbYes Then
    lvSql = "Cn_Ventas_Revierte_Voucher_Letras '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "','" & GridEX1.Value(GridEX1.Columns("Secuencia").Index) & "'"
    ExecuteCommandSQL cCONNECT, lvSql
    Buscar
  End If
Case Is = "SALIR"
  Unload Me
End Select

Exit Sub

errores:
    errores err.Number
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

On Error GoTo errores

Dim lvSql As String

  Select Case ActionName
  Case Is = "ACEPTAR"
   If GridEX1.Value(GridEX1.Columns("Fec_Cobranza").Index) <> dtpFecha Then
    If MsgBox("ESTA SEGURO DE CAMBIAR LA FECHA DE CANCELACION", vbYesNo, "AVISO") = vbYes Then
      lvSql = "Cn_Ventas_Cambio_Fecha_Cancelacion '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "','" & GridEX1.Value(GridEX1.Columns("Secuencia").Index) & "','" & dtpFecha & "'"
      ExecuteCommandSQL cCONNECT, lvSql
      Buscar
      Call FunctButt2_ActionClick(0, 0, "CANCELAR")
    End If
   End If
  Case Is = "CANCELAR"
    FunctButt1.Visible = True
    GridEX1.Enabled = True
    frFecha.Visible = False
  End Select
  
Exit Sub

errores:
    errores err.Number
End Sub
