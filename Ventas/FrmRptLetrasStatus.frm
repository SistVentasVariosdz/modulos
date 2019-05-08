VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmRptLetrasStatus 
   Caption         =   "Reporte de Status Letras"
   ClientHeight    =   7770
   ClientLeft      =   525
   ClientTop       =   1800
   ClientWidth     =   15030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   15030
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   570
      Left            =   6000
      TabIndex        =   6
      Top             =   7200
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   1005
      Custom          =   $"FrmRptLetrasStatus.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1200
      ControlHeigth   =   550
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   1455
      Left            =   20
      TabIndex        =   8
      Top             =   0
      Width           =   15015
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Height          =   615
         Left            =   7800
         TabIndex        =   17
         Top             =   720
         Width           =   6615
         Begin VB.TextBox Txt_DesUsuario 
            Height          =   285
            Left            =   2160
            TabIndex        =   19
            Top             =   240
            Width           =   4335
         End
         Begin VB.TextBox Txt_Cod_Usuario 
            Height          =   285
            Left            =   960
            TabIndex        =   18
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "Vendedor"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   690
         End
      End
      Begin VB.Frame frCliente 
         BackColor       =   &H00C0FFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   7485
         Begin VB.TextBox txtNum_Ruc 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   960
            MaxLength       =   11
            TabIndex        =   14
            Top             =   240
            Width           =   1200
         End
         Begin VB.TextBox txtDes_Anexo 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   2160
            MaxLength       =   30
            TabIndex        =   13
            Top             =   240
            Width           =   4290
         End
         Begin VB.TextBox txtCod_TipAnxo 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   6960
            MaxLength       =   1
            TabIndex        =   12
            Text            =   "C"
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "Nro Ruc:"
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   240
            Width           =   645
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Tipo :"
            Height          =   255
            Left            =   6480
            TabIndex        =   15
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.CheckBox chkFecha 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Desde :"
         Height          =   255
         Left            =   5640
         TabIndex        =   2
         Top             =   255
         Width           =   855
      End
      Begin VB.TextBox TxtCod_Status 
         Height          =   285
         Left            =   810
         MaxLength       =   1
         TabIndex        =   0
         Text            =   "P"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox TxtDes_Status 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   495
         Left            =   13440
         TabIndex        =   5
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
      Begin MSComCtl2.DTPicker txtFec_Ini 
         Height          =   315
         Left            =   6630
         TabIndex        =   3
         Top             =   225
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   88932353
         CurrentDate     =   37543
      End
      Begin MSComCtl2.DTPicker txtFec_Fin 
         Height          =   315
         Left            =   8910
         TabIndex        =   4
         Top             =   225
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   88932353
         CurrentDate     =   37543
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Hasta :"
         Height          =   255
         Left            =   8280
         TabIndex        =   10
         Top             =   255
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Status :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   645
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5700
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1440
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   10054
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
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
      Column(1)       =   "FrmRptLetrasStatus.frx":0090
      Column(2)       =   "FrmRptLetrasStatus.frx":0158
      FormatStylesCount=   8
      FormatStyle(1)  =   "FrmRptLetrasStatus.frx":01FC
      FormatStyle(2)  =   "FrmRptLetrasStatus.frx":0334
      FormatStyle(3)  =   "FrmRptLetrasStatus.frx":03E4
      FormatStyle(4)  =   "FrmRptLetrasStatus.frx":0498
      FormatStyle(5)  =   "FrmRptLetrasStatus.frx":0570
      FormatStyle(6)  =   "FrmRptLetrasStatus.frx":0628
      FormatStyle(7)  =   "FrmRptLetrasStatus.frx":0708
      FormatStyle(8)  =   "FrmRptLetrasStatus.frx":07B4
      ImageCount      =   0
      PrinterProperties=   "FrmRptLetrasStatus.frx":0864
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   240
      Top             =   6120
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmRptLetrasStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCod_Anxo As String
Public codigo As String, Descripcion As String
Dim strSQL As String
Public stip_Trabajador As String
Public scod_trabajador As String
Public rstAux As ADODB.Recordset
Public opcion As String


            

Private Sub chkFecha_Click()
If chkFecha Then
  txtFec_Ini.Enabled = True
  txtFec_Fin.Enabled = True
Else
  txtFec_Ini.Enabled = False
  txtFec_Fin.Enabled = False
End If
End Sub

Private Sub Form_Load()
 txtFec_Ini = Date
 txtFec_Fin = Date
 Call TxtCod_Status_KeyPress(13)
End Sub

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

Dim oGroup As GridEX20.JSGroup
Dim dFecIni As String, dFecFin As String

On Error GoTo errCarga

If chkFecha Then
  dFecIni = "'" & txtFec_Ini & "'"
  dFecFin = "'" & txtFec_Fin & "'"
Else
  dFecIni = "NULL"
  dFecFin = "NULL"
End If

If Len(txtNum_Ruc) = 0 Then
    scodclienteAne = ""
End If


strSQL = "Cn_Ventas_Muestra_Letras_x_Estatus '" & TxtCod_Status & "'," & dFecIni & "," & dFecFin & ",'C','" & scodclienteAne & "','" & Left(Txt_Cod_Usuario, 1) & "','" & Right(Txt_Cod_Usuario, 4) & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

GridEX1.Columns("Cliente").Width = 3045
GridEX1.Columns("Ruc").Width = 1320

GridEX1.Columns("Letra").Width = 795
GridEX1.Columns("Fecha_Vencimiento").Width = 1065
GridEX1.Columns("Moneda").Width = 720
GridEX1.Columns("Fec_EmiDoc").Caption = "Fecha Emision"
GridEX1.Columns("Fec_EmiDoc").Width = 1065
GridEX1.Columns("Fecha_Vencimiento").Caption = "Fecha Vencimiento"
GridEX1.Columns("Saldo_Soles").Width = 1035
GridEX1.Columns("Saldo_Soles").Caption = "Saldo  Soles"
GridEX1.Columns("Saldo_Dolares").Width = 1035
GridEX1.Columns("Saldo_Dolares").Caption = "Saldo Dolares"
GridEX1.Columns("Banco").Width = 2505
GridEX1.Columns("Letra_Banco").Width = 1380

GridEX1.DefaultGroupMode = jgexDGMExpanded

GridEX1.BackColorRowGroup = &H80000005


Exit Sub
Resume
errCarga:
    ErrorHandler err, "Carga Grid"
End Sub

Sub Reporte()
On Error GoTo hand
Dim oo As Object
Dim strSQL As String
Dim sEmpresa As String
    strSQL = "SELECT DES_EMPRESA FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA='" & vemp & "'"
    sEmpresa = DevuelveCampo(strSQL, cCONNECT)

If GridEX1.RowCount = 0 Then Exit Sub

Set oo = CreateObject("excel.application")

oo.Workbooks.Open vRuta & "\RptLetrasxStatus.XLT"
oo.Visible = True
oo.displayalerts = False
oo.Run "reporte", GridEX1.ADORecordset, UCase(TxtDes_Status), sEmpresa


Set oo = Nothing

Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Call CARGA_GRID
End Sub

Private Sub Txt_Cod_Usuario_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Busca_Trabajador
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Txt_DesUsuario_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Busca_Trabajador
        SendKeys "{TAB}"
    End If
End Sub

Private Sub TxtCod_Status_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("flg_status_letra", "Descripcion", "Cn_Status_Letras where ", TxtCod_Status, TxtDes_Status, 1, Me)
End Sub

Private Sub txtDes_Anexo_KeyPress(KeyAscii As Integer)
    scodclienteAne = ""

  If KeyAscii = 13 Then Call Busca_Opcion_Anexo("Num_Ruc", "Des_Anexo", " Cn_AnexosContables Where cod_tipanex ='" & Trim(txtCod_TipAnxo.Text) & "' and ", txtNum_Ruc, txtDes_Anexo, 2, Me)

End Sub

Private Sub TxtDes_Status_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("flg_status_letra", "Descripcion", "Cn_Status_Letras where ", TxtCod_Status, TxtDes_Status, 2, Me)
End Sub


Public Sub Busca_Trabajador()
On Error GoTo Fin
Dim iCol As Long
      
strSQL = "Tg_Sm_Muestra_Operario_Caracteristica '001'"
    With frmBusqGeneralOperario
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("Codigo").Caption = "Codigo"
        .DGridLista.Columns("Codigo").Width = 900
        .DGridLista.Columns("Apellido_Paterno").Caption = "Ape Paterno"
        .DGridLista.Columns("Apellido_Paterno").Width = 1500
        .DGridLista.Columns("Apellido_Materno").Caption = "Ape Materno"
        .DGridLista.Columns("Apellido_Materno").Width = 1500
        .DGridLista.Columns("Nombre_Trabajador").Caption = "Nombres"
        .DGridLista.Columns("Nombre_Trabajador").Width = 1500
        
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If codigo <> "" And rstAux.RecordCount > 0 Then
            Txt_Cod_Usuario = Trim(rstAux!codigo)
            Txt_Cod_Usuario.Tag = Left(Trim(rstAux!codigo), 1)
            Txt_DesUsuario = Trim(rstAux!Apellido_Paterno) + " " + Trim(rstAux!Apellido_Materno) + " " + Trim(rstAux!Nombre_Trabajador)
            Txt_DesUsuario.Tag = Right(Trim(rstAux!codigo), 4)
            stip_Trabajador = Left(rstAux!codigo, 1)
            scod_trabajador = Right(rstAux!codigo, 4)
        End If
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Color (" & opcion & ")"
End Sub



Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  scodclienteAne = ""
  Call Busca_Opcion_Anexo("Num_Ruc", "Des_Anexo", " Cn_AnexosContables Where cod_tipanex ='" & Trim(txtCod_TipAnxo.Text) & "' and ", txtNum_Ruc, txtDes_Anexo, 1, Me)
  SendKeys "{TAB}"
 End If
End Sub
