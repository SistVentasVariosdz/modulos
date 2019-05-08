VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmChequesDiferidos 
   Caption         =   "Cheques Diferidos"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   9180
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCambioEstadoaCancelado 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cambio de Estado"
      Height          =   1785
      Left            =   2820
      TabIndex        =   18
      Top             =   5730
      Visible         =   0   'False
      Width           =   3540
      Begin NumBoxProject.NumBox inpFec_PagoReal 
         Height          =   285
         Left            =   1935
         TabIndex        =   19
         Top             =   450
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         TypeVal         =   3
         Mask            =   "99/99/9999"
         Formato         =   "dd/MM/yyyy"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   2
         Text            =   ""
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   0
      End
      Begin FunctionsButtons.FunctButt funCambioEstado 
         Height          =   510
         Left            =   480
         TabIndex        =   21
         Top             =   990
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   900
         Custom          =   "0~0~ACEPTAR~True~True~&Aceptar~0~0~1~~0~False~False~&Aceptar~~1~0~CANCELAR~True~True~&Cancelar~0~0~2~~0~False~False~&Cancelar~"
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fecha Real de Pago :"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   480
         Width           =   1560
      End
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   45
      TabIndex        =   16
      Top             =   -15
      Width           =   9045
      Begin VB.OptionButton opt 
         Caption         =   "Cancelados"
         Height          =   210
         Left            =   4635
         TabIndex        =   11
         Top             =   225
         Width           =   1410
      End
      Begin VB.OptionButton optDiferidos 
         Caption         =   "Pendientes"
         Height          =   210
         Left            =   2985
         TabIndex        =   10
         Top             =   225
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.TextBox txtCod_Origen 
         Height          =   285
         Left            =   810
         MaxLength       =   1
         TabIndex        =   0
         Text            =   "N"
         Top             =   210
         Width           =   375
      End
      Begin VB.TextBox txtDes_Origen 
         Height          =   285
         Left            =   1245
         TabIndex        =   1
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Origen :"
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Top             =   255
         Width           =   495
      End
   End
   Begin VB.Frame fraDatos 
      Height          =   1080
      Left            =   45
      TabIndex        =   13
      Top             =   570
      Width           =   9045
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   435
         Left            =   7500
         TabIndex        =   7
         Top             =   225
         Width           =   1395
      End
      Begin VB.OptionButton optFecha 
         Caption         =   "Fecha de Transacción"
         Height          =   330
         Left            =   3465
         TabIndex        =   15
         Top             =   615
         Width           =   1290
      End
      Begin VB.OptionButton optParteCobranza 
         Caption         =   "Número de Parte de Cobranza"
         Height          =   330
         Left            =   105
         TabIndex        =   14
         Top             =   615
         Width           =   1605
      End
      Begin VB.OptionButton optBanco 
         Caption         =   "Banco"
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Value           =   -1  'True
         Width           =   1515
      End
      Begin VB.TextBox TxtCod_Banco 
         Height          =   285
         Left            =   1725
         TabIndex        =   2
         Top             =   225
         Width           =   375
      End
      Begin VB.TextBox TxtDes_Banco 
         Height          =   285
         Left            =   2175
         TabIndex        =   3
         Top             =   225
         Width           =   5130
      End
      Begin VB.TextBox txtNum_Parte 
         Height          =   285
         Left            =   1725
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin NumBoxProject.NumBox inpFec_Inicio 
         Height          =   285
         Left            =   4785
         TabIndex        =   5
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         TypeVal         =   3
         Mask            =   "99/99/9999"
         Formato         =   "dd/MM/yyyy"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   2
         Text            =   ""
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   0
      End
      Begin NumBoxProject.NumBox inpFec_Final 
         Height          =   285
         Left            =   6090
         TabIndex        =   6
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         TypeVal         =   3
         Mask            =   "99/99/9999"
         Formato         =   "dd/MM/yyyy"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   2
         Text            =   ""
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   0
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5340
      Left            =   30
      TabIndex        =   8
      Top             =   1740
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   9419
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
      Column(1)       =   "frmChequesDiferidos.frx":0000
      Column(2)       =   "frmChequesDiferidos.frx":00C8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmChequesDiferidos.frx":016C
      FormatStyle(2)  =   "frmChequesDiferidos.frx":02A4
      FormatStyle(3)  =   "frmChequesDiferidos.frx":0354
      FormatStyle(4)  =   "frmChequesDiferidos.frx":0408
      FormatStyle(5)  =   "frmChequesDiferidos.frx":04E0
      FormatStyle(6)  =   "frmChequesDiferidos.frx":0598
      FormatStyle(7)  =   "frmChequesDiferidos.frx":0678
      FormatStyle(8)  =   "frmChequesDiferidos.frx":0724
      ImageCount      =   0
      PrinterProperties=   "frmChequesDiferidos.frx":07D4
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3345
      TabIndex        =   9
      Top             =   7200
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmChequesDiferidos.frx":09AC
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   7140
      Top             =   7350
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmChequesDiferidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sOpcion As String
Public codigo As String, Descripcion As String

Private Sub cmdBuscar_Click()
    Buscar
End Sub

Private Sub Form_Load()
    sOpcion = "1"
End Sub

Private Sub funCambioEstado_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            CambioEstadoaCancelado
        Case "CANCELAR"
            Me.fraCambioEstadoaCancelado.Visible = False
    End Select
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "CAMBIOESTADO"
            If GridEX1.RowCount = 0 Then Exit Sub
            If GridEX1.Value(GridEX1.Columns("flg_status_diferido").Index) = "P" Then
                Me.fraCambioEstadoaCancelado.Visible = True
            Else
                CambioEstadoaPlaneado
            End If
            
        Case "SALIR"
            Unload Me
    End Select
End Sub

Private Sub inpFec_Final_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub inpFec_Inicio_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub optBanco_Click()
    sOpcion = "1"
    TxtCod_Banco.SetFocus
End Sub

Private Sub optFecha_Click()
    sOpcion = "3"
    inpFec_Inicio.SetFocus
End Sub

Private Sub optParteCobranza_Click()
    sOpcion = "2"
    txtNum_Parte.SetFocus
End Sub


Sub Buscar()
Dim strSQL As String
Dim sFlg_Pendiente_Cancelado As String
On Error GoTo errores

If optDiferidos Then
    sFlg_Pendiente_Cancelado = "P"
Else
    sFlg_Pendiente_Cancelado = "C"
End If

strSQL = "CN_VENTAS_MUESTRA_CHEQUES_DIFERIDOS '$' , '$' ,'$','$', '$', '$' ,'$'"
strSQL = VBsprintf(strSQL, sOpcion, sFlg_Pendiente_Cancelado, txtCod_Origen.Text, TxtCod_Banco, txtNum_Parte, inpFec_Inicio.Text, inpFec_Final.Text)
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

Exit Sub
Resume
errores:
    errores Err.Number
End Sub

Private Sub inpFec_Emi_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtCod_Banco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Banco", "Nom_Banco", "Tg_Banco where ", TxtCod_Banco, TxtDes_Banco, 1, Me)
End Sub

Private Sub TxtDes_Banco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Banco", "Nom_Banco", "Tg_Banco where ", TxtCod_Banco, TxtDes_Banco, 2, Me)
End Sub

Private Sub txtDes_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Origen", "Des_Origen", " Cn_Origen where ", txtCod_Origen, txtDes_Origen, 2, Me)
End Sub

Private Sub txtCod_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Origen", "Des_Origen", " Cn_Origen where ", txtCod_Origen, txtDes_Origen, 1, Me)
End Sub

Private Sub txtNum_Parte_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub CambioEstadoaCancelado()
Dim strSQL As String
On Error GoTo errores

If inpFec_PagoReal.Text = "" Then
    Aviso "Fecha de Pago es Obligtoria", 2
    Exit Sub
End If

strSQL = "CN_VENTAS_CAMBIA_ESTADO_CHEQUES_DIFERIDOS '$' , '$' ,'$'"
strSQL = VBsprintf(strSQL, GridEX1.Value(GridEX1.Columns("FEC_TRANSACCION").Index), GridEX1.Value(GridEX1.Columns("SECUENCIA").Index), inpFec_PagoReal.Text)

ExecuteCommandSQL cCONNECT, strSQL

Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
Me.fraCambioEstadoaCancelado.Visible = False

Buscar
Exit Sub

Resume
errores:
    errores Err.Number
End Sub


Private Sub CambioEstadoaPlaneado()
Dim strSQL As String
On Error GoTo errores

strSQL = "CN_VENTAS_CAMBIA_ESTADO_CHEQUES_DIFERIDOS '$' , '$' ,''"
strSQL = VBsprintf(strSQL, GridEX1.Value(GridEX1.Columns("FEC_TRANSACCION").Index), GridEX1.Value(GridEX1.Columns("SECUENCIA").Index), inpFec_PagoReal.Text)

ExecuteCommandSQL cCONNECT, strSQL

Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
Me.fraCambioEstadoaCancelado.Visible = False


Buscar
Exit Sub

Resume
errores:
    errores Err.Number
End Sub

