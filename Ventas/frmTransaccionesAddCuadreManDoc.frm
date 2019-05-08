VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmTransaccionesAddCuadreManDoc 
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   1335
   ClientTop       =   1785
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   7710
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   7455
      Begin VB.TextBox txtDes_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Width           =   5265
      End
      Begin VB.TextBox txtNumeroPendiente 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Text            =   "0"
         Top             =   840
         Width           =   1245
      End
      Begin VB.TextBox txtCod_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6120
         MaxLength       =   4
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "C"
         Top             =   240
         Width           =   360
      End
      Begin VB.TextBox txtNum_Ruc 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4935
         MaxLength       =   11
         TabIndex        =   1
         Top             =   600
         Width           =   1545
      End
      Begin VB.TextBox txtDes_TipAnex 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   6
         Text            =   "C"
         Top             =   240
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label14 
         Caption         =   "Numero :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   870
         Width           =   675
      End
      Begin VB.Label Label28 
         Caption         =   "R.U.C."
         Height          =   255
         Left            =   4320
         TabIndex        =   9
         Top             =   615
         Width           =   495
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   285
         Width           =   570
      End
   End
   Begin GridEX20.GridEX gexGrid1 
      Height          =   4875
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1455
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   8599
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      BorderStyle     =   3
      GroupByBoxVisible=   0   'False
      RowHeaders      =   -1  'True
      BackColorBkg    =   -2147483628
      ColumnHeaderHeight=   285
      IntProp1        =   0
      ColumnsCount    =   3
      Column(1)       =   "frmTransaccionesAddCuadreManDoc.frx":0000
      Column(2)       =   "frmTransaccionesAddCuadreManDoc.frx":00F4
      Column(3)       =   "frmTransaccionesAddCuadreManDoc.frx":01E0
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmTransaccionesAddCuadreManDoc.frx":02AC
      FormatStyle(2)  =   "frmTransaccionesAddCuadreManDoc.frx":03E4
      FormatStyle(3)  =   "frmTransaccionesAddCuadreManDoc.frx":0494
      FormatStyle(4)  =   "frmTransaccionesAddCuadreManDoc.frx":0548
      FormatStyle(5)  =   "frmTransaccionesAddCuadreManDoc.frx":0620
      FormatStyle(6)  =   "frmTransaccionesAddCuadreManDoc.frx":06D8
      ImageCount      =   0
      PrinterProperties=   "frmTransaccionesAddCuadreManDoc.frx":07B8
   End
   Begin FunctionsButtons.FunctButt fncBuscar 
      Height          =   390
      Left            =   5115
      TabIndex        =   3
      Top             =   6480
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   661
      Custom          =   $"frmTransaccionesAddCuadreManDoc.frx":0990
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   350
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmTransaccionesAddCuadreManDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lvSql As String
Public codigo As String, Descripcion As String, strCod_Anxo As String
Public strStore_Carga As String, strCod_Moneda  As String, dFecha  As String

Public Sub Carga_Grid()

On Error GoTo hand

lvSql = strStore_Carga & "'" & txtCod_TipAne & "','" & strCod_Anxo & "','" & strCod_Moneda & "','" & dFecha & "'"

Set gexGrid1.ADORecordset = CargarRecordSetDesconectado(lvSql, cCONNECT)

With gexGrid1
  .Columns(2).Width = 1305
  .Columns(3).Width = 945
  .Columns(4).Width = 720
  .Columns(5).Width = 1065
  .Columns(6).Width = 1500
  .Columns(7).Width = 1365
  .Columns("Correlativo").Visible = False
'  .Columns("Monto_Origen1").Visible = False
'  .Columns("Cod_Cobranza").Visible = False
'  .Columns("Debe_Haber").Visible = False
  .Columns("Monto_Aceptado").Format = "###,###.00"
  .Columns("Monto_Origen").Format = "###,###.00"
'  .Columns("Observacion").Visible = False
End With

Exit Sub
Resume

hand:

errores Err.Number
End Sub

Private Sub fncBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

Select Case ActionName

Case Is = "ACEPTAR"
  With frmTransaccionesAddCuadreMan
    .txtCod_TipDoc.Text = Left(gexGrid1.Value(gexGrid1.Columns("Numero").Index), 2)
    .txtSer_Docum.Text = Mid(gexGrid1.Value(gexGrid1.Columns("Numero").Index), 4, 3)
    .txtNum_Docum.Text = Mid(gexGrid1.Value(gexGrid1.Columns("Numero").Index), 7, 8)
    .txtImporte.Text = gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index)
    .TxtTipo_Cambio.Text = gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index)
    .txtOtro_Tipo_Cambio.Text = gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio_Otro").Index)
    .txtOtro_Tipo_Cambio.Text = gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio_Otro").Index)
    .txtCod_Moneda.Text = gexGrid1.Value(gexGrid1.Columns("Moneda").Index)
    .strNum_Corre = gexGrid1.Value(gexGrid1.Columns("Correlativo").Index)
    .Calcula_Importe_Converido
  End With
End Select

Unload Me

End Sub

Private Sub txtCod_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAne, txtDes_TipAnex, 1, Me)
  End If
End Sub

Private Sub txtDes_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 2, Me)
    Carga_Grid
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 1, Me)
    Carga_Grid
  End If
End Sub

Private Sub txtNumeroPendiente_Change()
  Call gexGrid1.Find(gexGrid1.Columns("Numero").Index, jgexContains, txtNumeroPendiente)
End Sub
