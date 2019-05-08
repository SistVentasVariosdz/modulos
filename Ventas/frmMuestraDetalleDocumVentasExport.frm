VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMuestraDetalleDocumVentasExport 
   Caption         =   "Detalle"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPesos 
      Caption         =   "Registro de Pesos de Cajas"
      Height          =   2145
      Left            =   3405
      TabIndex        =   2
      Top             =   1875
      Visible         =   0   'False
      Width           =   3165
      Begin VB.TextBox txtPeso_Bruto 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1695
         TabIndex        =   3
         Tag             =   "SET"
         Text            =   "0"
         Top             =   255
         Width           =   1200
      End
      Begin VB.TextBox txtPeso_Neto 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1695
         TabIndex        =   5
         Tag             =   "SET"
         Text            =   "0"
         Top             =   705
         Width           =   1200
      End
      Begin FunctionsButtons.FunctButt FunctOKCancel 
         Height          =   510
         Left            =   330
         TabIndex        =   6
         Top             =   1335
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   "7~0~ACEPTAR~True~True~&Aceptar~0~0~4~~0~True~False~&Ok~~8~0~CANCELAR~True~True~&Cancelar~0~0~3~~0~False~True~&Cancel~"
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label lblPeso_Bruto 
         Caption         =   "Peso Bruto"
         Height          =   480
         Left            =   150
         TabIndex        =   7
         Tag             =   "PESO_BRUTO"
         Top             =   255
         Width           =   1500
      End
      Begin VB.Label lblPeso_Neto 
         Caption         =   "Peso Neto"
         Height          =   480
         Left            =   150
         TabIndex        =   4
         Tag             =   "PESO_NETO"
         Top             =   705
         Width           =   1500
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1860
      TabIndex        =   0
      Top             =   4050
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   900
      Custom          =   $"frmMuestraDetalleDocumVentasExport.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   3855
      Left            =   15
      TabIndex        =   1
      Top             =   15
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   6800
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmMuestraDetalleDocumVentasExport.frx":0150
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmMuestraDetalleDocumVentasExport.frx":04A2
      Column(2)       =   "frmMuestraDetalleDocumVentasExport.frx":056A
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmMuestraDetalleDocumVentasExport.frx":060E
      FormatStyle(2)  =   "frmMuestraDetalleDocumVentasExport.frx":0746
      FormatStyle(3)  =   "frmMuestraDetalleDocumVentasExport.frx":07F6
      FormatStyle(4)  =   "frmMuestraDetalleDocumVentasExport.frx":08AA
      FormatStyle(5)  =   "frmMuestraDetalleDocumVentasExport.frx":0982
      FormatStyle(6)  =   "frmMuestraDetalleDocumVentasExport.frx":0A3A
      FormatStyle(7)  =   "frmMuestraDetalleDocumVentasExport.frx":0B1A
      FormatStyle(8)  =   "frmMuestraDetalleDocumVentasExport.frx":0FD2
      ImageCount      =   1
      ImagePicture(1) =   "frmMuestraDetalleDocumVentasExport.frx":141E
      PrinterProperties=   "frmMuestraDetalleDocumVentasExport.frx":1770
   End
End
Attribute VB_Name = "frmMuestraDetalleDocumVentasExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strSql As String, Num_Corre As String


Public Function Buscar() As Boolean

On Error GoTo errores

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSql, cCONNECT)
GridEX1.Columns("T").Width = 240
GridEX1.Columns("Codigo").Width = 1455
GridEX1.Columns("Codigo").Caption = "Codigo"
GridEX1.Columns("Articulo").Width = 4455
GridEX1.Columns("Articulo").Caption = "Articulo"
GridEX1.Columns("Cantidad").Width = 765
GridEX1.Columns("Cantidad").Caption = "Cantidad"
GridEX1.Columns("Uni_Med").Width = 780
GridEX1.Columns("Uni_Med").Caption = "Uni Med"
GridEX1.Columns("Valor_Unitario").Width = 1125
GridEX1.Columns("Valor_Unitario").Caption = "Valor Unitario"
GridEX1.Columns("Valor_Venta").Width = 1005
GridEX1.Columns("Valor_Venta").Caption = "Valor Venta"
GridEX1.Columns("Num_Corre").Visible = False
GridEX1.Columns("Secuencia").Visible = False
GridEX1.Columns("Origen").Visible = False

Exit Function
Resume
errores:
    errores Err.Number
End Function
Private Function ifValidaDoc() As Boolean

Dim strMsg As String

strMsg = DevuelveCampo("Select dbo.ventas_Valida_Documento_Manuales_Det('" & Num_Corre & "')", cCONNECT)
If strMsg <> "" Then
  MsgBox strMsg, vbInformation, "AVISO"
  ifValidaDoc = False
  Exit Function
End If

ifValidaDoc = True

End Function

Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

Dim lvSql As String

On Error GoTo DrpDepurar

Select Case ActionName
Case Is = "ADICIONAR"
  
  If Not ifValidaDoc Then Exit Sub
  
  Load frmAdicionaDetalleDocum
  With frmAdicionaDetalleDocum
    .Caption = "Adicion " + Me.Caption
    .strNum_Corre_Detalle = Num_Corre
    .IntSencuencia = 0
    .StrOption = "I"
    .Show 1
    Buscar
    Call GridEX1.Find(GridEX1.Columns("Secuencia").Index, jgexEqual, .IntSencuencia)
  End With
Case Is = "MODIFICAR"

  If GridEX1.RowCount = 0 Then Exit Sub
  
  If Not ifValidaDoc Then Exit Sub
  
  Load frmAdicionaDetalleDocum
  With frmAdicionaDetalleDocum
    .Caption = "Modificar " + Me.Caption
    .strNum_Corre_Detalle = GridEX1.Value(GridEX1.Columns("Num_Corre").Index)
    .IntSencuencia = GridEX1.Value(GridEX1.Columns("Secuencia").Index)
    .StrOption = "U"
    .strNum_Corre_Detalle = GridEX1.Value(GridEX1.Columns("Num_Corre").Index)
    .txtTip_Item = GridEX1.Value(GridEX1.Columns("T").Index)
    .txtCod_Producto = GridEX1.Value(GridEX1.Columns("Codigo").Index)
    .TxtDescripcion = GridEX1.Value(GridEX1.Columns("Articulo").Index)
    .txtCantidad.Text = GridEX1.Value(GridEX1.Columns("Cantidad").Index)
    .txtUnida_Medida.Text = GridEX1.Value(GridEX1.Columns("Uni_Med").Index)
    .txtImp_Unitario.Text = GridEX1.Value(GridEX1.Columns("Valor_Unitario").Index)
    .txtImp_Total.Text = GridEX1.Value(GridEX1.Columns("Valor_Venta").Index)
    .txtPorc_Commision.Text = GridEX1.Value(GridEX1.Columns("Porcentaje_Commision").Index)
    .Show 1
    Buscar
    Call GridEX1.Find(GridEX1.Columns("Secuencia").Index, jgexEqual, .IntSencuencia)
  End With
Case Is = "ELIMINAR"

  If Not ifValidaDoc Then Exit Sub
  
  If GridEX1.RowCount = 0 Then Exit Sub
  If MsgBox("Esta Seguro de Eliminar este Registro", vbYesNo, "ADVERTENCIA") = vbYes Then
    lvSql = "Ventas_Up_Man_Detalle '" & "D" & "','" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "'," & GridEX1.Value(GridEX1.Columns("Secuencia").Index)
    ExecuteCommandSQL cCONNECT, lvSql
    Buscar
  End If
Case Is = "PESOS"
    Me.txtPeso_Bruto = GridEX1.Value(GridEX1.Columns("PESO_BRUTO").Index)
    Me.txtPeso_Neto = GridEX1.Value(GridEX1.Columns("PESO_NETO").Index)
    Me.fraPesos.Visible = True
    Me.txtPeso_Bruto.SetFocus
Case Is = "SALIR"
  Unload Me
End Select

Exit Sub

DrpDepurar:

errores Err.Number

End Sub



Private Sub FunctOKCancel_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            GrabarPesos
        Case "CANCELAR"
            Me.fraPesos.Visible = False
    End Select
End Sub

Private Sub txtPeso_Bruto_GotFocus()
    SelectionText txtPeso_Bruto
End Sub

Private Sub txtPeso_Bruto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtPeso_Neto_GotFocus()
    SelectionText txtPeso_Neto
End Sub

Private Sub txtPeso_Neto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub


Private Sub GrabarPesos()
On Error GoTo errores
Dim ssql As String

ssql = "CN_VENTAS_PRENDAS_PESOS '$','$','$','$'"
ssql = VBsprintf(ssql, Num_Corre, GridEX1.Value(GridEX1.Columns("secuencia").Index), CDbl(txtPeso_Bruto.Text), CDbl(txtPeso_Neto.Text))

ExecuteCommandSQL cCONNECT, ssql
Unload Me

Exit Sub
errores:
    errores Err.Number
End Sub
