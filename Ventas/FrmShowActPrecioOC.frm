VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmShowActPrecioOC 
   Caption         =   "Mostrar O/C Tejeduria por Actualizar"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   11985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Opciones Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11950
      Begin VB.TextBox txtSerOrdComp 
         Height          =   285
         Left            =   7080
         MaxLength       =   3
         TabIndex        =   12
         Top             =   680
         Width           =   570
      End
      Begin VB.TextBox txtCodOrdComp 
         Height          =   285
         Left            =   7665
         MaxLength       =   6
         TabIndex        =   11
         Top             =   680
         Width           =   1335
      End
      Begin VB.TextBox TxtNom_Cliente 
         Height          =   285
         Left            =   2625
         TabIndex        =   10
         Top             =   680
         Width           =   2775
      End
      Begin VB.TextBox TxtAbr_Cliente 
         Height          =   285
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   9
         Top             =   680
         Width           =   675
      End
      Begin VB.OptionButton OptCliente 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   495
         Left            =   10440
         TabIndex        =   7
         Top             =   360
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
      Begin VB.TextBox TxtCod_OrdTra 
         Height          =   285
         Left            =   6600
         TabIndex        =   6
         Top             =   330
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton Optot 
         Caption         =   "Por OT"
         Height          =   255
         Left            =   5520
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OptPrecio 
         Caption         =   "Ordenes sin Precio"
         Height          =   255
         Left            =   3240
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton OptTodas 
         Caption         =   "Todas las Ordenes Vigentes"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Orden Compra"
         Height          =   195
         Left            =   5760
         TabIndex        =   13
         Top             =   720
         Width           =   1020
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   540
      Left            =   3840
      TabIndex        =   1
      Top             =   6960
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   953
      Custom          =   $"FrmShowActPrecioOC.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1300
      ControlHeigth   =   520
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5700
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   10054
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
      Column(1)       =   "FrmShowActPrecioOC.frx":0124
      Column(2)       =   "FrmShowActPrecioOC.frx":01EC
      FormatStylesCount=   8
      FormatStyle(1)  =   "FrmShowActPrecioOC.frx":0290
      FormatStyle(2)  =   "FrmShowActPrecioOC.frx":03C8
      FormatStyle(3)  =   "FrmShowActPrecioOC.frx":0478
      FormatStyle(4)  =   "FrmShowActPrecioOC.frx":052C
      FormatStyle(5)  =   "FrmShowActPrecioOC.frx":0604
      FormatStyle(6)  =   "FrmShowActPrecioOC.frx":06BC
      FormatStyle(7)  =   "FrmShowActPrecioOC.frx":079C
      FormatStyle(8)  =   "FrmShowActPrecioOC.frx":0848
      ImageCount      =   0
      PrinterProperties=   "FrmShowActPrecioOC.frx":08F8
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   1080
      Top             =   6960
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmShowActPrecioOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public codigo As String, Descripcion As String, TipoAdd As String
Dim bCod_Cliente As String

Private Sub Form_Load()
  FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name) & "/SALIR"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACTUALIZAR"
    If gridex1.RowCount = 0 Then Exit Sub
    Load FrmActPrecioOC
    FrmActPrecioOC.vCod_Cliente = gridex1.Value(gridex1.Columns("cod_cliente_tex").Index)
    FrmActPrecioOC.vSer_OrdComp = gridex1.Value(gridex1.Columns("ser_ordcomp").Index)
    FrmActPrecioOC.vCod_OrdComp = gridex1.Value(gridex1.Columns("cod_ordcomp").Index)
    FrmActPrecioOC.vSec_OrdComp = gridex1.Value(gridex1.Columns("sec_ordcomp").Index)
    FrmActPrecioOC.lblCLIENTE = gridex1.Value(gridex1.Columns("nom_cliente").Index)
    FrmActPrecioOC.LblOC = gridex1.Value(gridex1.Columns("ord_compra").Index)
    FrmActPrecioOC.TxtPrecio = CDbl(gridex1.Value(gridex1.Columns("pre_unitario").Index))
    FrmActPrecioOC.Show vbModal
    Set FrmActPrecioOC = Nothing
    Call CARGA_GRID
Case "OTROS"
    If gridex1.RowCount = 0 Then Exit Sub
    Load FrmShowActPrecioOCOtrosClientes
    FrmShowActPrecioOCOtrosClientes.sCod_Cliente = gridex1.Value(gridex1.Columns("cod_cliente_tex").Index)
    FrmShowActPrecioOCOtrosClientes.sSer_OrdComp = gridex1.Value(gridex1.Columns("ser_ordcomp").Index)
    FrmShowActPrecioOCOtrosClientes.sCod_Ordcomp = gridex1.Value(gridex1.Columns("cod_ordcomp").Index)
    FrmShowActPrecioOCOtrosClientes.sSec_Ordcomp = gridex1.Value(gridex1.Columns("sec_ordcomp").Index)
    FrmShowActPrecioOCOtrosClientes.LblOrden = gridex1.Value(gridex1.Columns("ser_ordcomp").Index) & "-" & gridex1.Value(gridex1.Columns("cod_ordcomp").Index)
    FrmShowActPrecioOCOtrosClientes.LblSecuencia = gridex1.Value(gridex1.Columns("sec_ordcomp").Index)
    FrmShowActPrecioOCOtrosClientes.lblCLIENTE = DevuelveCampo("select abr_cliente  + '-' + nom_cliente from tx_cliente where cod_cliente_tex='" & gridex1.Value(gridex1.Columns("cod_cliente_tex").Index) & "'", cCONNECT)
    FrmShowActPrecioOCOtrosClientes.CARGA_GRID
    FrmShowActPrecioOCOtrosClientes.Show vbModal
    Set FrmShowActPrecioOCOtrosClientes = Nothing
Case "SALIR"
    Unload Me
End Select
End Sub


Sub CARGA_GRID()
Dim vopcion As String

If OptTodas Then
    vopcion = "1"
ElseIf OptPrecio Then
    vopcion = "2"
ElseIf Me.Optot Then
    vopcion = "3"
Else
    vopcion = "4"
    If Trim(txtSerOrdComp.Text) = "" Then
        MsgBox "No ha ingresado la Serie de la O.C.", vbCritical, "Busqueda"
        txtSerOrdComp.SetFocus
        Exit Sub
    End If
    If Trim(txtCodOrdComp.Text) = "" Then
        MsgBox "No ha ingresado el Codigo de la O.C.", vbCritical, "Busqueda"
        txtCodOrdComp.SetFocus
        Exit Sub
    End If
End If

strSQL = "SELECT cod_cliente_Tex FROM Tx_cliente WHERE abr_cliente = '" & Trim(Me.txtAbr_Cliente.Text) & "'"
bCod_Cliente = DevuelveCampo(strSQL, cCONNECT)

strSQL = "Ventas_ocs_sin_precio_tejeduria '" & vopcion & "','" & TxtCod_OrdTra & "','" & bCod_Cliente & "','" & txtSerOrdComp.Text & "','" & txtCodOrdComp.Text & "'"
Set gridex1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

gridex1.Columns("cod_moneda").Caption = "Moneda"
gridex1.Columns("Pre_Unitario").Width = 1000
gridex1.Columns("cod_moneda").Width = 680
gridex1.Columns("nom_cliente").Width = 2000
gridex1.Columns("cod_tela").Width = 950
gridex1.Columns("nombre_tela").Width = 2000
gridex1.Columns("ot").Width = 650
gridex1.Columns("Ord_Compra").Width = 1100
gridex1.Columns("cod_cliente_tex").Width = 0
gridex1.Columns("ser_ordcomp").Width = 0
gridex1.Columns("cod_ordcomp").Width = 0
gridex1.Columns("sec_ordcomp").Width = 0

gridex1.FrozenColumns = 7

End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Call CARGA_GRID
End Sub

Private Sub Optot_Click()
TxtCod_OrdTra.Text = ""
TxtCod_OrdTra.Visible = True
End Sub

Private Sub OptPrecio_Click()
TxtCod_OrdTra.Visible = False
End Sub

Private Sub OptTodas_Click()
TxtCod_OrdTra.Visible = False
End Sub


'Private Sub TxtAbr_Cliente_GotFocus()
'OptCliente.SetFocus
'End Sub

Private Sub TxtAbr_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtAbr_Cliente.Text) = "" Then
            BUSCA_CLIENTE (3)
        Else
            BUSCA_CLIENTE (1)
        End If
    End If
End Sub

Private Sub TxtCod_OrdTra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtCod_OrdTra.Text = Right("00000" & TxtCod_OrdTra.Text, 5)
    FunctButt2.SetFocus
End If
End Sub

Private Sub txtCodOrdComp_KeyPress(KeyAscii As Integer)
Dim varCliente As String
    If KeyAscii = 13 Then
        If Trim(txtSerOrdComp.Text) = "" Then
            MsgBox "Ingrese Serie de la O.C.", vbCritical, "Busqueda"
            txtSerOrdComp.SetFocus
        End If

        txtCodOrdComp.Text = Right("000000" & Trim(txtCodOrdComp.Text), 6)
        
        strSQL = "SELECT cod_cliente_tex FROM Tx_cliente WHERE abr_cliente = '" & Trim(Me.txtAbr_Cliente.Text) & "'"
        varCliente = DevuelveCampo(strSQL, cCONNECT)
        strSQL = "select count(*) from Tx_OrdComp where cod_cliente_tex='" & varCliente & "' and ser_ordcomp='" & Right("000" & Trim(txtSerOrdComp.Text), 3) & "' and cod_ordcomp='" & txtCodOrdComp & "'"
        If DevuelveCampo(strSQL, cCONNECT) = 0 Then
            MsgBox "La O.C. no existe", vbCritical, "Busqueda"
            txtCodOrdComp.SetFocus
            SelectionText txtCodOrdComp
        Else
            FunctButt2.SetFocus
        End If
'        FunctBuscar.SetFocus
    End If
End Sub

Private Sub TxtNom_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtNom_Cliente.Text) = "" Then
            BUSCA_CLIENTE (3)
        Else
            BUSCA_CLIENTE (2)
        End If
    End If
End Sub

Public Sub BUSCA_CLIENTE(Tipo As Integer)
    Select Case Tipo
        Case 1:
                    strSQL = "SELECT nom_cliente FROM Tx_cliente WHERE abr_cliente = '" & Trim(Me.txtAbr_Cliente.Text) & "'"
                    Me.txtNom_Cliente.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
                    txtSerOrdComp.SetFocus
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim RS As Object
                    Set RS = CreateObject("ADODB.Recordset")
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.SQuery = "SELECT abr_cliente AS 'Código', nom_cliente AS 'Descripción' FROM tx_cliente where nom_cliente like '%" & Trim(txtNom_Cliente.Text) & "%' order by abr_cliente"
                    Else
                        oTipo.SQuery = "SELECT abr_cliente AS 'Código', nom_cliente AS 'Descripción' FROM tx_cliente order by abr_cliente"
                    End If
                    
                    oTipo.CARGAR_DATOS
                    oTipo.Show 1
                    If codigo <> "" Then
                         Me.txtAbr_Cliente.Text = Trim(codigo)
                         Me.txtNom_Cliente.Text = Trim(Descripcion)
                         txtSerOrdComp.SetFocus
                         codigo = "": Descripcion = ""
                    End If
                    Set oTipo = Nothing
                    Set RS = Nothing
    End Select
End Sub

Private Sub txtSerOrdComp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSerOrdComp.Text = Right("000" & Trim(txtSerOrdComp.Text), 3)
        txtCodOrdComp.SetFocus
    End If
End Sub
