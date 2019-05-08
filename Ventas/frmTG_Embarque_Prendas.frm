VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmTG_Embarque_Prendas 
   Caption         =   "Detalle Embarque Prendas"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Lote P.O."
      Height          =   1695
      Left            =   60
      TabIndex        =   35
      Top             =   2280
      Width           =   6840
      Begin VB.TextBox txtarancelaria3 
         BackColor       =   &H8000000E&
         Height          =   285
         Left            =   5670
         TabIndex        =   44
         Top             =   975
         Width           =   915
      End
      Begin VB.TextBox txtarancelaria2 
         BackColor       =   &H8000000E&
         Height          =   285
         Left            =   3270
         TabIndex        =   42
         Top             =   975
         Width           =   915
      End
      Begin VB.TextBox txtarancelaria1 
         BackColor       =   &H8000000E&
         Height          =   285
         Left            =   720
         TabIndex        =   40
         Top             =   975
         Width           =   1635
      End
      Begin VB.TextBox txtCod_LotPurOrd 
         Height          =   285
         Left            =   3360
         TabIndex        =   4
         Top             =   540
         Width           =   480
      End
      Begin VB.TextBox txtCod_EstCli 
         Height          =   285
         Left            =   4830
         TabIndex        =   5
         Top             =   540
         Width           =   1755
      End
      Begin VB.TextBox txtCod_PurOrd 
         BackColor       =   &H8000000E&
         Height          =   285
         Left            =   780
         TabIndex        =   3
         Top             =   555
         Width           =   1875
      End
      Begin VB.TextBox txtAbr_Cliente 
         Height          =   300
         Left            =   780
         TabIndex        =   1
         Tag             =   "SET"
         Top             =   210
         Width           =   555
      End
      Begin VB.TextBox txtNom_Cliente 
         Height          =   300
         Left            =   1410
         TabIndex        =   2
         Tag             =   "SET"
         Top             =   210
         Width           =   5190
      End
      Begin VB.Label Label6 
         Caption         =   "Num. Categoria Internacional"
         Height          =   555
         Left            =   4440
         TabIndex        =   45
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Sec. Partida Arancelaria"
         Height          =   555
         Left            =   2520
         TabIndex        =   43
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Num. Partida Arancelaria"
         Height          =   555
         Left            =   120
         TabIndex        =   41
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Estilo"
         Height          =   195
         Left            =   4125
         TabIndex        =   39
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label3 
         Caption         =   "Lote"
         Height          =   195
         Left            =   2730
         TabIndex        =   38
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "P.O."
         Height          =   195
         Left            =   150
         TabIndex        =   37
         Top             =   600
         Width           =   315
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente"
         Height          =   330
         Left            =   135
         TabIndex        =   36
         Tag             =   "COD_TIPANEX"
         Top             =   255
         Width           =   615
      End
   End
   Begin VB.Frame fraNP 
      Caption         =   "NP"
      Height          =   2265
      Left            =   60
      TabIndex        =   31
      Top             =   15
      Width           =   6825
      Begin VB.Frame Frame4 
         Caption         =   "Buscar por"
         Height          =   495
         Left            =   120
         TabIndex        =   58
         Top             =   480
         Width           =   4455
         Begin VB.OptionButton optcliente 
            Caption         =   "Cliente/Tempoarda"
            Height          =   195
            Left            =   2400
            TabIndex        =   60
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optnp 
            Caption         =   "NP"
            Height          =   195
            Left            =   1200
            TabIndex        =   59
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame FrmCliente 
         Height          =   1215
         Left            =   120
         TabIndex        =   49
         Top             =   960
         Visible         =   0   'False
         Width           =   6615
         Begin VB.TextBox txtestilo 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1035
            TabIndex        =   61
            Top             =   840
            Width           =   1755
         End
         Begin VB.TextBox txtCod_TemCli 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1035
            MaxLength       =   3
            TabIndex        =   56
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton cmdBusca_Temporada 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   55
            Top             =   480
            Width           =   300
         End
         Begin VB.TextBox txtNom_TemCli 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2100
            TabIndex        =   54
            Top             =   480
            Width           =   2325
         End
         Begin VB.TextBox txtcliente 
            Height          =   285
            Left            =   1035
            TabIndex        =   52
            Top             =   135
            Width           =   735
         End
         Begin VB.TextBox txtDes_Cliente 
            Height          =   285
            Left            =   2100
            TabIndex        =   51
            Top             =   135
            Width           =   2325
         End
         Begin VB.CommandButton cmdBusCliente 
            Caption         =   "..."
            Height          =   285
            Left            =   1800
            TabIndex        =   50
            Tag             =   "..."
            Top             =   135
            Width           =   300
         End
         Begin VB.Label Label10 
            Caption         =   "Estilo"
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   900
            Width           =   645
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Temporada"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   525
            Width           =   810
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.Frame FrmNP 
         Height          =   855
         Left            =   120
         TabIndex        =   46
         Top             =   960
         Width           =   6615
         Begin VB.TextBox txtDes_OrdPro 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1620
            TabIndex        =   47
            Top             =   360
            Width           =   4815
         End
         Begin VB.TextBox txtCod_OrdPro 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   630
            MaxLength       =   5
            TabIndex        =   0
            Top             =   360
            Width           =   915
         End
         Begin VB.Label lblOrdPro 
            AutoSize        =   -1  'True
            Caption         =   "N/P:"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   405
            Width           =   345
         End
      End
      Begin VB.TextBox txtNom_Fabrica 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1335
         TabIndex        =   33
         Top             =   165
         Width           =   2490
      End
      Begin VB.TextBox txtCod_Fabrica 
         Height          =   285
         Left            =   840
         TabIndex        =   32
         Top             =   165
         Width           =   480
      End
      Begin VB.Label lblFabrica 
         Caption         =   "Fábrica"
         Height          =   195
         Left            =   165
         TabIndex        =   34
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.Frame fraProgramado 
      Caption         =   "Detalle Programado"
      Height          =   2970
      Left            =   60
      TabIndex        =   24
      Top             =   4035
      Width           =   3030
      Begin VB.TextBox txtCubicaje_Prog 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1680
         TabIndex        =   11
         Tag             =   "SET"
         Text            =   "0"
         Top             =   2505
         Width           =   1200
      End
      Begin VB.TextBox txtPeso_Neto_Prog 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1680
         TabIndex        =   10
         Tag             =   "SET"
         Text            =   "0"
         Top             =   2055
         Width           =   1200
      End
      Begin VB.TextBox txtPeso_Bruto_Prog 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1680
         TabIndex        =   9
         Tag             =   "SET"
         Text            =   "0"
         Top             =   1605
         Width           =   1200
      End
      Begin VB.TextBox txtNum_Cajas_Prog 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1680
         TabIndex        =   8
         Tag             =   "SET"
         Text            =   "0"
         Top             =   1155
         Width           =   1200
      End
      Begin VB.TextBox txtPre_Unitario 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1680
         TabIndex        =   7
         Tag             =   "SET"
         Text            =   "0"
         Top             =   705
         Width           =   1200
      End
      Begin VB.TextBox txtNum_Prendas_Prog 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1695
         TabIndex        =   6
         Tag             =   "SET"
         Text            =   "0"
         Top             =   270
         Width           =   1200
      End
      Begin VB.Label lblCubicaje_Prog 
         Caption         =   "Cubicaje Prog"
         Height          =   405
         Left            =   135
         TabIndex        =   30
         Tag             =   "CUBICAJE_PROG"
         Top             =   2505
         Width           =   1500
      End
      Begin VB.Label lblPeso_Neto_Prog 
         Caption         =   "Peso Neto Prog"
         Height          =   480
         Left            =   135
         TabIndex        =   29
         Tag             =   "PESO_NETO_PROG"
         Top             =   2055
         Width           =   1500
      End
      Begin VB.Label lblPeso_Bruto_Prog 
         Caption         =   "Peso Bruto"
         Height          =   480
         Left            =   135
         TabIndex        =   28
         Tag             =   "PESO_BRUTO_PROG"
         Top             =   1605
         Width           =   1500
      End
      Begin VB.Label lblNum_Cajas_Prog 
         Caption         =   "Num Cajas "
         Height          =   480
         Left            =   135
         TabIndex        =   27
         Tag             =   "NUM_CAJAS_PROG"
         Top             =   1155
         Width           =   1500
      End
      Begin VB.Label lblPre_Unitario 
         Caption         =   "Precio Unitario"
         Height          =   480
         Left            =   135
         TabIndex        =   26
         Tag             =   "PRE_UNITARIO"
         Top             =   705
         Width           =   1500
      End
      Begin VB.Label lblNum_Prendas_Prog 
         Caption         =   "Prendas "
         Height          =   480
         Left            =   135
         TabIndex        =   25
         Tag             =   "NUM_PRENDAS_PROG"
         Top             =   255
         Width           =   1500
      End
   End
   Begin VB.Frame fraReal 
      Caption         =   "Detalle Real"
      Enabled         =   0   'False
      Height          =   2970
      Left            =   3765
      TabIndex        =   13
      Top             =   4035
      Width           =   3105
      Begin VB.TextBox txtCubicaje 
         Height          =   300
         Left            =   1680
         TabIndex        =   23
         Tag             =   "SET"
         Top             =   2520
         Width           =   1200
      End
      Begin VB.TextBox txtPeso_Neto 
         Height          =   300
         Left            =   1680
         TabIndex        =   21
         Tag             =   "SET"
         Top             =   2070
         Width           =   1200
      End
      Begin VB.TextBox txtPeso_Bruto 
         Height          =   300
         Left            =   1680
         TabIndex        =   19
         Tag             =   "SET"
         Top             =   1620
         Width           =   1200
      End
      Begin VB.TextBox txtNum_Cajas 
         Height          =   300
         Left            =   1680
         TabIndex        =   17
         Tag             =   "SET"
         Top             =   1170
         Width           =   1200
      End
      Begin VB.TextBox txtNum_Prendas 
         Height          =   300
         Left            =   1680
         TabIndex        =   15
         Tag             =   "SET"
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label lblCubicaje 
         Caption         =   "Cubicaje"
         Height          =   390
         Left            =   135
         TabIndex        =   22
         Tag             =   "CUBICAJE"
         Top             =   2535
         Width           =   1500
      End
      Begin VB.Label lblPeso_Neto 
         Caption         =   "Peso Neto"
         Height          =   480
         Left            =   135
         TabIndex        =   20
         Tag             =   "PESO_NETO"
         Top             =   2070
         Width           =   1500
      End
      Begin VB.Label lblPeso_Bruto 
         Caption         =   "Peso Bruto"
         Height          =   480
         Left            =   135
         TabIndex        =   18
         Tag             =   "PESO_BRUTO"
         Top             =   1620
         Width           =   1500
      End
      Begin VB.Label lblNum_Cajas 
         Caption         =   "Num. Cajas"
         Height          =   480
         Left            =   135
         TabIndex        =   16
         Tag             =   "NUM_CAJAS"
         Top             =   1170
         Width           =   1500
      End
      Begin VB.Label lblNum_Prendas 
         Caption         =   "Num_Prendas"
         Height          =   480
         Left            =   135
         TabIndex        =   14
         Tag             =   "NUM_PRENDAS"
         Top             =   240
         Width           =   1500
      End
   End
   Begin FunctionsButtons.FunctButt FunctOKCancel 
      Height          =   510
      Left            =   2235
      TabIndex        =   12
      Top             =   7125
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmTG_Embarque_Prendas.frx":0000
      Orientacion     =   0
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmTG_Embarque_Prendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sTituliAbrOP  As String
Public lNum_Embarque As Long
Public codigo As String, TipoAdd As String, Tipoa As String
Public Descripcion As String, Tipob As String, estado As String
Public sAccion As String
Public lSec_Embarque As Integer
Public oParent As Object
Dim strSQL As String
Dim scliente As String

Private Sub cmdBusca_Temporada_Click()
 Call BUSCA_TEMPORADA
End Sub

Private Sub cmdBusCliente_Click()
    Dim oTipo As New frmBusqGeneral
    Dim RS As Object
    Set RS = CreateObject("ADODB.Recordset")
    Set oTipo.oParent = Me
    oTipo.SQuery = "SELECT Abr_Cliente as Código, nom_cliente as Descripción FROM TG_Cliente ORDER BY Abr_Cliente"
    oTipo.CARGAR_DATOS
    oTipo.Show 1
    If codigo <> "" Then
        txtcliente.Text = codigo
        txtDes_Cliente.Text = Descripcion
        txtCod_TemCli.Enabled = True
        txtNom_TemCli.Enabled = True
        cmdBusca_Temporada.Enabled = True
        txtCod_TemCli.SetFocus
        codigo = ""
    End If
    Set oTipo = Nothing
    Set RS = Nothing
End Sub

Private Sub Form_Load()
    Dim sSQL As String
    VerificaFabrica txtCod_Fabrica, txtNom_Fabrica
    sTituliAbrOP = DevuelveCampo("select Titulo_Abr_Orden from TG_Control", cCONNECT)
    lblOrdPro.Caption = sTituliAbrOP
    
End Sub
Private Sub VerificaFabrica(ByRef objFabrica As TextBox, ByRef objNombreFabrica As TextBox)
On Error GoTo errorx
    Dim sSQL As String
    Dim iRet As String
    
    sSQL = "SELECT count(*) FROM TG_Fabrica "
    iRet = DevuelveCampo(sSQL, cCONNECT)
    If iRet = 1 Then
        sSQL = "SELECT Cod_Fabrica FROM TG_Fabrica "
        objFabrica.Text = DevuelveCampo(sSQL, cCONNECT)
        
        sSQL = "SELECT Nom_Fabrica FROM TG_Fabrica "
        objNombreFabrica.Text = DevuelveCampo(sSQL, cCONNECT)
        objFabrica.Enabled = False
        objNombreFabrica.Enabled = False
        
    End If
Exit Sub
errorx:
    errores err.Number
    
End Sub


Private Sub FunctOKCancel_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            GrabarLote
        Case "CANCELAR"
            Unload Me
    End Select
End Sub

Private Sub optCliente_Click()
    FrmNP.Visible = False
    FrmCliente.Visible = True
    txtcliente.SetFocus
End Sub

Private Sub optnp_Click()
    FrmNP.Visible = True
    FrmCliente.Visible = False
    txtCod_OrdPro.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

End Sub

Private Sub txtcliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtcliente.Text) = "" Then
            cmdBusCliente_Click
        Else
            strSQL = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE Abr_Cliente LIKE '" & Trim(txtcliente.Text) & "%'"
            txtDes_Cliente.Text = DevuelveCampo(strSQL, cCONNECT)
            txtCod_TemCli.Enabled = True
            txtNom_TemCli.Enabled = True
            cmdBusca_Temporada.Enabled = True
            txtCod_TemCli.SetFocus


        End If
    End If
End Sub

Private Sub txtCod_EstCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCod_LotPurOrd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        BuscaLotePO 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCod_OrdPro_KeyPress(KeyAscii As Integer)
    Dim iLen As Integer
    Dim sSQL As String
    
    
    
    If KeyAscii = vbKeyReturn Then
        If RTrim(txtCod_OrdPro.Text) <> "" Then
            
            txtCod_OrdPro.Text = LPad(txtCod_OrdPro, 5, "0")
        
            If BuscaPedido(txtCod_OrdPro.Text) Then
                txtCod_LotPurOrd.SetFocus
            End If
        End If
    End If

End Sub

Private Function BuscaPedido(ByVal sCod_Pedido As String) As Boolean
On Error GoTo errorx
    Dim sSQL As String
    Dim mRs As ADODB.Recordset
    
    sSQL = "SM_MUESTRA_Cod_OrdPro '" & txtCod_Fabrica.Text & "', '" & txtCod_OrdPro.Text & "'"
    Set mRs = GetDataSet(cCONNECT, sSQL)
    
    If mRs.EOF Then
        Aviso RTrim(lblOrdPro.Caption) & " NO EXISTE", 1
        txtCod_OrdPro.SetFocus
        mRs.Close
        Set mRs = Nothing
        Exit Function
    Else
        txtAbr_Cliente.Text = mRs!Abr_Cliente
        txtAbr_Cliente.Tag = mRs!Cod_Cliente
        txtNom_Cliente.Text = mRs!Nom_Cliente
        txtCod_PurOrd.Text = mRs!cod_purord
        
        txtCod_OrdPro.Text = mRs!Cod_OrdPro
        txtDes_OrdPro.Text = mRs!Des_EstPro
        
        txtCod_LotPurOrd.SetFocus
    End If
    mRs.Close
    Set mRs = Nothing
    BuscaPedido = True
    
Exit Function
errorx:
    errores err.Number
End Function



Public Sub BuscaLotePO(opcion As String)
On Error GoTo errx
Dim rstAux As ADODB.Recordset
Dim strSQL As String
Dim mRs  As ADODB.Recordset

    strSQL = "TG_EMBARQUE_MUESTRA_LOTE_PO_PARTIDA_ARANCELARIA '" & txtAbr_Cliente.Tag & "','" & txtCod_PurOrd & "' ,'" & txtCod_LotPurOrd & "' "
    
    txtCod_LotPurOrd = Trim(txtCod_LotPurOrd)
    txtCod_EstCli = Trim(txtCod_EstCli)
    txtarancelaria1 = Trim(txtarancelaria1)
    txtarancelaria2 = Trim(txtarancelaria2)
    txtarancelaria3 = Trim(txtarancelaria3)
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    
    'frmBusqGeneral3.gexLista.Columns("COD_PURORD").Width = 0
    frmBusqGeneral3.gexLista.Columns("Cod_LotPurOrd").Width = 570
    frmBusqGeneral3.gexLista.Columns("Cod_EstCli").Width = 1500
    frmBusqGeneral3.gexLista.Columns("Num_Partida_Arancelaria").Width = 1500
    frmBusqGeneral3.gexLista.Columns("Sec_Partida_Arancelaria").Width = 1500
    frmBusqGeneral3.gexLista.Columns("Num_Categoria_Internacional").Width = 1500
    
    'frmBusqGeneral3.gexLista.Columns("COD_PURORD").Caption = "Cod.PurOrd"
    frmBusqGeneral3.gexLista.Columns("Cod_LotPurOrd").Caption = "Cod.LotPurOrd"
    frmBusqGeneral3.gexLista.Columns("Cod_ESTCLI").Caption = "Estilo"
    frmBusqGeneral3.gexLista.Columns("Num_Partida_Arancelaria").Caption = "Num.PartidaArancelaria"
    frmBusqGeneral3.gexLista.Columns("Sec_Partida_Arancelaria").Caption = "Sec.PartidaArancelaria"
    frmBusqGeneral3.gexLista.Columns("Num_Categoria_Internacional").Caption = "Num.CategoriaInternacional"
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtCod_LotPurOrd = ""
    txtCod_EstCli = ""
    
    If codigo <> "" Then
        txtCod_LotPurOrd = codigo
        txtCod_EstCli = Descripcion
        txtarancelaria1 = TipoAdd
        txtarancelaria2 = Tipoa
        txtarancelaria3 = Tipob
        
        Set mRs = GetDataSet(cCONNECT, "SM_DATOS_PURORD_LOTE '" & txtAbr_Cliente.Tag & "','" & txtCod_PurOrd & "','" & txtCod_LotPurOrd & "','" & txtCod_EstCli & "'")
        If Not mRs Is Nothing Then
            txtNum_Cajas_Prog = mRs!Num_Cajas
            txtNum_Prendas_Prog = mRs!Num_Prendas
            mRs.Close
        End If
        Set mRs = Nothing
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    codigo = ""
    Descripcion = ""
Exit Sub
errx:
    errores err.Number
End Sub



Private Sub txtCod_TemCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_TemCli.Text) = "" Then
            Call BUSCA_TEMPORADA
        Else
            strSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"
            strSQL = "SELECT Nom_TemCli FROM TG_TemCli WHERE Cod_Cliente='" & DevuelveCampo(strSQL, cCONNECT) & "' AND Cod_TemCli='" & txtCod_TemCli.Text & "'"
            txtNom_TemCli.Text = DevuelveCampo(strSQL, cCONNECT)
            txtestilo.Enabled = True
            txtestilo.SetFocus
        End If
    End If

End Sub

Private Sub BUSCA_TEMPORADA()
    Dim oTipo As New frmBusqGeneral
    Dim RS As Object
    Set RS = CreateObject("ADODB.Recordset")
    Set oTipo.oParent = Me

    strSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"
    oTipo.SQuery = "SELECT  Cod_TemCli as Código, Nom_TemCli as Descripción FROM TG_TemCli WHERE Cod_Cliente='" & DevuelveCampo(strSQL, cCONNECT) & "'"

    oTipo.CARGAR_DATOS
    oTipo.Show 1
    If codigo <> "" Then
        txtCod_TemCli.Text = codigo
        txtNom_TemCli.Text = Descripcion
    End If
    Set oTipo = Nothing
    Set RS = Nothing
    txtestilo.Enabled = True
    txtestilo.SetFocus
End Sub

Private Sub txtCubicaje_Prog_GotFocus()
    SelectionText txtCubicaje_Prog
End Sub

Private Sub txtCubicaje_Prog_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub TxtDes_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtDes_Cliente) > 4 Then
            strSQL = "SELECT Abr_Cliente FROM TG_CLIENTE WHERE Nom_Cliente LIKE '" & Trim(txtDes_Cliente.Text) & "%'"
            txtcliente.Text = DevuelveCampo(strSQL, cCONNECT)
            strSQL = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"
            txtDes_Cliente.Text = DevuelveCampo(strSQL, cCONNECT)
            txtCod_TemCli.Enabled = True
            txtNom_TemCli.Enabled = True
            cmdBusca_Temporada.Enabled = True
            txtCod_TemCli.SetFocus


        Else
            MsgBox ("El Texto Ingresado debe contar con un mínimo de 5 caracteres")
            txtDes_Cliente.SetFocus
        End If
    End If

End Sub

Private Sub txtestilo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim oTipo As New frmBusqGeneral
    Dim RS As Object
    Set RS = CreateObject("ADODB.Recordset")
    Set oTipo.oParent = Me

    scliente = DevuelveCampo("select cod_cliente from tg_cliente where abr_cliente='" & txtcliente.Text & "'", cCONNECT)
    oTipo.SQuery = " SM_TG_EstCliEst_ViewxCliente '" & scliente & "', '" & txtCod_TemCli.Text & "'"

    oTipo.CARGAR_DATOS
    oTipo.Show 1
    If codigo <> "" Then
        txtestilo.Text = codigo
        Call busca_detalle(codigo)
    End If
    Set oTipo = Nothing
    Set RS = Nothing

    txtarancelaria1.SetFocus
End If
End Sub

Private Function busca_detalle(ByVal sestilo As String) As Boolean
On Error GoTo errorx

    Dim oTipo As New frmBusqGeneral_Lis
    Dim RS As Object
    Set RS = CreateObject("ADODB.Recordset")
    Set oTipo.oParent = Me

    oTipo.SQuery = "TG_MUESTRA_LOTES_CLIENTE_TEMPORADA_ESTILO '" & scliente & "', '" & txtCod_TemCli.Text & "','" & txtestilo.Text & "'"

    oTipo.CARGAR_DATOS
    oTipo.Show 1
    If codigo <> "" Then
        txtAbr_Cliente.Text = txtcliente.Text
        txtAbr_Cliente.Text = txtcliente.Text
        txtNom_Cliente.Text = txtDes_Cliente.Text
        txtCod_PurOrd.Text = codigo
        txtCod_LotPurOrd.Text = Descripcion
        txtCod_EstCli.Text = estado
    End If
    Set oTipo = Nothing
    Set RS = Nothing

    
Exit Function
errorx:
    errores err.Number
End Function

Private Sub txtNum_Cajas_Prog_GotFocus()
    SelectionText txtNum_Cajas_Prog
End Sub

Private Sub txtNum_Cajas_Prog_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtNum_Prendas_Prog_GotFocus()
    SelectionText txtNum_Prendas_Prog
End Sub

Private Sub txtNum_Prendas_Prog_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtPeso_Bruto_Prog_GotFocus()
    SelectionText txtPeso_Bruto_Prog
End Sub

Private Sub txtPeso_Bruto_Prog_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtPeso_Neto_Prog_GotFocus()
    SelectionText txtPeso_Neto_Prog
End Sub

Private Sub txtPeso_Neto_Prog_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtPre_Unitario_GotFocus()
SelectionText txtPre_Unitario
End Sub

Private Sub txtPre_Unitario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub GrabarLote()
On Error GoTo errx
Dim sSQL As String

sSQL = "TG_EMBARQUE_PRENDAS_MAN '$',$,$,'$','$','$','$',$,$,$,$,$,$,'$','$','$'"
  
sSQL = VBsprintf(sSQL, sAccion, lNum_Embarque, lSec_Embarque, txtAbr_Cliente.Tag, txtCod_PurOrd.Text, txtCod_LotPurOrd, txtCod_EstCli, txtNum_Prendas_Prog, txtPre_Unitario, txtNum_Cajas_Prog, txtPeso_Bruto_Prog, txtPeso_Neto_Prog, txtCubicaje_Prog, txtarancelaria1, txtarancelaria2, txtarancelaria3)
  

ExecuteCommandSQL cCONNECT, sSQL

oParent.Buscar
Unload Me

Exit Sub
errx:
    errores err.Number
End Sub
