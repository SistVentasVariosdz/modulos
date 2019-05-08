VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAddMovimAlm 
   Caption         =   "Movimientos de Almacen"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   615
      Left            =   4200
      TabIndex        =   18
      Top             =   4095
      Width           =   2505
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "0~0~ACEPTAR~True~True~&Aceptar~0~0~1~~0~False~False~&Aceptar~~1~0~CANCELAR~True~True~&Cancelar~0~0~2~~0~False~False~&Cancelar~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Fradetalle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3930
      Left            =   0
      TabIndex        =   20
      Tag             =   "Detail"
      Top             =   0
      Width           =   10935
      Begin VB.TextBox txtNumConosHilosCoser 
         Height          =   315
         Left            =   8430
         TabIndex        =   17
         Top             =   2790
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Frame FraSolicitante 
         Caption         =   "Solicitante"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4500
         TabIndex        =   40
         Top             =   3195
         Visible         =   0   'False
         Width           =   6375
         Begin VB.TextBox TxtTip_Trabajador 
            Height          =   285
            Left            =   1080
            MaxLength       =   1
            TabIndex        =   43
            Text            =   "O"
            Top             =   240
            Width           =   420
         End
         Begin VB.TextBox TxtCod_Trabajador 
            Height          =   285
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   42
            Top             =   240
            Width           =   780
         End
         Begin VB.TextBox TxtNom_Trabajador 
            Height          =   285
            Left            =   2400
            TabIndex        =   41
            Top             =   240
            Width           =   3825
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Codigo"
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.TextBox TxtCod_CenCosto 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         MaxLength       =   12
         TabIndex        =   13
         Top             =   600
         Width           =   765
      End
      Begin VB.TextBox TxtDes_CenCosto 
         Height          =   315
         Left            =   7620
         TabIndex        =   14
         Top             =   600
         Width           =   3045
      End
      Begin VB.TextBox TxtCod_Cliente 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         MaxLength       =   12
         TabIndex        =   11
         Top             =   280
         Width           =   765
      End
      Begin VB.TextBox TxtNom_cliente 
         Height          =   315
         Left            =   7620
         TabIndex        =   12
         Top             =   280
         Width           =   3045
      End
      Begin VB.TextBox TxtDes_TipMov 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   280
         Width           =   3285
      End
      Begin VB.TextBox TxtCod_TipMov 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         MaxLength       =   12
         TabIndex        =   0
         Top             =   280
         Width           =   765
      End
      Begin VB.ComboBox CmbOrdComp 
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   960
         Width           =   2355
      End
      Begin VB.TextBox TxtObservaciones 
         Height          =   645
         Left            =   1140
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   1950
         Width           =   4155
      End
      Begin VB.TextBox TxtOrdPro 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         MaxLength       =   5
         TabIndex        =   6
         Top             =   1290
         Width           =   945
      End
      Begin VB.TextBox Txtproveedor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         MaxLength       =   12
         TabIndex        =   3
         Top             =   960
         Width           =   1245
      End
      Begin VB.TextBox TxtDetalle 
         Height          =   315
         Left            =   2400
         TabIndex        =   4
         Top             =   960
         Width           =   2805
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   345
         Left            =   5250
         TabIndex        =   5
         Top             =   960
         Width           =   345
      End
      Begin VB.TextBox TxtGuia 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   8
         Top             =   1620
         Width           =   1935
      End
      Begin VB.TextBox txtNum_SecOrd 
         Height          =   315
         Left            =   3120
         TabIndex        =   7
         Top             =   1290
         Width           =   630
      End
      Begin VB.Frame fraDatosAdic 
         Caption         =   "Datos Adicionales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5760
         TabIndex        =   21
         Top             =   1650
         Visible         =   0   'False
         Width           =   4050
         Begin VB.TextBox txtDes_Color 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   24
            Top             =   600
            Width           =   2730
         End
         Begin VB.TextBox txtCod_TipOrdTra 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   23
            Top             =   240
            Width           =   525
         End
         Begin VB.TextBox txtCod_Ordtra 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1590
            TabIndex        =   22
            Top             =   240
            Width           =   2220
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Color :"
            Height          =   195
            Left            =   210
            TabIndex        =   26
            Top             =   660
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "O/T :"
            Height          =   195
            Left            =   195
            TabIndex        =   25
            Top             =   330
            Width           =   390
         End
      End
      Begin VB.CommandButton CmdOC 
         Height          =   330
         Left            =   8520
         Picture         =   "FrmAddMovimAlm.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1320
         Width           =   705
      End
      Begin VB.TextBox txtParteSalida 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3585
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1620
         Width           =   1710
      End
      Begin VB.TextBox txtGlosa_Hilado 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   16
         Top             =   2790
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DtFechaMov 
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Top             =   630
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   23855105
         CurrentDate     =   37270
      End
      Begin VB.Label lblConos 
         Caption         =   "Num de Conos de  Hilos de Coser"
         Height          =   195
         Left            =   5820
         TabIndex        =   45
         Top             =   2865
         Visible         =   0   'False
         Width           =   2520
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Index           =   0
         Left            =   5760
         TabIndex        =   39
         Top             =   405
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Mov:"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   38
         Top             =   405
         Width           =   720
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   60
         TabIndex        =   37
         Tag             =   "Hilado :"
         Top             =   1035
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro Costo:"
         Height          =   195
         Index           =   3
         Left            =   5760
         TabIndex        =   36
         Top             =   705
         Width           =   960
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Mov:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   60
         TabIndex        =   35
         Tag             =   "Hilado :"
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Orden Comp:"
         Height          =   195
         Index           =   5
         Left            =   5760
         TabIndex        =   34
         Top             =   1035
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones:"
         Height          =   195
         Index           =   6
         Left            =   60
         TabIndex        =   33
         Top             =   2040
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Orden Prod:"
         Height          =   195
         Index           =   7
         Left            =   60
         TabIndex        =   32
         Top             =   1365
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Guia:"
         Height          =   195
         Index           =   8
         Left            =   60
         TabIndex        =   31
         Top             =   1695
         Width           =   375
      End
      Begin VB.Label lblsecuencia 
         AutoSize        =   -1  'True
         Caption         =   "Secuencia:"
         Height          =   195
         Left            =   2310
         TabIndex        =   30
         Top             =   1350
         Width           =   810
      End
      Begin VB.Label Label4 
         Caption         =   "O.C. Relacionadas:"
         Height          =   255
         Left            =   6960
         TabIndex        =   29
         Top             =   1395
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "P/S:"
         Height          =   195
         Index           =   4
         Left            =   3180
         TabIndex        =   28
         Top             =   1680
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Glosa Hilado:"
         Height          =   195
         Index           =   9
         Left            =   60
         TabIndex        =   27
         Top             =   2865
         Width           =   945
      End
   End
End
Attribute VB_Name = "FrmAddMovimAlm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public Codigo As String, Descripcion As String, TipoAdd As String
Public vCod_Almacen As String, Num_MovStk As String
Public varCod_Fabrica As String, varNum_SecOrd As String, Accion As String
Public sCod_AlmacenOrigen As String, sNum_MovStkOrigen As String
Dim Tip_Accion As String, Cod_TipOrdPro As String, Cod_TipAnx As String, Cod_ClaOrdComp As String, Cod_ClaMov As String, Cod_Fabrica As String, Flg_Rollo As String
Public Tip_presentacion As String, Tip_item As String, Estado As String
Public vCod_Cliente As String, Cod_TipOrdTra As String
Public Paso As Boolean, vOk As Boolean
Public vFlg_Almacen_Tejeduria As String
Public Almacen

Private Sub CmbOrdComp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub



Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    If Validar = False Then Exit Sub
    Call Grabar
Case "CANCELAR"
    vOk = False
    Unload Me
End Select
End Sub

Public Sub BUSCA_Num_SecOrd()
    
    Dim oTipo As New frmBusqNum_SecOrd
    Dim Rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    
    oTipo.sQuery = "EXEC UP_SEL_CFORDPRO_Num_SecOrd '" & varCod_Fabrica & "','" & Me.TxtOrdPro.Text & "'"
    
    oTipo.Cargar_Datos
    oTipo.Show 1
    If varNum_SecOrd <> "" Then
        Me.txtNum_SecOrd.Text = varNum_SecOrd
        varNum_SecOrd = ""
    End If
    Set oTipo = Nothing
    Set Rs = Nothing
    
    If TxtGuia.Enabled Then TxtGuia.SetFocus
    
End Sub

Sub Grabar()
On Error GoTo errGrabar

vCod_Cliente = DevuelveCampo("SELECT COD_CLIENTE FROM TG_CLIENTE WHERE ABR_CLIENTE='" & Trim(TxtCod_Cliente.Text) & "'", cConnect)

Call Datos_TipoMov

strSQL = "UP_Lg_Movstk '" & UCase(Accion) & "','" & vCod_Almacen & "','" & Num_MovStk & "','" & _
         DtFechaMov.Value & "','" & vusu & "','" & Txtproveedor & "','" & _
         Trim(TxtCod_CenCosto.Text) & "','" & TxtCod_TipMov & "','" & Trim(Mid(Me.CmbOrdComp, 1, 3)) & "','" & _
         Trim(Mid(Me.CmbOrdComp, 5, 6)) & "','" & vCod_Cliente & "','" & Cod_TipOrdPro & "','" & _
         TxtOrdPro & "','" & TxtObservaciones & "','" & TxtGuia & "','" & Cod_Fabrica & "'" & _
         IIf(Me.txtNum_SecOrd.Visible, ",'" & Me.txtNum_SecOrd.Text & "'", ",'0'") & ",'" & Trim(Me.txtCod_TipOrdTra.Text) & "','" & Trim(Me.txtCod_Ordtra.Text) & "','" & txtGlosa_Hilado.Text & "','" & sCod_AlmacenOrigen & "','" & sNum_MovStkOrigen & "','" & _
         TxtTip_Trabajador.Text & "','" & TxtCod_Trabajador.Text & "'," & txtNumConosHilosCoser.Text & ""
                  
Call ExecuteSQL(cConnect, strSQL)
vOk = True
Unload Me
Exit Sub
errGrabar:
    vOk = False
    ErrorHandler err, "Grabar"
End Sub

Sub Datos_TipoMov()
Dim sFlg_Partida_Generada As Variant, vcod_cencost As String, _
    sFlg_Ot_Tejeduria_Generada As String

    
    Tip_Accion = DevuelveCampo("select tip_accion from lg_tiposmov where Cod_TipMov='" & Trim(TxtCod_TipMov.Text) & "'", cConnect)
    Cod_TipOrdPro = DevuelveCampo("select Cod_TipOrdPro from lg_tiposmov where Cod_TipMov='" & Trim(TxtCod_TipMov.Text) & "'", cConnect)
    Cod_TipAnx = DevuelveCampo("select isnull(Cod_TipAnx,'') from lg_tiposmov where Cod_TipMov='" & Trim(TxtCod_TipMov.Text) & "'", cConnect)
    Cod_ClaOrdComp = DevuelveCampo("select rtrim(Cod_ClaOrdComp) from lg_tiposmov where Cod_TipMov='" & Trim(TxtCod_TipMov.Text) & "'", cConnect)
    Cod_ClaMov = DevuelveCampo("select rtrim(Cod_ClaMov) from lg_tiposmov where Cod_TipMov='" & Trim(TxtCod_TipMov.Text) & "'", cConnect)
    Cod_Fabrica = DevuelveCampo("select rtrim(Cod_Fabrica ) from lg_tiposmov where Cod_TipMov='" & Trim(TxtCod_TipMov.Text) & "'", cConnect)
    Flg_Rollo = DevuelveCampo("select isnull(Flg_Rollo,'') from lg_tiposmov where Cod_TipMov='" & Trim(TxtCod_TipMov.Text) & "'", cConnect)
    vcod_cencost = DevuelveCampo("select isnull(cod_cencost,'') from lg_tiposmov where Cod_TipMov='" & Trim(TxtCod_TipMov.Text) & "'", cConnect)
    Tip_Relacion = DevuelveCampo("select isnull(Tip_Relacion,'') from lg_tiposmov where Cod_TipMov='" & Trim(TxtCod_TipMov) & "'", cConnect)
    sFlg_Ot_Tejeduria_Generada = DevuelveCampo("select isnull(Flg_Ot_Tejeduria_Generada, '') FROM lg_tiposmov where Cod_TipMov = '" & Trim(TxtCod_TipMov) & "'", cConnect)
    'TxtCod_CenCosto.Text = vcod_cencost
    If Cod_ClaMov = "S" Then
        FraSolicitante.Visible = True
    Else
        TxtTip_Trabajador.Text = ""
        TxtCod_Trabajador.Text = ""
        TxtNom_Trabajador.Text = ""
        FraSolicitante.Visible = False
    End If
    If Trim(vcod_cencost) <> "" Then
        TxtCod_CenCosto.Enabled = False
        TxtDes_CenCosto.Enabled = False
        TxtCod_CenCosto.Text = vcod_cencost
        TxtDes_CenCosto.Text = DevuelveCampo("select des_cencost from tg_cencosto where cod_cencost ='" & vcod_cencost & "'", cConnect)
    Else
        TxtCod_CenCosto.Enabled = True
        TxtDes_CenCosto.Enabled = True
        TxtDes_CenCosto.Text = ""
    End If
    
    strSQL = "SELECT ISNULL(Flg_SecOrd,'') FROM lg_tiposmov WHERE Cod_TipMov = '" & Trim(TxtCod_TipMov.Text) & "'"
    If DevuelveCampo(strSQL, cConnect) = "*" Then
        lblsecuencia.Visible = True
        Me.txtNum_SecOrd.Visible = True
        
        strSQL = "SELECT Cod_Fabrica FROM lg_tiposmov WHERE Cod_TipMov = '" & Trim(TxtCod_TipMov.Text) & "'"
        varCod_Fabrica = DevuelveCampo(strSQL, cConnect)
    Else
        lblsecuencia.Visible = False
        Me.txtNum_SecOrd.Visible = False
        varCod_Fabrica = ""
    End If

If Cod_ClaMov = "S" And Tip_Accion = "E" Then TxtGuia.Enabled = False
'Estas son funcionalidades nuevas anadidas
strSQL = "SELECT flg_partida_generada FROM LG_TIPOSMOV WHERE Cod_TipMov = '" & _
         Trim(TxtCod_TipMov.Text) & "'"
         
sFlg_Partida_Generada = DevuelveCampo(strSQL, cConnect)
sFlg_Partida_Generada = IIf(IsNull(sFlg_Partida_Generada), "", CStr(sFlg_Partida_Generada))
Me.fraDatosAdic.Visible = False

Me.txtCod_Ordtra.Text = ""
Me.txtCod_TipOrdTra = ""
Me.txtDes_Color = ""

'If sTipo = "I" Then
    strSQL = "SELECT COUNT(*) FROM LG_ALMACEN WHERE Cod_Almacen = '" & vCod_Almacen & "' AND (Tip_Item = 'T' OR Tip_Item = 'H') AND Tip_Presentacion = 'C'"
    If DevuelveCampo(strSQL, cConnect) Then
        strSQL = "SELECT Cod_ClaOrdComp FROM LG_TIPOSMOV " & _
                 "WHERE Cod_TipMov = '" & Trim(TxtCod_TipMov.Text) & "' " & _
                 "AND (Tip_Item = 'T' OR Tip_Item = 'H')  " & _
                 "AND (Flg_Partidas_Tinto = 'S' OR Flg_Partida_Generada = 'S')"
        strSQL = DevuelveCampo(strSQL, cConnect)
        If Trim(strSQL) <> "" Then
            strSQL = "SELECT COUNT(*) FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp = '" & strSQL & "' AND (Tip_Item = 'T' OR Tip_Item = 'H' ) AND Tip_Presentacion = 'T' AND Cod_Protex IS NOT NULL"
            If DevuelveCampo(strSQL, cConnect) > 0 Then
                Me.fraDatosAdic.Visible = True
                
                If Accion = "I" Then
                    'Aqui ponemos el levantamiento del formulario
                    Load frmMovAlmacenAnexo
                    frmMovAlmacenAnexo.varCod_ClaOrdComp = DevuelveCampo("select rtrim(Cod_ClaOrdComp) from lg_tiposmov where Cod_TipMov='" & Trim(TxtCod_TipMov.Text) & "'", cConnect)
                    frmMovAlmacenAnexo.varCod_Fabrica = DevuelveCampo("select rtrim(Cod_Fabrica ) from lg_tiposmov where Cod_TipMov='" & Trim(TxtCod_TipMov.Text) & "'", cConnect)
                    frmMovAlmacenAnexo.varCod_Clamov = DevuelveCampo("select rtrim(Cod_ClaMov) from lg_tiposmov where Cod_TipMov='" & Trim(TxtCod_TipMov.Text) & "'", cConnect)
                    Set frmMovAlmacenAnexo.oParent = Me
                    frmMovAlmacenAnexo.varTip_Item = Tip_item
                    If Tip_item = "H" Then
                        frmMovAlmacenAnexo.txtCod_TipOrdTra1er = "TH"
                        frmMovAlmacenAnexo.txtCod_TipOrdTra2da = "TH"
                        frmMovAlmacenAnexo.txtCod_TipOrdTraPar = "TH"
                    Else
                        frmMovAlmacenAnexo.txtCod_TipOrdTra1er = "TI"
                        frmMovAlmacenAnexo.txtCod_TipOrdTra2da = "TI"
                        frmMovAlmacenAnexo.txtCod_TipOrdTraPar = "TI"
                    End If
                    
                    ' Nueva Opción
                    If sFlg_Partida_Generada = "S" Then
                        frmMovAlmacenAnexo.opt2doEnvio = True
                        frmMovAlmacenAnexo.opt1erEnvio.Visible = False
                        frmMovAlmacenAnexo.opt2doEnvio.Visible = False
                        frmMovAlmacenAnexo.fra2doEnvio.Caption = ""
                    End If
                    
                    frmMovAlmacenAnexo.Show 1
                    
                    'Aqui bloquearemos algunos campos
                    TxtCod_TipMov.Enabled = False
                    TxtDes_TipMov.Enabled = False
                    
                    DtFechaMov.Enabled = False
                    Txtproveedor.Enabled = False
                    TxtDetalle.Enabled = False
                    Command1.Enabled = False
                    TxtOrdPro.Enabled = False
                    
                    TxtCod_Cliente.Enabled = False
                    TxtNom_cliente.Enabled = False
                    
                    TxtCod_CenCosto.Enabled = False
                    TxtDes_CenCosto.Enabled = False
                    
                    CmbOrdComp.Enabled = False
                    
                    If TxtGuia.Enabled Then
                        TxtGuia.SetFocus
                    End If
                Else
                    Me.txtCod_TipOrdTra.Enabled = False
                    Me.txtCod_Ordtra.Enabled = False
                    Me.txtDes_Color.Enabled = False
                End If
            End If
        Else
            strSQL = ""
        End If
    Else
        
    End If
End Sub


Function Validar() As Boolean
Validar = True

If Trim(TxtCod_TipMov.Text) = "" Then
        MsgBox "Seleccione un tipo de movimiento", vbInformation
        Validar = False
        TxtCod_TipMov.SetFocus
        Exit Function
End If

If Tip_Accion = "I" Then
    If Cod_TipOrdPro = "" Then
        If Trim(TxtCod_CenCosto.Text) = "" Then
            MsgBox "Seleccione Centro de Costo", vbInformation
            Validar = False
            TxtCod_CenCosto.SetFocus
            Exit Function
        End If
    Else
        If Cod_TipOrdPro = "CF" Then
            If DevuelveCampo("select count(*) from es_ordpro where cod_ordpro ='" & Ceros(TxtOrdPro) & "'", cConnect) <= 0 Then
                MsgBox "La Orden de Produccion no existe", vbInformation
                Validar = False
                Exit Function
            End If
        Else
            If Cod_TipOrdPro = "CO" Then
                If DevuelveCampo("select count(*) from CO_ordpro where CO_codordpro ='" & Ceros(TxtOrdPro) & "'", cConnect) <= 0 Then
                    MsgBox "La Orden de Corte no existe", vbInformation
                    Validar = False
                    Exit Function
                End If
            End If
        End If
    End If
    
    'Aqui validamos cuando es del tipo cf_*
    strSQL = "SELECT ISNULL(Flg_SecOrd,'') FROM lg_tiposmov WHERE Cod_TipMov = '" & Trim(TxtCod_TipMov.Text) & "'"
    If DevuelveCampo(strSQL, cConnect) = "*" Then
        strSQL = "SELECT COUNT(*) FROM CF_ORDPRO WHERE Cod_Fabrica = '" & varCod_Fabrica & "' AND Cod_OrdPro = '" & Me.TxtOrdPro & "'"
        If DevuelveCampo(strSQL, cConnect) = 0 Then
            MsgBox "El código no existe. Sirvase verificar", vbInformation, "Mensaje"
            Validar = False
            Exit Function
        End If
        
        If Trim(Me.txtNum_SecOrd.Text) = "" Then
            MsgBox "El código de secuencia no puede estar vacio. Sirvase verificar", vbInformation, "Mensaje"
            Validar = False
            Exit Function
        End If
        
        strSQL = "SELECT COUNT(*) FROM CF_ORDPRO WHERE Cod_Fabrica = '" & varCod_Fabrica & "' AND Cod_OrdPro = '" & Me.TxtOrdPro.Text & "' AND Num_SecOrd = '" & Me.txtNum_SecOrd.Text & "'"
        If DevuelveCampo(strSQL, cConnect) = 0 Then
            MsgBox "El código de secuencia no existe. Sirvase verificar", vbInformation, "Mensaje"
            Validar = False
            Exit Function
        End If
    End If
End If

If MovIngresoTintoPropiaEditable(vCod_Almacen, Num_MovStk, Trim(TxtCod_TipMov.Text), Accion, Txtproveedor.Text, TxtGuia.Text) <> "0" Then
    If Tip_item = "T" And Tip_presentacion = "T" Then
        If sCod_AlmacenOrigen = "" Then
            Aviso "No se ha seleccionado Guia Correctamente", 3
            Validar = False
            Exit Function
        End If
    End If
End If

If Cod_TipAnx = "P" Then
    If Me.Txtproveedor = "" Then
        MsgBox "Debe ingresar un Proveedor", vbInformation
            Validar = False
            Txtproveedor.SetFocus
            Exit Function
    End If
ElseIf Cod_TipAnx = "C" Then
    If Trim(TxtCod_Cliente.Text) = "" Then
        MsgBox "Debe seleccionar un Cliente", vbInformation
            Validar = False
            TxtCod_Cliente.SetFocus
            Exit Function
    End If
End If

If Cod_ClaOrdComp <> "" Then
    If Me.CmbOrdComp = "" Then
        MsgBox "Debe seleccionar una Orden de Compra", vbInformation
            Validar = False
            Exit Function
    End If
End If
End Function

Public Sub CARGA_ORDCOMP()
LlenaCombo CmbOrdComp, "EXEC sm_muestra_ordenes_compra_segun_tipomov '" & _
           Trim(TxtCod_TipMov.Text) & "','" & Txtproveedor.Text & "','" & _
           Trim(TxtCod_Cliente.Text) & "', '" & IIf(Estado = "NUEVO", "I", _
           "") & "'", cConnect
End Sub

Private Sub CmbOrdComp_GotFocus()
    Call CARGA_ORDCOMP
End Sub

Private Sub CmdOC_Click()
    If Trim(Me.Txtproveedor.Text) = "" Then
        MsgBox "No tiene Proveedor", vbInformation
        Exit Sub
    End If

    If Trim(Me.TxtGuia.Text) = "" Then
        MsgBox "Ingrese la guia", vbInformation
        If TxtGuia.Enabled Then Me.TxtGuia.SetFocus
        Exit Sub
    End If

    Load FrmMuestraGuias
    With FrmMuestraGuias
        .vCod_Almacen = vCod_Almacen
        .vCod_Proveedor = Trim(Txtproveedor.Text)
        .vNum_Guia = Trim(TxtGuia.Text)
        .CARGA_GRID
        .Show 1
    End With
End Sub

 

 

Private Sub TxtCod_CenCosto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BuscaCentro_Costo(1)
End If
End Sub

Private Sub TxtCod_Cliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BuscaCliente(1)
End If
End Sub

Private Sub TxtCod_TipMov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BuscaTipoMov(1)
End If
End Sub

Private Sub TxtCod_TipMov_LostFocus()
Call Datos_TipoMov
End Sub

Private Sub TxtCod_Trabajador_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BUSCATRABAJADOR(2)
End If
End Sub

Private Sub TxtDes_CenCosto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BuscaCentro_Costo(2)
End If
End Sub

Private Sub TxtDes_TipMov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BuscaTipoMov(2)
End If
End Sub

Private Sub TxtDes_TipMov_LostFocus()
Call Datos_TipoMov
End Sub

Private Sub TxtDetalle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If DevuelveCampo("select count(*) from lg_proveedor where Des_Proveedor like '%" & TxtDetalle & "%'", cConnect) > 0 Then
    Set frmBusqGeneral.oParent = Me
        frmBusqGeneral.sQuery = "select Cod_Proveedor as Codigo ,Des_Proveedor as Nombre from lg_proveedor where Des_Proveedor like '%" & Trim(TxtDetalle) & "%' "
        frmBusqGeneral.Cargar_Datos
        frmBusqGeneral.Show 1
        Me.Txtproveedor = Codigo
        TxtDetalle = Descripcion
        If TxtOrdPro.Enabled Then
            Me.TxtOrdPro.SetFocus
        End If
    Else
        Txtproveedor = DevuelveCampo("Select Cod_Proveedor from lg_proveedor where Des_Proveedor='" & TxtDetalle & "'", cConnect)
        If TxtOrdPro.Enabled Then
            Me.TxtOrdPro.SetFocus
        End If
    End If
End If
End Sub



Private Sub txtGlosa_Hilado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtGuia_KeyPress(KeyAscii As Integer)
    Dim sMovIngresoTintoPropiaEditable As String
    
    If KeyAscii = 13 Then
        If RTrim(TxtGuia.Text) = "" Then
            sMovIngresoTintoPropiaEditable = MovIngresoTintoPropiaEditable(vCod_Almacen, Num_MovStk, Trim(TxtCod_TipMov.Text), Accion, Txtproveedor.Text, TxtGuia.Text)
            If sMovIngresoTintoPropiaEditable = "1" Or sMovIngresoTintoPropiaEditable = "2" Then
                If ShowGuiasTintoPropia Then
                    Me.TxtObservaciones.SetFocus
                End If
            End If
        Else
            Me.TxtObservaciones.SetFocus
        End If
    End If
End Sub

Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BuscaCliente(2)
End If
End Sub

Private Sub TxtNom_Trabajador_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BUSCATRABAJADOR(3)
End If
End Sub

Private Sub txtNum_SecOrd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.txtNum_SecOrd.Text = "" Then
            Call BUSCA_Num_SecOrd
        End If
        If TxtGuia.Enabled Then Me.TxtGuia.SetFocus
    End If
End Sub

Private Sub txtNumConosHilosCoser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub
 

Private Sub TxtObservaciones_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub



Private Sub TxtOrdPro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtOrdPro = Format(TxtOrdPro.Text, "00000")
    TxtOrdPro = Trim(DevuelveCampo("Select dbo.uf_devuelvecodigo(5," & IIf(Trim(TxtOrdPro) = "", 0, TxtOrdPro) & ")", cConnect))
    If Cod_TipOrdPro = "CO" Then
        If DevuelveCampo("select count(*) from CO_ordpro where CO_codordpro ='" & Ceros(Me.TxtOrdPro) & "'", cConnect) <= 0 Then
            MsgBox "La Orden de Corte no existe", vbInformation
            TxtOrdPro = ""
            Exit Sub
        End If
    ElseIf Cod_TipOrdPro = "CF" Then
        If DevuelveCampo("select count(cod_ordpro) from es_ordpro where cod_ordpro='" & TxtOrdPro & "'", cConnect) <= 0 Then
            MsgBox "El codigo no existe", vbInformation
            TxtOrdPro = ""
            Exit Sub
        End If
    End If
    
    strSQL = "SELECT ISNULL(Flg_SecOrd,'') FROM lg_tiposmov WHERE Cod_TipMov = '" & Trim(TxtCod_TipMov.Text) & "'"
    varCod_Fabrica = DevuelveCampo("SELECT COD_FABRICA FROM LG_TIPOSMOV WHERE COD_TIPMOV='" & Trim(TxtCod_TipMov.Text) & "'", cConnect)
    If DevuelveCampo(strSQL, cConnect) = "*" Then
        'lblsecuencia.Visible = True
        'Me.txtNum_SecOrd.Visible = True
        strSQL = "SELECT COUNT(*) FROM CF_ORDPRO WHERE Cod_Fabrica = '" & varCod_Fabrica & "' AND Cod_OrdPro = '" & Me.TxtOrdPro & "'"
        If DevuelveCampo(strSQL, cConnect) = 0 Then
            MsgBox "El código no existe. Sirvase verificar", vbInformation, "Mensjae"
        Else
            Call BUSCA_Num_SecOrd
        End If
    End If
    'Me.TxtGuia.SetFocus
    If TxtGuia.Enabled Then Me.TxtGuia.SetFocus
End If
End Sub

Private Sub Txtproveedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Txtproveedor.Text) = "" Then
        MsgBox "Ingrese el codigo", vbInformation
        Exit Sub
    End If
    Txtproveedor = DevuelveCampo("select dbo.uf_devuelvecodigo(12," & Txtproveedor & ")", cConnect)
    If ExisteCampo("Cod_Proveedor", "lg_proveedor", Txtproveedor, cConnect, True) Then
        TxtDetalle = DevuelveCampo("Select Des_Proveedor from lg_proveedor where Cod_Proveedor='" & Txtproveedor & "'", cConnect)
        If TxtOrdPro.Enabled Then
            Me.TxtOrdPro.SetFocus
        End If
    Else
        MsgBox "El codigo no existe", vbInformation
    End If
End If
End Sub

Private Function MovIngresoTintoPropiaEditable(sCod_Almacen As String, sNum_MovStk As String, sCod_TipMov As String, sAccion As String, SCOD_PROVEEDOR As String, sNumGuia As String) As String
On Error GoTo errx
Dim mRs As ADODB.Recordset
Dim sSQL As String

sSQL = "UP_VERIFICA_INGRESO_TINTOPROPIA_EDIT '$','$','$','$','$', '$'"
sSQL = VBsprintf(sSQL, sCod_Almacen, sNum_MovStk, sCod_TipMov, sAccion, SCOD_PROVEEDOR, sNumGuia)

Set mRs = GetRecordset(cConnect, sSQL)

TxtGuia.Locked = False

If Not mRs.EOF Then
    'If mRs!Retorno = "1" Or mRs!Retorno = "2" Then
    '    TxtGuia.Locked = True
    'End If

    If mRs!Retorno = "2" Then
        TxtGuia.Enabled = True
    End If

    If mRs!Retorno = "3" Then
        TxtGuia.Enabled = False
    End If

    MovIngresoTintoPropiaEditable = mRs!Retorno

End If
mRs.Close
Set mRs = Nothing
Exit Function
errx:
    ErrorHandler err, "MovIngresoTintoPropiaEditable"
End Function

Private Function ShowGuiasTintoPropia() As Boolean
On Error GoTo errx
Set frmBusqGeneral.oParent = Me
frmBusqGeneral.sQuery = "exec LG_MUESTRA_GUIAS_POR_RECIBIR_TINTO"
frmBusqGeneral.Cargar_Datos
frmBusqGeneral.Show vbModal
Set frmBusqGeneral = Nothing

If Codigo <> "" Then
    Me.sCod_AlmacenOrigen = Codigo
    Me.sNum_MovStkOrigen = Descripcion
    ShowGuiasTintoPropia = True
Else
    Me.sCod_AlmacenOrigen = ""
    Me.sNum_MovStkOrigen = ""
    Me.TxtGuia.Text = ""
End If
Exit Function
errx:
    ErrorHandler err, "ShowGuiasTintoPropia"
End Function

Sub Deshabilita()
Me.TxtCod_CenCosto.Enabled = False
Me.TxtDes_CenCosto.Enabled = False
Me.TxtCod_Cliente.Enabled = False
Me.TxtNom_cliente.Enabled = False

Me.CmbOrdComp.Enabled = False
Txtproveedor.Enabled = False
TxtDetalle.Enabled = False
TxtCod_TipMov.Enabled = False
TxtDes_TipMov.Enabled = False

Me.DtFechaMov.Enabled = False
Me.TxtObservaciones.Enabled = False
TxtOrdPro.Enabled = False
txtNum_SecOrd.Enabled = False
Command1.Enabled = False
TxtGuia.Enabled = False
txtParteSalida.Enabled = False
txtGlosa_Hilado.Enabled = False
'CmdOC.Enabled = False
End Sub

Sub Habilita()
Dim vAux As Variant, sCod_TipAccion As String

If Accion = "I" Then
    Command1.Enabled = True
    TxtOrdPro.Enabled = True
    txtNum_SecOrd.Enabled = True
    Me.TxtCod_CenCosto.Enabled = True
    Me.TxtDes_CenCosto.Enabled = True
    Me.TxtCod_Cliente.Enabled = True
    Me.TxtNom_cliente.Enabled = True
    Me.CmbOrdComp.Enabled = True
    Txtproveedor.Enabled = True
    TxtDetalle.Enabled = True
    Me.TxtCod_TipMov.Enabled = True
    Me.TxtDes_TipMov.Enabled = True
    Me.DtFechaMov.Enabled = True
    Me.txtGlosa_Hilado.Enabled = True
    Me.TxtGuia.Locked = False
End If

Me.TxtObservaciones.Enabled = True

strSQL = "SELECT Tip_Accion FROM lg_tiposmov WHERE Cod_TipMov = '" & Trim(TxtCod_TipMov.Text) & "'"
sCod_TipAccion = Trim(DevuelveCampo(strSQL, cConnect))
If sCod_TipAccion <> "E" Then
    TxtCod_CenCosto.Enabled = True
    TxtDes_CenCosto.Enabled = True
End If
TxtGuia.Enabled = True

strSQL = "SELECT Cod_ClaMov FROM lg_tiposmov WHERE Cod_TipMov = '" & Trim(TxtCod_TipMov.Text) & "'"
If Trim(DevuelveCampo(strSQL, cConnect)) = "S" And sCod_TipAccion = "E" Then
    TxtGuia.Enabled = False
End If
'CmdOC.Enabled = True

'strSQL = "SELECT Tip_Accion FROM lg_tiposmov WHERE Cod_TipMov = '" & Right(Me.CmbTipMov.Text, 3) & "'"
'If Trim(DevuelveCampo(strSQL, cConnect)) <> "E" Then
'    CmbCentCosto.Enabled = True
'End If

strSQL = "SELECT ISNULL(Cod_TipOrdPro,'')  FROM lg_tiposmov WHERE Cod_TipMov = '" & Trim(TxtCod_TipMov.Text) & "'"
If Trim(UCase(DevuelveCampo(strSQL, cConnect))) = "CF" Then
    Txtproveedor.Enabled = True
    TxtDetalle.Enabled = True
    Command1.Enabled = True
    TxtCod_CenCosto.Enabled = True
    TxtDes_CenCosto.Enabled = True
End If

If Accion = "U" Then
    MovIngresoTintoPropiaEditable vCod_Almacen, Num_MovStk, Trim(TxtCod_TipMov), Accion, Txtproveedor.Text, TxtGuia.Text
End If

If Accion <> "I" Then
  sCod_AlmacenOrigen = Me.sCod_AlmacenOrigen
  sNum_MovStkOrigen = Me.sNum_MovStkOrigen
End If
End Sub

Sub Limpia()
Me.TxtCod_CenCosto.Text = ""
Me.TxtDes_CenCosto.Text = ""
Me.TxtCod_Cliente.Text = ""
Me.TxtNom_cliente.Text = ""
Me.CmbOrdComp.ListIndex = -1
Txtproveedor = ""
TxtDetalle = ""
Me.TxtCod_TipMov.Text = ""
Me.TxtDes_TipMov.Text = ""
Me.DtFechaMov.Value = Date
Me.TxtObservaciones = ""
TxtOrdPro = ""
txtNum_SecOrd.Text = ""
TxtGuia = ""
txtGlosa_Hilado.Text = ""
End Sub

'Sub LlenarCombos()
''LlenaCombo CmbAlmacen, "Select Nom_Almacen+space(100)+Cod_Almacen from lg_almacen order by 1", cCONNECT
'LlenaCombo CmbAlmacen, "Select a.Nom_Almacen+space(100)+ a.Cod_Almacen from lg_almacen a, lg_segalm b  where a.cod_almacen=b.cod_almacen and b.cod_usuario='" & vusu & "' order by 1", cConnect
'LlenaCombo Me.CmbOrdComp, "select Ser_OrdComp + rtrim(Cod_OrdComp)+' - '+isnull(rtrim(cod_grupo),'') from Lg_OrdComp order by 1", cConnect
''LlenaCombo Me.CmbOrdComp, "select Cod_OrdComp+space(100)+Ser_OrdComp from Lg_OrdComp order by 1", cCONNECT
'End Sub

Public Sub BuscaTipoMov(opcion As Integer)
Dim rstAux As ADODB.Recordset
On Error GoTo Fin
Dim iCol As Long
    
    strSQL = "SELECT Cod_TipMov as Codigo, Des_TipMov as Descripcion FROM Lg_TiposMov WHERE Cod_TipMov in (select Cod_TipMov from lg_tipmovialm where Cod_Almacen='" & vCod_Almacen & "') AND "
    
    Select Case opcion
    Case 1: strSQL = strSQL & "cod_tipmov like '%" & Trim(TxtCod_TipMov.Text) & "%' order by Des_tipmov"
    Case 2: strSQL = strSQL & "des_tipmov like '%" & Trim(TxtDes_TipMov.Text) & "%' order by des_tipmov"
    End Select
    
    With frmBusqGeneral
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        .Caption = "Seleccionar Tipo Movimiento"
        Codigo = ".."
        Set rstAux = .gexList.ADORecordset
        
        .gexList.Columns("Codigo").Width = 700
        .gexList.Columns("Descripcion").Width = 5000
        
        If rstAux.RecordCount = 1 Then
            Codigo = rstAux!Codigo
            Descripcion = rstAux!Descripcion
        Else
            If rstAux.RecordCount > 1 Then
                .Show vbModal
            End If
        End If
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            TxtCod_TipMov = Codigo '.gexList.Value(.gexList.Columns("codigo").Index) 'Trim(rstAux!Codigo)
            TxtDes_TipMov = Descripcion '.gexList.Value(.gexList.Columns("Descripcion").Index) 'Trim(rstAux!Descripcion)
            DtFechaMov.SetFocus
        End If
    End With
    Codigo = "": Descripcion = ""
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    Num_Conos
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Busca Tipo Movimiento (" & opcion & ")"
End Sub

Public Sub Num_Conos()

    Dim valor As Integer
    valor = DevuelveCampo("select dbo.LG_VALIDA_SI_ES_TIPMOV_AVIOS_OP('" & vCod_Almacen & "','" & TxtCod_TipMov & "')", cConnect)
    

    If valor = 1 Then
    lblConos.Visible = True
    txtNumConosHilosCoser.Visible = True
    Else
    txtNumConosHilosCoser.Text = 0
    lblConos.Visible = False
    txtNumConosHilosCoser.Visible = False
    End If
    
'    If valor = 0 Then
'    lblConos.Visible = True
'    txtNumConosHilosCoser.Visible = True
'    End If
End Sub


Private Sub Command1_Click()
Set frmBusqGeneral.oParent = Me
frmBusqGeneral.sQuery = "select Cod_Proveedor as Codigo ,Des_Proveedor as Nombre from lg_proveedor order by 2"
frmBusqGeneral.Cargar_Datos
frmBusqGeneral.Show 1

Me.Txtproveedor = Codigo
TxtDetalle = Descripcion
If TxtOrdPro.Enabled Then
    Me.TxtOrdPro.SetFocus
End If

End Sub

Public Sub BuscaCliente(opcion As Integer)
Dim rstAux As ADODB.Recordset
On Error GoTo Fin
Dim iCol As Long
    
    strSQL = "SELECT abr_cliente as Codigo, nom_cliente as Descripcion FROM tg_cliente WHERE "
    
    Select Case opcion
    Case 1: strSQL = strSQL & "ABR_cliente like '%" & Trim(TxtCod_Cliente.Text) & "%' order by ABR_cliente"
    Case 2: strSQL = strSQL & "nom_cliente like '%" & Trim(TxtNom_cliente.Text) & "%' order by nom_cliente"
    End Select
    
    With frmBusqGeneral
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        .Caption = "Seleccionar Cliente"
        Set rstAux = .gexList.ADORecordset
        
        .gexList.Columns("Codigo").Width = 700
        .gexList.Columns("Descripcion").Width = 5000
        
        If rstAux.RecordCount = 1 Then
            Codigo = rstAux!Codigo
            Descripcion = rstAux!Descripcion
        Else
            If rstAux.RecordCount > 1 Then
                .Show vbModal
            End If
        End If
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            TxtCod_Cliente.Text = Codigo '.gexList.Value(.gexList.Columns("codigo").Index)
            TxtNom_cliente = Descripcion '.gexList.Value(.gexList.Columns("Descripcion").Index)
            vCod_Cliente = DevuelveCampo("SELECT COD_CLIENTE FROM TG_CLIENTE WHERE ABR_CLIENTE='" & Trim(TxtCod_Cliente.Text) & "'", cConnect)
            TxtCod_CenCosto.SetFocus
        End If
    End With
    Codigo = "": Descripcion = ""
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
    "Busca Cliente (" & opcion & ")"
End Sub

Public Sub BuscaCentro_Costo(opcion As Integer)
Dim rstAux As ADODB.Recordset
On Error GoTo Fin
Dim iCol As Long
    
    strSQL = "SELECT Cod_CenCost as Codigo, Des_CenCost as Descripcion FROM tg_cencosto WHERE "
    
    Select Case opcion
    Case 1: strSQL = strSQL & "Cod_CenCost like '%" & Trim(TxtCod_CenCosto.Text) & "%' order by Cod_CenCost"
    Case 2: strSQL = strSQL & "Des_CenCost like '%" & Trim(TxtDes_CenCosto.Text) & "%' order by Des_CenCost"
    End Select
    
    With frmBusqGeneral
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        .Caption = "Seleccionar Centro Costo"
        Codigo = ".."
        Set rstAux = .gexList.ADORecordset
        
        .gexList.Columns("Codigo").Width = 700
        .gexList.Columns("Descripcion").Width = 5000
        
        If rstAux.RecordCount = 1 Then
            Codigo = rstAux!Codigo
            Descripcion = rstAux!Descripcion
        Else
            If rstAux.RecordCount > 1 Then
                .Show vbModal
            End If
        End If
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            TxtCod_CenCosto = Codigo '.gexList.Value(.gexList.Columns("codigo").Index)
            TxtDes_CenCosto = Descripcion '.gexList.Value(.gexList.Columns("Descripcion").Index)
            'DtFechaMov.SetFocus
        End If
    End With
    Codigo = "": Descripcion = ""
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
    "Busca Cliente (" & opcion & ")"
End Sub

Private Sub TxtTip_Trabajador_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BUSCATRABAJADOR(1)
End If
End Sub

Sub BUSCATRABAJADOR(Tipo As Integer)
Dim oTipo As New frmBusqGeneral5
Dim Rs As New ADODB.Recordset
Dim sfabrica As String
'sfabrica = "002"

Set oTipo.oParent = Me

oTipo.sQuery = "select a.tip_trabajador as Tipo, a.cod_trabajador as Codigo, (LTRIM(RTRIM(apellido_paterno))  +  ' ' +  LTRIM(RTRIM(apellido_materno)) + ' ' + LTRIM(RTRIM(Nombre_trabajador))) as Nombre from tg_operario A, Tg_Operario_Hialpesa_Caracteristica b where a.cod_fabrica = b.cod_fabrica and a.tip_trabajador = b.tip_trabajador and a.cod_trabajador = b.cod_trabajador and b.cod_caracteristica_operario in ('001') "
If Tipo = 1 Then
    oTipo.sQuery = oTipo.sQuery & "and a.tip_trabajador ='" & TxtTip_Trabajador.Text & "' and a.cod_fabrica='" & Cod_Fabrica & "' "
ElseIf Tipo = 2 Then
    oTipo.sQuery = oTipo.sQuery & "and a.tip_trabajador like '%" & TxtTip_Trabajador.Text & "%' and a.cod_trabajador like '%" & TxtCod_Trabajador.Text & "%' and a.cod_fabrica='" & Cod_Fabrica & "' "
Else
    oTipo.sQuery = oTipo.sQuery & "and apellido_paterno  +  apellido_materno + Nombre_trabajador like '%" & Trim(TxtNom_Trabajador.Text) & "%' and a.cod_fabrica='" & Cod_Fabrica & "' "
End If

oTipo.Caption = "Buscar Trabajador"
oTipo.Cargar_Datos

oTipo.gexLista.Columns("Tipo").Width = 600
oTipo.gexLista.Columns("Codigo").Width = 1000
oTipo.gexLista.Columns("nombre").Width = 5000

If oTipo.gexLista.RowCount > 1 Then
    oTipo.Show vbModal
Else
    Codigo = oTipo.gexLista.Value(oTipo.gexLista.Columns("Tipo").Index)
    Descripcion = oTipo.gexLista.Value(oTipo.gexLista.Columns("codigo").Index)
    TipoAdd = oTipo.gexLista.Value(oTipo.gexLista.Columns("nombre").Index)
End If

If Trim(Codigo) <> "" Then
    TxtTip_Trabajador.Text = Codigo
    TxtCod_Trabajador.Text = Descripcion
    TxtNom_Trabajador.Text = TipoAdd
    Codigo = "": Descripcion = "": TipoAdd = ""
    FunctButt1.SetFocus
End If

Unload oTipo
Set oTipo = Nothing
Set Rs = Nothing
End Sub

