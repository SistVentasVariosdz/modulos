VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form FrmAddMovimAlmDet 
   Caption         =   "Mantenimiento Detalle Movimiento"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3817
      TabIndex        =   14
      Top             =   3315
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmAddMovimAlmDet.frx":0000
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
      Height          =   3225
      Left            =   0
      TabIndex        =   17
      Tag             =   "Detail"
      Top             =   0
      Width           =   10095
      Begin VB.TextBox TxtEstilo 
         Height          =   300
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1995
         Width           =   3030
      End
      Begin VB.TextBox TxtOP 
         Height          =   285
         Left            =   960
         MaxLength       =   5
         TabIndex        =   32
         Top             =   1995
         Width           =   735
      End
      Begin VB.TextBox TxtGlosa 
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
         Height          =   555
         Left            =   1560
         MaxLength       =   7
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   2490
         Visible         =   0   'False
         Width           =   5625
      End
      Begin VB.CommandButton cmdTransfMismoItem 
         Caption         =   "Transferir el Mismo Item a otro Almacén"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   7365
         TabIndex        =   29
         Top             =   1800
         Width           =   2670
      End
      Begin VB.TextBox TxtPeso 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         Left            =   4320
         TabIndex        =   11
         Text            =   "0"
         Top             =   1290
         Width           =   945
      End
      Begin VB.TextBox TxtCod_Maquina 
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
         Left            =   930
         MaxLength       =   7
         TabIndex        =   12
         Top             =   1630
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox TxtDes_Maquina 
         Height          =   315
         Left            =   1920
         TabIndex        =   13
         Top             =   1630
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.TextBox TxtCod_EstCli 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         Left            =   6360
         TabIndex        =   26
         Top             =   960
         Width           =   2625
      End
      Begin VB.TextBox TxtCod_Medida 
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
         Left            =   6360
         MaxLength       =   8
         TabIndex        =   6
         Top             =   630
         Width           =   825
      End
      Begin VB.TextBox TxtDes_Medida 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   7200
         TabIndex        =   7
         Top             =   630
         Width           =   2745
      End
      Begin VB.TextBox TxtDes_comb 
         Height          =   315
         Left            =   6960
         TabIndex        =   3
         Top             =   300
         Width           =   2985
      End
      Begin VB.TextBox TxtCod_Comb 
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
         Left            =   6375
         MaxLength       =   8
         TabIndex        =   2
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox Txtcod_Destino 
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
         Left            =   930
         MaxLength       =   7
         TabIndex        =   8
         Top             =   960
         Width           =   945
      End
      Begin VB.TextBox TxtDes_Destino 
         Height          =   315
         Left            =   1920
         TabIndex        =   9
         Top             =   960
         Width           =   3315
      End
      Begin VB.TextBox TxtCantidad 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         Left            =   930
         TabIndex        =   10
         Text            =   "0"
         Top             =   1290
         Width           =   945
      End
      Begin VB.TextBox TxtDetalle 
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         Top             =   630
         Width           =   3315
      End
      Begin VB.TextBox CmbColor 
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
         Left            =   945
         MaxLength       =   7
         TabIndex        =   4
         Top             =   630
         Width           =   945
      End
      Begin VB.TextBox TxtItem 
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
         Left            =   930
         MaxLength       =   8
         TabIndex        =   0
         Top             =   300
         Width           =   945
      End
      Begin VB.TextBox TxtDesitem 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   300
         Width           =   3345
      End
      Begin VB.TextBox TxtCodProv 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         Left            =   6360
         TabIndex        =   15
         Top             =   1290
         Width           =   945
      End
      Begin VB.CommandButton CmdTransferir 
         Caption         =   "Transferir a otro Item"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   7365
         TabIndex        =   16
         Top             =   1320
         Width           =   2670
      End
      Begin VB.Label lblop 
         Caption         =   "OP"
         Height          =   240
         Index           =   6
         Left            =   270
         TabIndex        =   34
         Top             =   2040
         Width           =   345
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Glosa Combinacion"
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
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   2550
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Peso (kgs.):"
         Height          =   195
         Index           =   5
         Left            =   3360
         TabIndex        =   28
         Top             =   1395
         Width           =   840
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Maquina"
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
         Index           =   0
         Left            =   240
         TabIndex        =   27
         Top             =   1695
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Medida:"
         Height          =   195
         Index           =   4
         Left            =   5370
         TabIndex        =   25
         Top             =   720
         Width           =   570
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Color"
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
         Left            =   240
         TabIndex        =   24
         Tag             =   "Hilado :"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estilo:"
         Height          =   195
         Index           =   3
         Left            =   5400
         TabIndex        =   23
         Top             =   1035
         Width           =   420
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
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
         Left            =   240
         TabIndex        =   22
         Tag             =   "Hilado :"
         Top             =   1035
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   21
         Top             =   405
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Combinacion:"
         Height          =   195
         Index           =   0
         Left            =   5400
         TabIndex        =   20
         Top             =   405
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   19
         Top             =   1365
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cod Prov.:"
         Height          =   195
         Index           =   2
         Left            =   5400
         TabIndex        =   18
         Top             =   1365
         Width           =   750
      End
   End
End
Attribute VB_Name = "FrmAddMovimAlmDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public Codigo As String, Descripcion As String
Public Estado As String
Public var_tipo As String
Public Cod_Almacen As String, Num_MovStk As String, Ser_OrdComp As String, Cod_OrdComp As String, cod_tipmov As String
Dim itemtransferir  As String, combtransferir  As String, colortransferir  As String, destinotransferir  As String, _
    estilotransferir  As String, medidatransferir  As String, cod_provtransferir As String
Public varNum_SecOrd As String, vFlg_Almacen_Tejeduria As String
Public Num_Secuencia
Public Sec_OrdComp As String, Cant As Double, item As String, Color As String, CombinacionX As String, TallaX As String, varTallaProv As String
Dim Rs As New ADODB.Recordset
Public Paso As Boolean, vOk As Boolean
Public oParent As Object
Public sflg_adicionales As String

Private Sub CmdTransferir_Click()
    Load FrmTranferirA
    
    FrmTranferirA.itemAnt = Me.TxtItem.Text
    FrmTranferirA.combAnt = Trim(TxtCod_Comb.Text)
    FrmTranferirA.colorAnt = Trim(Me.CmbColor.Text)
    FrmTranferirA.estiloAnt = Trim(TxtCod_EstCli.Text)
    FrmTranferirA.medidaAnt = Trim(TxtCod_Medida.Text)
    FrmTranferirA.destinoAnt = Trim(Txtcod_Destino)
    FrmTranferirA.cod_provAnt = Me.TxtCodProv.Text
    
    FrmTranferirA.var_tipo = Me.var_tipo
    FrmTranferirA.Cod_Almacen = Me.Cod_Almacen
    FrmTranferirA.Num_MovStk = Me.Num_MovStk
    FrmTranferirA.Ser_OrdComp = Me.Ser_OrdComp
    FrmTranferirA.Cod_OrdComp = Me.Cod_OrdComp
    FrmTranferirA.cod_tipmov = Me.cod_tipmov
    LlenaCombo FrmTranferirA.CmbDestino, "select des_destino+space(100)+cod_destino from tg_destino order by 1", cConnect
    
    If UCase(vFlg_Almacen_Tejeduria) = "S" Then
        FrmTranferirA.CmbDestino.Visible = False
        FrmTranferirA.TxtItem.Text = Me.TxtItem
        FrmTranferirA.TxtDesitem.Text = Me.TxtDesitem
        FrmTranferirA.CmbColor.Visible = False
        FrmTranferirA.TxtDetalle.Visible = False
        FrmTranferirA.CmbTalla.Visible = False
        FrmTranferirA.CmbEstilo.Visible = False
        FrmTranferirA.TxtCodProv.Visible = False
        FrmTranferirA.Label1(4).Visible = False
        FrmTranferirA.Label1(3).Visible = False
        FrmTranferirA.Label1(2).Visible = False
        FrmTranferirA.Etiqueta(5).Visible = False
        FrmTranferirA.Etiqueta(3).Visible = False
        FrmTranferirA.Command1.Visible = False
    Else
        LlenaCombo FrmTranferirA.CmbEstilo, "select rtrim(cod_estcli)+'  -  '+des_estcli+space(100)+cod_estcli from tg_estcli order by 1", cConnect
    End If
    
    FrmTranferirA.Show vbModal
    
    itemtransferir = FrmTranferirA.item
    combtransferir = FrmTranferirA.comb
    colortransferir = FrmTranferirA.Color
    destinotransferir = FrmTranferirA.Destino
    estilotransferir = FrmTranferirA.Estilo
    medidatransferir = FrmTranferirA.medida
    cod_provtransferir = FrmTranferirA.cod_prov
    
    Set FrmTranferirA = Nothing
End Sub

Sub Habilita()
CmbColor.Enabled = True
TxtDetalle.Enabled = True
Me.TxtCod_Comb.Enabled = True
Me.TxtDes_comb.Enabled = True
Me.Txtcod_Destino.Enabled = True
Me.TxtDes_Destino.Enabled = True
TxtCod_EstCli.Enabled = True
TxtItem.Enabled = True
TxtDesitem.Enabled = True
Me.TxtCod_Medida.Enabled = True
Me.TxtCantidad.Enabled = True
Me.TxtCodProv.Enabled = True
'Command1.Enabled = True
'Command2.Enabled = True
Me.TxtCod_Maquina.Enabled = True
Me.TxtDes_Maquina.Enabled = True
Me.TxtPeso.Enabled = True
Me.TxtGlosa.Enabled = True
End Sub

Sub Deshabilita()
'Command1.Enabled = False
TxtDetalle.Enabled = False
Me.CmbColor.Enabled = False
Me.TxtCod_Comb.Enabled = False
Me.TxtDes_comb.Enabled = False
Me.Txtcod_Destino.Enabled = False
Me.TxtDes_Destino.Enabled = False
Me.TxtCod_EstCli.Enabled = False
TxtItem.Enabled = False
TxtDesitem.Enabled = False
'Command2.Enabled = False
Me.TxtCod_Medida.Enabled = False
Me.TxtCantidad.Enabled = False
Me.TxtCodProv.Enabled = False
Me.TxtCod_Maquina.Enabled = False
Me.TxtDes_Maquina.Enabled = False
Me.TxtPeso.Enabled = False
Me.TxtGlosa.Enabled = False
Me.TxtOP.Enabled = False
Me.TxtEstilo.Enabled = False
End Sub

Sub Limpia()
TxtDetalle = ""
CmbColor = ""
Me.TxtCod_Comb.Text = ""
Me.TxtDes_comb.Text = ""
Me.Txtcod_Destino.Text = ""
Me.TxtDes_Destino.Text = ""
Me.TxtCod_EstCli.Text = ""
TxtItem = ""
TxtDesitem = ""
Me.TxtCod_Medida.Text = ""
Me.TxtCantidad = "0"
Me.TxtCodProv = ""
Me.TxtCod_Maquina.Text = ""
Me.TxtDes_Maquina.Text = ""
Me.TxtPeso.Text = "0"
Me.TxtGlosa.Text = ""
End Sub


Private Sub CmbColor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        'CmbColor = DevuelveCampo("select dbo.uf_devuelvecodigo(6," & CmbColor & ")", cConnect)
'        If ExisteCampo("Cod_color", "lb_color", CmbColor, cConnect, True) Then
'            TxtDetalle = DevuelveCampo("Select Des_color from lb_color where Cod_color='" & CmbColor & "'", cConnect)
'        Else
'            MsgBox "El codigo no existe", vbInformation
'        End If
    Call busca_color(1)
End If

End Sub

'Public Sub CmbCombinacion_DropDown()
'LlenaCombo Me.CmbCombinacion, "select Des_Comb + space(100) + Cod_Comb from lg_itemcomb where cod_item='" & Me.TxtItem & "'", cConnect
'End Sub


Sub busca_color(Tipo As Integer)
Codigo = ""
Descripcion = ""
Set frmBusqGeneral.oParent = Me
If Tipo = 1 Then
    frmBusqGeneral.sQuery = "select Cod_color as Codigo ,Des_color as Nombre from lb_color where cod_color like '%" & Trim(CmbColor.Text) & "%' order by 1"
Else
    frmBusqGeneral.sQuery = "select Cod_color as Codigo ,Des_color as Nombre from lb_color where Des_color like '%" & Trim(TxtDetalle.Text) & "%' order by 2"
End If
frmBusqGeneral.Cargar_Datos
frmBusqGeneral.Show 1

CmbColor = Codigo
TxtDetalle = Descripcion
TxtCod_Medida.SetFocus

End Sub

Private Sub Command2_Click()
Set frmBusqGeneral.oParent = Me

frmBusqGeneral.sQuery = "select Cod_item as Codigo ,Des_item as Nombre from lg_item where Des_item<>'' order by 2"
frmBusqGeneral.Cargar_Datos
frmBusqGeneral.Show 1

TxtItem = Codigo
TxtDesitem = Descripcion
End Sub

Private Sub cmdTransfMismoItem_Click()
    itemtransferir = TxtItem.Text   'FrmTranferirA.item
    combtransferir = TxtCod_Comb.Text  ' FrmTranferirA.comb
    colortransferir = CmbColor.Text '  FrmTranferirA.Color
    destinotransferir = Txtcod_Destino.Text  ' FrmTranferirA.Destino
    estilotransferir = TxtCod_EstCli.Text   ' FrmTranferirA.Estilo
    medidatransferir = TxtCod_Medida.Text    'FrmTranferirA.medida
    cod_provtransferir = TxtCodProv.Text   'FrmTranferirA.cod_prov
End Sub

Private Sub Form_Load()
If sflg_adicionales = "*" Then
    TxtOP.Visible = True
    lblop(6).Visible = True
    TxtEstilo.Visible = True

Else

    TxtOP.Visible = False
    lblop(6).Visible = False
    TxtEstilo.Visible = False
End If

itemtransferir = ""
combtransferir = ""
colortransferir = ""
destinotransferir = ""
estilotransferir = ""
medidatransferir = ""
cod_provtransferir = ""
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    Call Grabar
Case "CANCELAR"
    Unload Me
End Select
End Sub

Private Sub TxtCod_Comb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Comb(1)
End If
End Sub

Private Sub Txtcod_Destino_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Destino(1)
End If
End Sub

Private Sub TxtCod_EstCli_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Estilo
End If
End Sub

Private Sub TxtCod_Maquina_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Maquina
End If
End Sub

Private Sub TxtCod_Medida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Medida(1)
End If
End Sub

Private Sub TxtDes_comb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Comb(2)
End If
End Sub

Private Sub TxtDes_Destino_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Destino(2)
End If
End Sub

Private Sub TxtDes_Maquina_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Maquina
End If
End Sub

Private Sub TxtDetalle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'    If DevuelveCampo("select count(*) from lb_color where Des_color like '" & TxtDetalle & "%'", cConnect) > 0 Then
'        frmBusqGeneral.sQuery = "select Cod_color as Codigo ,Des_color as Nombre from lb_color where Des_color like '" & TxtDetalle & "%' "
'        frmBusqGeneral.CARGAR_DATOS
'        frmBusqGeneral.Show 1
'        CmbColor = Codigo
'        TxtDetalle = Descripcion
'
'    Else
'        CmbColor = DevuelveCampo("Select Cod_color from lb_color where Des_color='" & TxtDetalle & "'", cConnect)
'    End If
    Call busca_color(2)
End If
End Sub

Private Sub TxtItem_KeyPress(KeyAscii As Integer)
On Error GoTo hand
Dim Temp As String
Dim strSQL As String

If KeyAscii = 13 Then
    Codigo = ""
    Descripcion = ""
    If Len(TxtItem.Text) < 3 Then
         Set frmBusqGeneral.oParent = Me
         frmBusqGeneral.sQuery = "select Cod_Item AS Codigo,des_item as Descripcion from lg_item where Cod_item like '" & TxtItem & "%' order by cod_item"
         frmBusqGeneral.Cargar_Datos
         frmBusqGeneral.Show 1
         TxtDesitem = Descripcion
         TxtItem = Codigo
         Temp = TxtItem
         TxtCod_Comb.SetFocus
         If Codigo <> "" Then
            GoTo otro
         Else
            Exit Sub
         End If
    End If

    If Len(TxtItem.Text) > 2 Then Temp = Trim(DevuelveCampo("Select dbo.uf_devuelvecodigo(6," & IIf(Trim(TxtItem) = "", 0, Mid(TxtItem, 3)) & ")", cConnect))

    Temp = Left(TxtItem, 2) & Temp
        If DevuelveCampo("select count(*) from lg_item where cod_item ='" & Temp & "'", cConnect) > 0 Then
            Me.TxtDesitem = DevuelveCampo("select Des_Item from lg_item where cod_item ='" & Temp & "'", cConnect)
            TxtItem = Temp
            TxtCod_Comb.SetFocus
        Else
            MsgBox "Codigo no existe", vbInformation
            Me.TxtDesitem = ""
            Exit Sub
        End If
otro:
    strSQL = "EXEC UP_VERIFICA_MOV_AVIO_SAL '" & Me.cod_tipmov & "'"
    If Val(DevuelveCampo(strSQL, cConnect)) = 1 Then
    
        varTallaProv = ""
    
        Load frmListaStocksAvios
        frmListaStocksAvios.varCOD_ALMACEN = Me.Cod_Almacen
        frmListaStocksAvios.varCod_Item = Me.TxtItem.Text
        frmListaStocksAvios.Caption = "Stocks del : " & Me.TxtItem & " - " & Me.TxtDesitem
        Set frmListaStocksAvios.oParent = Me
        frmListaStocksAvios.CARGA_GRID
        frmListaStocksAvios.Show 1
        
        Set frmListaStocksAvios = Nothing
    
        TxtCantidad.SetFocus
    End If
End If
Exit Sub
hand:
End Sub

Private Sub TxtCantidad_GotFocus()
    TxtCantidad.SelStart = 0
    TxtCantidad.SelLength = Len(TxtCantidad.Text)
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    SoloNumeros TxtCantidad, KeyAscii, True, 3, 9
End If
End Sub



Private Sub TxtCodProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If var_tipo = "E" Then
        MUESTRA_AYUDA
    ElseIf var_tipo = "S" Then
        MUESTRA_AYUDA_SALIDA
    End If
End If
End Sub

Private Sub TxtDesitem_KeyPress(KeyAscii As Integer)
On Error GoTo hand
Dim Temp As String
If KeyAscii = 13 Then
    
       If DevuelveCampo("select count(*) from lg_item where des_item like '%" & TxtDesitem & "%'", cConnect) > 1 Then
       
            If Len(Trim(Me.TxtDesitem.Text)) < 3 Then
                MsgBox "La consulta requiere como mínimo 3 caracteres. Sirvase verificar", vbInformation, "Mensaje"
                TxtDesitem.SetFocus
                Exit Sub
            End If
       
            Set frmBusqGeneral.oParent = Me
            frmBusqGeneral.sQuery = "select Cod_Item AS Codigo,des_item as Descripcion from lg_item where des_item like '%" & TxtDesitem & "%'"
            frmBusqGeneral.Cargar_Datos
            frmBusqGeneral.Show 1
            TxtDesitem = Descripcion
            TxtItem = Codigo
            Temp = TxtItem
            
            
        Else
            TxtItem = DevuelveCampo("select cod_item from lg_item where des_item like'" & TxtDesitem & "%'", cConnect)

            Temp = TxtItem
        End If
                
End If
Exit Sub
hand:
End Sub

Sub MUESTRA_AYUDA()
Set frmBusqGeneral2.oParent = Me
Codigo = ""
Descripcion = ""
frmBusqGeneral2.sQuery = "MUESTRA_AYUDA_PROV_ENTRADA '" & Me.Cod_Almacen & "','" & _
                            Me.Num_MovStk & "','" & Me.Ser_OrdComp & "','" & Me.Cod_OrdComp & "','" & _
                            Me.TxtItem & "','" & Me.TxtCod_Comb & "','" & Me.CmbColor & "','" & Trim(TxtCod_Medida) & "','" & _
                            Trim(TxtCod_EstCli) & "','" & Trim(Txtcod_Destino.Text) & "'"

frmBusqGeneral2.Cargar_Datos
frmBusqGeneral2.Show 1
If Codigo <> "" Then
    TxtCodProv = Trim(Codigo)
End If
End Sub

Function VALIDA_PROV() As Boolean
    Set Rs = Nothing
    Rs.CursorLocation = adUseClient
    Rs.Open "VALIDA_PROV '" & Me.Cod_Almacen & "','" & _
                            Me.Num_MovStk & "','" & Me.Ser_OrdComp & "','" & Me.Cod_OrdComp & "','" & _
                            Me.TxtItem & "','" & Me.TxtCod_Comb & "','" & Me.CmbColor & "','" & Trim(TxtCod_Medida) & "','" & _
                            Trim(TxtCod_EstCli) & "','" & Trim(Txtcod_Destino) & "','" & TxtCodProv & "'", cConnect, 3, 3
    
    If Rs.RecordCount <= 0 Then
        'MsgBox "Cod. Prov. no valido", vbInformation
        VALIDA_PROV = False
    Else
        VALIDA_PROV = True
    End If
End Function

Function VALIDA_PROV_SALIDA() As Boolean
    Set Rs = Nothing
    Rs.CursorLocation = adUseClient
    Rs.Open "VALIDA_PROV_SALIDAS '" & Cod_Almacen & "','" & _
                            Me.TxtItem & "','" & Me.TxtCod_Comb & "','" & Me.CmbColor & "','" & Trim(TxtCod_Medida) & "','" & _
                            Trim(TxtCod_EstCli) & "','" & Trim(Txtcod_Destino) & "','" & TxtCodProv & "'", cConnect, 3, 3
    If Rs.RecordCount <= 0 Then
        'MsgBox "Cod. Prov. no valido", vbInformation
        VALIDA_PROV_SALIDA = False
    Else
        VALIDA_PROV_SALIDA = True
    End If
End Function

Sub MUESTRA_AYUDA_SALIDA()
Set frmBusqGeneral2.oParent = Me
        Codigo = ""
        Descripcion = ""

         frmBusqGeneral2.sQuery = "MUESTRA_AYUDA_PROV_SALIDAS '" & Cod_Almacen & "','" & _
                            Me.TxtItem & "','" & Me.TxtCod_Comb & "','" & Me.CmbColor & "','" & Trim(TxtCod_Medida) & "','" & _
                            Trim(TxtCod_EstCli) & "','" & Trim(Txtcod_Destino) & "'"
                                    
         'frmBusqGeneral2.sQuery = "select Cod_Prov from lg_ordcompitem where COD_ITEM = '" & TxtItem & "' AND COD_COLOR='" & CmbColor & "' AND COD_COMB='" & CmbCombinacion & "' AND COD_TALLA='" & CmbTalla & "' AND COD_DESTINO='" & CmbDestino & "' AND COD_ESTCLI='" & CmbEstilo & "' and cod_prov<>''"
         frmBusqGeneral2.Cargar_Datos
         frmBusqGeneral2.Show 1
         If Codigo <> "" Then
            TxtCodProv = Trim(Codigo)
        End If
End Sub

Function Valida_Data_Transferencia() As Boolean
If CmdTransferir.Enabled = True Then
    If itemtransferir & combtransferir & colortransferir & destinotransferir & estilotransferir & medidatransferir & cod_provtransferir = "" Then
        Valida_Data_Transferencia = False
    Else
        Valida_Data_Transferencia = True
    End If
Else
    Valida_Data_Transferencia = True
End If
End Function

Sub Grabar()
On Error GoTo errGrabar
Dim mRs As ADODB.Recordset
Dim sFlg_norealizado As String

If Me.var_tipo = "E" Then
    If Trim(TxtCodProv) <> "" Then
        sFlg_norealizado = RTrim(FixNulos(DevuelveCampo("SELECT flg_norealizado from lg_tiposmov WHERE cod_tipmov = '" & cod_tipmov & "'", cConnect), vbString))
        If sFlg_norealizado = "" Then
            Set mRs = GetRecordset(cConnect, "SELECT COD_PROV FROM LG_ORDCOMPITEM WHERE SER_ORDCOMP ='" & Ser_OrdComp & "' AND COD_ORDCOMP = '" & Cod_OrdComp & "' and COD_PROV = '" & TxtCodProv & "'")
'            If mRs.EOF Then
'                MsgBox "Cod. Prov. no valido", vbInformation
'                Exit Sub
'            End If
            mRs.Close
            Set mRs = Nothing
        Else
            If Not ExisteCampo("cod_prov", "LG_STOCKSITEM_PROV", TxtCodProv, cConnect) Then
                MsgBox "Cod. Prov. no valido", vbInformation
                Exit Sub
            End If
        End If
        'If Not ExisteCampo("cod_prov", "LG_ORDCOMPITEM", TxtCodProv, cConnect) Then
        '    MsgBox "Cod. Prov. no valido", vbInformation
        '    Exit Sub
        'End If
    End If
ElseIf Me.var_tipo = "S" Then
    If Trim(TxtCodProv) <> "" Then
        If Not ExisteCampo("cod_prov", "LG_STOCKSITEM_PROV", TxtCodProv, cConnect) Then
            MsgBox "Cod. Prov. no valido", vbInformation
            Exit Sub
        Else
            If VALIDA_PROV_SALIDA = False Then
                MsgBox "Cod. Prov. no válido", vbInformation
                Exit Sub
            End If
        End If
    End If
End If

If Trim(TxtItem) = "" Then MsgBox "Debe seleccionar un item", vbInformation: Exit Sub
'If Estado = "NUEVO" Then
'    Datos "i", True
'Else
'    Datos "a", True
'End If

'If Reg.RecordCount > 0 Then
    
'End If
If Trim(TxtPeso.Text) = "" Then
    TxtPeso.Text = "0"
End If

If Estado = "U" Then
    strSQL = "UP_ACTUALIZA_STOCKS_ITEM '" & Cod_Almacen & "','" & Num_MovStk & "','" & _
                item & "','" & CombinacionX & "','" & _
                Trim(CmbColor) & "','" & TallaX & "','" & Trim(Txtcod_Destino) & "','" & _
                Trim(TxtCod_EstCli) & "','" & Num_Secuencia & "'," & Cant & "," & CDbl(TxtCantidad.Text) & ",'M','" & Sec_OrdComp & "','" & Me.varNum_SecOrd & "','" & vusu & "','" & TxtCodProv.Text & "','" & _
                combtransferir & "','" & colortransferir & "','" & medidatransferir & "','" & destinotransferir & "','" & estilotransferir & "','" & cod_provtransferir & "','S','" & itemtransferir & "','" & Trim(TxtCod_Maquina.Text) & "'," & CDbl(TxtPeso.Text) & ",'','" & Trim(TxtOP.Text) & "'"
ElseIf Estado = "D" Then
    strSQL = "UP_ACTUALIZA_STOCKS_ITEM '" & Cod_Almacen & "','" & Num_MovStk & "','" & _
                item & "','" & CombinacionX & "','" & _
                Trim(CmbColor) & "','" & TallaX & "','" & Trim(Txtcod_Destino) & "','" & _
                Trim(TxtCod_EstCli) & "','" & Num_Secuencia & "'," & TxtCantidad & ",0,'E','" & Sec_OrdComp & "','" & Me.varNum_SecOrd & "','" & vusu & "','" & TxtCodProv.Text & "','" & Trim(TxtCod_Maquina.Text) & "'"
ElseIf Estado = "I" Then
    If Me.varTallaProv <> "" Then
        TallaX = varTallaProv
    Else
        TallaX = Trim(TxtCod_Medida.Text)
    End If
    If Valida_Data_Transferencia = False Then
            MsgBox "No hay data de transferencia", vbInformation
            Exit Sub
    End If
       
      
    strSQL = "UP_ACTUALIZA_STOCKS_ITEM '" & Cod_Almacen & "','" & Num_MovStk & "','" & _
                TxtItem & "','" & Trim(TxtCod_Comb) & "','" & _
                Trim(CmbColor) & "','" & TallaX & "','" & Trim(Txtcod_Destino) & "','" & _
                TxtCod_EstCli.Text & "','" & Num_Secuencia & "',0," & TxtCantidad & ",'I','" & Sec_OrdComp & "','','" & vusu & "','" & TxtCodProv.Text & "','" & _
                combtransferir & "','" & colortransferir & "','" & medidatransferir & "','" & destinotransferir & "','" & estilotransferir & "','" & cod_provtransferir & "','S','" & _
                itemtransferir & "','" & Trim(TxtCod_Maquina.Text) & "'," & CDbl(TxtPeso.Text) & ",'" & Trim(TxtGlosa.Text) & "','" & Trim(TxtOP.Text) & "'"
    Me.varTallaProv = ""
End If
Call ExecuteSQL(cConnect, strSQL)
If Estado <> "I" Then
    Unload Me
Else
    oParent.Datos "v", False
    Limpia
    TxtItem.SetFocus
End If
Exit Sub
errGrabar:
    ErrorHandler err, "Grabar"
End Sub

Sub Busca_Destino(Tipo As Integer)
Codigo = ""
Descripcion = ""
Set frmBusqGeneral.oParent = Me
If Tipo = 1 Then
    frmBusqGeneral.sQuery = "select Cod_Destino as Codigo ,Des_Destino as Nombre from tg_destino where cod_destino like '%" & Trim(Txtcod_Destino.Text) & "%' order by 1"
Else
    frmBusqGeneral.sQuery = "select Cod_Destino as Codigo ,Des_Destino as Nombre from tg_destino where des_destino like '%" & Trim(TxtDes_Destino.Text) & "%' order by 2"
End If
frmBusqGeneral.Cargar_Datos
frmBusqGeneral.Show 1

Txtcod_Destino = Codigo
TxtDes_Destino = Descripcion
TxtCod_EstCli.SetFocus
End Sub

Sub Busca_Comb(Tipo As Integer)
Codigo = ""
Descripcion = ""
Set frmBusqGeneral.oParent = Me
If Tipo = 1 Then
    frmBusqGeneral.sQuery = "select Cod_Comb as Codigo ,Des_Comb as Nombre from lg_itemcomb where cod_item='" & Trim(TxtItem.Text) & "' and Cod_Comb like '%" & Trim(TxtCod_Comb.Text) & "%' order by 1"
Else
    frmBusqGeneral.sQuery = "select Cod_Comb as Codigo ,Des_Comb as Nombre from lg_itemcomb where cod_item='" & Trim(TxtItem.Text) & "' and Des_Comb like '%" & Trim(TxtDes_comb.Text) & "%' order by 2"
End If

frmBusqGeneral.Cargar_Datos
frmBusqGeneral.Show 1

TxtCod_Comb = Codigo
TxtDes_comb = Descripcion
SendKeys "{TAB}"
'CmbColor.SetFocus
End Sub

Sub Busca_Medida(Tipo As Integer)
Dim Tot As Integer
Tot = DevuelveCampo("Select count(*) from lg_itemmed where cod_item='" & TxtItem & "'", cConnect)

Codigo = ""
Descripcion = ""
Set frmBusqGeneral.oParent = Me
If Tot > 0 Then
    If Tipo = 1 Then
        frmBusqGeneral.sQuery = "select Cod_Medida as Codigo ,Descripcion as Nombre from lg_itemmed where cod_item='" & Trim(TxtItem.Text) & "' and Cod_Medida like '%" & Trim(TxtCod_Medida.Text) & "%' order by 1"
    End If
Else
    frmBusqGeneral.sQuery = "select Cod_talla as Codigo, Cod_talla from tg_talla where cod_talla like '" & Trim(TxtCod_Medida) & "%'"
End If

frmBusqGeneral.Cargar_Datos
frmBusqGeneral.Show 1
TxtCod_Medida = Codigo
TxtDes_Medida = Descripcion
Txtcod_Destino.SetFocus
End Sub

Sub Busca_Estilo()
Codigo = ""
Descripcion = ""
Set frmBusqGeneral.oParent = Me
frmBusqGeneral.sQuery = "select cod_estcli as Codigo ,des_estcli as Nombre from tg_estcli where cod_estcli='%" & Trim(TxtCod_EstCli.Text) & "%' order by 1"
frmBusqGeneral.Cargar_Datos
frmBusqGeneral.Show 1

TxtCod_EstCli = Codigo
TxtCantidad.SetFocus

End Sub

Sub Deshabilita_Tej()
Me.CmbColor.Enabled = False
Me.TxtDetalle.Enabled = False
Me.Txtcod_Destino.Enabled = False
Me.TxtDes_Destino.Enabled = False
Me.TxtCod_Medida.Enabled = False
Me.TxtCod_EstCli.Enabled = False
Me.TxtCodProv.Enabled = False
End Sub

Sub Busca_Maquina()
Codigo = ""
Descripcion = ""
Set frmBusqGeneral.oParent = Me
'If Tipo = 1 Then
frmBusqGeneral.sQuery = "TJ_MUESTRA_MAQUINAS_TEJEDURIA_PROPIAS"
'Else
'    frmBusqGeneral.sQuery = "select Cod_Destino as Codigo ,Des_Destino as Nombre from tg_destino where des_destino like '%" & Trim(TxtDes_Destino.Text) & "%' order by 2"
'End If
frmBusqGeneral.Cargar_Datos
frmBusqGeneral.gexList.Columns("Codigo").Width = 0
frmBusqGeneral.gexList.Columns("Descripcion").Width = 3200
frmBusqGeneral.Show 1

TxtCod_Maquina = Codigo
TxtDes_Maquina = Descripcion
SendKeys "{TAB}"
'TxtCod_EstCli.SetFocus
End Sub

Private Sub TxtOP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Dim sCod_Fabrica As String

        strSQL = "select cod_fabrica from tg_fabrica where abr_fabrica='HIA'"
        sCod_Fabrica = DevuelveCampo(strSQL, cConnect)

        TxtOP.Text = Format(Trim(TxtOP.Text), "00000")
        If DevuelveCampo("select count(*) from es_Ordpro where cod_fabrica='" & sCod_Fabrica & "' AND cod_ordpro = '" & TxtOP.Text & "'", cConnect) > 0 Then
            strSQL = "SELECT cod_fabrica FROM TG_FABRICA WHERE Abr_Fabrica = 'HIA'"
            Me.TxtEstilo.Text = DevuelveCampo("SELECT b.Des_EstPro FROM   ES_OrdPro  a , ES_EstPRo b WHERE  a.Cod_EstPro = b.Cod_EstPRo AND a.Cod_Fabrica= '" & DevuelveCampo(strSQL, cConnect) & "' AND a.Cod_OrdPro = '" & TxtOP.Text & "'", cConnect)
            Me.FunctButt1.SetFocus
        Else
            MsgBox "Codigo de " & TxtOP.Text & " no existe", vbInformation, Me.Caption
        End If
    End If
End Sub

Private Sub TxtPeso_GotFocus()
SelectionText TxtPeso
End Sub

Private Sub TxtPeso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtPeso, KeyAscii, True, 2)
End If
End Sub
