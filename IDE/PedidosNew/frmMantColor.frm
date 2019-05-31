VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Begin VB.Form frmMantColor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Colores"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Colours"
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   840
      TabIndex        =   12
      Top             =   6570
      Width           =   1935
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "frmMantColor.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Anterior"
         Top             =   105
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "frmMantColor.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Siguiente"
         Top             =   105
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   0
         Picture         =   "frmMantColor.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Primero"
         Top             =   105
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "frmMantColor.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ultimo"
         Top             =   105
         Width           =   495
      End
   End
   Begin VB.Frame Fralista 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   90
      TabIndex        =   10
      Tag             =   "List"
      Top             =   -15
      Width           =   7140
      Begin MSDataGridLib.DataGrid DGridlista 
         Height          =   2580
         Left            =   105
         TabIndex        =   11
         Top             =   300
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   4551
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "cod_colcli"
            Caption         =   "Code"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Nom_colcli"
            Caption         =   "Name"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "num_ordcolcli"
            Caption         =   "Order"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Por_AdicProd"
            Caption         =   "% Adicional "
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Pre_AdicProd"
            Caption         =   "Prend Adici."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Identificador_Color_Cliente"
            Caption         =   "Id. Color"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "NumLotePaking"
            Caption         =   "Nro P. List"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column06 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Fradetalle 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3510
      Left            =   90
      TabIndex        =   1
      Tag             =   "Detail"
      Top             =   3015
      Width           =   7140
      Begin VB.TextBox txtNroLotePack 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2145
         MaxLength       =   30
         TabIndex        =   30
         Top             =   2880
         Width           =   1785
      End
      Begin VB.CheckBox chktodos 
         Caption         =   "Cambios se aplican a todos los colores"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3600
         TabIndex        =   29
         Top             =   2640
         Width           =   3135
      End
      Begin VB.TextBox txtIdentificador_Color_Cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4410
         MaxLength       =   20
         TabIndex        =   27
         Top             =   2160
         Width           =   2145
      End
      Begin VB.TextBox txtPre_AdicProd 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2145
         MaxLength       =   60
         TabIndex        =   26
         Text            =   "0"
         Top             =   2520
         Width           =   1260
      End
      Begin VB.TextBox txtPor_AdicProd 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2145
         MaxLength       =   60
         TabIndex        =   25
         Text            =   "0.00"
         Top             =   2160
         Width           =   1260
      End
      Begin VB.CommandButton cmdBuscaColor 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   330
         Left            =   3735
         TabIndex        =   22
         Tag             =   "..."
         Top             =   1440
         Width           =   270
      End
      Begin VB.TextBox txtNomEstCli 
         Enabled         =   0   'False
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
         Left            =   2895
         TabIndex        =   19
         Top             =   1095
         Width           =   3045
      End
      Begin VB.TextBox txtIdEstCli 
         Enabled         =   0   'False
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
         Left            =   2145
         TabIndex        =   18
         Top             =   1095
         Width           =   750
      End
      Begin VB.TextBox txtidPO 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2145
         TabIndex        =   17
         Top             =   750
         Width           =   3810
      End
      Begin VB.TextBox txtNomcolor 
         Enabled         =   0   'False
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
         Left            =   4005
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtIdcolor 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
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
         Left            =   2145
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1440
         Width           =   1590
      End
      Begin VB.TextBox txtidcliente 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
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
         Left            =   4485
         MaxLength       =   5
         TabIndex        =   4
         Top             =   390
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtabrecli 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
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
         Left            =   2145
         MaxLength       =   5
         TabIndex        =   3
         Top             =   390
         Width           =   1575
      End
      Begin VB.TextBox txtOrdColor 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2145
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1770
         Width           =   3810
      End
      Begin VB.Label Etiqueta 
         Caption         =   "Nro Lote Packing  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   6
         Left            =   300
         TabIndex        =   31
         Tag             =   "Garment Aditional:"
         Top             =   2940
         Width           =   1500
      End
      Begin VB.Label Etiqueta 
         Caption         =   "Identificador Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   5
         Left            =   3495
         TabIndex        =   28
         Tag             =   "Garment Aditional:"
         Top             =   2160
         Width           =   930
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "% Adicional a Producir :"
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
         Index           =   4
         Left            =   300
         TabIndex        =   24
         Tag             =   "% Aditional:"
         Top             =   2190
         Width           =   1725
      End
      Begin VB.Label Etiqueta 
         Caption         =   "Prendas Adicionales a Producir :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   300
         TabIndex        =   23
         Tag             =   "Garment Aditional:"
         Top             =   2490
         Width           =   1755
      End
      Begin VB.Label Label2 
         Caption         =   "Estilo:"
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
         Left            =   270
         TabIndex        =   21
         Tag             =   "Style:"
         Top             =   1110
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "# PO :"
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
         Left            =   270
         TabIndex        =   20
         Tag             =   "# P.O.:"
         Top             =   750
         Width           =   1080
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Orden :"
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
         Left            =   270
         TabIndex        =   9
         Tag             =   "Order :"
         Top             =   1845
         Width           =   1080
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Id Color :"
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
         Left            =   270
         TabIndex        =   8
         Tag             =   "Colour Id:"
         Top             =   1485
         Width           =   1080
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Id Cliente :"
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
         Left            =   270
         TabIndex        =   7
         Tag             =   "Customer Id:"
         Top             =   420
         Width           =   1080
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2835
      TabIndex        =   0
      Top             =   6630
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   953
      Custom          =   $"frmMantColor.frx":05C8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmMantColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Rs_Carga As New ADODB.Recordset

Public sCod_Cliente, sCod_PurOrd, sCod_EstCli As String

Public oParent As Object

Public Codigo, Descripcion As String

Private boolPuedeModificarMerma

Private Sub cmdBuscaColor_Click()
    Load frmListaGeneral
    Set frmListaGeneral.oParent = Me
    frmListaGeneral.sQuery = "SELECT cod_colcli as Codigo, nom_colcli as Descripcion " & "FROM TG_ColCli WHERE cod_cliente = '" & txtIdCliente.Text & "'"
    Call frmListaGeneral.Cargar_Datos
    frmListaGeneral.Show 1

    If Codigo <> "" Then
        txtIdcolor.Text = Codigo
        txtNomcolor.Text = Descripcion
        Codigo = ""
    End If

End Sub

Private Sub cmdFirst_Click()

    If Not Rs_Carga.BOF Then
        Rs_Carga.MoveFirst
    End If

End Sub

Private Sub cmdLast_Click()

    If Not Rs_Carga.EOF Then
        Rs_Carga.MoveLast
    End If

End Sub

Private Sub cmdNext_Click()

    If Not Rs_Carga.EOF Then
        Rs_Carga.MoveNext
    End If

End Sub

Private Sub cmdPrevious_Click()

    If Not Rs_Carga.BOF Then
        Rs_Carga.MovePrevious
    End If

End Sub

Private Sub DGridlista_Click()

    If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
        txtIdcolor.Text = Rs_Carga!COD_COLCLI
        txtNomcolor.Text = Rs_Carga!nom_colcli
        txtOrdColor.Text = Rs_Carga!num_ordcolcli
        txtPor_AdicProd.Text = Rs_Carga!Por_AdicProd
        txtPre_AdicProd.Text = Rs_Carga!Pre_AdicProd
        txtNroLotePack.Text = Rs_Carga!NumLotePaking
        txtIdentificador_Color_Cliente.Text = Trim(Rs_Carga!Identificador_Color_Cliente)
        DESHABILITA_DATOS
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
    Else
        HabilitaMant Me.MantFunc1, "ADICIONAR"
    End If

End Sub

Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
    Avanza (KeyCode)
End Sub

Private Sub DGridlista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
        txtIdcolor.Text = Rs_Carga!COD_COLCLI
        txtNomcolor.Text = Rs_Carga!nom_colcli
        txtOrdColor.Text = Rs_Carga!num_ordcolcli
        txtPor_AdicProd.Text = FixNulos(Rs_Carga!Por_AdicProd, vbDouble)
        txtPre_AdicProd.Text = FixNulos(Rs_Carga!Pre_AdicProd, vbLong)
        txtIdentificador_Color_Cliente.Text = Trim(Rs_Carga!Identificador_Color_Cliente)
        DESHABILITA_DATOS
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
    Else
        HabilitaMant Me.MantFunc1, "ADICIONAR"
    End If

End Sub

Private Sub Form_Load()
    Call FormSet(Me)

    'Me.MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    FormateaGrid Me.DGridlista
    HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
    HABILITAR_MERMA (False)
End Sub

Sub SALVAR_DATOS()

    Dim Rs_DATOS As New ADODB.Recordset, stodo As String

    If chkTodos.value = 1 Then
        stodo = "S"
    
    Else
        stodo = "N"

    End If

    B_SQL = "UP_MAN_TG_PURORDCOL1 '" & txtIdCliente.Text & "','" & txtidPO.Text & "','" & txtIdEstCli.Text & "','" & txtIdcolor.Text & "', " & txtOrdColor.Text & "," & txtPor_AdicProd & "," & txtPre_AdicProd & ",'" & txtIdCliente & "','U','" & stodo & "','" & txtNroLotePack & "'"

    ExecuteCommandSQL cCONNECT, B_SQL

    Dim amensaje As New clsMessages

    amensaje.Codigo = MESSAGECODE.kMESSAGE_INF_DATA_SAVE
    Informa "", amensaje
End Sub

Sub LIMPIAR_DATOS()
    txtIdcolor.Text = ""
    txtNomcolor.Text = ""
    txtOrdColor.Text = ""
    txtPor_AdicProd.Text = ""
    txtPre_AdicProd.Text = ""
    txtIdentificador_Color_Cliente.Text = ""
    txtOrdColor.Enabled = True
    txtPor_AdicProd.Enabled = True
    txtPre_AdicProd.Enabled = True
    cmdBuscaColor.Enabled = True
    txtNroLotePack.Text = ""
    txtNroLotePack.Enabled = True
    txtIdentificador_Color_Cliente.Enabled = True

    cmdBuscaColor.SetFocus
End Sub

Public Sub VALIDAR_ACCESO_MERMA()

    On Error GoTo errorsql:

    Dim strSql As String

    Dim rsTemp As New ADODB.Recordset

    boolPuedeModificarMerma = False
    strSql = "select cod_usuario from tg_purordcol_UsuarioMerma where cod_usuario = '" & vusu & "'"

    rsTemp.ActiveConnection = cCONNECT
    rsTemp.CursorLocation = adUseClient
    rsTemp.CursorType = adOpenStatic
    rsTemp.LockType = adLockReadOnly
    rsTemp.Open strSql

    If Not rsTemp.EOF Then
        boolPuedeModificarMerma = True
    End If

    Exit Sub

errorsql:
    MsgBox "Error al cargar permisos " & Err.Description, vbOKOnly + vbCritical, Me.Caption
End Sub

Sub Cargar_Datos()
    B_SQL = "SG_Lista_Color '" & txtIdCliente.Text & "', '" & txtidPO.Text & "','" & txtIdEstCli.Text & "'"
    Rs_Carga.ActiveConnection = cCONNECT
    Rs_Carga.CursorLocation = adUseClient
    Rs_Carga.CursorType = adOpenStatic
    Rs_Carga.LockType = adLockReadOnly
    Rs_Carga.Open B_SQL
    Set DGridlista.DataSource = Rs_Carga

    If Not Rs_Carga.EOF Then
        txtIdcolor.Text = Rs_Carga!COD_COLCLI
        txtNomcolor.Text = Rs_Carga!nom_colcli
        txtOrdColor.Text = Rs_Carga!num_ordcolcli
        txtPor_AdicProd.Text = FixNulos(Rs_Carga!Por_AdicProd, vbDouble)
        txtPre_AdicProd.Text = FixNulos(Rs_Carga!Pre_AdicProd, vbLong)
        txtNroLotePack.Text = Rs_Carga!NumLotePaking
        txtIdentificador_Color_Cliente.Text = Trim(Rs_Carga!Identificador_Color_Cliente)
    End If

    DESHABILITA_DATOS
End Sub

Sub RECARGAR_DATOS()
    Rs_Carga.Close
    Cargar_Datos
End Sub

Sub BUSCA_COLOR()

    Dim Rs_busca As New ADODB.Recordset

    If txtIdCliente.Text <> "" Then
        B_SQL = "SELECT * FROM " & "TG_PurOrdCol " & "WHERE cod_cliente = '" & txtIdCliente.Text & "' " & "AND   cod_purord  = '" & txtidPO.Text & "' " & "AND   cod_estcli  = '" & txtIdEstCli.Text & "' " & "AND   cod_colcli  = '" & txtIdcolor.Text & "'"
        Rs_busca.ActiveConnection = cCONNECT
        Rs_busca.CursorType = adOpenStatic
        Rs_busca.Open B_SQL

        If Not Rs_busca.EOF Then
            txtIdcolor.Text = Rs_busca!COD_COLCLI
            txtOrdColor.Text = Rs_busca!num_ordcolcli
            DESHABILITA_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridlista.Enabled = True
        End If

        Rs_busca.Close
        Set Rs_busca = Nothing
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Rs_Carga = Nothing
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, _
                                  ByVal ActionType As Integer, _
                                  ByVal ActionName As String)

    Select Case ActionName

        Case "GRABAR"

            If VALIDA_DATOS Then
                SALVAR_DATOS
                RECARGAR_DATOS
                DGridlista.Enabled = True
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                cmdBuscaColor.Enabled = False
                HABILITAR_MERMA (False)
            End If

        Case "ADICIONAR"
            LIMPIAR_DATOS
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridlista.Enabled = False
            HABILITAR_MERMA (True)

        Case "MODIFICAR"
            txtOrdColor.Enabled = True
            HABILITAR_MERMA (True)
            txtIdentificador_Color_Cliente.Enabled = True
            txtNroLotePack.Enabled = True
            chkTodos.Enabled = True
            txtOrdColor.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridlista.Enabled = False

        Case "SALIR"
            Unload Me

        Case "DESHACER"
            LIMPIAR_DATOS
            RECARGAR_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridlista.Enabled = True
            cmdBuscaColor.Enabled = False
            HABILITAR_MERMA (False)
    End Select

End Sub

Private Sub HABILITAR_MERMA(bool As Boolean)
    txtPor_AdicProd.Enabled = IIf(boolPuedeModificarMerma, bool, False)
    txtPre_AdicProd.Enabled = IIf(boolPuedeModificarMerma, bool, False)
        
End Sub

Private Sub txtIdcolor_Change()

    If txtIdcolor.Text <> "" Then
        BUSCA_COLOR
        BUSCA_TIPOCOLOR
    End If

End Sub

Private Sub txtIdentificador_Color_Cliente_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then Avanza (KeyCode)
End Sub

Private Sub txtNomcolor_KeyDown(KeyCode As Integer, Shift As Integer)
    Avanza (KeyCode)
End Sub

Function VALIDA_DATOS() As Boolean

    Dim amensaje As clsMessages

    Set amensaje = New clsMessages
    VALIDA_DATOS = True

    If Len(Trim(txtOrdColor.Text)) = 0 Then
        amensaje.Codigo = MESSAGECODE.kMESSAGE_ERR_VALIDA_DES_COLOR
        VALIDA_DATOS = False
    End If

    If Len(Trim(txtIdcolor)) = 0 Then
        amensaje.Codigo = MESSAGECODE.kMESSAGE_ERR_VALIDA_COD_COLOR
        VALIDA_DATOS = False
    End If

    If Not VALIDA_DATOS Then
    
        amensaje.ShowMesage iLanguage
    End If

End Function

Sub DESHABILITA_DATOS()
    txtNomcolor.Enabled = False
    txtOrdColor.Enabled = False
    txtIdentificador_Color_Cliente.Enabled = False
    cmdBuscaColor.Enabled = False
    txtNroLotePack.Enabled = False
    'txtPor_AdicProd.Enabled = False
    'txtPre_AdicProd.Enabled = False
End Sub

Sub BUSCA_TIPOCOLOR()

    Dim Rs_busca As New ADODB.Recordset

    If txtIdcolor.Text <> "" Then
        B_SQL = "SELECT * FROM TG_colcli WHERE " & "cod_cliente = '" & txtIdCliente.Text & "' AND " & "cod_colcli  = '" & txtIdcolor.Text & "' "
        Rs_busca.ActiveConnection = cCONNECT
        Rs_busca.CursorType = adOpenStatic
        Rs_busca.Open B_SQL

        If Not Rs_busca.EOF Then
            txtNomcolor.Text = Rs_busca!nom_colcli
        Else
            txtNomcolor.Text = ""
        End If

        Rs_busca.Close
        Set Rs_busca = Nothing
    End If

End Sub

Public Sub Inicializar()
    txtIdCliente.Text = sCod_Cliente
    txtidPO.Text = sCod_PurOrd
    txtIdEstCli.Text = sCod_EstCli
End Sub

Private Sub txtOrdColor_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 13
            Avanza (KeyCode)

        Case Else
            SoloNumeros txtOrdColor, KeyAscii, True, 4, 0
    End Select

End Sub

Private Sub txtPor_AdicProd_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 13
            Avanza (KeyCode)

        Case Else
            SoloNumeros txtPor_AdicProd, KeyAscii, True, 4, 0
    End Select

End Sub

Private Sub txtPre_AdicProd_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 13
            Avanza (KeyCode)

        Case Else
            SoloNumeros txtPre_AdicProd, KeyAscii, True, 4, 0
    End Select

End Sub

