VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "mantenimientos.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmOrdComp 
   Caption         =   "Ordenes de Compra"
   ClientHeight    =   8820
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11085
   Icon            =   "frmOrdComp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   11085
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   525
      Left            =   3690
      TabIndex        =   23
      Top             =   8175
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmOrdComp.frx":030A
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin VB.Frame fraOpciones 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5955
      Left            =   9600
      TabIndex        =   22
      Top             =   1200
      Width           =   1455
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   4920
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   8678
         Custom          =   $"frmOrdComp.frx":04B0
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1150
         ControlHeigth   =   500
         ControlSeparator=   50
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   630
         Left            =   120
         TabIndex        =   70
         Top             =   5160
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   900
         Custom          =   "0~0~CERRAR~Verdadero~Verdadero~&Cerrar~0~0~1~~0~Falso~Falso~&Cerrar~"
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
   End
   Begin VB.Frame fraDetalle 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   21
      Top             =   3960
      Width           =   9465
      Begin VB.TextBox txtSer_ordcomp_Tinto 
         Height          =   285
         Left            =   6000
         MaxLength       =   3
         TabIndex        =   75
         Top             =   3720
         Width           =   1005
      End
      Begin VB.TextBox txtcod_ordcomp_Tinto 
         Height          =   285
         Left            =   7095
         MaxLength       =   6
         TabIndex        =   74
         Top             =   3720
         Width           =   2265
      End
      Begin VB.TextBox txtAbr_Cliente 
         Height          =   315
         Left            =   1650
         TabIndex        =   72
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox txtNom_Cliente 
         Height          =   315
         Left            =   2400
         TabIndex        =   71
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox TxtNumImportacion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   68
         Top             =   2475
         Width           =   840
      End
      Begin VB.TextBox TxtCodProv 
         Height          =   315
         Left            =   5010
         TabIndex        =   66
         Top             =   2475
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.ComboBox CboEstado 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   2835
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.ComboBox cboCod_CenCost 
         Height          =   315
         Left            =   5010
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   2130
         Width           =   4335
      End
      Begin VB.TextBox txtPorc_IGV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7005
         TabIndex        =   40
         Top             =   690
         Width           =   675
      End
      Begin VB.ComboBox cboCod_ProTex 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   2130
         Width           =   1590
      End
      Begin VB.CommandButton cmdBuscaGrupo 
         Caption         =   "..."
         Height          =   330
         Left            =   2985
         TabIndex        =   55
         Top             =   2115
         Width           =   330
      End
      Begin VB.TextBox txtCod_Grupo 
         Height          =   315
         Left            =   1710
         MaxLength       =   8
         TabIndex        =   54
         Top             =   2130
         Visible         =   0   'False
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker dtpFec_Entrega_Fin 
         Height          =   315
         Left            =   5010
         TabIndex        =   52
         Top             =   1770
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   69468163
         CurrentDate     =   37267
      End
      Begin MSComCtl2.DTPicker dtpFec_Entrega_Inicio 
         Height          =   315
         Left            =   1710
         TabIndex        =   50
         Top             =   1770
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   69468163
         CurrentDate     =   37267
      End
      Begin VB.ComboBox cboCod_ClaOrdComp 
         Height          =   315
         Left            =   5010
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   1410
         Width           =   4365
      End
      Begin VB.ComboBox cboCod_StaOrdComp 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   1410
         Width           =   2040
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   465
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   62
         Top             =   3255
         Width           =   7665
      End
      Begin VB.ComboBox cboCod_LugEntr 
         Height          =   315
         Left            =   5010
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   1050
         Width           =   4365
      End
      Begin VB.ComboBox cboCod_Moneda 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   1050
         Width           =   2040
      End
      Begin VB.ComboBox cboCod_Descuento 
         Height          =   315
         Left            =   5010
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   690
         Width           =   1260
      End
      Begin VB.ComboBox cboCod_CondVent 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   690
         Width           =   2040
      End
      Begin VB.TextBox txtDes_Proveedor 
         Height          =   315
         Left            =   6240
         MaxLength       =   50
         TabIndex        =   34
         Top             =   330
         Width           =   3165
      End
      Begin VB.TextBox txtCod_Proveedor 
         Height          =   315
         Left            =   5010
         MaxLength       =   12
         TabIndex        =   33
         Top             =   330
         Width           =   1200
      End
      Begin VB.TextBox txtCod_OrdComp 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1710
         TabIndex        =   31
         Top             =   330
         Width           =   1560
      End
      Begin VB.TextBox TxtDes_Grupo 
         Height          =   315
         Left            =   3285
         MaxLength       =   50
         TabIndex        =   56
         Top             =   2130
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Orden:"
         Height          =   195
         Left            =   5520
         TabIndex        =   76
         Top             =   3765
         Width           =   480
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Textil"
         Height          =   195
         Left            =   480
         TabIndex        =   73
         Top             =   3840
         Width           =   900
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "C. de Costo :"
         Height          =   195
         Left            =   3870
         TabIndex        =   59
         Top             =   2205
         Width           =   915
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Num. Importación:"
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   2550
         Width           =   1290
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Prov :"
         Height          =   195
         Left            =   3870
         TabIndex        =   67
         Top             =   2550
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Total/Parcial:"
         Height          =   195
         Left            =   105
         TabIndex        =   64
         Top             =   2940
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Proceso Textíl :"
         Height          =   195
         Left            =   105
         TabIndex        =   57
         Top             =   2205
         Width           =   1125
      End
      Begin VB.Label Label17 
         Caption         =   "Grupo :"
         Height          =   225
         Left            =   120
         TabIndex        =   53
         Top             =   2205
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "F. Entrega Fin :"
         Height          =   195
         Left            =   3870
         TabIndex        =   51
         Top             =   1845
         Width           =   1080
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "F. Entrega Inicio :"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   1845
         Width           =   1245
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Clase OC :"
         Height          =   195
         Left            =   3870
         TabIndex        =   47
         Top             =   1485
         Width           =   750
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Estado de la OC :"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   1485
         Width           =   1245
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones :"
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   3375
         Width           =   1155
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Lugar Entrega :"
         Height          =   195
         Left            =   3870
         TabIndex        =   43
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Moneda :"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   1110
         Width           =   675
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "I.G.V.:"
         Height          =   195
         Left            =   6405
         TabIndex        =   39
         Top             =   765
         Width           =   450
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Dsctos :"
         Height          =   195
         Left            =   3870
         TabIndex        =   37
         Top             =   765
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cond. Venta :"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Orden Compra :"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   405
         Width           =   1110
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor :"
         Height          =   195
         Left            =   3870
         TabIndex        =   32
         Top             =   405
         Width           =   825
      End
   End
   Begin VB.Frame FraBuscar 
      Caption         =   "Buscar Por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   10890
      Begin VB.Frame fraoptions 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   330
         Left            =   360
         TabIndex        =   17
         Top             =   160
         Width           =   7335
         Begin VB.OptionButton optProveedor 
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   2205
            TabIndex        =   20
            Top             =   120
            Width           =   1425
         End
         Begin VB.OptionButton optEstado 
            Caption         =   "Estado"
            Height          =   195
            Left            =   3885
            TabIndex        =   19
            Top             =   120
            Width           =   1185
         End
         Begin VB.OptionButton optOrdCompra 
            Caption         =   "Orden de Compra"
            Height          =   195
            Left            =   45
            TabIndex        =   18
            Top             =   120
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin FunctionsButtons.FunctButt FunctBuscar 
         Height          =   495
         Left            =   9480
         TabIndex        =   16
         Top             =   480
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
      Begin VB.Frame FraOrdComp 
         Height          =   645
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   7455
         Begin VB.TextBox txtCodOrdComp 
            Height          =   285
            Left            =   4275
            MaxLength       =   6
            TabIndex        =   7
            Top             =   270
            Width           =   1425
         End
         Begin VB.TextBox txtSerOrdComp 
            Height          =   285
            Left            =   1500
            MaxLength       =   3
            TabIndex        =   4
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Numero"
            Height          =   195
            Left            =   3075
            TabIndex        =   6
            Top             =   345
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Serie"
            Height          =   195
            Left            =   300
            TabIndex        =   3
            Top             =   315
            Width           =   360
         End
      End
      Begin VB.Frame FraEstado 
         Height          =   640
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   7455
         Begin VB.TextBox txtCodStaOrdComp 
            Height          =   285
            Left            =   1500
            MaxLength       =   1
            TabIndex        =   13
            Top             =   270
            Width           =   1005
         End
         Begin VB.TextBox txtDesStaOrdComp 
            Height          =   285
            Left            =   2880
            MaxLength       =   50
            TabIndex        =   15
            Top             =   255
            Width           =   4200
         End
         Begin VB.CommandButton cmdBusEstado 
            Caption         =   "..."
            Height          =   330
            Left            =   2520
            TabIndex        =   14
            Tag             =   "..."
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label2 
            Caption         =   "Estado :"
            Height          =   240
            Left            =   300
            TabIndex        =   12
            Top             =   330
            Width           =   690
         End
      End
      Begin VB.Frame FraProveedor 
         Height          =   640
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   7455
         Begin VB.TextBox txtDesProveedor 
            Height          =   285
            Left            =   2865
            MaxLength       =   50
            TabIndex        =   11
            Top             =   270
            Width           =   4155
         End
         Begin VB.TextBox txtCodProveedor 
            Height          =   285
            Left            =   1500
            MaxLength       =   12
            TabIndex        =   9
            Top             =   270
            Width           =   1365
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor :"
            Height          =   195
            Left            =   300
            TabIndex        =   8
            Top             =   270
            Width           =   825
         End
      End
   End
   Begin VB.Frame FraLista 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   150
      TabIndex        =   0
      Top             =   1200
      Width           =   9435
      Begin GridEX20.GridEX gexLista 
         Height          =   2325
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   9240
         _ExtentX        =   16298
         _ExtentY        =   4101
         Version         =   "2.0"
         AllowRowSizing  =   -1  'True
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         EmptyRows       =   -1  'True
         HeaderStyle     =   3
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         DataMode        =   1
         GridLines       =   1
         ColumnHeaderHeight=   285
         IntProp7        =   0
         ColumnsCount    =   6
         Column(1)       =   "frmOrdComp.frx":080E
         Column(2)       =   "frmOrdComp.frx":0962
         Column(3)       =   "frmOrdComp.frx":0A96
         Column(4)       =   "frmOrdComp.frx":0BDE
         Column(5)       =   "frmOrdComp.frx":0C82
         Column(6)       =   "frmOrdComp.frx":0D26
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmOrdComp.frx":0DCA
         FormatStyle(2)  =   "frmOrdComp.frx":0F02
         FormatStyle(3)  =   "frmOrdComp.frx":0FB2
         FormatStyle(4)  =   "frmOrdComp.frx":1066
         FormatStyle(5)  =   "frmOrdComp.frx":113E
         FormatStyle(6)  =   "frmOrdComp.frx":11F6
         ImageCount      =   0
         PrinterProperties=   "frmOrdComp.frx":12D6
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   1635
      TabIndex        =   24
      Top             =   8115
      Width           =   1965
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1455
         Picture         =   "frmOrdComp.frx":14AE
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Ultimo"
         Top             =   105
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   975
         Picture         =   "frmOrdComp.frx":1620
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Siguiente"
         Top             =   105
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   495
         Picture         =   "frmOrdComp.frx":1792
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Anterior"
         Top             =   105
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   15
         Picture         =   "frmOrdComp.frx":1904
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Primero"
         Top             =   105
         Width           =   495
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   11640
      Top             =   7080
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmOrdComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSql As String
Dim Rs_Lista As ADODB.Recordset
Dim sTipo As String
Dim opcion As Integer
Public codigo, Descripcion As String
'VAriables del Form
Public varCod_TipRequ As Integer
Dim varSer_OrdComp As String
Dim varProvCod_ClaOrdComp As String
Dim varFlg_Requerimiento As Boolean
'Variables para la impresion
Public varCadena_Familias As String
Public varCancelImpresion As Integer
Dim sTituliAbrOP As String
Public varAyuda As Integer
Public scliente As String

Dim Pregunta As Variant
Public rstAux As ADODB.Recordset

Dim vTip_Item, vTip_Presentacion, vCod_ProTex As String

Private Sub cboCod_ClaOrdComp_Click()
   Dim varCod_Protex As String
   Dim varTip_Item As String
   Dim varTip_Presentacion As String
    'si no tiene proceso relacionado entonces es un proceso post tenido
    strSql = "SELECT ISNULL(Cod_Protex,'') FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
    varCod_Protex = DevuelveCampo(strSql, cConnect)
    If Trim(varCod_Protex) = "" Then
        If sTipo = "I" Or sTipo = "U" Then
            cboCod_ProTex.Enabled = True
        End If
        
        strSql = "SELECT Tip_Item FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
        varTip_Item = DevuelveCampo(strSql, cConnect)
        
        strSql = "SELECT Tip_Presentacion FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
        varTip_Presentacion = DevuelveCampo(strSql, cConnect)
        
        If varTip_Item = "T" And varTip_Presentacion = "T" Then
            strSql = "SELECT Des_ProTex + SPACE(100) + Cod_ProTex FROM TX_PROCESOS WHERE Flg_TejTen = 'T' AND Flg_principal = ''"
            Call LlenaCombo(cboCod_ProTex, strSql, cConnect)
        Else
            If varTip_Item = "T" And varTip_Presentacion = "C" Then
                strSql = "SELECT Des_ProTex + SPACE(100) + Cod_ProTex FROM TX_PROCESOS WHERE Flg_TejTen = 'J' AND Flg_principal = ''"
                Call LlenaCombo(cboCod_ProTex, strSql, cConnect)
            Else
                cboCod_ProTex.Clear
            End If
        End If
    Else
        'Aqui llenamos los codigos de los procesos textiles
        strSql = "SELECT Des_ProTex + SPACE(100) + Cod_ProTex FROM TX_PROCESOS WHERE Cod_ProTex = '" & varCod_Protex & "'"
        Call LlenaCombo(cboCod_ProTex, strSql, cConnect)
        cboCod_ProTex.Enabled = False
       cboCod_ProTex.ListIndex = 0
    End If
   
    
    strSql = "SELECT Flg_Requerimiento FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
    If DevuelveCampo(strSql, cConnect) = "S" Then
        strSql = "SELECT Cod_TipRequ FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
        varCod_TipRequ = DevuelveCampo(strSql, cConnect)
    End If
    
'    Strsql = "SELECT Flg_Requerimiento FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
'    If DevuelveCampo(Strsql, cCONNECT) = "S" Then
'        txtCod_Grupo.Enabled = True
'        TxtDes_Grupo.Enabled = True
'        cmdBuscaGrupo.Enabled = True
'
'        varFlg_Requerimiento = True
'        'ProvCod_ClaOrdComp = Right(cboCod_ClaOrdComp.Text, 2)
'
'        Strsql = "SELECT Cod_TipRequ FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
'        varCod_TipRequ = DevuelveCampo(Strsql, cCONNECT)
'
'    Else
'        txtCod_Grupo.Text = ""
'        TxtDes_Grupo.Text = ""
'        txtCod_Grupo.Enabled = False
'        TxtDes_Grupo.Enabled = False
'        cmdBuscaGrupo.Enabled = False
'
'        varFlg_Requerimiento = False
'    End If
'
'    If sTipo = "" Then
'        txtCod_Grupo.Enabled = False
'        TxtDes_Grupo.Enabled = False
'        cmdBuscaGrupo.Enabled = False
'    End If
    
End Sub



Private Sub cmdFirst_Click()
    If Not Rs_Lista.BOF Then
        Rs_Lista.MoveFirst
    End If
End Sub

Private Sub cmdLast_Click()
    If Not Rs_Lista.EOF Then
        Rs_Lista.MoveLast
    End If
End Sub

Private Sub cmdNext_Click()
    If Not Rs_Lista.EOF Then
        Rs_Lista.MoveNext
        If Rs_Lista.EOF Then
            Rs_Lista.MoveLast
        End If
    End If
End Sub

Private Sub cmdPrevious_Click()
    If Not Rs_Lista.BOF Then
        Rs_Lista.MovePrevious
        If Rs_Lista.BOF Then
            Rs_Lista.MoveFirst
        End If
    End If
End Sub

Sub LIMPIAR_DATOS()
    
    txtCod_OrdComp.Text = ""
    txtCod_Proveedor.Text = ""
    txtDes_Proveedor.Text = ""
    TxtNumImportacion.Text = ""
    txtcod_ordcomp_Tinto.Text = ""
    txtSer_ordcomp_Tinto.Text = ""
    
    cboCod_CondVent.ListIndex = -1
    cboCod_Descuento.ListIndex = -1
    
    cboCod_Moneda.ListIndex = -1
    cboCod_LugEntr.ListIndex = -1
    txtObservaciones.Text = ""
    cboCod_StaOrdComp.ListIndex = -1
    cboCod_ClaOrdComp.ListIndex = -1
    dtpFec_Entrega_Inicio.Value = Date
    dtpFec_Entrega_Fin.Value = Date
    cboCod_CenCost.ListIndex = -1
    txtCod_Grupo.Text = ""
    TxtDes_Grupo.Text = ""
    cboCod_ProTex.ListIndex = -1
    CboEstado.ListIndex = -1

    'Aqui llenamos a los valores por defecto
    strSql = "SELECT Porc_IGV FROM TG_IGV WHERE ANO=YEAR(GETDATE()) AND MES=RIGHT('0'+CONVERT(VARCHAR,MONTH(GETDATE())),2) "
    txtPorc_IGV.Text = DevuelveCampo(strSql, cConnect)

End Sub

Sub CARGA_COMBOS()

    'Aqui llenamos las condiciones de Venta
    strSql = "SELECT Des_CondVent + SPACE(100)+ Cod_CondVent FROM LG_CONDVENT"
    Call LlenaCombo(cboCod_CondVent, strSql, cConnect)
    
    'Aqui llenamos los Descuentos
    strSql = "SELECT CONVERT(VARCHAR,Porcentaje1) + ' - '+ CONVERT(VARCHAR,Porcentaje2) + SPACE(100) + COD_DESCUENTO FROM LG_DSCTOS"
    Call LlenaCombo(cboCod_Descuento, strSql, cConnect)
    
    'Aqui llenamos las Monedas
    strSql = "SELECT Nom_Moneda + SPACE(100) + Cod_Moneda FROM TG_MONEDA"
    Call LlenaCombo(cboCod_Moneda, strSql, cConnect)
    
    
    strSql = "SELECT Des_LugEntr + SPACE(100) + Cod_LugEntr FROM LG_LUGENTR"
    Call LlenaCombo(cboCod_LugEntr, strSql, cConnect)
    
    strSql = "SELECT Des_StaOrdComp + SPACE(100) + Cod_StaOrdComp FROM LG_STAORDCOMP"
    Call LlenaCombo(cboCod_StaOrdComp, strSql, cConnect)
    
    strSql = "SELECT a.Des_ClaOrdComp + SPACE(100) + a.Cod_ClaOrdComp FROM LG_CLAORDCOMP a,lg_segordcomp b where a.cod_claordcomp = b.cod_claordcomp and b.cod_usuario ='" & vusu & "'"
    Call LlenaCombo(cboCod_ClaOrdComp, strSql, cConnect)
    
    'Aqui llenamos los codigos de los procesos textiles
    strSql = "SELECT Des_ProTex + SPACE(100) + Cod_ProTex FROM TX_PROCESOS"
    Call LlenaCombo(cboCod_ProTex, strSql, cConnect)
    
    'Aqui llenamos nos centros de costo
    strSql = "SELECT Des_CenCost + SPACE(100) + Cod_CenCost FROM TG_CENCOSTO"
    Call LlenaCombo(cboCod_CenCost, strSql, cConnect)
    
    'Aqui llenamos los estados (Total y parcial)
    
    strSql = "SELECT DESCRIPCION + SPACE(100) + FLG_TOTAL_PARCIAL FROM LG_MODORDCOMP"
    Call LlenaCombo(CboEstado, strSql, cConnect)
End Sub

Function VALIDA_DATOS() As Boolean
    Dim NombreTabla As String
    Dim CodigoTabla As String
    

    VALIDA_DATOS = True
    If sTipo <> "D" Then
'
'        If sTipo = "I" Then
'            If ExisteCampo("Cod_StaOrdComp", "Lg_StaOrdComp", Trim(txtcod_StaOrdComp.Text), cCONNECT, True) Then
'                MsgBox "El código de Status de Orden de Compra ya se encuentra registrado. Sirvase verificar", vbInformation, "Status de Orden de Compra"
'                txtcod_StaOrdComp.SetFocus
'                VALIDA_DATOS = False
'                Exit Function
'            End If
'        End If
'
'        If Trim(txtcod_StaOrdComp.Text) = "" Then
'            MsgBox "El código de Status de Orden de Compra no puede estar vacío. Sirvase verificar", vbInformation, "Ordenes de Compra"
'            txtcod_StaOrdComp.Text = ""
'            txtcod_StaOrdComp.SetFocus
'            VALIDA_DATOS = False
'            Exit Function
'        End If
'
'        If Trim(txtDes_StaOrdComp.Text) = "" Then
'            MsgBox "La descripción de Status de Orden de Compra no puede estar vacío. Sirvase verificar", vbInformation, "Ordenes de Compra"
'            txtDes_StaOrdComp.Text = ""
'            txtDes_StaOrdComp.SetFocus
'            VALIDA_DATOS = False
'            Exit Function
'        End If

        If Trim(txtCod_Proveedor.Text) = "" Then
            MsgBox "El Código de Proveedor no puede estar vacio. Sirvase verificar", vbInformation, "Ordenes de Compra"
            txtCod_Proveedor.Text = ""
            txtCod_Proveedor.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
        
        strSql = "SELECT count(*) FROM LG_PROVEEDOR WHERE Cod_Proveedor = '" & txtCod_Proveedor.Text & "'"
        If DevuelveCampo(strSql, cConnect) = "0" Then
            MsgBox "El código de proveedor ingresado no es válido. Sirvase verificar", vbInformation, "Ordenes de Compra"
            txtCod_Proveedor.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
    
        If Trim(cboCod_Descuento.Text) = "" Then
            MsgBox "El descuento no puede estar vacio. Sirvase verificar", vbInformation, "Ordenes de Compra"
            cboCod_Descuento.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
    
    
        If Trim(cboCod_CondVent.Text) = "" Then
            MsgBox "La condición de venta no puede estar vacia. Sirvase verificar", vbInformation, "Ordenes de Compra"
            cboCod_CondVent.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
        
        If Trim(cboCod_Moneda.Text) = "" Then
            MsgBox "La moneda no puede estar vacia. Sirvase verificar", vbInformation, "Ordenes de Compra"
            cboCod_Moneda.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
        
        If Trim(cboCod_LugEntr.Text) = "" Then
            MsgBox "El campo lugar de entrega no puede estar vacio. Sirvase verificar", vbInformation, "Ordenes de Compra"
            cboCod_LugEntr.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
        
        If Trim(cboCod_ClaOrdComp.Text) = "" Then
            MsgBox "La clase de orden de compra no puede estar vacia. Sirvase verificar", vbInformation, "Ordenes de Compra"
            cboCod_ClaOrdComp.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If

        If dtpFec_Entrega_Fin.Value < dtpFec_Entrega_Inicio.Value Then
            MsgBox "La fecha de entrega final no puede ser menor a la inicial. Sirvase verificar", vbInformation, "Ordenes de Compra"
            dtpFec_Entrega_Fin.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If

        'Preguntamos por la variable si es requerida o no
        strSql = "SELECT Flg_Requerimiento FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
        
        If DevuelveCampo(strSql, cConnect) <> "S" Then
            If Trim(cboCod_CenCost.Text) = "" Then
                MsgBox "El centro de costo no puede estar vacio. Sirvase verificar", vbInformation, "Ordenes de Compra"
                cboCod_CenCost.SetFocus
                VALIDA_DATOS = False
                Exit Function
            End If
        End If
        
        If varFlg_Requerimiento = True Then
        
            If Trim(txtCod_Grupo.Text) = "" Then
                MsgBox "El grupo no puede estar vacio. Sirvase verificar", vbInformation, "Ordenes de Compra"
                txtCod_Grupo.Text = ""
                txtCod_Grupo.SetFocus
                VALIDA_DATOS = False
                Exit Function
            End If
            
            'Como el grupo puede ser textil o log, determinamos primero de quien se trata
            strSql = "SELECT Tip_Grupo FROM LG_TIPREQ WHERE Cod_TipRequ='" & varCod_TipRequ & "'"
            If DevuelveCampo(strSql, cConnect) = "I" Then
                NombreTabla = "ES_GRUPOLOG"
                CodigoTabla = "Cod_GrupoLog"
            Else
                NombreTabla = "ES_GRUPOTEX"
                CodigoTabla = "Cod_GrupoTex"
            End If
            'Una vez determ el grupo preguntamos si el codigo existe
            strSql = "SELECT count(*) FROM " & NombreTabla & " WHERE " & CodigoTabla & " = '" & txtCod_Grupo.Text & "'"
            If DevuelveCampo(strSql, cConnect) = "0" Then
                MsgBox "El codigo de grupo ingresado no es válido. Sirvase verificar", vbInformation, "Ordenes de Compra"
                txtCod_Grupo.SetFocus
                VALIDA_DATOS = False
                Exit Function
            End If
        End If

        strSql = "SELECT Tip_Item FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
        If DevuelveCampo(strSql, cConnect) <> "I" Then
            If Trim(cboCod_ProTex.Text) = "" Then
                MsgBox "El proceso textil no puede estar vacio. Sirvase verificar", vbInformation, "Ordenes de Compra"
                cboCod_ProTex.SetFocus
                VALIDA_DATOS = False
                Exit Function
            End If
        End If

    Else
        'Aqui se valida que no tenga registros dependientes
        strSql = "SELECT COUNT(*) FROM LG_ORDCOMPITEM WHERE Ser_OrdComp='" & gexLista.Value(gexLista.Columns("Ser_OrdComp").Index) & "' AND Cod_OrdComp='" & gexLista.Value(gexLista.Columns("Cod_OrdComp").Index) & "'"
        If DevuelveCampo(strSql, cConnect) > 0 Then
            MsgBox "El registro seleccionado posee registros relacionados. Sirvase verificar", vbInformation, "Ordenes de Compra"
            VALIDA_DATOS = False
            Exit Function
        End If
    End If
End Function

Sub CARGA_DATOS()

    If Not Rs_Lista.EOF And Not Rs_Lista.BOF Then
        
        varSer_OrdComp = gexLista.Value(gexLista.Columns("Ser_OrdComp").Index)
        txtCod_OrdComp.Text = gexLista.Value(gexLista.Columns("Cod_OrdComp").Index)
        txtPorc_IGV.Text = gexLista.Value(gexLista.Columns("I.G.V.").Index)
        txtObservaciones.Text = gexLista.Value(gexLista.Columns("Observaciones").Index)
        TxtCodProv.Text = gexLista.Value(gexLista.Columns("Observaciones").Index)
        dtpFec_Entrega_Inicio.Value = gexLista.Value(gexLista.Columns("F.Entrega Inicial").Index)
        TxtNumImportacion.Text = gexLista.Value(gexLista.Columns("Num_Importacion").Index)
        'dtpFec_Entrega_Fin.Value = gexLista.Value(gexLista.Columns("F.Entrega Final").Index)
        txtAbr_Cliente.Tag = gexLista.Value(gexLista.Columns("COD_CLIENTE").Index)
        txtAbr_Cliente = gexLista.Value(gexLista.Columns("ABR_CLIENTE").Index)
        txtNom_Cliente = gexLista.Value(gexLista.Columns("NOM_CLIENTE").Index)
        
        Call BuscaCombo(gexLista.Value(gexLista.Columns("Cod_CondVent").Index), 2, cboCod_CondVent)
        Call BuscaCombo(gexLista.Value(gexLista.Columns("Cod_Descuento").Index), 2, cboCod_Descuento)
        Call BuscaCombo(gexLista.Value(gexLista.Columns("Cod_CenCost").Index), 2, cboCod_CenCost)
        Call BuscaCombo(gexLista.Value(gexLista.Columns("Cod_Moneda").Index), 2, cboCod_Moneda)
        Call BuscaCombo(gexLista.Value(gexLista.Columns("Cod_LugEntr").Index), 2, cboCod_LugEntr)
        Call BuscaCombo(gexLista.Value(gexLista.Columns("Cod_StaOrdComp").Index), 2, cboCod_StaOrdComp)
        Call BuscaCombo(gexLista.Value(gexLista.Columns("Cod_ClaOrdComp").Index), 2, cboCod_ClaOrdComp)
        Call BuscaCombo(gexLista.Value(gexLista.Columns("Cod_ProTex").Index), 2, cboCod_ProTex)
        Call BuscaCombo(gexLista.Value(gexLista.Columns("Flg_Total_Parcial").Index), 2, CboEstado)
        
        txtCod_Proveedor.Text = gexLista.Value(gexLista.Columns("Cod_Proveedor").Index)
        Call BUSCA_PROVEEDOR(1, 2)
        txtCod_Grupo.Text = gexLista.Value(gexLista.Columns("Cod.Grupo").Index)
        Call BUSCA_GRUPO(1)
        
        txtSer_ordcomp_Tinto.Text = Trim(gexLista.Value(gexLista.Columns("ser_ordcomp_tex").Index))
        txtcod_ordcomp_Tinto.Text = Trim(gexLista.Value(gexLista.Columns("cod_ordcomp_tex").Index))
        
    End If
End Sub

Sub HABILITA_DATOS()
Dim RsDet As ADODB.Recordset
    If sTipo = "I" Then
        cboCod_ClaOrdComp.Enabled = True
        txtCod_Grupo.Enabled = True
        TxtDes_Grupo.Enabled = True
        cmdBuscaGrupo.Enabled = True
        txtSer_ordcomp_Tinto.Enabled = True
        txtcod_ordcomp_Tinto.Enabled = True
        
   Else
        Set RsDet = Nothing
        Set RsDet = New ADODB.Recordset
        RsDet.CursorLocation = adUseClient
        RsDet.Open "SELECT * FROM lg_ordcompitem WHERE Ser_OrdComp='" & Trim(gexLista.Value(gexLista.Columns("Ser_OrdComp").Index)) & "' AND Cod_OrdComp='" & Trim(gexLista.Value(gexLista.Columns("Cod_OrdComp").Index)) & "'", cConnect
        
        If RsDet.RecordCount = 0 Then
            txtCod_Grupo.Enabled = True
            TxtDes_Grupo.Enabled = True
            cmdBuscaGrupo.Enabled = True
'        Else
'            txtCod_Grupo.Enabled = False
'            txtDes_Grupo.Enabled = False
'            cmdBuscaGrupo.Enabled = False
        End If
        
        txtSer_ordcomp_Tinto.Enabled = True
        txtcod_ordcomp_Tinto.Enabled = True
        
    End If
    
    txtCod_Proveedor.Enabled = True
    txtDes_Proveedor.Enabled = True
    cboCod_CondVent.Enabled = True
    cboCod_Descuento.Enabled = True
    cboCod_Moneda.Enabled = True
    cboCod_LugEntr.Enabled = True
    txtObservaciones.Enabled = True
        
    cboCod_CenCost.Enabled = True
    cboCod_ProTex.Enabled = True
    
    dtpFec_Entrega_Fin.Enabled = True
    dtpFec_Entrega_Inicio.Enabled = True
End Sub

Sub INHABILITA_DATOS()
    
    txtCod_Proveedor.Enabled = False
    txtDes_Proveedor.Enabled = False
    cboCod_CondVent.Enabled = False
    cboCod_Descuento.Enabled = False
    cboCod_Moneda.Enabled = False
    cboCod_LugEntr.Enabled = False
    txtObservaciones.Enabled = False
    cboCod_StaOrdComp.Enabled = False
    cboCod_ClaOrdComp.Enabled = False
    cboCod_CenCost.Enabled = False
    txtCod_Grupo.Enabled = False
    TxtDes_Grupo.Enabled = False
    cmdBuscaGrupo.Enabled = False
    cboCod_ProTex.Enabled = False
    CboEstado.Enabled = False
    dtpFec_Entrega_Fin.Enabled = False
    dtpFec_Entrega_Inicio.Enabled = False

End Sub

Sub CARGA_GRID()
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.ActiveConnection = cConnect
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    'Esta cadena es para devolver el Codigo de Cliente
    strSql = "EXEC UP_SEL_ORDCOMP " & CStr(opcion) & ",'" & Trim(txtSerOrdComp.Text) & "','" & Trim(txtCodOrdComp.Text) & "','" & Trim(txtCodProveedor.Text) & "','" & Trim(txtCodStaOrdComp.Text) & "','','" & vusu & "','',''"
    
    Rs_Lista.Open strSql
    Set gexLista.ADORecordset = Rs_Lista

    If Rs_Lista.RecordCount > 0 Then
        gexLista.Enabled = True
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Call CARGA_DATOS
    Else
        gexLista.Enabled = False
        HabilitaMant Me.MantFunc1, "ADICIONAR"
        Call LIMPIAR_DATOS
    End If
    CONFIGURAR_GRID
End Sub

Private Sub CONFIGURAR_GRID()
    gexLista.Columns("Ser_OrdComp").Visible = False
    gexLista.Columns("Cod_OrdComp").Visible = False
    gexLista.Columns("Cod_Proveedor").Visible = False
    gexLista.Columns("Cod_CondVent").Visible = False
    gexLista.Columns("Cod_LugEntr").Visible = False
    gexLista.Columns("Cod_StaOrdComp").Visible = False
    gexLista.Columns("Cod_ClaOrdComp").Visible = False
    gexLista.Columns("Cod_ProTex").Visible = False
    gexLista.Columns("Cod_CenCost").Visible = False
    gexLista.Columns("Cod_Moneda").Visible = False
    gexLista.Columns("Cod_Descuento").Visible = False
    gexLista.Columns("Observaciones").Visible = False
    gexLista.Columns("Flg_Total_Parcial").Visible = False
    
    gexLista.Columns("Proveedor").Width = 2500
    gexLista.Columns("I.G.V.").Width = 700
    gexLista.Columns("O.C.").Width = 1100
    gexLista.Columns("Descuentos").Width = 900
    gexLista.Columns("Cod.Grupo").Width = 900
    gexLista.Columns("Moneda").Width = 2000
    gexLista.Columns("L.Entrega").Width = 2000
    gexLista.Columns("CondVenta").Width = 2000
End Sub

Sub CAMBIO_ESTADO()
    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Dim strSql As String
    
    Con.ConnectionString = cConnect
    Con.Open
    
        Con.BeginTrans

        strSql = "EXEC UP_MAN_ORDCOMPCAMBIOESTADO '" & _
        gexLista.Value(gexLista.Columns("Ser_OrdComp").Index) & "','" & _
        gexLista.Value(gexLista.Columns("Cod_OrdComp").Index) & "','" & _
        vusu & "'"
        
        Con.Execute strSql

        Con.CommitTrans
        'Dim amensaje As New clsMensaje
        'amensaje.Codigo = CodeMsg.KMESSAGE_INF_DATA_SAVE
        'Informa "", amensaje
        
        MsgBox "El cambio de estado resultó exitoso.", vbOKOnly, "Ordenes de Compra"
        
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub
Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    Dim Rs As New ADODB.Recordset
    On Error GoTo Salvar_DatosErr
    Dim strSql As String
    
    'Con.ConnectionString = cConnect
    'Con.Open
    
        'Con.BeginTrans
    'If RTrim(txtPorc_IGV.Text) = "" Then
    '    txtPorc_IGV.Text = "0"
    'End If
        strSql = "EXEC UP_MAN_ORDCOMP '" & _
        sTipo & "','" & _
        varSer_OrdComp & "','" & _
        Trim(txtCod_OrdComp.Text) & "','" & _
        Trim(txtCod_Proveedor.Text) & "','" & _
        Right(cboCod_CondVent.Text, 3) & "','" & _
        Right(cboCod_Descuento.Text, 3) & "','" & _
        Trim(txtPorc_IGV.Text) & "','" & _
        Right(cboCod_Moneda.Text, 3) & "','" & _
        Right(cboCod_LugEntr.Text, 3) & "','" & _
        Trim(txtObservaciones.Text) & "','" & _
        Right(cboCod_StaOrdComp.Text, 1) & "','" & _
        Right(cboCod_ClaOrdComp.Text, 2) & "','" & _
        dtpFec_Entrega_Inicio.Value & "','" & _
        dtpFec_Entrega_Fin.Value & "','" & _
        Right(cboCod_CenCost.Text, 16) & "','" & _
        Trim(txtCod_Grupo.Text) & "','" & _
        Right(cboCod_ProTex.Text, 2) & "','" & _
        Right(CboEstado.Text, 1) & "','" & txtAbr_Cliente.Tag & "',0,'po','MOTS','respon','" & txtAbr_Cliente.Tag & "','" & Trim(txtSer_ordcomp_Tinto.Text) & "','" & Trim(txtcod_ordcomp_Tinto) & "'"
        
        If sTipo = "I" Then
            Rs.Open strSql, cConnect, adOpenStatic
            
            If Rs.RecordCount > 0 Then
                optOrdCompra.Value = True
                txtSerOrdComp.Text = Rs(0)
                txtCodOrdComp.Text = Rs(1)
                CARGA_GRID
            End If
        Else
            'Con.Execute strSql
            Call ExecuteSQL(cConnect, strSql)
        End If
       
        'Con.CommitTrans
        
        Dim amensaje As New clsMessages
        amensaje.codigo = CodeMsg.kMESSAGE_INF_DATA_SAVE
        Informa "", amensaje
        
'        If sTipo = "I" Then
'            optOrdCompra.Value = True
'            Strsql = "SELECT MAX(Ser_OrdComp) FROM lg_ordcomp"
'            txtSerOrdComp.Text = DevuelveCampo(Strsql, cCONNECT)
'            Strsql = "SELECT MAX(Cod_OrdComp) FROM lg_ordcomp WHERE Ser_OrdComp ='" & Trim(txtSerOrdComp.Text) & "'"
'            txtCodOrdComp.Text = DevuelveCampo(Strsql, cCONNECT)
'            CARGA_GRID
'        End If
        
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub
Sub ELIMINAR_DATOS()
 Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
   
    Con.ConnectionString = cConnect
    Con.Open
    Con.BeginTrans
       
        strSql = "EXEC UP_MAN_ORDCOMP '" & _
        sTipo & "','" & _
        varSer_OrdComp & "','" & _
        Trim(txtCod_OrdComp.Text) & "','" & _
        Trim(txtCod_Proveedor.Text) & "','" & _
        Right(cboCod_CondVent.Text, 3) & "','" & _
        Right(cboCod_Descuento.Text, 3) & "','" & _
        Trim(txtPorc_IGV.Text) & "','" & _
        Right(cboCod_Moneda.Text, 3) & "','" & _
        Right(cboCod_LugEntr.Text, 3) & "','" & _
        Trim(txtObservaciones.Text) & "','" & _
        Right(cboCod_StaOrdComp.Text, 1) & "','" & _
        Right(cboCod_ClaOrdComp.Text, 2) & "','" & _
        dtpFec_Entrega_Inicio.Value & "','" & _
        dtpFec_Entrega_Fin.Value & "','" & _
        Right(cboCod_CenCost.Text, 16) & "','" & _
        Trim(txtCod_Grupo.Text) & "','" & _
        Right(cboCod_ProTex.Text, 2) & "','" & _
        Right(CboEstado.Text, 1) & "'"
        
        Con.Execute strSql
    
    Con.CommitTrans
    
    Dim amensaje As New clsMessages
    amensaje.codigo = CodeMsg.kMESSAGE_INF_DATA_DELETE
    Informa "", amensaje
    
Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Eliminar_Datos"

End Sub

'Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    Call CARGA_DATOS
'End Sub

Sub BUSCA_PROVEEDOR(Tipo As Integer, Ubic As Integer)
    Select Case Tipo
        Case 1:
                If Ubic = 1 Then
                    strSql = "SELECT Des_Proveedor FROM LG_PROVEEDOR WHERE Cod_Proveedor = '" & txtCodProveedor.Text & "'"
                    txtDesProveedor.Text = Trim(DevuelveCampo(strSql, cConnect))
                    'Strsql = "SELECT Cod_Proveedor FROM LG_PROVEEDOR WHERE Des_Proveedor = '" & txtDesProveedor.Text & "'"
                    'txtCodProveedor.Text = Trim(DevuelveCampo(Strsql, cCONNECT))
                    FunctBuscar.SetFocus
                Else
                    strSql = "SELECT Des_Proveedor FROM LG_PROVEEDOR WHERE Cod_Proveedor = '" & txtCod_Proveedor.Text & "'"
                    txtDes_Proveedor.Text = Trim(DevuelveCampo(strSql, cConnect))
                    'Strsql = "SELECT Cod_Proveedor FROM LG_PROVEEDOR WHERE Des_Proveedor = '" & txtDes_Proveedor.Text & "'"
                    'txtCod_Proveedor.Text = Trim(DevuelveCampo(Strsql, cCONNECT))
                    If cboCod_CondVent.Enabled = True Then
                        
                        'Aqui poscionaremos por defecto al cond, venta del prov
                        strSql = "SELECT Cod_CondVENT FROM LG_PROVEEDOR WHERE Cod_Proveedor='" & txtCod_Proveedor.Text & "'"
                        Call BuscaCombo(DevuelveCampo(strSql, cConnect), 2, cboCod_CondVent)
                        strSql = "SELECT Cod_Descuento FROM LG_PROVEEDOR WHERE Cod_Proveedor='" & txtCod_Proveedor.Text & "'"
                        Call BuscaCombo(DevuelveCampo(strSql, cConnect), 2, cboCod_Descuento)
                        
                        cboCod_CondVent.SetFocus
                    End If
                End If
                'FunctBuscar.SetFocus
        Case 2:
                Dim oTipo As New frmBusqGeneral
                Dim Rs As New ADODB.Recordset
                Set oTipo.oParent = Me
                If Ubic = 1 Then
                    oTipo.sQuery = "SELECT Cod_Proveedor as Código, Des_Proveedor as Descripción FROM LG_PROVEEDOR WHERE Des_Proveedor like '%" & Trim(txtDesProveedor.Text) & "%'"
                Else
                    oTipo.sQuery = "SELECT Cod_Proveedor as Código, Des_Proveedor as Descripción FROM LG_PROVEEDOR WHERE Des_Proveedor like '%" & Trim(txtDes_Proveedor.Text) & "%'"
                End If
                oTipo.CARGAR_DATOS
                
                oTipo.Show 1
                If codigo <> "" Then
                    If Ubic = 1 Then
                        txtCodProveedor.Text = Trim(codigo)
                        txtDesProveedor.Text = Trim(Descripcion)
                        FunctBuscar.SetFocus
                        codigo = ""
                        Descripcion = ""
                    Else
                        txtCod_Proveedor.Text = Trim(codigo)
                        txtDes_Proveedor.Text = Trim(Descripcion)
                        
                        'Aqui posicionaremos por defecto al cond, venta del prov
                        strSql = "SELECT Cod_CondVENT FROM LG_PROVEEDOR WHERE Cod_Proveedor='" & txtCod_Proveedor.Text & "'"
                        Call BuscaCombo(DevuelveCampo(strSql, cConnect), 2, cboCod_CondVent)
                        strSql = "SELECT Cod_Descuento FROM LG_PROVEEDOR WHERE Cod_Proveedor='" & txtCod_Proveedor.Text & "'"
                        Call BuscaCombo(DevuelveCampo(strSql, cConnect), 2, cboCod_Descuento)
                        
                        cboCod_CondVent.SetFocus
                    End If
                End If
                Set oTipo = Nothing
                Set Rs = Nothing
                
    End Select
End Sub

Sub BUSCA_GRUPO(Tipo As Integer)
    Dim NombreTabla As String
    Dim CodigoTabla As String
    strSql = "SELECT Tip_Grupo FROM LG_TIPREQ WHERE Cod_TipRequ='" & varCod_TipRequ & "'"
    If DevuelveCampo(strSql, cConnect) = "I" Then
        NombreTabla = "ES_GRUPOLOG"
        CodigoTabla = "Cod_GrupoLog"
    Else
        NombreTabla = "ES_GRUPOTEX"
        CodigoTabla = "Cod_GrupoTex"
    End If
    
    
    Select Case Tipo
        Case 1:
                strSql = "SELECT Des_Grupo FROM " & NombreTabla & " WHERE " & CodigoTabla & " = '" & txtCod_Grupo.Text & "'"
                TxtDes_Grupo.Text = Trim(DevuelveCampo(strSql, cConnect))
                
                'Strsql = "SELECT " & CodigoTabla & " FROM " & NombreTabla & " WHERE Des_Grupo = '" & txtDes_Grupo.Text & "'"
                'txtCod_Grupo.Text = Trim(DevuelveCampo(Strsql, cCONNECT))
        Case 2, 3:
        
                Dim oTipo As New frmBusqGeneral
                Dim Rs As New ADODB.Recordset
                Set oTipo.oParent = Me
                
                If Tipo = 2 Then
                    oTipo.sQuery = "SELECT " & CodigoTabla & " as Código, Des_Grupo as Descripción FROM " & NombreTabla & " WHERE Des_Grupo LIKE '" & Trim(TxtDes_Grupo.Text) & "%'"
                Else
                    oTipo.sQuery = "SELECT " & CodigoTabla & " as Código, Des_Grupo as Descripción FROM " & NombreTabla
                End If
                
                oTipo.CARGAR_DATOS
                oTipo.Show 1
                If codigo <> "" Then
                    txtCod_Grupo.Text = Trim(codigo)
                    TxtDes_Grupo.Text = Trim(Descripcion)
                    If cboCod_ProTex.Enabled Then
                        cboCod_ProTex.SetFocus
                        codigo = ""
                        Descripcion = ""
                    End If
                End If
                Set oTipo = Nothing
                Set Rs = Nothing
    End Select
End Sub

Sub BUSCA_ESTADO(Tipo As Integer)
    'Dim TipGrupo As Integer
    'Strsql = ""
    'TipGrupo = DevuelveCampo(Strsql, cCONNECT)
    
    Select Case Tipo
        Case 1:
                strSql = "SELECT Des_StaOrdComp FROM LG_STAORDCOMP WHERE  Cod_StaOrdComp = '" & txtCodStaOrdComp.Text & "'"
                txtDesStaOrdComp.Text = Trim(DevuelveCampo(strSql, cConnect))
                strSql = "SELECT Cod_StaOrdComp FROM LG_STAORDCOMP WHERE Des_StaOrdComp = '" & txtDesStaOrdComp.Text & "'"
                txtCodStaOrdComp.Text = Trim(DevuelveCampo(strSql, cConnect))
                FunctBuscar.SetFocus
        Case 2, 3:
        
                Dim oTipo As New frmBusqGeneral
                Dim Rs As New ADODB.Recordset
                Set oTipo.oParent = Me
                
                If Tipo = 2 Then
                    oTipo.sQuery = "SELECT Cod_StaOrdComp as Código, Des_StaOrdComp as Descripción FROM LG_STAORDCOMP WHERE Des_StaOrdComp LIKE '" & txtDesStaOrdComp.Text & "%'"
                Else
                    oTipo.sQuery = "SELECT Cod_StaOrdComp as Código, Des_StaOrdComp as Descripción FROM LG_STAORDCOMP"
                End If
                
                oTipo.CARGAR_DATOS
                oTipo.Show 1
                If codigo <> "" Then
                    txtCodStaOrdComp.Text = Trim(codigo)
                    txtDesStaOrdComp.Text = Trim(Descripcion)
                    FunctBuscar.SetFocus
                    codigo = ""
                    Descripcion = ""
                End If
                Set oTipo = Nothing
                Set Rs = Nothing
    End Select
End Sub

'Private Sub Command1_Click()
'
'
'    If optExcel.Value = True Then
'        ReporteExcel
'    Else
'        ReporteCrystal
'    End If
'
'    frmImp.Visible = False
'    FunctButt1.Visible = True
'    fraDetalle.Enabled = True
'    gexLista.Enabled = True
'    MantFunc1.Visible = True
'    FunctButt2.Visible = True
'End Sub


Private Sub Form_Load()
    opcion = 1
'    Call FormateaGrid(DGridLista)
    Call CARGA_COMBOS
    Call CARGA_GRID
    Call INHABILITA_DATOS
    
    Me.MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    Me.FunctBuscar.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    Me.FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    Me.FunctButt2.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    
   
    sTituliAbrOP = RTrim(DevuelveCampo("select Titulo_Abr_Orden from TG_Control", cConnect))
   ' optOP.Caption = sTituliAbrOP
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Call CARGA_GRID
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
'Dim vRow As Long
Dim vOrdCompBusq As String
Dim oo As Object
On Error GoTo AceptaError:

    Dim varOrigen As String
    Dim varTipFabrica As Integer

    Dim varCambioEstado As Integer
    If Rs_Lista.EOF And Rs_Lista.EOF Then
        MsgBox "Debe seleccionar un registro, para poder acceder a esta opción. Sirvase verificar", vbInformation, "Ordenes de Compra"
        Exit Sub
    End If
    vOrdCompBusq = gexLista.Value(gexLista.Columns("O.C.").Index)
    'vRow = gexLista.RowIndex(gexLista.Row)
    Select Case ActionName
        Case "IMPRESION":
'                        Strsql = "SELECT Origen FROM LG_PROVEEDOR WHERE Cod_Proveedor = '" & Trim(txtCod_Proveedor.Text) & "'"
'                        varOrigen = DevuelveCampo(Strsql, cConnect)
'
'
'                        Strsql = "SELECT Tip_Item FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"

'                        Set oo = CreateObject("excel.application")
'                        If varOrigen = "N" Then
'                            oo.Workbooks.Open vRuta & "\RptOCompra.xlt"
'                        Else
'                            oo.Workbooks.Open vRuta & "\RptOCompraIng.xlt"
'                        End If
'                        oo.Visible = True
'                        oo.Run "REPORTE", gexLista.Value(gexLista.Columns("Ser_OrdComp").Index), gexLista.Value(gexLista.Columns("Cod_OrdComp").Index), vusu, txtCod_Grupo.Text & " - " & TxtDes_Grupo.Text, DevuelveCampo(Strsql, cConnect), vemp, Me.varCod_TipRequ, cConnect
'                        Screen.MousePointer = vbNormal
'                        oo.Visible = True
'                        Set oo = Nothing
            ReporteExcel
            'FunctButt1.Visible = False
            'fraDetalle.Enabled = False
            'gexLista.Enabled = False
            'frmImp.Visible = True
            'FunctButt2.Visible = False
            'MantFunc1.Visible = False
        Case "CAMBIOESTADO":
                    varCambioEstado = MsgBox("¿Esta usted seguro de cambiar el estado al registro seleccionado?", vbInformation + vbYesNo, "Ordenes de Compra")
                    If varCambioEstado = vbYes Then
                        Call CAMBIO_ESTADO
                        Call CARGA_GRID
                    End If
        Case "DETALLE":
                    strSql = "SELECT Tip_Item FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
                    Load frmOrdCompItem
                    frmOrdCompItem.Caption = "Detalles de la Orden de Compra: " & gexLista.Value(gexLista.Columns("Ser_OrdComp").Index) & " - " & gexLista.Value(gexLista.Columns("Cod_OrdComp").Index)
                    frmOrdCompItem.varTip_Presentacion = DevuelveCampo(strSql, cConnect)
                    frmOrdCompItem.varSer_OrdComp = gexLista.Value(gexLista.Columns("Ser_OrdComp").Index)
                    frmOrdCompItem.varCod_OrdComp = gexLista.Value(gexLista.Columns("Cod_OrdComp").Index)
                    frmOrdCompItem.varCod_ClaOrdComp = gexLista.Value(gexLista.Columns("Cod_ClaOrdComp").Index)
                    frmOrdCompItem.varPorc_IGV = gexLista.Value(gexLista.Columns("I.G.V.").Index)
                    frmOrdCompItem.varCod_Descuento = gexLista.Value(gexLista.Columns("Cod_Descuento").Index)
                    frmOrdCompItem.varCod_Proveedor = gexLista.Value(gexLista.Columns("Cod_Proveedor").Index)
                    frmOrdCompItem.varCod_StaOrdComp = gexLista.Value(gexLista.Columns("Cod_StaOrdComp").Index)
                    frmOrdCompItem.varCod_GrupoTex = gexLista.Value(gexLista.Columns("Cod.Grupo").Index)
                    frmOrdCompItem.varDes_Grupo = Trim(TxtDes_Grupo.Text)
                    frmOrdCompItem.varCod_TipRequ = varCod_TipRequ
                    frmOrdCompItem.CARGA_GRID
                    frmOrdCompItem.Show 1
         Case "HILREQ"
                    MUESTRA_HILOS
         Case "ENTDET"
                    EntregasDet
         Case "NUMIMPORT"
                    If gexLista.RowCount = 0 Then Exit Sub
                    If Trim(gexLista.Value(gexLista.Columns("Num_Importacion").Index)) <> "" Then
                        MsgBox "Número de Importacion ya generado", vbExclamation, "Orden Compra"
                        Exit Sub
                    End If
                    Load FrmImportaciones
                    FrmImportaciones.txtSerOrdComp.Text = Me.gexLista.Value(gexLista.Columns("Ser_OrdComp").Index)
                    FrmImportaciones.txtCod_OrdComp.Text = Me.gexLista.Value(gexLista.Columns("Cod_OrdComp").Index)
                    FrmImportaciones.txtSerOrdComp.Enabled = False
                    FrmImportaciones.txtCod_OrdComp.Enabled = False
                    FrmImportaciones.Show 1
                    Set FrmImportaciones = Nothing
                    Call CARGA_GRID
        Case "OTTEJ"
            If gexLista.RowCount = 0 Then Exit Sub
            Call Genera_Ots
        Case "IMPORDSERV"
            Call Rep_OrdServ
        Case "POSTTENIDO"
            If gexLista.RowCount = 0 Then Exit Sub
            vTip_Item = Trim(DevuelveCampo("select tip_item from lg_claordcomp where cod_claordcomp='" & gexLista.Value(gexLista.Columns("Cod_ClaOrdComp").Index) & "'", cConnect))
            vTip_Presentacion = Trim(DevuelveCampo("select tip_presentacion from lg_claordcomp where cod_claordcomp='" & gexLista.Value(gexLista.Columns("Cod_ClaOrdComp").Index) & "'", cConnect))
            vCod_ProTex = Trim(DevuelveCampo("select isnull(Cod_ProTex,'') from lg_claordcomp where cod_claordcomp='" & gexLista.Value(gexLista.Columns("Cod_ClaOrdComp").Index) & "'", cConnect))
            
            If vTip_Item = "T" And vTip_Presentacion = "T" And vCod_ProTex = "" Then
                Load FrmEnviosPostTenido
                FrmEnviosPostTenido.sSer_OrdComp = gexLista.Value(gexLista.Columns("Ser_OrdComp").Index)
                FrmEnviosPostTenido.sCod_OrdComp = gexLista.Value(gexLista.Columns("cod_OrdComp").Index)
                FrmEnviosPostTenido.LblSer_OrdComp = gexLista.Value(gexLista.Columns("Ser_OrdComp").Index)
                FrmEnviosPostTenido.LblCod_Ordcomp = gexLista.Value(gexLista.Columns("cod_OrdComp").Index)
                FrmEnviosPostTenido.CARGA_GRID
                FrmEnviosPostTenido.Show vbModal
                Set FrmEnviosPostTenido = Nothing
            End If
    End Select
    'gexLista.Row = vRow
    Call gexLista.Find(3, jgexEqual, vOrdCompBusq)
    Exit Sub
AceptaError:
    ErrorHandler Err, "Aceptar"
    Screen.MousePointer = vbNormal
    Set oo = Nothing

End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
CERRAR_ORDCOMP
End Sub

Private Sub gexLista_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    Call CARGA_DATOS
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim eliminar As Integer
    Dim vRow As Long
    vRow = gexLista.Row
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            LIMPIAR_DATOS
            HABILITA_DATOS
            CboEstado.Enabled = True
            txtCod_Proveedor.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            gexLista.Enabled = False
        Case "MODIFICAR"
        
            If Not Rs_Lista.EOF And Not Rs_Lista.BOF Then
                If gexLista.Value(gexLista.Columns("Cod_StaOrdComp").Index) <> "P" Then
                    MsgBox "El estado del registro no permite la modificación. Sirvase verificar", vbInformation, "Ordenes de Compra"
                    Exit Sub
                End If
            End If
        
            sTipo = "U"
            HABILITA_DATOS
            txtCod_Proveedor.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            gexLista.Enabled = False
        Case "ELIMINAR"
        
            If Not Rs_Lista.EOF And Not Rs_Lista.BOF Then
                If gexLista.Value(gexLista.Columns("Cod_StaOrdComp").Index) <> "P" Then
                    MsgBox "El estado del registro no permite la eliminación. Sirvase verificar", vbInformation, "Ordenes de Compra"
                    Exit Sub
                End If
            End If
        
            eliminar = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Combinación-Detalle")
            If eliminar = vbYes Then
                sTipo = "D"
                If VALIDA_DATOS Then
                    Call ELIMINAR_DATOS
                    Call CARGA_GRID
                    gexLista.Row = vRow - 1
                    sTipo = ""
                End If
            End If
        Case "GRABAR"
            If VALIDA_DATOS Then
                Call SALVAR_DATOS
                Call CARGA_GRID
                Call INHABILITA_DATOS
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                gexLista.Enabled = True
                If sTipo = "I" Then
                    gexLista.MoveLast
                    strSql = DevuelveCampo("select origen from lg_proveedor where cod_proveedor='" & txtCod_Proveedor.Text & "'", cConnect)
                    If strSql = "E" Then
                        Pregunta = MsgBox("¿Desea generar Num. Importación?", vbYesNo)
                        If Pregunta = vbYes Then
                            Call FunctButt1_ActionClick(0, 0, "NUMIMPORT")
                        End If
                    End If
                Else
                    gexLista.Row = vRow
                End If
                sTipo = ""
            End If
        Case "DESHACER"
            Call LIMPIAR_DATOS
            Call CARGA_DATOS
            Call INHABILITA_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            gexLista.Enabled = True
            sTipo = ""
         Case "SALIR"
            Unload Me
    End Select
End Sub

Private Sub OpGrupo_Click()
    FraOrdComp.Visible = False
    FraProveedor.Visible = False
    FraEstado.Visible = False
  

    opcion = 4
End Sub

'Private Sub OpLog_Click()
'    OpTex.Value = False
'End Sub

Private Sub optEstado_Click()
    FraOrdComp.Visible = False
    FraProveedor.Visible = False
    'FraGrupo.Visible = False
    FraEstado.Visible = True
    'fraOP.Visible = False
    'txtCod_OrdPro.Text = ""
    
    txtCodStaOrdComp.Text = ""
    txtDesStaOrdComp.Text = ""
    txtCodStaOrdComp.SetFocus
    opcion = 3
End Sub



Private Sub optOP_Click()

    FraProveedor.Visible = False
    FraOrdComp.Visible = False

    FraEstado.Visible = False
    'txtCod_OrdPro.Text = ""
    'txtCod_OrdPro.SetFocus
    opcion = 5
End Sub

Private Sub optOrdCompra_Click()
    FraProveedor.Visible = False
    FraEstado.Visible = False
    'FraGrupo.Visible = False
    FraOrdComp.Visible = True
   ' fraOP.Visible = False
   ' txtCod_OrdPro.Text = ""
    'txtDes_OrdPro.Text = ""
    
    txtSerOrdComp.Text = ""
    txtCodOrdComp.Text = ""
    txtSerOrdComp.SetFocus
    opcion = 1
End Sub

Private Sub optProveedor_Click()
    FraEstado.Visible = False
    FraOrdComp.Visible = False
   ' FraGrupo.Visible = False
    FraProveedor.Visible = True
    'fraOP.Visible = False
    'txtCod_OrdPro.Text = ""
    
    txtCodProveedor.Text = ""
    txtDesProveedor.Text = ""
    txtCodProveedor.SetFocus
     
    opcion = 2
End Sub



Private Sub txtAbr_Cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
        BuscaCliente 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtcod_ordcomp_Tinto_LostFocus()
    txtcod_ordcomp_Tinto.Text = Format(txtcod_ordcomp_Tinto, "000000")
End Sub

Private Sub txtCodOrdComp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCodOrdComp.Text = Right("000000" & Trim(txtCodOrdComp.Text), 6)
        FunctBuscar.SetFocus
    End If
End Sub

Private Sub txtCodOrdComp_LostFocus()
    txtCodOrdComp.Text = Right("000000" & Trim(txtCodOrdComp.Text), 6)
    FunctBuscar.SetFocus
End Sub

Private Sub txtCodProveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCodProveedor.Text) <> "" Then
            txtCodProveedor.Text = Right("000000000000" & txtCodProveedor.Text, 12)
            Call BUSCA_PROVEEDOR(1, 1)
        End If
    End If
End Sub


Private Sub txtDesProveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDesProveedor.Text) <> "" Then
            Call BUSCA_PROVEEDOR(2, 1)
        End If
    End If
End Sub

Private Sub Txtcod_Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_Proveedor.Text) <> "" Then
            txtCod_Proveedor.Text = Right("000000000000" & txtCod_Proveedor.Text, 12)
            Call BUSCA_PROVEEDOR(1, 2)
        End If
    End If
End Sub

Private Sub TxtDes_Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDes_Proveedor.Text) <> "" Then
            Call BUSCA_PROVEEDOR(2, 2)
        End If
    End If
End Sub

Private Sub txtCodStaOrdComp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCodStaOrdComp.Text) <> "" Then
            Call BUSCA_ESTADO(1)
        End If
    End If
End Sub
Private Sub txtDesStaOrdComp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDesStaOrdComp.Text) <> "" Then
            Call BUSCA_ESTADO(2)
        End If
    End If
End Sub

Private Sub cmdBusEstado_Click()
    Call BUSCA_ESTADO(3)
End Sub

Private Sub txtCod_Grupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_Grupo.Text) <> "" Then
            txtCod_Grupo.Text = Right("00000000" & txtCod_Grupo.Text, 8)
            Call BUSCA_GRUPO(1)
        End If
    End If

End Sub

Private Sub txtDes_Grupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(TxtDes_Grupo.Text) <> "" Then
            Call BUSCA_GRUPO(2)
        End If
    End If

End Sub

Private Sub cmdBuscaGrupo_Click()
    Call BUSCA_GRUPO(3)
End Sub

Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
        BuscaCliente 2
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtSer_ordcomp_Tinto_LostFocus()
txtSer_ordcomp_Tinto.Text = Format(txtSer_ordcomp_Tinto, "000")
End Sub


Private Sub txtSerOrdComp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSerOrdComp.Text = Right("000" & Trim(txtSerOrdComp.Text), 3)
        txtCodOrdComp.SetFocus
    End If
End Sub

Private Sub txtSerOrdComp_LostFocus()
    txtSerOrdComp.Text = Right("000" & Trim(txtSerOrdComp.Text), 3)
End Sub

Private Sub CERRAR_ORDCOMP()
    Dim Con As New ADODB.Connection
    Dim Message As Integer
    On Error GoTo Salvar_DatosErr
    Dim strSql As String
    
    Con.ConnectionString = cConnect
    Con.Open
    Message = MsgBox("¿Esta usted seguro que desea Abrir/Cerrar la O/C seleccionada?", vbInformation + vbYesNo, "Orden de Compra")
    If Message = vbYes Then
        Con.BeginTrans

        strSql = "EXEC UP_MAN_ORDCOMP_ABRIRCERRAR '" & _
        varSer_OrdComp & "','" & _
        Trim(txtCod_OrdComp.Text) & "','" & _
        vusu & "'"
        
        Con.Execute strSql

        Con.CommitTrans
        
        MsgBox "La Orden de Compra se Modificó satisfactoriamente", vbOKOnly, "Ordenes de Compra"
        CARGA_GRID
    End If
        
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "CERRAR_ORDCOMP"
End Sub

Sub MUESTRA_HILOS()
On Error GoTo Muestra_DatosErr
Dim Rs As New ADODB.Recordset

Rs.Open "select * from lg_claordcomp where cod_claOrdComp='" & gexLista.Value(gexLista.Columns("Cod_ClaOrdComp").Index) & "'", cConnect, adOpenStatic
If Rs.RecordCount Then
    If Rs.Fields("Tip_Item").Value = "T" And Rs.Fields("Tip_Presentacion").Value = "C" Then
        frmHiladosRequeridos.varSer_OrdComp = gexLista.Value(gexLista.Columns("Ser_OrdComp").Index)
        frmHiladosRequeridos.varCod_OrdComp = gexLista.Value(gexLista.Columns("Cod_OrdComp").Index)
        frmHiladosRequeridos.varCod_Proveedor = gexLista.Value(gexLista.Columns("Cod_Proveedor").Index)
        frmHiladosRequeridos.Show 1
    End If
End If

Set Rs = Nothing
    Exit Sub
Muestra_DatosErr:
    Set Rs = Nothing
    ErrorHandler Err, "MUESTRA_HILOS"
End Sub




Private Sub VerificaFabrica(ByRef objFabrica As TextBox, ByRef objNombreFabrica As TextBox)
    Dim sSQl As String
    Dim iRet As String
    
    sSQl = "SELECT count(*) FROM TG_Fabrica "
    iRet = DevuelveCampo(sSQl, cConnect)
    If iRet = 1 Then
        sSQl = "SELECT Cod_Fabrica FROM TG_Fabrica "
        objFabrica.Text = DevuelveCampo(sSQl, cConnect)
        
        sSQl = "SELECT Nom_Fabrica FROM TG_Fabrica "
        objNombreFabrica.Text = DevuelveCampo(sSQl, cConnect)
        objFabrica.Enabled = False
        objNombreFabrica.Enabled = False
        
    End If
End Sub








Sub GENERA_NUMIMPORTACION()
Dim amensaje As New clsMessages
    On Error GoTo Salvar_DatosErr
    
        strSql = "EXEC UP_MAN_ORDCOMP '" & _
        sTipo & "','" & _
        varSer_OrdComp & "','" & _
        Trim(txtCod_OrdComp.Text) & "','" & _
        Trim(txtCod_Proveedor.Text) & "','" & _
        Right(cboCod_CondVent.Text, 3) & "','" & _
        Right(cboCod_Descuento.Text, 3) & "','" & _
        Trim(txtPorc_IGV.Text) & "','" & _
        Right(cboCod_Moneda.Text, 3) & "','" & _
        Right(cboCod_LugEntr.Text, 3) & "','" & _
        Trim(txtObservaciones.Text) & "','" & _
        Right(cboCod_StaOrdComp.Text, 1) & "','" & _
        Right(cboCod_ClaOrdComp.Text, 2) & "','" & _
        dtpFec_Entrega_Inicio.Value & "','" & _
        dtpFec_Entrega_Fin.Value & "','" & _
        Right(cboCod_CenCost.Text, 16) & "','" & _
        Trim(txtCod_Grupo.Text) & "','" & _
        Right(cboCod_ProTex.Text, 2) & "','" & _
        Right(CboEstado.Text, 1) & "'"
        
        Call ExecuteSQL(cConnect, strSql)
        
        amensaje.codigo = CodeMsg.kMESSAGE_INF_DATA_SAVE
        Informa "", amensaje
    Exit Sub
Salvar_DatosErr:
    ErrorHandler Err, "GENERA_NUMIMPORTACION"
End Sub



Sub ReporteExcel()
    Dim varOrigen As String
    Dim varTipFabrica As Integer
On Error GoTo ErrReporte
        strSql = "SELECT Origen FROM LG_PROVEEDOR WHERE Cod_Proveedor = '" & Trim(txtCod_Proveedor.Text) & "'"
        varOrigen = DevuelveCampo(strSql, cConnect)

        strSql = "SELECT Tip_Item FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & Right(cboCod_ClaOrdComp.Text, 2) & "'"
        Dim oo As Object
        Set oo = CreateObject("excel.application")

        If varOrigen = "N" Then
            oo.Workbooks.Open vRuta & "\RptOCompra.xlt"
            'oo.Workbooks.Open App.Path & "\RptOCompra.xlt"
        Else
            oo.Workbooks.Open vRuta & "\RptOCompraIng.xlt"
            'oo.Workbooks.Open App.Path & "\RptOCompraIng.xlt"
        End If
        oo.Visible = True

        oo.Run "REPORTE", gexLista.Value(gexLista.Columns("Ser_OrdComp").Index), gexLista.Value(gexLista.Columns("Cod_OrdComp").Index), vusu, txtCod_Grupo.Text & " - " & TxtDes_Grupo.Text, DevuelveCampo(strSql, cConnect), vemp, Me.varCod_TipRequ, cConnect
        Screen.MousePointer = vbNormal
        oo.Visible = True
        Set oo = Nothing
        
Exit Sub
ErrReporte:
Set oo = Nothing
ErrorHandler Err, "Reporte Crystal"
End Sub



Private Sub EntregasDet()
Dim oo As Object
    
    If gexLista.RowCount = 0 Then Exit Sub
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\EntregasDet.xlt"
    oo.Visible = True
    'oo.Run "PRUEBA", CStr(varCod_Cliente), CStr(varCod_Fabrica), CStr(txtCod_EstCli.Text), CStr(txtAbr_Cliente.Text & " - " & txtNom_Cliente.Text), CStr(txtAbr_Fabrica.Text & " - " & txtNom_Fabrica.Text), CStr(txtCod_EstCli.Text & " - " & txtDes_EstCli.Text), cCONNECT
    oo.Run "REPORTE", gexLista.Value(gexLista.Columns("Ser_OrdComp").Index), gexLista.Value(gexLista.Columns("Cod_OrdComp").Index), "", cConnect
    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
End Sub

Sub Rep_OrdServ()
Dim oo As Object
Dim Rs As New ADODB.Recordset
On Error GoTo ErrReporte

strSql = "Tx_Muestra_Cabecera_OrdenServicio_Interno '" & gexLista.Value(gexLista.Columns("Ser_OrdComp").Index) & "','" & gexLista.Value(gexLista.Columns("Cod_OrdComp").Index) & "'"

Set Rs = Nothing
Rs.CursorLocation = adUseClient
Rs.Open strSql, cConnect

Set oo = CreateObject("excel.application")
oo.Workbooks.Open vRuta & "\RptOrdenServicioInterno.XLT"
oo.Visible = True
oo.DisplayAlerts = False
oo.Run "REPORTE", Rs.DataSource, gexLista.Value(gexLista.Columns("Ser_OrdComp").Index), gexLista.Value(gexLista.Columns("Cod_OrdComp").Index), cConnect
Screen.MousePointer = vbNormal

Set Rs = Nothing
Set oo = Nothing
Exit Sub
ErrReporte:
    Set Rs = Nothing
    ErrorHandler Err, "Rep. Orden Servicio"
End Sub

Sub Genera_Ots()
On Error GoTo errGenera_Ots
Dim Mensa As Variant
Mensa = MsgBox("¿Está seguro de Generar OTs?", vbYesNo)
If Mensa = vbNo Then Exit Sub
strSql = "TJ_GENERACION_OT_TEJEDURIA_AUTOMATICA_LG_ORDCOMPITEM '" & gexLista.Value(gexLista.Columns("Ser_OrdComp").Index) & "','" & gexLista.Value(gexLista.Columns("Cod_OrdComp").Index) & "'"
ExecuteSQL cConnect, strSql
Call CARGA_GRID
Exit Sub
errGenera_Ots:
    MsgBox Err.Description, vbCritical
End Sub


Public Sub BuscaCliente(opcion As Integer)
On Error GoTo Fin
Dim sSQl As String
Dim Cliente_control As String
Dim iCol As Long
    
    txtAbr_Cliente = Trim(txtAbr_Cliente)
    txtNom_Cliente = Trim(txtNom_Cliente)
    strSql = "SELECT Abr_Cliente, Nom_Cliente, Cod_Cliente_Tex, Res_Cliente, Num_Ruc, " & _
    "Cod_Moneda, Peso_por_Rollo, Flg_ClientePropio FROM TX_CLIENTE WHERE "
    
    Select Case opcion
    Case 1: strSql = strSql & "Abr_Cliente like '" & txtAbr_Cliente & IIf(txtAbr_Cliente = "", "%", "") & "'"
    Case 2: strSql = strSql & "Nom_Cliente like '%" & txtNom_Cliente & "%'"
    End Select
    strSql = strSql & " ORDER BY Abr_Cliente"
    
    txtAbr_Cliente = ""
    txtAbr_Cliente.Tag = ""
    txtNom_Cliente = ""
    
    With frmBusqGeneral_Cliente
        Set .oParent = Me
        .sQuery = strSql
        .CARGAR_DATOS
        codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("Abr_Cliente").Caption = "Abrev."
        .DGridLista.Columns("Abr_Cliente").Width = 700
        .DGridLista.Columns("Nom_Cliente").Caption = "Nombre Cliente"
        .DGridLista.Columns("Nom_Cliente").Width = 5000
        For iCol = 3 To .DGridLista.Columns.Count
            .DGridLista.Columns(iCol).Visible = False
        Next iCol
        
        
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If codigo <> "" And rstAux.RecordCount > 0 Then
            txtAbr_Cliente.Tag = Trim(rstAux!cod_cliente_tex)
            txtAbr_Cliente = Trim(rstAux!Abr_Cliente)
            txtNom_Cliente = Trim(rstAux!Nom_Cliente)
            scliente = Trim(rstAux!cod_cliente_tex)
        End If
        sSQl = "select Flg_CentroCosto from tx_cliente WHERE Cod_Cliente_Tex = '" & scliente & "'"
        Cliente_control = DevuelveCampo(sSQl, cConnect)
        
 
        
        
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
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Cliente (" & opcion & ")"
End Sub


