VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrdCompItemAddN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Orden de Compra"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDetalle 
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
      Height          =   3570
      Left            =   15
      TabIndex        =   1
      Top             =   2895
      Width           =   7425
      Begin VB.Frame fraPreRep 
         Caption         =   "Consulta Adicional del Item"
         Height          =   1470
         Left            =   2295
         TabIndex        =   43
         Top             =   -60
         Visible         =   0   'False
         Width           =   3780
         Begin VB.CommandButton cmdSalir 
            Caption         =   "Salir"
            Height          =   255
            Left            =   2955
            TabIndex        =   48
            Top             =   1170
            Width           =   720
         End
         Begin VB.TextBox txtComentario 
            Enabled         =   0   'False
            Height          =   465
            Left            =   1440
            MultiLine       =   -1  'True
            TabIndex        =   46
            Top             =   660
            Width           =   2235
         End
         Begin VB.TextBox txtPreRep 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   44
            Text            =   "0.00000"
            Top             =   225
            Width           =   1515
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Comentario :"
            Height          =   195
            Left            =   180
            TabIndex        =   47
            Top             =   675
            Width           =   885
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Precio Repos. $ :"
            Height          =   195
            Left            =   165
            TabIndex        =   45
            Top             =   285
            Width           =   1230
         End
      End
      Begin VB.TextBox TxtCodProv 
         Height          =   285
         Left            =   5115
         TabIndex        =   49
         Top             =   2625
         Width           =   1305
      End
      Begin VB.CommandButton cmdHelpPrecio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6930
         Picture         =   "frmOrdCompItemAddN.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Precios"
         Top             =   1620
         Width           =   375
      End
      Begin VB.CommandButton CmdHelp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6510
         Picture         =   "frmOrdCompItemAddN.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Compras"
         Top             =   1620
         Width           =   375
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   465
         Left            =   1305
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   2940
         Width           =   5580
      End
      Begin VB.TextBox txtDes_Color 
         Height          =   315
         Left            =   2745
         TabIndex        =   13
         Top             =   960
         Width           =   4125
      End
      Begin VB.CommandButton cmdBuscaColor 
         Caption         =   "..."
         Height          =   300
         Left            =   2460
         TabIndex        =   12
         Top             =   960
         Width           =   300
      End
      Begin VB.TextBox txtCod_Color 
         Height          =   285
         Left            =   1300
         TabIndex        =   11
         Top             =   960
         Width           =   1155
      End
      Begin VB.ComboBox cboCod_ItemProv 
         Height          =   315
         Left            =   1300
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   2610
         Width           =   2700
      End
      Begin VB.ComboBox cboCod_Descuento 
         Height          =   315
         Left            =   1300
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1950
         Width           =   1300
      End
      Begin VB.TextBox txtCod_UniMed 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1300
         TabIndex        =   15
         Top             =   1290
         Width           =   1305
      End
      Begin MSComCtl2.DTPicker dtpFec_Entrega_Fin 
         Height          =   315
         Left            =   5115
         TabIndex        =   29
         Top             =   2280
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         Format          =   83951617
         CurrentDate     =   37270
      End
      Begin MSComCtl2.DTPicker dtpFec_Entrega_Inicio 
         Height          =   315
         Left            =   1300
         TabIndex        =   27
         Top             =   2280
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         Format          =   83951617
         CurrentDate     =   37270
      End
      Begin VB.TextBox txtPorc_IGV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5115
         TabIndex        =   25
         Top             =   1950
         Width           =   1305
      End
      Begin VB.TextBox txtPre_Unitario 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5115
         TabIndex        =   21
         Text            =   "0.00000"
         Top             =   1620
         Width           =   1305
      End
      Begin VB.TextBox txtCan_Comprada 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1300
         TabIndex        =   19
         Text            =   "0"
         Top             =   1620
         Width           =   1300
      End
      Begin VB.ComboBox cboCod_Destino 
         Height          =   315
         Left            =   5115
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1290
         Width           =   2000
      End
      Begin VB.ComboBox cboCod_Talla 
         Height          =   315
         Left            =   5115
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   615
         Width           =   2000
      End
      Begin VB.ComboBox cboCod_Comb 
         Height          =   315
         Left            =   1300
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   630
         Width           =   2000
      End
      Begin VB.CommandButton cmdBuscaCodigo 
         Caption         =   "..."
         Height          =   300
         Left            =   2460
         TabIndex        =   4
         Top             =   300
         Width           =   300
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1300
         MaxLength       =   10
         TabIndex        =   3
         Top             =   300
         Width           =   1155
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   2745
         MaxLength       =   50
         TabIndex        =   5
         Top             =   300
         Width           =   4125
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Prov. :"
         Height          =   195
         Left            =   4095
         TabIndex        =   50
         Top             =   2655
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones :"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   3000
         Width           =   1155
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "F. Entrega Fin :"
         Height          =   195
         Left            =   4095
         TabIndex        =   28
         Top             =   2355
         Width           =   1080
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "F. Entrega Ini."
         Height          =   195
         Left            =   180
         TabIndex        =   26
         Top             =   2350
         Width           =   990
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "P.Unitario :"
         Height          =   195
         Left            =   4095
         TabIndex        =   20
         Top             =   1710
         Width           =   780
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor :"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Top             =   2680
         Width           =   825
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "I.G.V. :"
         Height          =   195
         Left            =   4095
         TabIndex        =   24
         Top             =   2055
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Descuento :"
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   2055
         Width           =   870
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cant. Comprar :"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   1710
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Talla/Medida:"
         Height          =   195
         Left            =   4095
         TabIndex        =   8
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Destino :"
         Height          =   195
         Left            =   4095
         TabIndex        =   16
         Top             =   1365
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "U. Medida :"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   1335
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Color :"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   1035
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Combinación :"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   700
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   370
         Width           =   585
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
      Height          =   2850
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   7380
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2565
         Left            =   90
         TabIndex        =   40
         Top             =   180
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4524
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
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
         ColumnCount     =   14
         BeginProperty Column00 
            DataField       =   "ITEM"
            Caption         =   "Item"
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
         BeginProperty Column01 
            DataField       =   "Combinacion"
            Caption         =   "Combinación"
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
         BeginProperty Column02 
            DataField       =   "Color"
            Caption         =   "Color"
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
         BeginProperty Column03 
            DataField       =   "Cod_Talla"
            Caption         =   "Talla/Medida"
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
            DataField       =   "Destino"
            Caption         =   "Destino"
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
            DataField       =   "Estilo_Cli"
            Caption         =   "Est. Cliente"
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
            DataField       =   "Cod_UniMed"
            Caption         =   "Uni. Med."
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
         BeginProperty Column07 
            DataField       =   "Pre_Unitario"
            Caption         =   "P. Unit."
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
         BeginProperty Column08 
            DataField       =   "Can_Requerida"
            Caption         =   "C. Requerida"
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
         BeginProperty Column09 
            DataField       =   "Can_Comprada"
            Caption         =   "C. Comprada"
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
         BeginProperty Column10 
            DataField       =   "Can_Recibida"
            Caption         =   "C. Recibida"
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
         BeginProperty Column11 
            DataField       =   "Fac_EquiProv"
            Caption         =   "F. Equivalencia"
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
         BeginProperty Column12 
            DataField       =   "Cod_ItemProv"
            Caption         =   "Item Prov."
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
         BeginProperty Column13 
            DataField       =   "Cod_Prov"
            Caption         =   "Cod. Prov."
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
            MarqueeStyle    =   2
            BeginProperty Column00 
               ColumnWidth     =   2055.118
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column13 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   540
      TabIndex        =   34
      Top             =   6450
      Width           =   1965
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   15
         Picture         =   "frmOrdCompItemAddN.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Primero"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   495
         Picture         =   "frmOrdCompItemAddN.frx":0786
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Anterior"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   975
         Picture         =   "frmOrdCompItemAddN.frx":08F8
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1455
         Picture         =   "frmOrdCompItemAddN.frx":0A6A
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Ultimo"
         Top             =   120
         Width           =   495
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   3045
      TabIndex        =   35
      Top             =   6540
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmOrdCompItemAddN.frx":0BDC
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmOrdCompItemAddN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public codigo As String
Public Descripcion As String
Public Cod_Anterior As String

Dim strSql As String
Dim Strsql_X As String
Dim Rs_Lista As ADODB.Recordset
Dim NombreTabla, CodigoTabla, DesTabla As String
Dim DesTabla2 As String
Dim sTipo As String
'Definicion de variables que seran pasadas por nuestro master
Public varSer_OrdComp, varCod_OrdComp, varSec_OrdComp As String
Public varTip_Presentacion, varCod_ClaOrdComp, varCod_Proveedor As String
Public varCod_StaOrdComp As String
Public varCod_Descuento As String
Public varPorc_IGV As Double
Public varCodProv As String
Public oParent As Object
Dim varRepetir_Precio As String
Dim sTip_Item As String
Dim Cadena  As String

Dim Rsx_Textil As ADODB.Recordset
Dim Xcod_cliente_tex, Xcod_ordcomp_tex, Xser_ordcomp_tex As String
Dim flg_requerimientos_Tex As String
Private blSWx As Boolean


Sub BUSCA_CODIGO(Tipo As Integer)

    Select Case varTip_Presentacion
        Case "I":
                    NombreTabla = "LG_ITEM"
                    CodigoTabla = "Cod_Item"
                    DesTabla = "Des_Item"
                    DesTabla2 = ""
        Case "T":
                    NombreTabla = "TX_TELA"
                    CodigoTabla = "Cod_Tela"
                    DesTabla = "Des_Tela"
                    DesTabla2 = ""
       Case "H":
                    NombreTabla = "IT_HILADO"
                    CodigoTabla = "Cod_HilTel"
                    DesTabla = "Des_HilTel"
                    DesTabla2 = ", COD_HILADO_ESTRUCTURADO"
    End Select
    
    Select Case Tipo
        Case 1:
                
                strSql = "SELECT " & DesTabla & " FROM " & NombreTabla & " WHERE " & CodigoTabla & " LIKE '" & txtCodigo.Text & "%'"
                txtDescripcion.Text = Trim(DevuelveCampo(strSql, cConnect))
                'Strsql = "SELECT " & CodigoTabla & " FROM " & NombreTabla & " WHERE " & DesTabla & " = '" & txtDescripcion.Text & "'"
                'txtCodigo.Text = Trim(DevuelveCampo(Strsql, cCONNECT))
        Case 2, 3, 4, 5:
        
                Dim oTipo As New frmBusqGeneral
                Dim Rs As New ADODB.Recordset
                Set oTipo.oParent = Me
                
                If Tipo = 2 Then
                    oTipo.sQuery = "SELECT " & CodigoTabla & " as Código, " & DesTabla & " as Descripción " & DesTabla2 & "   FROM " & NombreTabla & " WHERE " & DesTabla & " LIKE '%" & txtDescripcion.Text & "%'"
                End If
                
                If Tipo = 3 Then
                    oTipo.sQuery = "SELECT " & CodigoTabla & " as Código, " & DesTabla & " as Descripción " & DesTabla2 & "   FROM " & NombreTabla
                End If
                
                If Tipo = 4 Then
                    oTipo.sQuery = "SELECT " & CodigoTabla & " as Código, " & DesTabla & " as Descripción " & DesTabla2 & "   FROM " & NombreTabla & " WHERE " & CodigoTabla & " LIKE '" & txtCodigo.Text & "%'"
                End If
            
                If Tipo = 5 Then
                    oTipo.sQuery = "Exec tj_muestra_requerimientos_cliente_textil_oc '" & Xcod_cliente_tex & "','" & Xser_ordcomp_tex & "','" & Xcod_ordcomp_tex & "'"
                End If
                
                
                oTipo.CARGAR_DATOS
                'oTipo.DGridLista.Columns(2).Width = 5000
               ' oTipo.DGridLista.Columns(1).Width = 1000
                
                oTipo.Show 1
                If codigo <> "" Then
                    txtCodigo.Text = Trim(codigo)
                    txtDescripcion.Text = Trim(Descripcion)
                    Cod_Anterior = Trim(Cod_Anterior)
                    'cboCod_ProTex.SetFocus
                End If
                
                codigo = ""
                Descripcion = ""
                
                Set oTipo = Nothing
                Set Rs = Nothing
                
    End Select
    Call CARGA_COMBOS
    
End Sub

Sub BUSCA_COLOR(Tipo As Integer)
On Error GoTo fin

    Dim varProvTip_Presentacion As String
    Dim Busca As Boolean
    
    Busca = False
    strSql = "SELECT Tip_Presentacion FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp='" & varCod_ClaOrdComp & "'"
    varProvTip_Presentacion = DevuelveCampo(strSql, cConnect)
    If (varTip_Presentacion = "T" Or varTip_Presentacion = "H") And varProvTip_Presentacion = "T" Then
        Busca = True
    End If
    If varTip_Presentacion = "I" Then
        Busca = True
    End If
    
    Select Case Tipo
        Case 1:
                
                strSql = "SELECT Des_Color FROM LB_COLOR WHERE Cod_Color='" & Trim(txtCod_Color.Text) & "'"
                txtDes_Color.Text = Trim(DevuelveCampo(strSql, cConnect))
                'Strsql = "SELECT " & CodigoTabla & " FROM " & NombreTabla & " WHERE " & DesTabla & " = '" & txtDescripcion.Text & "'"
                'txtCodigo.Text = Trim(DevuelveCampo(Strsql, cCONNECT))
        Case 2, 3:
        
                Dim oTipo As New frmBusqGeneral
                Dim Rs As New ADODB.Recordset
                Set oTipo.oParent = Me
                
                If Tipo = 2 Then
                    oTipo.sQuery = "SELECT Cod_Color as Código, Des_Color as Descripción FROM LB_COLOR WHERE Des_Color LIKE '%" & Trim(txtDes_Color.Text) & "%'"
                Else
                    oTipo.sQuery = "SELECT Cod_Color as Código, Des_Color as Descripción FROM LB_COLOR"
                End If
                
                oTipo.CARGAR_DATOS
                oTipo.Show 1
                If codigo <> "" Then
                    txtCod_Color.Text = Trim(codigo)
                    txtDes_Color.Text = Trim(Descripcion)
                    'cboCod_ProTex.SetFocus
                End If
                Set oTipo = Nothing
                Set Rs = Nothing
                
                codigo = ""
                Descripcion = ""
    End Select
Exit Sub
fin:

MsgBox " inconvenientes para crear ordenes", vbCritical + vbOKOnly, "Mensaje"
    'Call CARGA_COMBOS
    
End Sub

Private Sub cmdBuscaCodigo_Click()
    Call BUSCA_CODIGO(3)
End Sub

Private Sub cmdBuscaColor_Click()
    Call BUSCA_COLOR(3)
End Sub

Private Sub cmdFirst_Click()
    If Not Rs_Lista.BOF Then
        Rs_Lista.MoveFirst
    End If
End Sub

Private Sub CmdHelp_Click()
    strSql = "Exec SM_AYUDA_PRECIO_ULT_COMPRAS '" & Trim(txtCodigo.Text) & "','" & Trim(Right(cboCod_Comb.Text, 3)) & "','" & txtCod_Color & "','" & Trim(Right(Me.cboCod_Talla.Text, 10)) & "','" & varSer_OrdComp & "','" & varCod_OrdComp & "'"
    
    Set frmBusqGeneral.oParent = Me
    frmBusqGeneral.sQuery = strSql
    frmBusqGeneral.CARGAR_DATOS
    frmBusqGeneral.Show 1
    If Descripcion <> "" Then
        txtPre_Unitario = Descripcion
    End If
    codigo = ""
    Descripcion = ""

End Sub

Private Sub cmdHelpPrecio_Click()
    Dim dPrecio As Double
    Dim sComentario As String
    Dim mRs As ADODB.Recordset
    
    strSql = "SELECT Tip_Item FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp = '" & varCod_ClaOrdComp & "'"
    sTip_Item = Trim(DevuelveCampo(strSql, cConnect))
    
    If sTip_Item = "I" Then
        strSql = "SM_AYUDA_PRECIO_REPOSIC_COM '" & txtCodigo.Text & "', '" & varCod_ClaOrdComp & "'"
        Set mRs = GetRecordset(cConnect, strSql)
        If Not mRs.EOF And Not mRs.BOF Then
            txtPreRep.Text = FixNulos(mRs!Rep_PreDol, vbDouble)
            txtComentario.Text = FixNulos(mRs!Comentario, vbString)
            fraPreRep.Visible = True
        End If
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
    
    txtCodigo.Text = ""
    txtDescripcion.Text = ""
    'txtCod_UniMed.Text = ""
    txtCan_Comprada.Text = "0.00"
    txtPre_Unitario.Text = "0.00000"
    txtCod_Color.Text = ""
    txtDes_Color.Text = ""
    txtObservaciones.Text = ""
    TxtCodProv.Text = ""
    
    cboCod_Descuento.ListIndex = -1
    cboCod_ItemProv.ListIndex = -1
    cboCod_Comb.ListIndex = -1
    cboCod_Talla.ListIndex = -1
    cboCod_Destino.ListIndex = -1

    dtpFec_Entrega_Fin.Value = Date
    dtpFec_Entrega_Inicio.Value = Date
    

    'Aqui llenamos a los valores por defecto
    txtPorc_IGV.Text = Format(varPorc_IGV, "#####0.00")
    Call BuscaCombo(varCod_Descuento, 2, cboCod_Descuento)
    
    varRepetir_Precio = ""
    
End Sub

Sub CARGA_COMBOS()

'    If varTip_Presentacion = "T" Or varTip_Presentacion = "H" Then
'
'    End If

    If varTip_Presentacion = "I" Or varTip_Presentacion = "T" Then
            
        'Aqui llenamos las combinaciones
        If varTip_Presentacion = "I" Then
            strSql = "SELECT Des_Comb + SPACE(100) + Cod_Comb  FROM LG_ITEMCOMB WHERE Cod_Item = '" & Trim(txtCodigo.Text) & "'"
        Else
            strSql = "SELECT Des_Comb + SPACE(100) + Cod_Comb  FROM TX_TELACOMB WHERE Cod_Tela = '" & Trim(txtCodigo.Text) & "'"
        End If
        Call LlenaCombo(cboCod_Comb, strSql, cConnect)
        cboCod_Comb.AddItem "[Ninguna]                 "
        
        'Aqui llenamos las Tallas
        LlenaComboTalla
    End If
    
    If varTip_Presentacion = "I" Then
        'Aqui llenamos los Destinos
        strSql = "SELECT Des_Destino + SPACE(100) + Cod_Destino FROM TG_DESTINO"
        Call LlenaCombo(cboCod_Destino, strSql, cConnect)
        cboCod_Destino.AddItem "[Ninguna]     "
        
        'Aqui llenamos la Unidad de MEdida
        strSql = "SELECT Cod_UniMed FROM LG_ITEM WHERE COD_ITEM='" & txtCodigo.Text & "'"
        txtCod_UniMed.Text = DevuelveCampo(strSql, cConnect)
        
    End If
        
    If varTip_Presentacion = "T" Then
        'Aqui llenamos la Unidad de MEdida
        strSql = "SELECT Cod_UniMed FROM TX_TELA WHERE COD_tELA='" & txtCodigo.Text & "'"
        txtCod_UniMed.Text = DevuelveCampo(strSql, cConnect)
        
    End If
    
    strSql = "SELECT Cod_ItemProv + ' : ' + CONVERT(VARCHAR, Fac_EquiProv) + ' - ' + Cod_UniMedProv FROM LG_ITEMPROV WHERE Cod_Item='" & Trim(txtCodigo.Text) & "' AND Cod_Proveedor='" & varCod_Proveedor & "'"
    Call LlenaCombo(cboCod_ItemProv, strSql, cConnect)
    
End Sub

Sub LlenaComboTalla()
On Error GoTo hand
Dim Rs As New ADODB.Recordset

Rs.Open "select * from lg_itemmed where cod_item='" & txtCodigo.Text & "'", cConnect, adOpenStatic

 If varTip_Presentacion = "I" Or varTip_Presentacion = "T" Then
    If Rs.RecordCount > 0 Then
        strSql = "select descripcion + space(100) + cod_medida from lg_itemmed where cod_item='" & txtCodigo.Text & "'"
    Else
        strSql = "SELECT Cod_Talla + space(100) + Cod_Talla FROM TG_TALLA"
    End If
    
    Call LlenaCombo(cboCod_Talla, strSql, cConnect)
    cboCod_Talla.AddItem "[Ninguna]                                    "
 
 End If
 
Exit Sub
hand:
    Set Rs = Nothing
    ErrorHandler Err, "LlenaComboTalla"
End Sub

Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
    If sTipo <> "D" Then

        If sTipo = "I" Then
        
            strSql = "SELECT count(*) FROM " & NombreTabla & " WHERE " & CodigoTabla & " = '" & txtCodigo.Text & "'"
            If DevuelveCampo(strSql, cConnect) = 0 Then
                MsgBox "El código ingresado no es valido. Sirvase verificar", vbInformation, "Ordenes de Compra"
                txtCodigo.SetFocus
                VALIDA_DATOS = False
                Exit Function
            End If
        End If
'
'        If Trim(txtcod_StaOrdComp.Text) = "" Then
'            MsgBox "El código de Status de Orden de Compra no puede estar vacío. Sirvase verificar", vbInformation, "Status de Orden de Compra"
'            txtcod_StaOrdComp.Text = ""
'            txtcod_StaOrdComp.SetFocus
'            VALIDA_DATOS = False
'            Exit Function
'        End If
'
'        If Trim(txtDes_StaOrdComp.Text) = "" Then
'            MsgBox "La descripción de Status de Orden de Compra no puede estar vacío. Sirvase verificar", vbInformation, "Status de Orden de Compra"
'            txtDes_StaOrdComp.Text = ""
'            txtDes_StaOrdComp.SetFocus
'            VALIDA_DATOS = False
'            Exit Function
'        End If


        If Trim(txtCodigo.Text) = "" Then
            MsgBox "El código no puede estar vacio. Sirvase verificar", vbInformation, "Ordenes de Compra"
            txtCodigo.Text = ""
            txtCodigo.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If

        If dtpFec_Entrega_Fin.Value < dtpFec_Entrega_Inicio.Value Then
            MsgBox "La fecha de entrega final no puede ser menor que la inicial. Sirvase verificar", vbInformation, "Ordenes de Compra"
            dtpFec_Entrega_Fin.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If


        If Val(txtCan_Comprada.Text) <= 0 Then
            MsgBox "La cantidad a comprar no puede ser cero. Sirvase verificar", vbInformation, "Ordenes de Compra"
            txtCan_Comprada.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
        
        If Val(txtPre_Unitario.Text) <= 0 Then
            MsgBox "El precio del producto no puede ser cero. Sirvase verificar", vbInformation, "Ordenes de Compra"
            txtPre_Unitario.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If

    Else
        'Aqui se valida que no tenga registros dependientes
        strSql = "SELECT COUNT(*) FROM LG_ORDCOMPITEMREQ WHERE Ser_OrdComp = '" & varSer_OrdComp & "' AND Cod_OrdComp = '" & varCod_OrdComp & "' AND Sec_OrdComp = '" & Rs_Lista("Sec_OrdComp").Value & "'"
        If DevuelveCampo(strSql, cConnect) > 0 Then
            MsgBox "El registro no puede ser eliminado por que posee registros relacionados. Sirvase verificar", vbInformation, "Ordenes de Compra"
            VALIDA_DATOS = False
            Exit Function
        End If
    End If
End Function

Sub CARGA_DATOS()

    If Not Rs_Lista.EOF And Not Rs_Lista.BOF Then
        
        varSec_OrdComp = Rs_Lista("Sec_OrdComp").Value
        txtCodigo.Text = Rs_Lista("Codigo").Value
        Call BUSCA_CODIGO(1)
        
        'Esto es por si no es la cantidad comprada sino la requerida
        'txtCan_Comprada.Text = Format(Rs_Lista("Can_Requerida").Value, "#####0.00")
        
        txtCan_Comprada.Text = Format(Rs_Lista("Can_Comprada").Value, "#####0.00")
        txtPre_Unitario.Text = Format(Rs_Lista("Pre_Unitario").Value, "#####0.00000")
        txtPorc_IGV.Text = Format(Rs_Lista("Porc_IGV").Value, "##0")
        TxtCodProv.Text = RTrim(Rs_Lista("Cod_Prov").Value)
        If IsNull(Rs_Lista("Observaciones").Value) Then
            txtObservaciones.Text = ""
        Else
            txtObservaciones.Text = Trim(Rs_Lista("Observaciones").Value)
        End If
        
        'txtCod_ItemProv.Text = Rs_Lista("Cod_ItemProv").Value
        
        txtCod_Color.Text = Trim(Rs_Lista("Cod_Color").Value)
        Call BUSCA_COLOR(1)
       
        Call BuscaCombo(Rs_Lista("Cod_Descuento").Value, 2, cboCod_Descuento)
        Call BuscaCombo(Rs_Lista("Cod_Comb").Value, 2, cboCod_Comb)
        Call BuscaCombo(Rs_Lista("Cod_ItemProv").Value, 1, cboCod_ItemProv)
        Call BuscaCombo(Rs_Lista("Cod_Talla").Value, 2, cboCod_Talla)
        Call BuscaCombo(Rs_Lista("Cod_Destino").Value, 2, cboCod_Destino)
        
        dtpFec_Entrega_Inicio.Value = Rs_Lista("Fec_Entrega_Inicio").Value
        dtpFec_Entrega_Fin.Value = Rs_Lista("Fec_Entrega_Fin").Value
        
    End If
End Sub

Sub HABILITA_DATOS()
    
    If sTipo = "I" Then
        txtCodigo.Enabled = True
        txtDescripcion.Enabled = True
        cboCod_Comb.Enabled = True
        txtCod_Color.Enabled = True
        txtDes_Color.Enabled = True
        cboCod_Talla.Enabled = True
        cboCod_Destino.Enabled = True
        cmdBuscaCodigo.Enabled = True
        cmdBuscaColor.Enabled = True
        
    End If
    TxtCodProv.Enabled = True
    txtPre_Unitario.Enabled = True
    txtCan_Comprada.Enabled = True
    cboCod_Descuento.Enabled = True
    txtPorc_IGV.Enabled = True
    cboCod_ItemProv.Enabled = True
    dtpFec_Entrega_Inicio.Enabled = True
    dtpFec_Entrega_Fin.Enabled = True
    txtObservaciones.Enabled = True
    
    
End Sub

Sub INHABILITA_DATOS()

    txtCodigo.Enabled = False
    txtDescripcion.Enabled = False
    txtCan_Comprada.Enabled = False
    txtPre_Unitario.Enabled = False
    cboCod_Descuento.Enabled = False
    txtPorc_IGV.Enabled = False
    TxtCodProv.Enabled = False
    cboCod_ItemProv.Enabled = False
    cboCod_Comb.Enabled = False
    txtCod_Color.Enabled = False
    txtDes_Color.Enabled = False
    cboCod_Talla.Enabled = False
    cboCod_Destino.Enabled = False
    dtpFec_Entrega_Inicio.Enabled = False
    dtpFec_Entrega_Fin.Enabled = False

    cmdBuscaCodigo.Enabled = False
    cmdBuscaColor.Enabled = False
    
    txtObservaciones.Enabled = False
End Sub

Sub CARGA_GRID()
    Dim vBookm As Variant
    
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.ActiveConnection = cConnect
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    If DGridLista.Row > 0 Then
        vBookm = DGridLista.Bookmark
    End If
    
    'Esta cadena es para devolver el Codigo de Cliente
    strSql = "EXEC UP_SEL_ORDCOMPITEM '" & varTip_Presentacion & "','" & varSer_OrdComp & "','" & varCod_OrdComp & "'"
    
    Rs_Lista.Open strSql
    Set DGridLista.DataSource = Rs_Lista
    DGridLista.Refresh
    
    If sTipo <> "D" And sTipo <> "" Then
        If Not IsEmpty(vBookm) Then
            DGridLista.Bookmark = vBookm
        End If
    End If
    
    If Rs_Lista.RecordCount > 0 Then
        DGridLista.Enabled = True
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Call CARGA_DATOS
    Else
        DGridLista.Enabled = False
        HabilitaMant Me.MantFunc1, "ADICIONAR"
        Call LIMPIAR_DATOS
    End If
End Sub

Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Dim strSql As String
    
    Con.ConnectionString = cConnect
    Con.Open
    
        Con.BeginTrans

        strSql = "EXEC UP_MAN_ORDCOMPITEMN '" & _
        sTipo & "','" & _
        varSer_OrdComp & "','" & _
        varCod_OrdComp & "','" & _
        varSec_OrdComp & "','" & _
        Trim(txtCodigo.Text) & "','" & _
        Right(cboCod_Comb.Text, 3) & "','" & _
        txtCod_Color.Text & "','" & _
        Trim(Right(cboCod_Talla.Text, 10)) & "','" & _
        Right(cboCod_Destino.Text, 3) & "','" & _
        "" & "','" & _
        Right(cboCod_Descuento.Text, 3) & "'," & _
        txtPorc_IGV.Text & ",'" & _
        dtpFec_Entrega_Inicio.Value & "','" & _
        dtpFec_Entrega_Fin.Value & "','" & _
        Trim(txtPre_Unitario.Text) & "','" & _
        Trim(txtCan_Comprada.Text) & "','" & _
        Mid(cboCod_ItemProv.Text, 1, 15) & "','" & _
        Trim(txtObservaciones.Text) & "','" & _
        varRepetir_Precio & "','" & _
        Trim(TxtCodProv.Text) & "','S','" & txtCodigo.Text & "'"
        
        Con.Execute strSql

        Con.CommitTrans
        Dim amensaje As New clsMessages
        amensaje.codigo = CodeMsg.kMESSAGE_INF_DATA_SAVE
        Informa "", amensaje
        
'        If sTipo = "I" Then
'            optOrdCompra.Value = True
'            txtSerOrdComp.Text = varSer_OrdComp
'            txtCodOrdComp.Text = txtCod_OrdComp.Text
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
       
        strSql = "EXEC UP_MAN_ORDCOMPITEMN '" & _
        sTipo & "','" & _
        varSer_OrdComp & "','" & _
        varCod_OrdComp & "','" & _
        varSec_OrdComp & "','" & _
        Trim(txtCodigo.Text) & "','" & _
        Right(cboCod_Comb.Text, 3) & "','" & _
        txtCod_Color.Text & "','" & _
        cboCod_Talla.Text & "','" & _
        Right(cboCod_Destino.Text, 3) & "','" & _
        "" & "','" & _
        Right(cboCod_Descuento.Text, 3) & "'," & _
        txtPorc_IGV.Text & ",'" & _
        dtpFec_Entrega_Inicio.Value & "','" & _
        dtpFec_Entrega_Fin.Value & "','" & _
        Trim(txtPre_Unitario.Text) & "','" & _
        Trim(txtCan_Comprada.Text) & "','" & _
        Mid(cboCod_ItemProv.Text, 1, 15) & "','" & _
        Trim(txtObservaciones.Text) & "','" & _
        varRepetir_Precio & "','" & _
        Trim(TxtCodProv.Text) & "'"
        
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

Private Sub cmdSalir_Click()
    fraPreRep.Visible = False
End Sub

Private Sub DGridLista_ColEdit(ByVal ColIndex As Integer)
    'If ColIndex = 13 Then
    'End If
End Sub

Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call CARGA_DATOS
End Sub

Private Sub Form_Load()
    Call FormateaGrid(DGridLista)
    Call CARGA_COMBOS
    
    Me.MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    
    'Aqui llenamos los Descuentos
    strSql = "SELECT CONVERT(VARCHAR,Porcentaje1) + ' - '+ CONVERT(VARCHAR,Porcentaje2) + SPACE(100) + COD_DESCUENTO FROM LG_DSCTOS"
    Call LlenaCombo(cboCod_Descuento, strSql, cConnect)
    
    Xcod_cliente_tex = ""
    Xcod_ordcomp_tex = ""
    Xser_ordcomp_tex = ""
    
    If vemp = "01" Or vemp = "04" Then
            flg_requerimientos_Tex = DevuelveCampo("select dbo.Valida_Si_Posee_Cliente_Textil('" & varSer_OrdComp & "','" & varCod_OrdComp & "')", cConnect)
            If flg_requerimientos_Tex = "S" Then
                Strsql_X = " select cod_cliente_tex, ser_ordcomp_tex, cod_ordcomp_tex from lg_ordComp where ser_OrdComp= '" & varSer_OrdComp & "' and Cod_OrdComp = '" & varCod_OrdComp & "'"
                Set Rsx_Textil = New ADODB.Recordset
                Set Rsx_Textil.DataSource = CargarRecordSetDesconectado(Strsql_X, cConnect)
                Xcod_cliente_tex = Rsx_Textil.Fields("cod_cliente_tex").Value
                Xser_ordcomp_tex = Rsx_Textil.Fields("ser_ordcomp_tex").Value
               Xcod_ordcomp_tex = Rsx_Textil.Fields("cod_ordcomp_tex").Value
            End If
   End If
   
    Call INHABILITA_DATOS
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim eliminar As Integer
    
    If varCod_StaOrdComp <> "P" And ActionName <> "SALIR" Then
        MsgBox "El estado del registro no permite modificación alguna. Sirvase verificar", vbInformation, "Ordenes de Compra"
        Exit Sub
    End If
    
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            LIMPIAR_DATOS
            HABILITA_DATOS
            'txtCod_Proveedor.SetFocus
            txtCodigo.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
            CmdHelp.Enabled = True
            cmdHelpPrecio.Enabled = True
        Case "MODIFICAR"
            sTipo = "U"
            HABILITA_DATOS
            'txtCod_Proveedor.SetFocus
            txtPre_Unitario.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
            CmdHelp.Enabled = True
            cmdHelpPrecio.Enabled = True
        Case "ELIMINAR"
            eliminar = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Combinación-Detalle")
            If eliminar = vbYes Then
                sTipo = "D"
                If VALIDA_DATOS Then
                    varRepetir_Precio = ""
                    Call ELIMINAR_DATOS
                    Call CARGA_GRID
                    sTipo = ""
                End If
            End If
        Case "GRABAR"
            If VALIDA_DATOS Then
            
                Cadena = DevuelveCampo("SELECT COD_FAMITEM FROM LG_ITEM WHERE COD_ITEM='" & txtCodigo & "'", cConnect)
                If Cadena = "HI" Then
                    If TxtCodProv.Text = "" Then
                        MsgBox "Debe ingresar el campo Cod. Prov."
                    Else
                        eliminar = MsgBox("Desea copiar el precio a los items relacionados?", vbYesNo + vbDefaultButton2, "Mensaje")
                        If eliminar = vbYes Then
                            varRepetir_Precio = "S"
                        Else
                            varRepetir_Precio = "N"
                        End If
                    
                        Call SALVAR_DATOS
                        Call CARGA_GRID
                        Call INHABILITA_DATOS
                        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                        DGridLista.Enabled = True
                        sTipo = ""
                        CmdHelp.Enabled = False
                        cmdHelpPrecio.Enabled = False
                        End If
                Else
                    eliminar = MsgBox("Desea copiar el precio a los items relacionados?", vbYesNo + vbDefaultButton2, "Mensaje")
                    If eliminar = vbYes Then
                        varRepetir_Precio = "S"
                    Else
                        varRepetir_Precio = "N"
                    End If
                
                    Call SALVAR_DATOS
                    Call CARGA_GRID
                    Call INHABILITA_DATOS
                    HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                    DGridLista.Enabled = True
                    sTipo = ""
                    CmdHelp.Enabled = False
                    cmdHelpPrecio.Enabled = False
                End If
            End If
        Case "DESHACER"
            Call LIMPIAR_DATOS
            Call CARGA_DATOS
            Call INHABILITA_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridLista.Enabled = True
            CmdHelp.Enabled = False
            cmdHelpPrecio.Enabled = False
            sTipo = ""
         Case "SALIR"
            Unload Me
    End Select
End Sub

Private Sub txtCan_Comprada_KeyPress(KeyAscii As Integer)
    SoloNumeros txtCan_Comprada, KeyAscii, True, 2, 9
End Sub

Private Sub txtCan_Comprada_LostFocus()
    If Trim(txtCan_Comprada.Text) = "" Then
        txtCan_Comprada.Text = "0"
    End If
End Sub

Private Sub txtCod_Color_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_Color.Text) <> "" Then
            Call BUSCA_COLOR(1)
        End If
    End If
End Sub
Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
    If Trim(Xcod_ordcomp_tex) = "" And Trim(Xser_ordcomp_tex) = "" Then
    
        If Trim(txtCodigo.Text) <> "" Then
            If varTip_Presentacion = "I" Or varTip_Presentacion = "T" Or varTip_Presentacion = "H" Then
                If Len(txtCodigo.Text) = 2 Then
                    Call BUSCA_CODIGO(4)
                Else
                    If varTip_Presentacion <> "H" Then
                        txtCodigo.Text = CompletaCodigo(Trim(txtCodigo.Text), 8, 2)
                        Call BUSCA_CODIGO(1)
                    Else
                        Call BUSCA_CODIGO(4)
                    End If
                End If
            End If
        End If
        
    Else
         Call BUSCA_CODIGO(5)
    End If
    End If
    
End Sub

Private Sub txtDes_Color_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        If Trim(txtDes_Color.Text) <> "" Then
            Call BUSCA_COLOR(2)
        End If
    End If
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDescripcion.Text) <> "" Then
            Call BUSCA_CODIGO(2)
        End If
    End If
End Sub

Private Sub txtPorc_IGV_KeyPress(KeyAscii As Integer)
    SoloNumeros txtPorc_IGV, KeyAscii, True, 2, 4
End Sub

Private Sub txtPre_Unitario_KeyPress(KeyAscii As Integer)
    SoloNumeros txtPre_Unitario, KeyAscii, True, 5, 8
    
End Sub

Private Sub txtPre_Unitario_LostFocus()
    If Trim(txtPre_Unitario.Text) = "" Then
        txtPre_Unitario.Text = "0.00000"
    Else
        txtPre_Unitario.Text = Format(txtPre_Unitario.Text, "#####000.00000")
    End If
End Sub

Public Function CompletaCodigo(CodOrigen As String, longcodfinal As Integer, PosfinalCod As Integer) As String
' CodOrigen     = Es el codigo que sera pasado por parametro
' LongCodFinal  = Es el tamaño del Codigo a devolver
' PosFinalCod   = Es la posicion de la 1era parte del codigo
    Dim Contador As Integer
    CompletaCodigo = Mid(CodOrigen, 1, PosfinalCod)
    For Contador = 1 To longcodfinal - Len(CodOrigen)
        CompletaCodigo = CompletaCodigo & "0"
    Next
    Contador = Len(CodOrigen) - PosfinalCod
    If Contador < 0 Then
        Contador = 0
    End If
    CompletaCodigo = CompletaCodigo & Right(CodOrigen, Contador)
End Function
