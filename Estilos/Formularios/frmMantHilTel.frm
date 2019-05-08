VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMantHilTel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hilados"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   Icon            =   "frmMantHilTel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Hilados"
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   615
      TabIndex        =   37
      Top             =   7725
      Width           =   2085
      Begin VB.CommandButton cmdLast 
         Height          =   510
         Left            =   1545
         Picture         =   "frmMantHilTel.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Ultimo"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   510
         Left            =   15
         Picture         =   "frmMantHilTel.frx":0A3C
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Primero"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   510
         Left            =   1050
         Picture         =   "frmMantHilTel.frx":0BAE
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Siguiente"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   510
         Left            =   510
         Picture         =   "frmMantHilTel.frx":0D20
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Anterior"
         Top             =   0
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
      Height          =   2070
      Left            =   90
      TabIndex        =   35
      Tag             =   "List"
      Top             =   1305
      Width           =   6840
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   1590
         Left            =   5535
         TabIndex        =   40
         Top             =   270
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   2805
         Custom          =   $"frmMantHilTel.frx":0E92
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1100
         ControlHeigth   =   450
         ControlSeparator=   110
      End
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   1725
         Left            =   240
         TabIndex        =   36
         Top             =   255
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3043
         _Version        =   393216
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "Cod_organizacion"
            Caption         =   "Código"
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
            DataField       =   "Nom_organizacion"
            Caption         =   "Descripción"
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
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   "Precio ($)"
            Caption         =   "Precio $"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1980.284
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1814.74
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
      Height          =   4185
      Left            =   90
      TabIndex        =   29
      Tag             =   "Detail"
      Top             =   3405
      Width           =   6840
      Begin VB.TextBox TxtDes_Hilado 
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
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   45
         Top             =   3570
         Width           =   3960
      End
      Begin VB.TextBox TxtCodHilado 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   44
         Top             =   3570
         Width           =   1080
      End
      Begin VB.ComboBox cboTip_Solicitud_Hilado 
         Height          =   315
         Left            =   1695
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   3210
         Width           =   3105
      End
      Begin VB.TextBox txtPrecio 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1710
         TabIndex        =   3
         Top             =   990
         Width           =   795
      End
      Begin VB.TextBox txtTit_Hilado 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1695
         TabIndex        =   8
         Top             =   2085
         Width           =   780
      End
      Begin VB.CommandButton cmdBusHilado 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   330
         Left            =   2520
         TabIndex        =   9
         Tag             =   "..."
         Top             =   2085
         Width           =   315
      End
      Begin VB.TextBox txtDesUniMed 
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
         Left            =   2865
         MaxLength       =   30
         TabIndex        =   18
         Top             =   2835
         Width           =   3855
      End
      Begin VB.CommandButton cmdBusUM 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   330
         Left            =   2520
         TabIndex        =   17
         Tag             =   "..."
         Top             =   2820
         Width           =   315
      End
      Begin VB.TextBox txtIdUniMed 
         BackColor       =   &H80000004&
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
         Left            =   1695
         MaxLength       =   2
         TabIndex        =   16
         Top             =   2820
         Width           =   780
      End
      Begin VB.TextBox txtDes_TitHilado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2910
         TabIndex        =   10
         Top             =   2085
         Width           =   3825
      End
      Begin VB.TextBox txtNumCabo 
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
         Left            =   1695
         TabIndex        =   12
         Top             =   2445
         Width           =   885
      End
      Begin VB.TextBox txtIdFamHil 
         BackColor       =   &H80000004&
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
         Left            =   1695
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1380
         Width           =   780
      End
      Begin VB.CommandButton cmdBusFam 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   330
         Left            =   2520
         TabIndex        =   5
         Tag             =   "..."
         Top             =   1365
         Width           =   315
      End
      Begin VB.TextBox txtIdHilado 
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
         Left            =   1695
         MaxLength       =   8
         TabIndex        =   1
         Top             =   285
         Width           =   1080
      End
      Begin VB.TextBox txtDesFamHil 
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
         Left            =   2865
         MaxLength       =   30
         TabIndex        =   30
         Top             =   1380
         Width           =   3855
      End
      Begin VB.TextBox txtDesHilTel 
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
         Left            =   1710
         TabIndex        =   2
         Top             =   630
         Width           =   5025
      End
      Begin VB.TextBox txtNumOcu 
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
         Left            =   1695
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1740
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtIdCtaCon 
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
         Left            =   1665
         TabIndex        =   7
         Top             =   1740
         Width           =   2295
      End
      Begin VB.TextBox txtMtsKgr 
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
         Left            =   4425
         TabIndex        =   14
         Top             =   2445
         Width           =   2295
      End
      Begin VB.Label lblCodEstructurado 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Estructurado :"
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
         Index           =   2
         Left            =   315
         TabIndex        =   47
         Tag             =   "Family :"
         Top             =   3630
         Width           =   1395
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Solicitud"
         Height          =   195
         Left            =   270
         TabIndex        =   42
         Top             =   3270
         Width           =   960
      End
      Begin VB.Label Label8 
         Caption         =   "Precio Unit.$:"
         Height          =   225
         Left            =   300
         TabIndex        =   41
         Top             =   1050
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "Unid.Med. :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   15
         Tag             =   "UM :"
         Top             =   2820
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Tit. Hilado :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   39
         Tag             =   "Tit. Hilado :"
         Top             =   2145
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Mts / Kgr :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3420
         TabIndex        =   13
         Tag             =   "Mts / Kgr :"
         Top             =   2490
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "# Cabo :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   11
         Tag             =   "# Cabo :"
         Top             =   2475
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Cta.Con. :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   38
         Tag             =   "Cta.Con. :"
         Top             =   1770
         Width           =   735
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   300
         TabIndex        =   34
         Tag             =   "Code"
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Familia :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   300
         TabIndex        =   33
         Tag             =   "Family :"
         Top             =   1425
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Descripción :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   32
         Tag             =   "Description :"
         Top             =   675
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "# Ocur. :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   31
         Tag             =   "# Ocur. :"
         Top             =   1770
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar por "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   90
      TabIndex        =   0
      Tag             =   "Find"
      Top             =   45
      Width           =   6840
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   975
         Left            =   4905
         TabIndex        =   28
         Top             =   195
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   1720
         Custom          =   $"frmMantHilTel.frx":0FA0
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1100
         ControlHeigth   =   450
         ControlSeparator=   50
      End
      Begin VB.OptionButton optHilado 
         Caption         =   "Hilado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2265
         TabIndex        =   24
         Tag             =   "Hilado"
         Top             =   315
         Width           =   1035
      End
      Begin VB.OptionButton optFamilia 
         Caption         =   "Familia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   975
         TabIndex        =   23
         Tag             =   "Family"
         Top             =   315
         Value           =   -1  'True
         Width           =   1035
      End
      Begin MSDataListLib.DataCombo dcbFamHil 
         Height          =   330
         Left            =   990
         TabIndex        =   26
         Top             =   675
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtDesHilado 
         Height          =   330
         Left            =   990
         TabIndex        =   27
         Top             =   675
         Width           =   3360
      End
      Begin VB.Label lblOpcion 
         Caption         =   "Familia :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   255
         TabIndex        =   25
         Tag             =   "Family :"
         Top             =   720
         Width           =   945
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2910
      TabIndex        =   46
      Top             =   7665
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantHilTel.frx":1031
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmMantHilTel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public Codigo, Descripcion As String
Dim sTipo As String
Dim Rs_Carga As New ADODB.Recordset
Dim varBusqueda As String
Dim StrSQL As String
Private Sub cmdBusFam_Click()
Dim oTipo As New frmBusqGeneral
Dim Rs As New ADODB.Recordset
Set oTipo.oParent = Me
oTipo.sQuery = "SELECT cod_famhiltel as Codigo, des_famhiltel as Descripcion FROM IT_FamHil"
oTipo.Cargar_Datos
oTipo.Show 1
If Codigo <> "" Then
    txtIdFamHil.Text = Codigo
    txtDesFamHil.Text = Descripcion
    Codigo = ""
End If
Set oTipo = Nothing
End Sub
Private Sub cmdBusFam_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub

Private Sub cmdBusHilado_Click()
Dim oUniMed As New frmBusqGeneral
Dim Rs As New ADODB.Recordset
Set oUniMed.oParent = Me
oUniMed.sQuery = "SELECT Tit_Hilado as Código, Des_TitHilado as Descripcion FROM It_TitHil"
oUniMed.Cargar_Datos
oUniMed.Show 1
If Codigo <> "" Then
    txtTit_Hilado.Text = Codigo
    txtDes_TitHilado.Text = Descripcion
    Codigo = ""
End If
Set oUniMed = Nothing

End Sub

Private Sub cmdBusUM_Click()
Dim oUniMed As New frmBusqGeneral
Dim Rs As New ADODB.Recordset
Set oUniMed.oParent = Me
oUniMed.sQuery = "SELECT cod_unimed as Codigo, des_unimed as Descripcion FROM TG_UniMed"
oUniMed.Cargar_Datos
oUniMed.Show 1
If Codigo <> "" Then
    txtIdUniMed.Text = Codigo
    txtDesUniMed.Text = Descripcion
    Codigo = ""
End If
Set oUniMed = Nothing
End Sub
Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "BUSCAR"
        If optFamilia.Value = True Then
            If Len(Trim(dcbFamHil.BoundText)) = 0 Then
                MsgBox "Debe seleccionar Familia", vbInformation, "Mantenimiento Hilados"
                Exit Sub
            End If
        Else
            If Len(Trim(txtDesHilado.Text)) = 0 Then
                MsgBox "Debe ingresar una Descripción", vbInformation, "Mantenimiento Hilados"
                Exit Sub
            End If
        End If
        Carga_Datos
    Case "FAMILIAS"
        Dim oFamilia As New frmMantFamHil
        Set oFamilia.oParent = Me
        oFamilia.Show 1
        Set oFamilia = Nothing
        Llena_Combo
End Select
End Sub
Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

'Almacenamos el codigo del hilado para efectuar busquedas
varBusqueda = Me.txtIdHilado

Select Case ActionName
    Case "MATPRIMA"
        Dim oComp As New frmMantComHil
        Set oComp.oParent = Me
        oComp.txtIdHilado = txtIdHilado.Text
        oComp.txtDesHilado = txtDesHilTel.Text
        oComp.Cargar_Datos
        oComp.Show 1
        Set oComp = Nothing
    Case "SOLICITUDES"
        Select Case Rs_Carga("Tip_Solicitud_hilado").Value
            Case "0"
                        MsgBox "El tipo de solicitud elegida no permite acceder a esta opción. Sirvase verificar", vbInformation, "Mensaje"
                        Exit Sub
            Case "1"
                        Load frmSolicitudHiladoTipo1
                        frmSolicitudHiladoTipo1.varcod_hiltel = Rs_Carga("Cod_HilTel").Value
                        frmSolicitudHiladoTipo1.Cargar_Datos
                        frmSolicitudHiladoTipo1.lblTipo.Caption = "Tipo " & CStr(Rs_Carga("Tip_Solicitud_hilado").Value) & " : " & UCase(Trim(Mid(Me.cboTip_Solicitud_Hilado.Text, 1, Len(Me.cboTip_Solicitud_Hilado.Text) - 2)))
                        frmSolicitudHiladoTipo1.Show 1
                        
                        Set frmSolicitudHiladoTipo1 = Nothing
            Case "2"
                        Load frmSolicitudHiladoTipo2
                        frmSolicitudHiladoTipo2.varcod_hiltel = Rs_Carga("Cod_HilTel").Value
                        frmSolicitudHiladoTipo2.Cargar_Datos
                        frmSolicitudHiladoTipo2.lblTipo.Caption = "Tipo " & CStr(Rs_Carga("Tip_Solicitud_hilado").Value) & " : " & UCase(Trim(Mid(Me.cboTip_Solicitud_Hilado.Text, 1, Len(Me.cboTip_Solicitud_Hilado.Text) - 2)))
                        frmSolicitudHiladoTipo2.Show 1
                        
                        Set frmSolicitudHiladoTipo2 = Nothing
            Case "3"
                        Load frmSolicitudHiladoTipo2
                        frmSolicitudHiladoTipo2.varcod_hiltel = Rs_Carga("Cod_HilTel").Value
                        frmSolicitudHiladoTipo2.Cargar_Datos
                        frmSolicitudHiladoTipo2.lblTipo.Caption = "Tipo " & CStr(Rs_Carga("Tip_Solicitud_hilado").Value) & " : " & UCase(Trim(Mid(Me.cboTip_Solicitud_Hilado.Text, 1, Len(Me.cboTip_Solicitud_Hilado.Text) - 2)))
                        frmSolicitudHiladoTipo2.Show 1
                        
                        Set frmSolicitudHiladoTipo2 = Nothing
            Case "4"
                        Load frmSolicitudHiladoTipo4
                        frmSolicitudHiladoTipo4.varcod_hiltel = Rs_Carga("Cod_HilTel").Value
                        frmSolicitudHiladoTipo4.Cargar_Datos
                        frmSolicitudHiladoTipo4.lblTipo.Caption = "Tipo " & CStr(Rs_Carga("Tip_Solicitud_hilado").Value) & " : " & UCase(Trim(Mid(Me.cboTip_Solicitud_Hilado.Text, 1, Len(Me.cboTip_Solicitud_Hilado.Text) - 2)))
                        frmSolicitudHiladoTipo4.Show 1
                        
                        Set frmSolicitudHiladoTipo4 = Nothing
            Case "5"
        End Select
    Case "IMPRIME":
                    Call REPORTE
End Select
End Sub
Private Sub optFamilia_Click()
    dcbFamHil.Visible = True
    txtDesHilado.Visible = False
    lblOpcion.Caption = optFamilia.Caption
End Sub
Private Sub optHilado_Click()
    dcbFamHil.Visible = False
    txtDesHilado.Visible = True
    lblOpcion.Caption = optHilado.Caption
End Sub
Sub Llena_Combo()
Dim Rs_tipPO As New ADODB.Recordset
On Error GoTo Llena_ComboError
Rs_tipPO.ActiveConnection = cCONNECT
Rs_tipPO.CursorLocation = adUseClient
Rs_tipPO.CursorType = adOpenStatic
Rs_tipPO.Source = "SELECT * FROM IT_FamHil"
Rs_tipPO.Open
Set dcbFamHil.RowSource = Rs_tipPO
dcbFamHil.BoundColumn = "cod_famhiltel"
dcbFamHil.ListField = "des_famhiltel"
Exit Sub
Llena_ComboError:
    ErrorHandler Err, "Procedimiento LlenaCombo"
    Set Rs_tipPO = Nothing
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
Sub Carga_Datos()
    Dim StrSQL As String
    On Error GoTo Cargar_DatosErr
    If optFamilia.Value = True Then
        StrSQL = "SG_Act_Hilado '','','" & dcbFamHil.BoundText & "',0,'','','','',0,0,'L',''"
    Else
        StrSQL = "SG_Act_Hilado '','" & txtDesHilado.Text & "','',0,'','','','',0,0,'L',''"
    End If
    Set Rs_Carga = Nothing
    Rs_Carga.ActiveConnection = cCONNECT
    Rs_Carga.CursorType = adOpenStatic
    Rs_Carga.CursorLocation = adUseClient
    Rs_Carga.LockType = adLockReadOnly
    Rs_Carga.Open StrSQL
    Set DGridLista.DataSource = Rs_Carga
    DGridLista_RowColChange 0, 0
    If Rs_Carga.RecordCount > 0 Then
        'HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        HabilitaMant Me.FunctButt2, "MATPRIMA/SOLICITUDES/IMPRIME"
        
        Call BuscaCampo(Rs_Carga, "cod_hiltel", varBusqueda)
        
    Else
        LIMPIAR_DATOS
        DESHABILITA_DATOS
        'HabilitaMant Me.MantFunc1, "ADICIONAR"
        HabilitaMant Me.FunctButt2, ""
    End If
    Exit Sub
Cargar_DatosErr:
    Set Rs_Carga = Nothing
    ErrorHandler Err, "Cargar_Datos"
End Sub
Private Sub Form_Load()
'cCONNECT = "Provider=sqloledb;Server=SERVIDOR;Database=lives;UID=sa;pwd=;"
'cSEGURIDAD = "Provider=sqloledb;Server=servidor;Database=seguridad;UID=sa;pwd=;"
Call FormSet(Me)
FormateaGrid Me.DGridLista
DGridLista.Columns(0).DataField = "cod_hiltel"
DGridLista.Columns(1).DataField = "des_hiltel"
DGridLista.Columns(2).DataField = "des_famhiltel"
DGridLista.Columns(3).DataField = "Precio ($)"
DGridLista.Columns(2).Caption = "Familia"
HabilitaMant Me.MantFunc1, ""
HabilitaMant Me.FunctButt2, ""
Llena_Combo
'MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
'FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
FunctButt2.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)

StrSQL = "SELECT Des_Solicitud_Hilado + SPACE(100) + Tip_Solicitud_Hilado  FROM IT_TIPOS_SOLICITUD"
Call LlenaCombo(Me.cboTip_Solicitud_Hilado, StrSQL, cCONNECT)

End Sub
Sub SALVAR_DATOS()
Dim Con As New ADODB.Connection
Dim sUniMed, sFamHil As String
On Error GoTo Salvar_DatosErr
Con.ConnectionString = cCONNECT
Con.Open
If txtIdHilado.Text <> "" Then
    If Len(Trim(txtIdUniMed.Text)) = 0 Then
        sUniMed = "Null"
    Else
        sUniMed = "'" & txtIdUniMed.Text & "'"
    End If
    If Len(Trim(txtIdFamHil.Text)) = 0 Then
        sFamHil = "Null"
    Else
        sFamHil = "'" & txtIdFamHil.Text & "'"
    End If
    Con.Execute "SG_Act_Hilado '" & _
    txtIdHilado.Text & "','" & txtDesHilTel.Text & "'," & _
    sFamHil & "," & Val(txtNumOcu.Text) & "," & _
    sUniMed & ",'" & txtIdCtaCon.Text & "','" & _
    txtTit_Hilado.Text & "','" & txtDes_TitHilado.Text & "'," & _
    Val(txtNumCabo.Text) & "," & _
    Val(txtMtsKgr.Text) & ",'" & sTipo & "','" & TxtCodHilado.Text & "'," & txtPrecio & ",'" & Right(Me.cboTip_Solicitud_Hilado.Text, 1) & "','" & _
    vusu & "'"
    
    
    'Almacenamos el codigo del hilado para efectuar busquedas
    varBusqueda = Me.txtIdHilado
    
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
    Informa "", amensaje
End If
LIMPIAR_DATOS
RECARGAR_DATOS
Exit Sub
Salvar_DatosErr:
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub
Sub ELIMINAR_DATOS()
Dim Con As New ADODB.Connection
On Error GoTo Eliminar_DatosErr
Con.ConnectionString = cCONNECT
Con.Open
If txtIdFamHil.Text <> "" Then
    Con.BeginTrans
    Con.Execute "SG_Act_Hilado '" & txtIdHilado.Text & "','','',0,'','','','',0,0,'D',''"
    Con.CommitTrans
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_DELETE
    Informa "", amensaje
End If
LIMPIAR_DATOS
RECARGAR_DATOS
Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Eliminar_Datos"
End Sub
Sub LIMPIAR_DATOS()
    txtIdHilado.Text = ""
    txtDesHilTel.Text = ""
    txtIdFamHil.Text = ""
    txtDesFamHil.Text = ""
    txtTit_Hilado.Text = ""
    txtDes_TitHilado.Text = ""
    txtNumOcu.Text = ""
    txtNumCabo.Text = ""
    txtMtsKgr.Text = ""
    txtIdUniMed.Text = ""
    txtIdCtaCon.Text = ""
    TxtCodHilado.Text = ""
    txtPrecio.Text = 0
    
    If Me.cboTip_Solicitud_Hilado.ListCount > 0 Then
        Me.cboTip_Solicitud_Hilado.ListIndex = 0
    End If
    
End Sub
Private Sub DGridLista_Click()
If Rs_Carga.State <> 1 Then
    Exit Sub
End If
If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
    txtIdHilado.Text = Rs_Carga!Cod_HilTel
    txtDesHilTel.Text = Rs_Carga!des_hiltel
    txtIdFamHil.Text = Rs_Carga!cod_famhiltel
    txtTit_Hilado.Text = Rs_Carga!Tit_Hilado
    txtDes_TitHilado.Text = Rs_Carga!Des_TitHil
    txtDesFamHil.Text = Rs_Carga!des_famhiltel
    txtNumOcu.Text = Rs_Carga!num_ocurrencia
    txtNumCabo.Text = Rs_Carga!num_cabo
    txtMtsKgr.Text = Rs_Carga!mtskgr
    txtIdUniMed.Text = Rs_Carga!cod_uniMed
    txtIdCtaCon.Text = Rs_Carga!cod_ctacon
    txtPrecio.Text = Rs_Carga![Precio ($)]
    TxtCodHilado.Text = Rs_Carga!Cod_hilado_Estructurado
    Call BuscaCombo(Rs_Carga("Tip_Solicitud_hilado").Value, 2, Me.cboTip_Solicitud_Hilado)
    DESHABILITA_DATOS
End If
End Sub
Sub HABILITA_DATOS()
    txtIdHilado.Enabled = True
    If sTipo = "I" Then
        txtDesHilTel.Enabled = True
    End If
    txtIdFamHil.Enabled = True
    txtTit_Hilado.Enabled = True
    
    txtDes_TitHilado.Enabled = True
    cmdBusHilado.Enabled = True
    txtNumOcu.Enabled = True
    txtNumCabo.Enabled = True
    txtMtsKgr.Enabled = True
    txtIdUniMed.Enabled = True
    txtIdCtaCon.Enabled = True
    txtPrecio.Enabled = True
    Me.cboTip_Solicitud_Hilado.Enabled = True
End Sub
Sub DESHABILITA_DATOS()
    txtIdHilado.Enabled = False
    txtDesHilTel.Enabled = False
    txtIdFamHil.Enabled = False
    TxtCodHilado.Enabled = False
    txtTit_Hilado.Enabled = False
    txtDes_TitHilado.Enabled = False
    cmdBusHilado.Enabled = False
    txtNumOcu.Enabled = False
    txtNumCabo.Enabled = False
    txtMtsKgr.Enabled = False
    txtIdUniMed.Enabled = False
    txtIdCtaCon.Enabled = False
    txtPrecio.Enabled = False
    TxtCodHilado.Enabled = False
    Me.cboTip_Solicitud_Hilado.Enabled = False
End Sub
Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
'AVANZA (KeyCode)
End Sub
Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Rs_Carga.State <> 1 Then
    Exit Sub
End If
If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
    txtIdHilado.Text = Rs_Carga!Cod_HilTel
    txtDesHilTel.Text = Rs_Carga!des_hiltel
    txtIdFamHil.Text = Rs_Carga!cod_famhiltel
    txtTit_Hilado.Text = Rs_Carga!Tit_Hilado
    txtDes_TitHilado.Text = Rs_Carga!Des_TitHil
    txtDesFamHil.Text = Rs_Carga!des_famhiltel
    txtNumOcu.Text = Rs_Carga!num_ocurrencia
    txtNumCabo.Text = Rs_Carga!num_cabo
    txtMtsKgr.Text = Rs_Carga!mtskgr
    txtIdUniMed.Text = Rs_Carga!cod_uniMed
    txtIdCtaCon.Text = Rs_Carga!cod_ctacon
    txtPrecio.Text = Rs_Carga![Precio ($)]
    TxtCodHilado.Text = Rs_Carga!Cod_hilado_Estructurado
    Call BuscaCombo(Rs_Carga("Tip_Solicitud_hilado").Value, 2, Me.cboTip_Solicitud_Hilado)
    DESHABILITA_DATOS
End If
End Sub
Sub RECARGAR_DATOS()
Rs_Carga.Close
Carga_Datos
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set Rs_Carga = Nothing
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub
Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "ADICIONAR"
        
            TxtCodHilado.Enabled = True
            sTipo = "I"
            LIMPIAR_DATOS
            HABILITA_DATOS
            cmdBusFam.Enabled = True
            cmdBusUM.Enabled = True
            If optFamilia.Value = True Then
                txtIdFamHil.Text = dcbFamHil.BoundText
                Busca_FamHil
            End If
            txtIdUniMed.Text = "KG"
            Busca_UniMed
            txtIdHilado.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        
    Case "MODIFICAR"
        sTipo = "U"
        HABILITA_DATOS
        txtIdHilado.Enabled = False
        cmdBusFam.Enabled = True
        cmdBusUM.Enabled = True
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        DGridLista.Enabled = False
    Case "ELIMINAR"
        ELIMINAR_DATOS
    Case "GRABAR"
        If TxtCodHilado.Text = "" Then
            MsgBox "Debe ingresar codigo estructurado", vbInformation
        'ElseIf DevuelveCampo("SELECT COUNT(*) FROM IT_HILADO_TEJEDURIA WHERE COD_HILTEL_TEJEDURIA='" & TxtCodHilado & "'", cCONNECT) = 0 Then
        '        MsgBox "Codigo estructurado NO EXISTE. DEBE DE CREARSE PRIMERO EN TEJEDURIA", vbInformation
        'Else
        ElseIf DevuelveCampo("SELECT COUNT(*) FROM IT_HILADO WHERE COD_HILADO_ESTRUCTURADO='" & TxtCodHilado & "'", cCONNECT) = 0 Then
                MsgBox "Codigo estructurado NO EXISTE.", vbInformation
        Else
            TxtDes_Hilado.Text = DevuelveCampo("select des_hiltel_Tejeduria from IT_Hilado_tejeduria where cod_hiltel_tejeduria ='" & TxtCodHilado & "'", cCONNECT)
            If VALIDA_DATOS Then
                SALVAR_DATOS
                RECARGAR_DATOS
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                DGridLista.Enabled = True
                cmdBusFam.Enabled = False
                cmdBusUM.Enabled = False
            End If
        End If
    Case "DESHACER"
        LIMPIAR_DATOS
        RECARGAR_DATOS
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        DGridLista.Enabled = True
        cmdBusFam.Enabled = False
        cmdBusUM.Enabled = False
    Case "SALIR"
        Unload Me
End Select
End Sub
Function VALIDA_DATOS() As Boolean
Dim aMess(4)
Dim amensaje As clsMessages
Set amensaje = New clsMessages
VALIDA_DATOS = True
If Len(Trim(txtDesHilTel.Text)) = 0 Then
   MsgBox "Ingrese la descripcion", vbInformation, Me.Caption
   VALIDA_DATOS = False
End If
If Len(Trim(txtIdHilado.Text)) = 0 Then
   MsgBox "Ingrese el Codigo", vbInformation, Me.Caption
   VALIDA_DATOS = False
End If

If Len(Trim(TxtCodHilado.Text)) = 0 Then
   MsgBox "Ingrese Codigo Hilado", vbInformation, Me.Caption
   VALIDA_DATOS = False
End If

If Not VALIDA_DATOS Then
    LoadMessage aMess, amensaje.Codigo
    amensaje.ShowMesage (iLanguage)
    Exit Function
End If

If Trim(txtTit_Hilado.Text) = "" Then
    Call MsgBox("El Tit. Hilado no puede estar vacio. Sirvase verificar", vbInformation, "Hilados")
    VALIDA_DATOS = False
End If

StrSQL = "SELECT count(*) from IT_TITHIL WHERE Tit_Hilado='" & Trim(txtTit_Hilado.Text) & "'"
If DevuelveCampo(StrSQL, cCONNECT) < 1 Then
    Call MsgBox("El Tit. Hilado ingresado no es válido. Sirvase verificar", vbInformation, "Hilados")
    VALIDA_DATOS = False
End If

If Me.cboTip_Solicitud_Hilado.Text = "" Then
    MsgBox "El Tipo de solicitud no puede estar vacia. Sirvase verificar", vbInformation, "Mensaje"
    VALIDA_DATOS = False
    Me.cboTip_Solicitud_Hilado.SetFocus
    Exit Function
End If

End Function

Private Sub TxtCodHilado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    MantFunc1.SetFocus
End If
End Sub

Private Sub txtDesHilTel_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub
Private Sub txtIdHilado_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub
Private Sub txtIdHilado_LostFocus()
Busca_Hilado
End Sub
Private Sub txtIdFamHil_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub
Private Sub txtIdFamHil_LostFocus()
If Len(Trim(txtIdFamHil)) <> 0 Then
    Busca_FamHil
End If
End Sub
Sub Busca_FamHil()
Dim Rs_busca As New ADODB.Recordset
On Error GoTo Busca_FuncionErr
B_sql = "SELECT * FROM IT_FamHil " & _
"WHERE cod_famhiltel = '" & txtIdFamHil.Text & "'"
Rs_busca.ActiveConnection = cCONNECT
Rs_busca.CursorType = adOpenStatic
Rs_busca.Open B_sql
If Not Rs_busca.EOF Then
    txtDesFamHil.Text = Rs_busca!des_famhiltel
Else
    txtDesFamHil.Text = ""
    txtIdFamHil.Text = ""
End If
Rs_busca.Close
Set Rs_busca = Nothing
Exit Sub
Busca_FuncionErr:
    Set Rs_busca = Nothing
    ErrorHandler Err, "Busca_Acceso"
End Sub
Sub Busca_Hilado()
Dim Rs_busca As New ADODB.Recordset
On Error GoTo Busca_FuncionErr
B_sql = "SELECT * FROM IT_Hilado " & _
"WHERE cod_hiltel = '" & txtIdHilado.Text & "'"
Rs_busca.ActiveConnection = cCONNECT
Rs_busca.CursorType = adOpenStatic
Rs_busca.Open B_sql
If Not Rs_busca.EOF Then
    txtDesHilTel.Text = Rs_busca!des_hiltel
    txtIdFamHil.Text = Rs_busca!cod_famhiltel
    txtTit_Hilado.Text = Rs_busca!Tit_Hilado
    txtDes_TitHilado.Text = Rs_busca!Des_TitHil
    txtPrecio.Text = Rs_busca!rep_predolar
    
    Busca_FamHil
    DESHABILITA_DATOS
    HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
    DGridLista.Enabled = True
End If
Rs_busca.Close
Set Rs_busca = Nothing
Exit Sub
Busca_FuncionErr:
    Set Rs_busca = Nothing
    ErrorHandler Err, "Busca_Acceso"
End Sub
Private Sub txtIdUniMed_LostFocus()
If Len(Trim(txtIdUniMed.Text)) <> 0 Then
    Busca_UniMed
End If
End Sub
Private Sub txtMtsKgr_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub
Private Sub txtMtsKgr_KeyPress(KeyAscii As Integer)
SoloNumeros txtMtsKgr, KeyAscii, True, 2, 3
End Sub
Private Sub txtNumCabo_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub
Private Sub txtNumCabo_KeyPress(KeyAscii As Integer)
SoloNumeros txtNumCabo, KeyAscii, False, 0, 3
End Sub
Private Sub txtNumOcu_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub
Private Sub txtNumOcu_KeyPress(KeyAscii As Integer)
SoloNumeros txtNumOcu, KeyAscii, False, 0, 1
End Sub
Sub Busca_UniMed()
Dim Rs_busca As New ADODB.Recordset
On Error GoTo Busca_UniMedErr
B_sql = "SELECT * FROM TG_UniMed " & _
"WHERE cod_unimed = '" & txtIdUniMed.Text & "'"
Rs_busca.ActiveConnection = cCONNECT
Rs_busca.CursorType = adOpenStatic
Rs_busca.Open B_sql
If Not Rs_busca.EOF Then
    txtDesUniMed.Text = Rs_busca!des_unimed
Else
    txtDesUniMed.Text = ""
    txtIdUniMed.Text = ""
End If
Rs_busca.Close
Set Rs_busca = Nothing
Exit Sub
Busca_UniMedErr:
    Set Rs_busca = Nothing
    ErrorHandler Err, "Busca_UniMed"
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    Call SoloNumeros(txtPrecio, KeyAscii, True, 2)
End Sub

Public Sub REPORTE()
On Error GoTo ErrorImpresion
    Dim oo As Object
    Set oo = CreateObject("excel.application")
    'oo.Workbooks.Open App.Path & "\RptSolicitudesHil.xlt"
    oo.workbooks.Open vRuta & "\RptSolicitudesHil.xlt"
    oo.Visible = True
    oo.run "REPORTE", Rs_Carga("Cod_HilTel").Value, CStr(Rs_Carga("Tip_Solicitud_hilado").Value), "Tipo " & CStr(Rs_Carga("Tip_Solicitud_hilado").Value) & " : " & UCase(Trim(Mid(Me.cboTip_Solicitud_Hilado.Text, 1, Len(Me.cboTip_Solicitud_Hilado.Text) - 2))), cCONNECT
    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte de Tarifado de Operaciones " & Err.Description, vbCritical, "Impresion"
End Sub
