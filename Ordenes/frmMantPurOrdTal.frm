VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Begin VB.Form frmMantPurOrdTal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de PO/Talla"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   8115
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   915
      TabIndex        =   5
      Top             =   6840
      Width           =   1965
      Begin VB.CommandButton cmdClose 
         Height          =   495
         Left            =   4575
         Picture         =   "frmMantPurOrdTal.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Cerrar"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNew 
         Height          =   495
         Left            =   2160
         Picture         =   "frmMantPurOrdTal.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Nuevo"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdEdit 
         Height          =   495
         Left            =   2640
         Picture         =   "frmMantPurOrdTal.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Editar"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   495
         Left            =   3120
         Picture         =   "frmMantPurOrdTal.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Eliminar"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdSave 
         Height          =   495
         Left            =   3600
         Picture         =   "frmMantPurOrdTal.frx":05C8
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Grabar"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdUndo 
         Height          =   495
         Left            =   4080
         Picture         =   "frmMantPurOrdTal.frx":073A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Deshacerundo"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "frmMantPurOrdTal.frx":08AC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Ultimo"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   0
         Picture         =   "frmMantPurOrdTal.frx":0A1E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Primero"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "frmMantPurOrdTal.frx":0B90
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "frmMantPurOrdTal.frx":0D02
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Anterior"
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame Fradetalle 
      BackColor       =   &H00FFC0C0&
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
      Height          =   3285
      Left            =   30
      TabIndex        =   3
      Tag             =   "Detail"
      Top             =   3570
      Width           =   8025
      Begin VB.TextBox txtDes_Adicional_Partida 
         Height          =   285
         Left            =   1455
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1965
         Width           =   6390
      End
      Begin VB.TextBox txtComposicion 
         Height          =   285
         Left            =   1455
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2265
         Width           =   6390
      End
      Begin VB.TextBox txtNum_Partida_Arancelaria 
         Height          =   285
         Left            =   1455
         TabIndex        =   31
         Top             =   1350
         Width           =   1620
      End
      Begin VB.TextBox txtDes_Partida_Arancelaria 
         Height          =   285
         Left            =   3075
         TabIndex        =   30
         Top             =   1350
         Width           =   4740
      End
      Begin VB.TextBox txtSec_Partida_Arancelaria 
         Height          =   285
         Left            =   1455
         TabIndex        =   29
         Top             =   1650
         Width           =   600
      End
      Begin VB.TextBox txtDes_SecPartida_Arancelaria 
         Height          =   285
         Left            =   2040
         TabIndex        =   28
         Top             =   1650
         Width           =   5790
      End
      Begin VB.TextBox txtNum_Categoria_Internacional 
         Height          =   285
         Left            =   1455
         TabIndex        =   27
         Top             =   2595
         Width           =   600
      End
      Begin VB.TextBox txtDes_Categoria_Internacional 
         Height          =   285
         Left            =   2040
         TabIndex        =   26
         Top             =   2595
         Width           =   5820
      End
      Begin VB.TextBox txtNum_Partida_Arancelaria_Exterior 
         Height          =   285
         Left            =   1455
         TabIndex        =   25
         Top             =   2910
         Width           =   1740
      End
      Begin VB.TextBox TxtCod_Talla 
         Height          =   285
         Left            =   5010
         TabIndex        =   24
         Top             =   615
         Width           =   915
      End
      Begin VB.TextBox TxtCod_PurOrd 
         Height          =   285
         Left            =   1455
         TabIndex        =   23
         Top             =   615
         Width           =   2535
      End
      Begin VB.TextBox TxtCod_EstCli 
         Height          =   285
         Left            =   5010
         TabIndex        =   22
         Top             =   315
         Width           =   2535
      End
      Begin VB.TextBox TxtNom_Cliente 
         Height          =   285
         Left            =   1455
         TabIndex        =   21
         Top             =   315
         Width           =   2535
      End
      Begin VB.TextBox txtOrden 
         Height          =   315
         Left            =   1455
         MaxLength       =   3
         TabIndex        =   1
         Top             =   915
         Width           =   870
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Descrip. Adicional"
         Height          =   195
         Left            =   60
         TabIndex        =   39
         Top             =   2010
         Width           =   1275
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Composición"
         Height          =   195
         Left            =   60
         TabIndex        =   38
         Top             =   2310
         Width           =   900
      End
      Begin VB.Line Line1 
         X1              =   30
         X2              =   8970
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Partida Arancelaria"
         Height          =   195
         Left            =   60
         TabIndex        =   35
         Top             =   1395
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Categ Internacional"
         Height          =   195
         Left            =   60
         TabIndex        =   34
         Top             =   2640
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sec. Partida"
         Height          =   195
         Left            =   30
         TabIndex        =   33
         Top             =   1695
         Width           =   870
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Part H.T.S."
         Height          =   240
         Left            =   60
         TabIndex        =   32
         Top             =   2932
         Width           =   810
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Orden:"
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
         Left            =   60
         TabIndex        =   20
         Tag             =   "Number:"
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Talla :"
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
         Left            =   4290
         TabIndex        =   19
         Tag             =   "Size :"
         Top             =   645
         Width           =   420
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
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
         Index           =   2
         Left            =   4290
         TabIndex        =   18
         Tag             =   "Style:"
         Top             =   352
         Width           =   420
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "PO :"
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
         Left            =   60
         TabIndex        =   17
         Tag             =   "PO :"
         Top             =   645
         Width           =   300
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cliente:"
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
         Left            =   60
         TabIndex        =   4
         Tag             =   "Client:"
         Top             =   352
         Width           =   525
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
      Height          =   3585
      Left            =   0
      TabIndex        =   0
      Tag             =   "List"
      Top             =   -15
      Width           =   8040
      Begin MSDataGridLib.DataGrid DGridlista 
         Height          =   3285
         Left            =   60
         TabIndex        =   2
         Top             =   225
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   5794
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648384
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
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "cliente"
            Caption         =   "Cliente"
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
            DataField       =   "po"
            Caption         =   "P.O"
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
            DataField       =   "estilo"
            Caption         =   "Estilo"
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
            DataField       =   "talla"
            Caption         =   "Talla"
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
         BeginProperty Column04 
            DataField       =   "orden"
            Caption         =   "Orden"
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
         BeginProperty Column05 
            DataField       =   "Num_Partida_Arancelaria"
            Caption         =   "Num_Partida_Arancelaria"
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
            DataField       =   "Des_Partida_Arancelaria"
            Caption         =   "Des_Partida_Arancelaria"
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
            DataField       =   "Sec_Partida_Arancelaria"
            Caption         =   "Sec_Partida_Arancelaria"
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
            DataField       =   "Des_Partida"
            Caption         =   "Des_Partida"
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
            DataField       =   "Des_Adicional_Partida"
            Caption         =   "Des_Adicional_Partida"
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
            DataField       =   "Composicion"
            Caption         =   "Composicion"
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
            DataField       =   "Num_Categoria_Internacional"
            Caption         =   "Num_Categoria_Internacional"
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
            DataField       =   "Num_Partida_Arancelaria_Exteri"
            Caption         =   "Num_Partida_Arancelaria_Exterior"
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
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1260.284
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
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
            EndProperty
         EndProperty
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2895
      TabIndex        =   16
      Top             =   6915
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantPurOrdTal.frx":0E74
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmMantPurOrdTal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstTalla As New ADODB.Recordset
Dim TmpCodigo As String
Dim Estado As String
Dim Strsql As String
Public Cliente As String
Public PO As String
Public Estilo As String
Public Codigo As String, Descripcion As String, TipoAdd As String, TipoAdd2 As String

Sub Habilita(pTipo As Byte)
If pTipo = 1 Then 'nada habilitado
    txtNom_Cliente.Enabled = False
    txtCod_EstCli.Enabled = False
    txtCod_PurOrd.Enabled = False
    txtCod_Talla.Enabled = False
    txtOrden.Enabled = False
    txtNum_Partida_Arancelaria.Enabled = False
    txtDes_Partida_Arancelaria.Enabled = False
    txtSec_Partida_Arancelaria.Enabled = False
    txtDes_SecPartida_Arancelaria.Enabled = False
    txtDes_Adicional_Partida.Enabled = False
    txtComposicion.Enabled = False
    txtNum_Categoria_Internacional.Enabled = False
    txtNum_Partida_Arancelaria_Exterior.Enabled = False
    
ElseIf pTipo = 2 Then 'nuevo
    'TxtCod_Talla.Enabled = False
    txtOrden.Enabled = True
    
    txtNum_Partida_Arancelaria.Enabled = True
    txtDes_Partida_Arancelaria.Enabled = True
    txtSec_Partida_Arancelaria.Enabled = True
    txtDes_SecPartida_Arancelaria.Enabled = True
    txtDes_Adicional_Partida.Enabled = True
    txtComposicion.Enabled = True
    txtNum_Categoria_Internacional.Enabled = True
    txtNum_Partida_Arancelaria_Exterior.Enabled = True
    
ElseIf pTipo = 3 Then 'editar
    txtNom_Cliente.Enabled = False
    txtCod_EstCli.Enabled = False
    txtCod_PurOrd.Enabled = False
    txtCod_Talla.Enabled = False
    txtOrden.Enabled = True
  
    txtNum_Partida_Arancelaria.Enabled = True
    txtDes_Partida_Arancelaria.Enabled = True
    txtSec_Partida_Arancelaria.Enabled = True
    txtDes_SecPartida_Arancelaria.Enabled = True
    txtDes_Adicional_Partida.Enabled = True
    txtComposicion.Enabled = True
    txtNum_Categoria_Internacional.Enabled = True
    txtNum_Partida_Arancelaria_Exterior.Enabled = True
  
End If
End Sub

Sub Limpia()
'Me.cmbTalla.ListIndex = -1
Me.txtOrden = ""
End Sub

Sub ACTUALIZAR(pCliente As String, pPO As String, pEstilo As String, pTalla As String, pOrden As Integer, Num_Partida_Arancelaria As String, Sec_Partida_Arancelaria As String, Num_Categoria_Internacional As String, Num_Partida_Arancelaria_Exterior As String)
On Error GoTo hand:
B_db.Execute "SP_Tg_Purordtal 'modificar','" & pCliente & "','" & pPO & "','" & pEstilo & "','" & pTalla & "'," & pOrden & " ,'" & Num_Partida_Arancelaria & "','" & Sec_Partida_Arancelaria & "','" & Num_Categoria_Internacional & "','" & Num_Partida_Arancelaria_Exterior & "'"
Cargar_Data
Exit Sub
hand:
    ErrorHandler Err, "ACTUALIZAR"
End Sub

Sub Cargar_Data()
On Error GoTo hand
Set RstTalla = Nothing
RstTalla.CursorType = adOpenStatic
RstTalla.CursorLocation = adUseClient
Screen.MousePointer = vbHourglass
RstTalla.Open "SP_Tg_Purordtal 'ver','" & Cliente & "','" & Me.PO & "','" & Estilo & "','',0", cCONNECT
Screen.MousePointer = vbDefault
Set Me.DGridlista.DataSource = Nothing
If RstTalla.RecordCount > 0 Then
    Set DGridlista.DataSource = RstTalla
    DGridlista_RowColChange 0, 0
End If
Exit Sub
hand:
    ErrorHandler Err, "CARGAR_DATA"
End Sub

Sub INSERTAR(pCliente As String, pPO As String, pEstilo As String, pTalla As String, pOrden As Integer, Num_Partida_Arancelaria As String, Sec_Partida_Arancelaria As String, Num_Categoria_Internacional As String, Num_Partida_Arancelaria_Exterior As String)
On Error GoTo hand:
B_db.Execute "SP_Tg_Purordtal 'Insertar','" & pCliente & "','" & pPO & "','" & pEstilo & "','" & pTalla & "'," & pOrden & " ,'" & Num_Partida_Arancelaria & "','" & Sec_Partida_Arancelaria & "','" & Num_Categoria_Internacional & "','" & Num_Partida_Arancelaria_Exterior & "'"
Cargar_Data
Exit Sub
hand:
ErrorHandler Err, "INSERTAR"
End Sub

Private Sub cmdFirst_Click()
RstTalla.MoveFirst
End Sub

Private Sub cmdLast_Click()
RstTalla.MoveLast
End Sub

Private Sub cmdNext_Click()
If Not RstTalla.EOF Then RstTalla.MoveNext
End Sub

Private Sub cmdPrevious_Click()
If Not RstTalla.BOF Then RstTalla.MovePrevious
End Sub

Private Sub DGridlista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If (RstTalla.EOF = False) Then
        If (RstTalla.BOF = False) Then
            txtNom_Cliente.Text = Me.DGridlista.Columns(0).Text
            txtCod_PurOrd.Text = Me.DGridlista.Columns(1).Text
            txtCod_EstCli.Text = Me.DGridlista.Columns(2).Text
            'BuscaCombo DGridlista.Columns(3).Text, 1, Me.cmbTalla
            txtCod_Talla.Text = DGridlista.Columns(3).Text
            txtOrden = DGridlista.Columns(4).Text
            
            txtNum_Partida_Arancelaria.Text = Trim(DGridlista.Columns(5).Text)
            txtDes_Partida_Arancelaria.Text = Trim(DGridlista.Columns(6).Text)
            txtSec_Partida_Arancelaria.Text = Trim(DGridlista.Columns(7).Text)
            txtDes_SecPartida_Arancelaria.Text = Trim(DGridlista.Columns(8).Text)
            txtDes_Adicional_Partida.Text = Trim(DGridlista.Columns(9).Text)
            txtComposicion.Text = Trim(DGridlista.Columns(10).Text)
           
            txtNum_Categoria_Internacional.Text = Trim(DGridlista.Columns(11).Text)
            txtNum_Partida_Arancelaria_Exterior.Text = Trim(DGridlista.Columns(12).Text)
        End If
    End If
End Sub


Private Sub Form_Load()
FormateaGrid Me.DGridlista
Call FormSet(Me)
Set B_db = Nothing
B_db.Open cCONNECT
'Me.MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
HabilitaMant Me.MantFunc1, "MODIFICAR"
Limpia

'LlenaCombo Me.cmbCliente, "Select nom_cliente +space(100)+cod_cliente from tg_cliente order by nom_cliente", cCONNECT
'LlenaCombo Me.cmbEstilo, "select des_estcli +space(100)+cod_estcli from tg_estcli order by des_estcli", cCONNECT
'LlenaCombo Me.cmbPO, "select cod_purord from tg_purord order by 1", cCONNECT
'LlenaCombo Me.cmbTalla, "select * from tg_talla order by 1", cCONNECT

Habilita 1
'CARGAR_DATA

End Sub


Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo hand:
Select Case ActionName
    Case "ADICIONAR"
        Estado = "Adicionar"
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        DGridlista.Enabled = False
        Habilita 2
        Limpia
    Case "MODIFICAR"
        If RstTalla.RecordCount > 0 Then
            Habilita 3
            Estado = "Modificar"
            txtOrden.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        End If
    Case "ELIMINAR"
'         ELIMINAR txtCodigo
'         txtCodigo.Enabled = False
    Case "GRABAR"
            
            'HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            HabilitaMant Me.MantFunc1, "MODIFICAR"
            If Estado = "Modificar" Then
                ACTUALIZAR Me.Cliente, Me.PO, Me.Estilo, txtCod_Talla, txtOrden, txtNum_Partida_Arancelaria.Text, txtSec_Partida_Arancelaria.Text, txtNum_Categoria_Internacional.Text, txtNum_Partida_Arancelaria_Exterior.Text
            Else
                INSERTAR Me.Cliente, Me.PO, Me.Estilo, txtCod_Talla, txtOrden, txtNum_Partida_Arancelaria.Text, txtSec_Partida_Arancelaria.Text, txtNum_Categoria_Internacional.Text, txtNum_Partida_Arancelaria_Exterior.Text
            End If
            Habilita 1
            DGridlista.Enabled = True
    Case "DESHACER"
        Cargar_Data
        Habilita 1
        HabilitaMant Me.MantFunc1, "MODIFICAR"
        DGridlista.Enabled = True
    Case "SALIR"
        Unload Me
End Select
Exit Sub
hand:
ErrorHandler Err, "MantFunc1_ActionClick"
End Sub



Private Sub txtOrden_KeyPress(KeyAscii As Integer)
SoloNumeros txtOrden, KeyAscii, False, 0, 3
End Sub


Private Sub txtNum_Partida_Arancelaria_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaPartidaArancelaria 1
        SendKeys "{TAB}"
    End If
End Sub


Public Sub BuscaPartidaArancelaria(opcion As String)
Dim rstAux As ADODB.Recordset

    Strsql = "SELECT Num_Partida_Arancelaria, Des_Partida_Arancelaria FROM TG_PARTIDA_ARANCELARIA WHERE "
    
    txtNum_Partida_Arancelaria = Trim(txtNum_Partida_Arancelaria)
    txtDes_Partida_Arancelaria = Trim(txtDes_Partida_Arancelaria)
    
    Select Case opcion
    Case 1: Strsql = Strsql & "Num_Partida_Arancelaria like '%" & txtNum_Partida_Arancelaria & "%'"
    Case 2: Strsql = Strsql & "Des_Partida_Arancelaria LIKE '%" & txtDes_Partida_Arancelaria & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.sQuery = Strsql
    frmBusqGeneral3.Cargar_Datos
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    
    frmBusqGeneral3.gexLista.Columns("Num_Partida_Arancelaria").Width = 2000
    frmBusqGeneral3.gexLista.Columns("Des_Partida_Arancelaria").Width = 7000
    
    frmBusqGeneral3.gexLista.Columns("Num_Partida_Arancelaria").Caption = "Partida"
    frmBusqGeneral3.gexLista.Columns("Des_Partida_Arancelaria").Caption = "Descrip."
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.value = True
    End If
    
    txtNum_Partida_Arancelaria = ""
    txtDes_Partida_Arancelaria = ""
    
    If Codigo <> "" Then
        txtNum_Partida_Arancelaria = Codigo
        txtDes_Partida_Arancelaria = Descripcion
        
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    Codigo = ""
    Descripcion = ""
End Sub

Private Sub txtDes_Partida_Arancelaria_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaPartidaArancelaria 1
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtSec_Partida_Arancelaria_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaPartidaArancelaria_Detalle 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDes_SecPartida_Arancelaria_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Public Sub BuscaPartidaArancelaria_Detalle(opcion As String)
Dim rstAux As ADODB.Recordset

    Strsql = "SELECT Sec_Partida_Arancelaria, Des_Partida , Des_Adicional_Partida, Composicion FROM TG_PARTIDA_ARANCELARIA_DETALLE   WHERE Num_Partida_Arancelaria = '" & txtNum_Partida_Arancelaria & "' AND "
    
    txtSec_Partida_Arancelaria = Trim(txtSec_Partida_Arancelaria)
    txtDes_SecPartida_Arancelaria = Trim(txtDes_SecPartida_Arancelaria)
    
    Select Case opcion
    Case 1: Strsql = Strsql & "Sec_Partida_Arancelaria like '%" & txtSec_Partida_Arancelaria & "%'"
    Case 2: Strsql = Strsql & "Des_Partida  LIKE '%" & txtDes_SecPartida_Arancelaria & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.sQuery = Strsql
    frmBusqGeneral3.Cargar_Datos
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    
    frmBusqGeneral3.gexLista.Columns("Sec_Partida_Arancelaria").Width = 500
    frmBusqGeneral3.gexLista.Columns("Des_Partida").Width = 3000
    frmBusqGeneral3.gexLista.Columns("Des_Adicional_Partida").Width = 3000
    
    frmBusqGeneral3.gexLista.Columns("Sec_Partida_Arancelaria").Caption = "Sec.Partida"
    frmBusqGeneral3.gexLista.Columns("Des_Partida").Caption = "Descrip."
    frmBusqGeneral3.gexLista.Columns("Des_Adicional_Partida").Caption = "Descrip.Adic"
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.value = True
    End If
    
    txtSec_Partida_Arancelaria = ""
    txtDes_SecPartida_Arancelaria = ""
    
    If Codigo <> "" Then
        txtSec_Partida_Arancelaria = Codigo
        txtDes_SecPartida_Arancelaria = Descripcion
        txtDes_Adicional_Partida = TipoAdd
        txtComposicion = TipoAdd2
    End If
    
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    Codigo = ""
    Descripcion = ""
End Sub


Private Sub txtNum_Categoria_Internacional_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        BuscaCategoria 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDES_Categoria_Internacional_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        BuscaCategoria 2
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtNum_Partida_Arancelaria_Exterior_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Public Sub BuscaCategoria(opcion As String)
Dim rstAux As ADODB.Recordset

    Strsql = "SELECT Num_Categoria_Internacional, Des_Categoria_Internacional FROM TG_Categoria_Estilo WHERE "
    
    txtNum_Categoria_Internacional = Trim(txtNum_Categoria_Internacional)
    txtDes_Categoria_Internacional = Trim(txtDes_Categoria_Internacional)
    
    Select Case opcion
    Case 1: Strsql = Strsql & "Num_Categoria_Internacional like '%" & txtNum_Categoria_Internacional & "%'"
    Case 2: Strsql = Strsql & "Des_Categoria_Internacional LIKE '%" & txtDes_Categoria_Internacional & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.sQuery = Strsql
    frmBusqGeneral3.Cargar_Datos
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    
    frmBusqGeneral3.gexLista.Columns("Num_Categoria_Internacional").Width = 1500
    frmBusqGeneral3.gexLista.Columns("Des_Categoria_Internacional").Width = 7000
    
    frmBusqGeneral3.gexLista.Columns("Num_Categoria_Internacional").Caption = "Categoria"
    frmBusqGeneral3.gexLista.Columns("Des_Categoria_Internacional").Caption = "Descrip."
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.value = True
    End If
    
    txtNum_Categoria_Internacional = ""
    txtDes_Categoria_Internacional = ""
    
    If Codigo <> "" Then
        txtNum_Categoria_Internacional = Codigo
        txtDes_Categoria_Internacional = Descripcion
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    Codigo = ""
    Descripcion = ""
End Sub

