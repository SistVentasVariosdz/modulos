VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmLetraAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Añadir Facturas"
   ClientHeight    =   6915
   ClientLeft      =   405
   ClientTop       =   465
   ClientWidth     =   9840
   Icon            =   "frmLetraAdd.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   9840
   Begin VB.Frame frFacturas 
      Caption         =   "<"
      ForeColor       =   &H8000000D&
      Height          =   6855
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   9585
      Begin VB.TextBox txtNumeroxCancelar 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   885
         TabIndex        =   44
         Text            =   "0"
         Top             =   3420
         Width           =   1245
      End
      Begin VB.CommandButton cmdBackAll 
         BackColor       =   &H8000000D&
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5355
         TabIndex        =   43
         Top             =   3330
         Width           =   675
      End
      Begin VB.CommandButton CmdBack 
         BackColor       =   &H8000000D&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4635
         TabIndex        =   42
         Top             =   3330
         Width           =   675
      End
      Begin VB.CommandButton CmdNextAll 
         BackColor       =   &H8000000D&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3555
         TabIndex        =   41
         Top             =   3330
         Width           =   675
      End
      Begin VB.CommandButton CmdNext 
         BackColor       =   &H8000000D&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2835
         TabIndex        =   40
         Top             =   3330
         Width           =   675
      End
      Begin VB.TextBox TxtMonto1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   7830
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "0.00"
         Top             =   3420
         Width           =   1455
      End
      Begin VB.TextBox TxtMonto2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   7830
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "0.00"
         Top             =   6465
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   675
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   9555
         Begin VB.Frame frOpciones 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   120
            TabIndex        =   52
            Top             =   120
            Width           =   5535
            Begin VB.TextBox txtOrden 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   4440
               TabIndex        =   55
               Top             =   120
               Width           =   960
            End
            Begin VB.TextBox txtDes_TipVenta 
               Height          =   285
               Left            =   1800
               TabIndex        =   54
               Top             =   120
               Width           =   1905
            End
            Begin VB.TextBox txtCod_TipVenta 
               Height          =   285
               Left            =   1200
               MaxLength       =   4
               TabIndex        =   53
               Text            =   " "
               Top             =   120
               Width           =   480
            End
            Begin VB.Label Label9 
               Caption         =   "Tipo Venta :"
               Height          =   255
               Left            =   120
               TabIndex        =   57
               Top             =   135
               Width           =   915
            End
            Begin VB.Label Label10 
               Caption         =   "Orden :"
               Height          =   255
               Left            =   3840
               TabIndex        =   56
               Top             =   120
               Width           =   555
            End
         End
         Begin VB.TextBox txtNumeroPendiente 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   990
            TabIndex        =   50
            Text            =   "0"
            Top             =   240
            Width           =   1245
         End
         Begin FunctionsButtons.FunctButt fncBuscar 
            Height          =   375
            Left            =   6840
            TabIndex        =   16
            Top             =   210
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   661
            Custom          =   $"frmLetraAdd.frx":030A
            Orientacion     =   0
            Style           =   0
            Language        =   0
            TypeImageList   =   0
            ControlWidth    =   1155
            ControlHeigth   =   350
            ControlSeparator=   110
         End
         Begin VB.Label lbSeleccionar 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   "Opciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   5760
            MouseIcon       =   "frmLetraAdd.frx":0397
            MousePointer    =   99  'Custom
            TabIndex        =   58
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Label14 
            Caption         =   "Numero :"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   270
            Width           =   675
         End
      End
      Begin GridEX20.GridEX gexGrid2 
         Height          =   2475
         Left            =   120
         TabIndex        =   45
         Top             =   3900
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   4366
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         BackColorBkg    =   -2147483628
         ColumnHeaderHeight=   285
         IntProp1        =   0
         ColumnsCount    =   3
         Column(1)       =   "frmLetraAdd.frx":06A1
         Column(2)       =   "frmLetraAdd.frx":0795
         Column(3)       =   "frmLetraAdd.frx":0881
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmLetraAdd.frx":094D
         FormatStyle(2)  =   "frmLetraAdd.frx":0A85
         FormatStyle(3)  =   "frmLetraAdd.frx":0B35
         FormatStyle(4)  =   "frmLetraAdd.frx":0BE9
         FormatStyle(5)  =   "frmLetraAdd.frx":0CC1
         FormatStyle(6)  =   "frmLetraAdd.frx":0D79
         ImageCount      =   0
         PrinterProperties=   "frmLetraAdd.frx":0E59
      End
      Begin GridEX20.GridEX gexGrid1 
         Height          =   2475
         Left            =   120
         TabIndex        =   46
         Top             =   720
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   4366
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         BackColorBkg    =   -2147483628
         ColumnHeaderHeight=   285
         IntProp1        =   0
         ColumnsCount    =   3
         Column(1)       =   "frmLetraAdd.frx":1031
         Column(2)       =   "frmLetraAdd.frx":1125
         Column(3)       =   "frmLetraAdd.frx":1211
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmLetraAdd.frx":12DD
         FormatStyle(2)  =   "frmLetraAdd.frx":1415
         FormatStyle(3)  =   "frmLetraAdd.frx":14C5
         FormatStyle(4)  =   "frmLetraAdd.frx":1579
         FormatStyle(5)  =   "frmLetraAdd.frx":1651
         FormatStyle(6)  =   "frmLetraAdd.frx":1709
         ImageCount      =   0
         PrinterProperties=   "frmLetraAdd.frx":17E9
      End
      Begin VB.Label Label15 
         Caption         =   "Numero :"
         Height          =   255
         Left            =   165
         TabIndex        =   49
         Top             =   3450
         Width           =   675
      End
      Begin VB.Label Label5 
         Caption         =   "Total a Cancelar :"
         Height          =   255
         Left            =   6150
         TabIndex        =   48
         Top             =   6480
         Width           =   1395
      End
      Begin VB.Label Label17 
         Caption         =   "Total Pendiente :"
         Height          =   285
         Left            =   6510
         TabIndex        =   47
         Top             =   3435
         Width           =   1305
      End
   End
   Begin VB.Frame frLetra 
      Height          =   6855
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   9615
      Begin VB.TextBox txtTercero_NomAnexo 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3885
         MaxLength       =   30
         TabIndex        =   9
         Top             =   5760
         Width           =   4665
      End
      Begin VB.TextBox txtTercero_Des_TipAnex 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4245
         MaxLength       =   11
         TabIndex        =   36
         Top             =   5760
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtTercero_NumRuc 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1035
         MaxLength       =   11
         TabIndex        =   8
         Top             =   5760
         Width           =   1545
      End
      Begin VB.TextBox txtTercero_CodTipAnexo 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3480
         MaxLength       =   4
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "C"
         Top             =   5760
         Width           =   360
      End
      Begin VB.Frame frOrigen 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   240
         TabIndex        =   33
         Top             =   6240
         Visible         =   0   'False
         Width           =   5055
         Begin VB.TextBox txtCod_Origen 
            Height          =   315
            Left            =   810
            TabIndex        =   10
            Top             =   0
            Width           =   735
         End
         Begin VB.TextBox txtDes_Origen 
            Height          =   315
            Left            =   1680
            TabIndex        =   11
            Top             =   0
            Width           =   3015
         End
         Begin VB.Label Label12 
            Caption         =   "Origen :"
            Height          =   255
            Left            =   0
            TabIndex        =   34
            Top             =   30
            Width           =   885
         End
      End
      Begin VB.TextBox txtTotLetras 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   4440
         Width           =   1455
      End
      Begin VB.TextBox TxtDes_Moneda 
         Height          =   315
         Left            =   5670
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox TxtGlosa 
         Height          =   345
         Left            =   1035
         TabIndex        =   7
         Top             =   5310
         Width           =   7530
      End
      Begin VB.TextBox TxtCod_Moneda 
         Height          =   315
         Left            =   4800
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtNumero_Propio 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Left            =   1650
         TabIndex        =   21
         Top             =   240
         Width           =   1230
      End
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   120
         TabIndex        =   18
         Top             =   690
         Width           =   9375
         Begin VB.TextBox txtTotFact 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   28
            Text            =   "0.00"
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtFacturas 
            Enabled         =   0   'False
            Height          =   585
            Left            =   1110
            MultiLine       =   -1  'True
            TabIndex        =   19
            Top             =   240
            Width           =   7050
         End
         Begin VB.CommandButton cmdFacturas 
            Caption         =   "&Documentos"
            Height          =   585
            Left            =   8160
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Total Monto Facturas :"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   975
            Width           =   1515
         End
         Begin VB.Label Label4 
            Caption         =   "Documentos :"
            Height          =   225
            Left            =   120
            TabIndex        =   20
            Top             =   300
            Width           =   1005
         End
      End
      Begin VB.TextBox txtCod_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3480
         MaxLength       =   4
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   4890
         Width           =   360
      End
      Begin VB.TextBox txtNum_Ruc 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1035
         MaxLength       =   11
         TabIndex        =   5
         Top             =   4890
         Width           =   1545
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   7080
         TabIndex        =   12
         Top             =   6210
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmLetraAdd.frx":19C1
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   420
         Left            =   8280
         TabIndex        =   22
         Top             =   3930
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         Custom          =   "0~0~ELIMINAR~Verdadero~Verdadero~&Eliminar~0~0~1~~0~Falso~Falso~&Eliminar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   300
         ControlSeparator=   110
      End
      Begin GridEX20.GridEX gexLetra 
         Height          =   2070
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   8070
         _ExtentX        =   14235
         _ExtentY        =   3651
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         GridLineStyle   =   2
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowDelete     =   -1  'True
         BorderStyle     =   2
         GroupByBoxVisible=   0   'False
         NewRowPos       =   1
         ImageWidth      =   0
         ImageHeight     =   0
         DataMode        =   1
         AllowAddNew     =   -1  'True
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "frmLetraAdd.frx":1A57
         Column(2)       =   "frmLetraAdd.frx":1B1F
         FormatStylesCount=   9
         FormatStyle(1)  =   "frmLetraAdd.frx":1BC3
         FormatStyle(2)  =   "frmLetraAdd.frx":1CFB
         FormatStyle(3)  =   "frmLetraAdd.frx":1DAB
         FormatStyle(4)  =   "frmLetraAdd.frx":1E5F
         FormatStyle(5)  =   "frmLetraAdd.frx":1F37
         FormatStyle(6)  =   "frmLetraAdd.frx":1FEF
         FormatStyle(7)  =   "frmLetraAdd.frx":20CF
         FormatStyle(8)  =   "frmLetraAdd.frx":24DB
         FormatStyle(9)  =   "frmLetraAdd.frx":28EB
         ImageCount      =   0
         PrinterProperties=   "frmLetraAdd.frx":2A73
      End
      Begin VB.TextBox txtDes_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3885
         MaxLength       =   30
         TabIndex        =   6
         Top             =   4890
         Width           =   4665
      End
      Begin VB.TextBox txtDes_TipAnex 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4320
         MaxLength       =   11
         TabIndex        =   31
         Top             =   4890
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   2805
         TabIndex        =   37
         Top             =   5805
         Width           =   645
      End
      Begin VB.Label Label11 
         Caption         =   "Ruc Tercero"
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   5655
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Total Monto Letras  :"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   4455
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ultimo Nro Letra :"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   270
         Width           =   1230
      End
      Begin VB.Label Label7 
         Caption         =   "Glosa :"
         Height          =   225
         Left            =   240
         TabIndex        =   26
         Top             =   5370
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Moneda :"
         Height          =   255
         Left            =   3390
         TabIndex        =   25
         Top             =   270
         Width           =   885
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Aval :"
         Height          =   195
         Left            =   3000
         TabIndex        =   24
         Top             =   4935
         Width           =   405
      End
      Begin VB.Label Label28 
         Caption         =   "R.U.C."
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   4905
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmLetraAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public varCod_anxo As String
Public varCod_TipAnex As String
Public strOption As String
'Public Cod_TipDoc As String

Public codigo As String
Public Descripcion As String

Dim RsGrid1 As ADODB.Recordset
Dim RsGrid2 As ADODB.Recordset

Public Num_Correlativo As String, strCod_Anxo As String, strTercero_Cod_Anxo As String, intCod_Grupo As Integer, strNum_Corre_Let_Renov As String

Dim Num_Voucher As String
Dim Cod_Subdiario As String

Dim bNext As Boolean

Public sSLIMConnect As String, strRenovacion As String


Private Sub CmdBack_Click()
    If gexGrid2.RowCount = 0 Then Exit Sub
    RsGrid1.AddNew
    RsGrid1.Fields("Correlativo").Value = gexGrid2.Value(gexGrid1.Columns("Correlativo").Index)
    RsGrid1.Fields("Numero").Value = gexGrid2.Value(gexGrid1.Columns("Numero").Index)
    RsGrid1.Fields("Fecha").Value = gexGrid2.Value(gexGrid1.Columns("Fecha").Index)
    RsGrid1.Fields("Tipo_Cambio").Value = gexGrid2.Value(gexGrid1.Columns("Tipo_Cambio").Index)
    RsGrid1.Fields("Moneda").Value = gexGrid2.Value(gexGrid1.Columns("Moneda").Index)
    RsGrid1.Fields("Imp_Total").Value = gexGrid2.Value(gexGrid1.Columns("Imp_Total").Index)
    RsGrid1.Fields("Monto_Origen1").Value = gexGrid2.Value(gexGrid1.Columns("Monto_Origen1").Index)
    RsGrid1.Fields("Monto_Origen").Value = gexGrid2.Value(gexGrid1.Columns("Monto_Origen").Index)
    RsGrid1.Fields("Monto_Aceptado").Value = gexGrid2.Value(gexGrid1.Columns("Monto_Aceptado").Index)
    
    RsGrid1.Update

    RsGrid2.MoveFirst
    Call BuscaCampo(RsGrid2, "Correlativo", gexGrid2.Value(gexGrid2.Columns("Correlativo").Index))

    RsGrid2.Delete

    Set gexGrid1.ADORecordset = RsGrid1
    ConfigurarGrid gexGrid1
    Set gexGrid2.ADORecordset = RsGrid2
    ConfigurarGrid gexGrid2

    If RsGrid1.RecordCount = 0 Then
        TxtMonto1.Text = "0.00"
    Else
        TxtMonto1 = CALCULA_MONTO_TOTAL(RsGrid1)
    End If

    If RsGrid2.RecordCount = 0 Then
        TxtMonto2.Text = "0.00"
    Else
        TxtMonto2 = CALCULA_MONTO_TOTAL(RsGrid2)
    End If

End Sub

Private Sub cmdBackAll_Click()
Dim i As Integer

    If gexGrid2.RowCount > 0 Then
        VB.Screen.MousePointer = 11
        gexGrid2.Redraw = False
        gexGrid1.Redraw = False
        gexGrid2.MoveFirst

        For i = 1 To gexGrid2.RowCount
            CmdBack_Click
        Next


        gexGrid2.Redraw = True
        gexGrid1.Redraw = True

        VB.Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdFacturas_Click()
  If DevuelveCampo("select count(*) from tg_moneda where cod_moneda='" & Trim(TxtCod_Moneda.Text) & "'", cCONNECT) = 1 Then
    frLetra.Visible = False
    frFacturas.Top = 0
    frFacturas.Visible = True
  Else
    MsgBox "Seleccione una Moneda Valida", vbInformation, Me.Caption
    TxtCod_Moneda.SetFocus
  End If
End Sub

Private Sub cmdNext_Click()

    If gexGrid1.RowCount = 0 Then Exit Sub
    
    If gexGrid1.EditMode = jgexEditModeOn Then
      MsgBox "Salga del Modo de Edicion de la Grilla" & vbCr & "Haga Click en la columna Numero", vbInformation, "IMPORTANTE"
      Exit Sub
    End If
    
    RsGrid2.AddNew
    RsGrid2.Fields("Correlativo").Value = gexGrid1.Value(gexGrid1.Columns("Correlativo").Index)
    RsGrid2.Fields("Numero").Value = gexGrid1.Value(gexGrid1.Columns("Numero").Index)
    RsGrid2.Fields("Fecha").Value = gexGrid1.Value(gexGrid1.Columns("Fecha").Index)
    RsGrid2.Fields("Moneda").Value = gexGrid1.Value(gexGrid1.Columns("Moneda").Index)
    RsGrid2.Fields("Imp_Total").Value = gexGrid1.Value(gexGrid1.Columns("Imp_Total").Index)
    RsGrid2.Fields("Tipo_Cambio").Value = gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index)
    RsGrid2.Fields("Monto_Origen1").Value = gexGrid1.Value(gexGrid1.Columns("Monto_Origen1").Index)
    RsGrid2.Fields("Monto_Origen").Value = gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index)
    RsGrid2.Fields("Monto_Aceptado").Value = gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index)

    RsGrid2.Update

    RsGrid1.MoveFirst
    Call BuscaCampo(RsGrid1, "Correlativo", gexGrid1.Value(gexGrid1.Columns("Correlativo").Index))

    RsGrid1.Delete

    Set gexGrid1.ADORecordset = RsGrid1
    ConfigurarGrid gexGrid1
    Set gexGrid2.ADORecordset = RsGrid2
    ConfigurarGrid gexGrid2

    If RsGrid1.RecordCount = 0 Then
        TxtMonto1.Text = "0.00"
    Else
        TxtMonto1 = CALCULA_MONTO_TOTAL(RsGrid1)
    End If

    If RsGrid2.RecordCount = 0 Then
        TxtMonto2.Text = "0.00"
    Else
        TxtMonto2 = CALCULA_MONTO_TOTAL(RsGrid2)
    End If
End Sub

Private Sub CmdNextAll_Click()
Dim i As Integer
    If gexGrid1.RowCount > 0 Then
        gexGrid2.Redraw = False
        gexGrid1.Redraw = False
        gexGrid1.MoveFirst

        For i = 1 To gexGrid1.RowCount
            cmdNext_Click
        Next


        gexGrid2.Redraw = True
        gexGrid1.Redraw = True

    End If
End Sub

Private Sub fncBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

Select Case ActionName
  Case "BUSCAR"
    CARGA_FACTURAS_PENDIENTES
    frOpciones.Visible = False
  Case "CERRAR"
    frFacturas.Visible = False
    frLetra.Visible = True
    Resumen_Facturas
    Calcula_Monto_Letra
    gexLetra.SetFocus
End Select
End Sub

Private Sub Form_Load()
  txtNumero_Propio = DevuelveCampo("select dbo.uf_StrZero (Ult_Letra_Por_Cobrar + 1,6) from cn_control", cCONNECT)
  sSLIMConnect = "Provider=IBM.UniOLEDB.1;Password=activity;Persist Security Info=True;User ID=planilla;Data Source=HIALPESA_H;Location=C:\ACTIVITY_HIALPESA"
  Carga_Temporal
  bNext = True
End Sub

Sub Carga_Temporal()
Set gexLetra.ADORecordset = CargarRecordSetDesconectado("Ventas_Letras_A_Generar", cCONNECT)
gexLetra.Columns("Imp_Total").Format = "###,###.00"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim sMessage As Long
Dim strSQL As String

On Error GoTo dprError

Select Case ActionName
  Case "ACEPTAR"
    If CDbl(txtTotFact) = 0 Then
      MsgBox "Tiene q Asignar algunas Facturas", vbInformation, "AVISO"
      Exit Sub
    End If
    
    If CDbl(txtTotLetras) = 0 Then
      MsgBox "Tiene q Ingresar algunas Letras", vbInformation, "AVISO"
      Exit Sub
    End If
    
    If frOrigen.Visible And txtCod_Origen = "" Then
      MsgBox "Tiene q seleccionar un origen", vbInformation, "AVISO"
      txtCod_Origen.SetFocus
      Exit Sub
    End If
    
    If CDbl(txtTotFact) <> CDbl(txtTotLetras) Then
      MsgBox "Monto Total Facturas : " & txtTotFact & vbCr & "Monto Total Letras     : " & txtTotLetras & vbCr & " NO Cuadra", vbInformation, "AVISO"
      Exit Sub
    End If
    
    
    If MsgBox("Esta seguro de generar " & gexLetra.RowCount & " letras ", vbYesNo, "IMPORTANTE") = vbYes Then
      If lfSalvar_Datos Then Unload Me
    End If
   
  Case "CANCELAR"
    If MsgBox("Esta seguro de generar de Cancelar esta operacion ", vbYesNo, "IMPORTANTE") = vbYes Then
      Unload Me
    End If
End Select

Exit Sub

dprError:

errores err.Number

End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
  gexLetra.Delete
  Calcula_Monto_Letra
End Sub

Private Sub gexGrid1_AfterColEdit(ByVal ColIndex As Integer)
  Select Case ColIndex
  
  Case Is = gexGrid1.Columns("Monto_Aceptado").Index
  
    If CDbl(gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index)) <= 0 Then
      MsgBox "El Monto del documento debe ser Mayor a Cero", vbInformation, "AVISO"
      gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) = gexGrid1.Value(gexGrid1.Columns("Monto_Origen1").Index)
      Exit Sub
    End If
    
    If TxtCod_Moneda <> gexGrid1.Value(gexGrid1.Columns("Moneda").Index) Then
      gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) = IIf(gexGrid1.Value(gexGrid1.Columns("Moneda").Index) = "SOL", gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) * gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index), Format(gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) / gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index), "###,###.00"))
    Else
      gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) = gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index)
    End If

    If CDbl(gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index)) > CDbl(gexGrid1.Value(gexGrid1.Columns("Monto_Origen1").Index)) Then
      MsgBox "El Monto del documento es Mayor a monto Pendiente", vbInformation, "AVISO"
      gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) = gexGrid1.Value(gexGrid1.Columns("Monto_Origen1").Index)
      
      If TxtCod_Moneda <> gexGrid1.Value(gexGrid1.Columns("Moneda").Index) Then
        gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) = IIf(gexGrid1.Value(gexGrid1.Columns("Moneda").Index) = "SOL", gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) / gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index), Format(gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) * gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index), "###,###.00"))
      Else
        gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) = gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index)
      End If
      
      SendKeys "{ENTER}"

      Exit Sub
    End If
    
    SendKeys "{ENTER}"
    
  Case Is = gexGrid1.Columns("Tipo_Cambio").Index
  
      If TxtCod_Moneda <> gexGrid1.Value(gexGrid1.Columns("Moneda").Index) Then
        gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) = IIf(TxtCod_Moneda = "SOL", gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) * gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index), Format(gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) / gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index), "###,###.00"))
      Else
        gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) = gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index)
      End If
      
      SendKeys "{ENTER}"
    
  End Select
End Sub

Private Sub gexGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
  Select Case ColIndex
  Case Is = gexGrid1.Columns("Monto_Aceptado").Index
    Cancel = False
  Case Is = gexGrid1.Columns("Tipo_Cambio").Index
    Cancel = False
  Case Else
    Cancel = True
  End Select
End Sub

Private Sub gexGrid1_DblClick()

If gexGrid1.RowCount = 0 Then Exit Sub

Load frmMuestraDetalleDocumVentas
With frmMuestraDetalleDocumVentas
  .Caption = Trim(Me.Caption) & "  " & gexGrid1.Value(gexGrid1.Columns("Numero").Index)
  .strSQL = "Ventas_Muestra_Detalle_Factura_Items '" & gexGrid1.Value(gexGrid1.Columns("Correlativo").Index) & "'"
  .Num_Corre = gexGrid1.Value(gexGrid1.Columns("Correlativo").Index)
  .Buscar
  .FunctButt1.Visible = False
  .Show 1
End With

End Sub

Private Sub gexLetra_AfterColEdit(ByVal ColIndex As Integer)
Select Case ColIndex
Case Is = gexLetra.Columns("Nro").Index
  gexLetra.Value(gexLetra.Columns("Fec_Emidoc").Index) = Date
  Calcula_Monto_Letra
Case Is = gexLetra.Columns("Dias").Index
  If IIf(IsNull(gexLetra.Value(gexLetra.Columns("dias").Index)), 0, 1) <> 0 Then gexLetra.Value(gexLetra.Columns("Fec_VenDoc").Index) = gexLetra.Value(gexLetra.Columns("Fec_Emidoc").Index) + CInt(gexLetra.Value(gexLetra.Columns("dias").Index))
Case Is = gexLetra.Columns("Fec_EmiDoc").Index
  If IIf(IsNull(gexLetra.Value(gexLetra.Columns("dias").Index)), 0, 1) <> 0 Then gexLetra.Value(gexLetra.Columns("Fec_VenDoc").Index) = gexLetra.Value(gexLetra.Columns("Fec_Emidoc").Index) + CInt(gexLetra.Value(gexLetra.Columns("dias").Index))
Case Is = gexLetra.Columns("Fec_VenDoc").Index
  gexLetra.Value(gexLetra.Columns("Dias").Index) = gexLetra.Value(gexLetra.Columns("Fec_VenDoc").Index) - gexLetra.Value(gexLetra.Columns("Fec_EmiDoc").Index)
Case Is = gexLetra.Columns("Imp_Total").Index
  Calcula_Total_Letra
End Select
  
End Sub

Private Sub gexLetra_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
Select Case ColIndex
Case Is = gexLetra.Columns("Nro").Index
  Cancel = False
Case Is = gexLetra.Columns("Fec_EmiDoc").Index
  If IsEmpty(gexLetra.Value(gexLetra.Columns("Nro").Index)) Then
    Cancel = True
  Else
    Cancel = False
  End If
Case Is = gexLetra.Columns("Fec_VenDoc").Index
  If IsEmpty(gexLetra.Value(gexLetra.Columns("Nro").Index)) Then
    Cancel = True
  Else
    Cancel = False
  End If
Case Is = gexLetra.Columns("Dias").Index
  If IsEmpty(gexLetra.Value(gexLetra.Columns("Nro").Index)) Then
    Cancel = True
  Else
    Cancel = False
  End If
Case Is = gexLetra.Columns("Imp_Total").Index
  If IsEmpty(gexLetra.Value(gexLetra.Columns("Nro").Index)) Then
    Cancel = True
  Else
    Cancel = False
  End If
Case Else
  Cancel = True
End Select
End Sub

Private Sub lbSeleccionar_Click()
  frOpciones.Visible = True
End Sub

Private Sub lbSeleccionar_DblClick()
  frOpciones.Visible = False
End Sub

Private Sub txtCod_Moneda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Moneda", "Nom_Moneda", "TG_Moneda where ", TxtCod_Moneda, TxtDes_Moneda, 1, Me)
End Sub

Private Sub txtCod_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Origen_Letra", "Descripcion", "cn_origen_letras where  flg_protesto='N' and cod_origen_letra<>'N' and ", txtCod_Origen, txtDes_Origen, 1, Me)
End Sub

Private Sub txtCod_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAne, txtDes_TipAnex, 1, Me)
End Sub

Private Sub txtCod_TipVenta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_Tipo_Venta", "Descripcion", "Cn_Tipos_Venta where ", txtCod_TipVenta, txtDes_TipVenta, 1, Me)
  End If
End Sub

Private Sub txtDes_Moneda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Moneda", "Nom_Moneda", "TG_Moneda where ", TxtCod_Moneda, TxtDes_Moneda, 2, Me)
End Sub

Public Sub CARGA_FACTURAS_PENDIENTES()
On Error GoTo hand
Dim SQL As String

If strOption = "U" Then
    TxtDes_Moneda.Text = DevuelveCampo("SELECT NOM_MONEDA FROM TG_MONEDA WHERE COD_MONEDA='" & Trim(TxtCod_Moneda.Text) & "'", cCONNECT)
End If

Set RsGrid1 = CreateObject("ADODB.Recordset")
RsGrid1.CursorLocation = adUseClient

If strOption = "U" Then
    SQL = "EXEC Ventas_Letras_Docum_Pedientes '" & varCod_TipAnex & "','" & varCod_anxo & "','" & TxtCod_Moneda & "','','','" & Num_Correlativo & "','" & strRenovacion & "'"
Else
    SQL = "EXEC Ventas_Letras_Docum_Pedientes '" & varCod_TipAnex & "','" & varCod_anxo & "','" & TxtCod_Moneda & "','" & txtCod_TipVenta & "','" & txtOrden & "','','" & strRenovacion & "'"
End If

Set RsGrid1 = CargarRecordSetDesconectado(SQL, cCONNECT)

Set gexGrid1.ADORecordset = RsGrid1
ConfigurarGrid gexGrid1

TxtMonto1 = 0
TxtMonto2 = 0
If RsGrid1.RecordCount Then

    TxtMonto1 = CALCULA_MONTO_TOTAL(RsGrid1)

    Set RsGrid2 = CreateObject("ADODB.Recordset")
    RsGrid2.CursorLocation = adUseClient
    Set RsGrid2.ActiveConnection = Nothing

    RsGrid2.Fields.Append RsGrid1.Fields("Correlativo").Name, RsGrid1.Fields(0).Type, RsGrid1.Fields(0).DefinedSize
    RsGrid2.Fields.Append RsGrid1.Fields("Numero").Name, RsGrid1.Fields(1).Type, RsGrid1.Fields(1).DefinedSize
    RsGrid2.Fields.Append RsGrid1.Fields("Fecha").Name, adDate
    RsGrid2.Fields.Append RsGrid1.Fields("Moneda").Name, RsGrid1.Fields(3).Type, RsGrid1.Fields(3).DefinedSize
    RsGrid2.Fields.Append RsGrid1.Fields("Imp_Total").Name, adDouble
    RsGrid2.Fields.Append RsGrid1.Fields("Tipo_Cambio").Name, adDouble
    RsGrid2.Fields.Append RsGrid1.Fields("Monto_Origen1").Name, adDouble
    RsGrid2.Fields.Append RsGrid1.Fields("Monto_Origen").Name, adDouble
    RsGrid2.Fields.Append RsGrid1.Fields("Monto_Aceptado").Name, adDouble
    RsGrid2.Open

    Set gexGrid2.ADORecordset = RsGrid2
    ConfigurarGrid gexGrid2
End If

Exit Sub
Resume
hand:
ErrorHandler err, "CARGA_FACTURAS_PENDIENTES"
End Sub

Private Sub txtDes_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Origen_Letra", "Descripcion", "cn_origen_letras where  flg_protesto='N' and cod_origen_letra<>'N' and ", txtCod_Origen, txtDes_Origen, 2, Me)
End Sub

Private Sub txtDes_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 2, Me)
End Sub

Private Sub txtDes_TipVenta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_Tipo_Venta", "Descripcion", "Cn_Tipos_Venta where ", txtCod_TipVenta, txtDes_TipVenta, 2, Me)
  End If
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Sub ConfigurarGrid(mGridEx As GridEx)
  mGridEx.Columns("Numero").Width = 1305
  mGridEx.Columns("Fecha").Width = 945
  mGridEx.Columns("Moneda").Width = 720
  mGridEx.Columns("Imp_Total").Width = 1065
  mGridEx.Columns("Tipo_Cambio").Width = 1065
  mGridEx.Columns("Monto_Origen").Width = 1500
  mGridEx.Columns("Monto_Aceptado").Width = 1365
  mGridEx.Columns("Correlativo").Visible = False
  mGridEx.Columns("Monto_Origen1").Visible = False
  mGridEx.Columns("Monto_Origen").Format = "###,###.00"
  mGridEx.Columns("Monto_Aceptado").Format = "###,###.00"
  mGridEx.Columns("Imp_ToTal").Format = "###,###.00"
End Sub

Private Function CALCULA_MONTO_TOTAL(ByVal mRs As ADODB.Recordset) As String
Dim Monto As Double
Dim i As Integer
    Monto = 0
    mRs.MoveFirst
    For i = 1 To mRs.RecordCount
        Monto = Monto + mRs.Fields("Monto_Aceptado").Value
        mRs.MoveNext
    Next
CALCULA_MONTO_TOTAL = Format(Monto, "###,###.00")
End Function

Sub Resumen_Facturas()

Dim Monto As Double

Dim i As Integer

    Monto = 0
    txtFacturas = ""
    
    If gexGrid2.RowCount > 0 Then
    
      gexGrid2.MoveFirst
      For i = 1 To gexGrid2.RowCount
          Monto = Monto + gexGrid2.Value(gexGrid2.Columns("Monto_Aceptado").Index)
          txtFacturas = txtFacturas + " " + gexGrid2.Value(gexGrid2.Columns("Numero").Index)
          gexGrid2.MoveNext
      Next
      
    End If

txtTotFact = Format(Monto, "###,###.00")

End Sub

Private Function lfSalvar_Datos() As Boolean

On Error GoTo hand

Dim RS As ADODB.Recordset
Dim SQL As String
Dim i As Integer

gexLetra.Redraw = False
gexLetra.MoveFirst
intCod_Grupo = 0

For i = 1 To gexLetra.RowCount
  SQL = "Ventas_Up_Man_Letras '" & strOption & "','','" & gexLetra.Value(gexLetra.Columns("Nro").Index) & "','" & varCod_TipAnex & "','" _
        & varCod_anxo & "','" & gexLetra.Value(gexLetra.Columns("Fec_EmiDoc").Index) & "','" _
        & gexLetra.Value(gexLetra.Columns("Fec_VenDoc").Index) & "','" & TxtCod_Moneda & "','" & Des_Apos(TxtGlosa) & "','" _
        & vusu & "','" & txtFacturas & "','" & IIf(strCod_Anxo = "", "", txtCod_TipAne) & "','" & strCod_Anxo & "'," & intCod_Grupo & "," _
        & gexLetra.Value(gexLetra.Columns("Imp_Total").Index) & ",'" & txtTercero_CodTipAnexo & "','" & strTercero_Cod_Anxo & "','" _
        & IIf(txtCod_Origen.Visible, txtCod_Origen, "N") & "'"
        
  Set RS = CargarRecordSetDesconectado(SQL, cCONNECT)
  
  If Not (RS.BOF Or RS.EOF) Then
    intCod_Grupo = RS!Cod_Grupo_Letra
    oParent.xCod_Grupo = intCod_Grupo
    oParent.strNum_Corre_Let = RS!Num_Corre
  End If
  
  gexLetra.MoveNext
Next

gexGrid2.MoveFirst
For i = 1 To gexGrid2.RowCount
  SQL = "Ventas_Up_Man_Letras_Facturas '" & gexGrid2.Value(gexGrid2.Columns("Correlativo").Index) & "'," & gexGrid2.Value(gexGrid2.Columns("Monto_Origen").Index) & "," & intCod_Grupo
  Set RS = CargarRecordSetDesconectado(SQL, cCONNECT)
  gexGrid2.MoveNext
  strNum_Corre_Let_Renov = gexGrid2.Value(gexGrid2.Columns("Correlativo").Index)
  oParent.strNum_Corre_Let_Renov = strNum_Corre_Let_Renov
Next


gexLetra.Redraw = True

lfSalvar_Datos = True

Exit Function

hand:

gexLetra.Redraw = True

errores err.Number
Set RS = Nothing

ExecuteCommandSQL cCONNECT, "Ventas_Revierte_Letras_Planeadas " & intCod_Grupo

lfSalvar_Datos = False

End Function

Sub Actualiza_Facturas_Letras(ByVal intCod_Grupo As String, ByVal strNum_Corre As String, ByVal Importe As Double)

On Error GoTo hand

Dim RS As ADODB.Recordset
Dim SQL As String
Dim mRs As ADODB.Recordset

    SQL = "Ventas_Up_Man_Letras_Facturas '" & strNum_Corre & "'," & Importe & "," & intCod_Grupo
    Set mRs = GetRecordset(cCONNECT, SQL)
    If Not mRs Is Nothing Then
        Do While Not mRs.EOF
            mRs.MoveNext
        Loop
    End If
    mRs.Close
    Set mRs = Nothing

Exit Sub

hand:

Error err.Number
Set RS = Nothing
End Sub

Private Sub txtNom_Tercero_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 1, Me)
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtNumeroPendiente_Change()
  Call gexGrid1.Find(gexGrid1.Columns("Numero").Index, jgexContains, txtNumeroPendiente)
End Sub

Private Sub txtNumeroPendiente_KeyPress(KeyAscii As Integer)
If gexGrid1.RowCount > 0 And KeyAscii = 13 Then
  If Not gexGrid1.Find(gexGrid1.Columns("Numero").Index, jgexContains, txtNumeroPendiente) Or txtNumeroPendiente = "" Then
    MsgBox "Factura no se encuentra en la Lista", vbInformation, "AVISO"
    Exit Sub
  End If
  Call cmdNext_Click
  txtNumeroPendiente = ""
End If
End Sub

Private Sub txtNumeroPendiente_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
  gexGrid1.SetFocus
End If

End Sub

Private Sub txtNumeroxCancelar_Change()
  Call gexGrid2.Find(gexGrid2.Columns("Numero").Index, jgexContains, txtNumeroxCancelar)
End Sub

Private Sub txtNumeroxCancelar_KeyPress(KeyAscii As Integer)
If gexGrid2.RowCount > 0 And KeyAscii = 13 Then
  If Not gexGrid2.Find(gexGrid2.Columns("Numero").Index, jgexContains, txtNumeroxCancelar) Or txtNumeroxCancelar = "" Then
    MsgBox "Factura no se encuentra en la Lista", vbInformation, "AVISO"
    Exit Sub
  End If
  Call CmdBack_Click
  txtNumeroxCancelar = ""
End If
End Sub

Sub VALIDA_DETALLE_LETRA()
On Error GoTo ERR1
Dim i As Integer
Dim strSQL As String
        gexGrid2.Redraw = False
        gexGrid2.MoveFirst
        For i = 1 To gexGrid2.RowCount
            gexGrid2.Row = i

            gexGrid2.MoveNext
        Next
        gexGrid2.Row = 1
        gexGrid2.Redraw = True
        Exit Sub
ERR1:
    err.Raise err.Number, err.Source, err.Description
End Sub

Sub Calcula_Monto_Letra()

Dim dbTotFac As Double, i As Integer, dbMontoLetra As Double, Nro_Letra As Integer, VrBookMark, dbTotLetras As Double

dbTotFac = CDbl(txtTotFact)
Nro_Letra = gexLetra.RowCount
txtTotLetras = 0

If dbTotFac = 0 Then Exit Sub
If Nro_Letra = 0 Then Exit Sub

VrBookMark = gexLetra.Row

gexLetra.MoveFirst
dbMontoLetra = 0

For i = 1 To Nro_Letra
  
  If i = Nro_Letra Then
    dbMontoLetra = dbTotFac - dbTotLetras
  Else
    dbMontoLetra = Format(dbTotFac / Nro_Letra, "###,###.00")
  End If

  dbTotLetras = dbTotLetras + dbMontoLetra
  
  gexLetra.Value(gexLetra.Columns("Imp_Total").Index) = dbMontoLetra
  gexLetra.MoveNext
Next i
  
gexLetra.Row = VrBookMark
txtTotLetras = Format(dbTotLetras, "###,###.00")

End Sub

Sub Calcula_Total_Letra()

Dim dbTotFac As Double, i As Integer, dbMontoLetra As Double, Nro_Letra As Integer, VrBookMark, dbTotLetras As Double

dbTotFac = CDbl(txtTotFact)
Nro_Letra = gexLetra.RowCount
txtTotLetras = 0

If dbTotFac = 0 Then Exit Sub
If Nro_Letra = 0 Then Exit Sub

VrBookMark = gexLetra.Row

gexLetra.MoveFirst
dbMontoLetra = 0

For i = 1 To Nro_Letra
  
  dbTotLetras = dbTotLetras + gexLetra.Value(gexLetra.Columns("Imp_Total").Index)
  
  gexLetra.MoveNext
  
Next i
  
gexLetra.Row = VrBookMark
txtTotLetras = Format(dbTotLetras, "###,###.00")

End Sub


Private Sub txtRuc_Tercero_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Sub Busca_Opcion_Anexo_Tercero(strCampo1 As String, strCampo2 As String, StrTabla As String, txtCod As TextBox, txtDes As TextBox, opcion As Integer, frmME As Form)

On Error GoTo Fin

Dim rstAux As ADODB.Recordset, strSQL As String
    strSQL = "select Cod_Anxo as Cod,Des_Anexo as Nombre,Num_Ruc as Ruc from cn_anexoscontables where cod_tipanex = '" & StrTabla & "' and "
    
    'StrSql = "Select " & strCampo1 & " AS Cod," & strCampo2 & " as Descripcion from " & StrTabla

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    Select Case opcion
    Case 1: strSQL = strSQL & strCampo1 & " like '%" & txtCod & "%'"
    Case 2: strSQL = strSQL & strCampo2 & " like '%" & txtDes & "%'"
    End Select
    txtCod = ""
    txtDes = ""
    With frmBusqGeneral
        Set .oParent = frmME
        .SQuery = strSQL
        .CARGAR_DATOS
        
        codigo = ""
        .DGridLista.Columns("Cod").Visible = False
        .DGridLista.Columns("Nombre").Width = 4575
        .DGridLista.Columns("RUC").Width = 1695
        Set rstAux = .DGridLista.ADORecordset
        
        If rstAux.RecordCount > 1 Then
          .Show vbModal
        Else
          frmME.codigo = ".."
        End If
        If frmME.codigo <> "" And rstAux.RecordCount > 0 Then
            frmME.strTercero_Cod_Anxo = Trim(rstAux!Cod)
            txtDes = Trim(rstAux!Nombre)
            txtCod = Trim(rstAux!Ruc)
            Select Case opcion
            Case 1: SendKeys "{TAB}"
            Case 2: SendKeys "{TAB}"
            End Select
        Else
            SendKeys "{TAB}"
        End If
        
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
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & opcion & ")"
End Sub

Private Sub txtTercero_CodTipAnexo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtTercero_CodTipAnexo, txtTercero_Des_TipAnex, 1, Me)
End Sub

Private Sub txtTercero_NomAnexo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion_Anexo_Tercero("Num_Ruc", "Des_Anexo", txtTercero_CodTipAnexo, txtTercero_NumRuc, txtTercero_NomAnexo, 2, Me)
End Sub

Private Sub txtTercero_NumRuc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion_Anexo_Tercero("Num_Ruc", "Des_Anexo", txtTercero_CodTipAnexo, txtTercero_NumRuc, txtTercero_NomAnexo, 1, Me)
End Sub
