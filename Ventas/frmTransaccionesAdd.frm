VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "numbox.ocx"
Begin VB.Form frmTransaccionesAdd 
   Caption         =   "Adicion Documento de Cobranza"
   ClientHeight    =   11115
   ClientLeft      =   2445
   ClientTop       =   1485
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   11160
   Begin VB.Frame frFacturas 
      ForeColor       =   &H8000000D&
      Height          =   7095
      Left            =   120
      TabIndex        =   26
      Top             =   7320
      Visible         =   0   'False
      Width           =   9705
      Begin VB.CheckBox chkmuestra_saldo_total 
         Caption         =   "Muestra Saldo Total"
         Height          =   255
         Left            =   7680
         TabIndex        =   67
         Top             =   240
         Width           =   1815
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
         Left            =   7905
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "0.00"
         Top             =   6645
         Width           =   1455
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
         Left            =   7905
         Locked          =   -1  'True
         TabIndex        =   54
         Text            =   "0.00"
         Top             =   3480
         Width           =   1455
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
         Left            =   2910
         TabIndex        =   30
         Top             =   3390
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
         Left            =   3630
         TabIndex        =   31
         Top             =   3390
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
         Left            =   4710
         TabIndex        =   32
         Top             =   3390
         Width           =   675
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
         Left            =   5430
         TabIndex        =   34
         Top             =   3390
         Width           =   675
      End
      Begin VB.TextBox txtNumeroPendiente 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3750
         TabIndex        =   28
         Text            =   "0"
         Top             =   240
         Width           =   1245
      End
      Begin VB.TextBox txtNumeroxCancelar 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   960
         TabIndex        =   29
         Text            =   "0"
         Top             =   3480
         Width           =   1245
      End
      Begin GridEX20.GridEX gexGrid2 
         Height          =   2595
         Left            =   195
         TabIndex        =   33
         Top             =   3960
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   4577
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
         Column(1)       =   "frmTransaccionesAdd.frx":0000
         Column(2)       =   "frmTransaccionesAdd.frx":00F4
         Column(3)       =   "frmTransaccionesAdd.frx":01E0
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmTransaccionesAdd.frx":02AC
         FormatStyle(2)  =   "frmTransaccionesAdd.frx":03E4
         FormatStyle(3)  =   "frmTransaccionesAdd.frx":0494
         FormatStyle(4)  =   "frmTransaccionesAdd.frx":0548
         FormatStyle(5)  =   "frmTransaccionesAdd.frx":0620
         FormatStyle(6)  =   "frmTransaccionesAdd.frx":06D8
         ImageCount      =   0
         PrinterProperties=   "frmTransaccionesAdd.frx":07B8
      End
      Begin GridEX20.GridEX gexGrid1 
         Height          =   2595
         Left            =   195
         TabIndex        =   56
         Top             =   660
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   4577
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
         Column(1)       =   "frmTransaccionesAdd.frx":0990
         Column(2)       =   "frmTransaccionesAdd.frx":0A84
         Column(3)       =   "frmTransaccionesAdd.frx":0B70
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmTransaccionesAdd.frx":0C3C
         FormatStyle(2)  =   "frmTransaccionesAdd.frx":0D74
         FormatStyle(3)  =   "frmTransaccionesAdd.frx":0E24
         FormatStyle(4)  =   "frmTransaccionesAdd.frx":0ED8
         FormatStyle(5)  =   "frmTransaccionesAdd.frx":0FB0
         FormatStyle(6)  =   "frmTransaccionesAdd.frx":1068
         ImageCount      =   0
         PrinterProperties=   "frmTransaccionesAdd.frx":1148
      End
      Begin FunctionsButtons.FunctButt fncBuscar 
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   165
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   661
         Custom          =   $"frmTransaccionesAdd.frx":1320
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   350
         ControlSeparator=   110
      End
      Begin MSComCtl2.DTPicker DTFecha 
         Height          =   285
         Left            =   5520
         TabIndex        =   68
         Top             =   240
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "MM / yyyy"
         Format          =   71630849
         CurrentDate     =   37987
      End
      Begin VB.Label Label19 
         Caption         =   "Dia"
         Height          =   255
         Left            =   5160
         TabIndex        =   69
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label17 
         Caption         =   "Total Pendiente :"
         Height          =   285
         Left            =   6585
         TabIndex        =   38
         Top             =   3495
         Width           =   1305
      End
      Begin VB.Label Label13 
         Caption         =   "Total a Cancelar :"
         Height          =   255
         Left            =   6225
         TabIndex        =   37
         Top             =   6660
         Width           =   1395
      End
      Begin VB.Label Label14 
         Caption         =   "Numero :"
         Height          =   255
         Left            =   3000
         TabIndex        =   36
         Top             =   270
         Width           =   675
      End
      Begin VB.Label Label15 
         Caption         =   "Numero :"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   3510
         Width           =   675
      End
   End
   Begin VB.Frame frTransacciones 
      BorderStyle     =   0  'None
      Height          =   7455
      Left            =   0
      TabIndex        =   39
      Top             =   -240
      Width           =   9735
      Begin VB.Frame Frame3 
         Height          =   3255
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   9615
         Begin VB.TextBox txtOtro_Tipo_Cambio 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8460
            TabIndex        =   64
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   600
            Width           =   825
         End
         Begin VB.TextBox txtTipo_Cambio 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4395
            TabIndex        =   63
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   600
            Width           =   825
         End
         Begin VB.TextBox txtCod_Origen 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6120
            MaxLength       =   1
            TabIndex        =   2
            Top             =   255
            Width           =   720
         End
         Begin VB.TextBox txtCuenta_Cod 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   6120
            MaxLength       =   11
            TabIndex        =   9
            Top             =   1320
            Width           =   720
         End
         Begin VB.CommandButton cmdObtieneComp 
            Caption         =   "....."
            Height          =   195
            Left            =   2040
            TabIndex        =   62
            ToolTipText     =   "Recupera el Ultimo Nro de Comprobante"
            Top             =   2175
            Width           =   375
         End
         Begin VB.TextBox txtDes_Moneda 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6900
            TabIndex        =   14
            Top             =   1710
            Width           =   2385
         End
         Begin VB.TextBox txtCuenta_Des 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   6915
            MaxLength       =   30
            TabIndex        =   10
            Top             =   1320
            Width           =   2370
         End
         Begin VB.TextBox txt_ImpTotal_Doc_Cobra 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2085
            MaxLength       =   15
            TabIndex        =   17
            Text            =   "0"
            Top             =   2490
            Width           =   1200
         End
         Begin VB.TextBox txtDes_Origen 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6915
            TabIndex        =   3
            Top             =   255
            Width           =   2370
         End
         Begin VB.CheckBox chkDiferido 
            Alignment       =   1  'Right Justify
            Caption         =   "&Diferido"
            Height          =   255
            Left            =   5400
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   2505
            Width           =   1455
         End
         Begin VB.TextBox txtDes_TipAne 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            TabIndex        =   5
            Top             =   1005
            Width           =   5265
         End
         Begin VB.TextBox TxtObservacion 
            Height          =   285
            Left            =   1320
            TabIndex        =   20
            Top             =   2820
            Width           =   7965
         End
         Begin VB.TextBox txtNum_DocCobra 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6120
            MaxLength       =   8
            TabIndex        =   16
            Top             =   2085
            Width           =   3150
         End
         Begin VB.TextBox txtSer_DocCobra 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            MaxLength       =   3
            TabIndex        =   15
            Top             =   2085
            Width           =   735
         End
         Begin VB.TextBox TxtDes_Banco 
            Height          =   285
            Left            =   2085
            TabIndex        =   8
            Top             =   1350
            Width           =   2415
         End
         Begin VB.TextBox TxtCod_Banco 
            Height          =   285
            Left            =   1320
            TabIndex        =   7
            Top             =   1350
            Width           =   735
         End
         Begin VB.TextBox txtCod_TipDocCobra 
            Height          =   285
            Left            =   1320
            MaxLength       =   3
            TabIndex        =   11
            Top             =   1725
            Width           =   735
         End
         Begin VB.TextBox txtDes_DocCobra 
            Height          =   285
            Left            =   2085
            TabIndex        =   12
            Top             =   1725
            Width           =   2415
         End
         Begin VB.TextBox txtCod_Moneda 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6120
            MaxLength       =   4
            TabIndex        =   13
            Top             =   1710
            Width           =   720
         End
         Begin VB.TextBox txtCod_TipAne 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6600
            MaxLength       =   4
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   "C"
            Top             =   1005
            Width           =   360
         End
         Begin VB.TextBox txtNum_Ruc 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   7740
            MaxLength       =   11
            TabIndex        =   6
            Top             =   1005
            Width           =   1545
         End
         Begin VB.TextBox txtCod_TipCobra 
            Height          =   285
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   0
            Top             =   255
            Width           =   735
         End
         Begin VB.TextBox txtDes_TipCobra 
            Height          =   285
            Left            =   2085
            TabIndex        =   1
            Top             =   255
            Width           =   3135
         End
         Begin VB.TextBox txtDes_TipAnex 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   4680
            MaxLength       =   4
            TabIndex        =   44
            Text            =   "C"
            Top             =   1005
            Visible         =   0   'False
            Width           =   360
         End
         Begin NumBoxProject.NumBox txtFecha 
            Height          =   285
            Left            =   1320
            TabIndex        =   4
            Top             =   615
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   503
            TypeVal         =   3
            Mask            =   "99/99/9999"
            Formato         =   "dd/MM/yyyy"
            AllowedMask     =   -1
            MaskLen         =   10
            Aling           =   2
            Text            =   ""
            CanEmpty        =   -1
            ShowError       =   0
            Locked          =   0   'False
            Enabled         =   -1  'True
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DecimalNumber   =   0
         End
         Begin NumBoxProject.NumBox txtFec_Diferido 
            Height          =   285
            Left            =   7905
            TabIndex        =   19
            Top             =   2490
            Visible         =   0   'False
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   503
            TypeVal         =   3
            Mask            =   "99/99/9999"
            Formato         =   "dd/MM/yyyy"
            AllowedMask     =   -1
            MaskLen         =   10
            Aling           =   2
            Text            =   ""
            CanEmpty        =   -1
            ShowError       =   0
            Locked          =   0   'False
            Enabled         =   -1  'True
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DecimalNumber   =   0
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Otro Tipo Cambio :"
            Height          =   195
            Left            =   6915
            TabIndex        =   66
            Top             =   645
            Width           =   1320
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio :"
            Height          =   195
            Left            =   3240
            TabIndex        =   65
            Top             =   645
            Width           =   975
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Numero :"
            Height          =   195
            Left            =   5400
            TabIndex        =   61
            Top             =   2130
            Width           =   645
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta :"
            Height          =   195
            Left            =   5400
            TabIndex        =   60
            Top             =   1335
            Width           =   600
         End
         Begin VB.Label lbDiferido 
            AutoSize        =   -1  'True
            Caption         =   "Fecha :"
            Height          =   195
            Left            =   7200
            TabIndex        =   59
            Top             =   2535
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.Label Label4 
            Caption         =   "Origen :"
            Height          =   255
            Left            =   5400
            TabIndex        =   58
            Top             =   270
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Imp Total Doc Cobranza:"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   2535
            Width           =   1770
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Observacion :"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   2805
            Width           =   990
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Serie :"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   2130
            Width           =   450
         End
         Begin VB.Label Label3 
            Caption         =   "Banco:"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   1380
            Width           =   645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tip Doc Cobra:"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   1770
            Width           =   1080
         End
         Begin VB.Label Label11 
            Caption         =   "Moneda :"
            Height          =   255
            Left            =   5400
            TabIndex        =   49
            Top             =   1740
            Width           =   735
         End
         Begin VB.Label Label28 
            Caption         =   "R.U.C."
            Height          =   255
            Left            =   7080
            TabIndex        =   48
            Top             =   1020
            Width           =   495
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Cliente :"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   1050
            Width           =   570
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cobranza:"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   300
            Width           =   1080
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fecha :"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   660
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   120
         TabIndex        =   40
         Top             =   3480
         Width           =   9615
         Begin VB.TextBox txtTotDoc 
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
            Left            =   8175
            Locked          =   -1  'True
            TabIndex        =   41
            Text            =   "0.00"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtDocumentos 
            Enabled         =   0   'False
            Height          =   585
            Left            =   1350
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   240
            Width           =   6810
         End
         Begin VB.CommandButton cmdDocumentos 
            Caption         =   "&Documentos"
            Height          =   345
            Left            =   8160
            TabIndex        =   21
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   "Documentos :"
            Height          =   225
            Left            =   120
            TabIndex        =   42
            Top             =   300
            Width           =   1005
         End
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   7200
         TabIndex        =   24
         Top             =   6960
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmTransaccionesAdd.frx":13AD
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   2460
         Left            =   120
         TabIndex        =   23
         Top             =   4440
         Visible         =   0   'False
         Width           =   11400
         _ExtentX        =   20108
         _ExtentY        =   4339
         Version         =   "2.0"
         HoldSortSettings=   -1  'True
         DefaultGroupMode=   1
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "frmTransaccionesAdd.frx":1440
         Column(2)       =   "frmTransaccionesAdd.frx":1508
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmTransaccionesAdd.frx":15AC
         FormatStyle(2)  =   "frmTransaccionesAdd.frx":16E4
         FormatStyle(3)  =   "frmTransaccionesAdd.frx":1794
         FormatStyle(4)  =   "frmTransaccionesAdd.frx":1848
         FormatStyle(5)  =   "frmTransaccionesAdd.frx":1920
         FormatStyle(6)  =   "frmTransaccionesAdd.frx":19D8
         FormatStyle(7)  =   "frmTransaccionesAdd.frx":1AB8
         FormatStyle(8)  =   "frmTransaccionesAdd.frx":1B64
         ImageCount      =   0
         PrinterProperties=   "frmTransaccionesAdd.frx":1C14
      End
   End
End
Attribute VB_Name = "frmTransaccionesAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public codigo As String, Descripcion As String, strOption As String, strCod_Anxo As String, lfSalvar As Boolean

Dim RsGrid1 As ADODB.Recordset
Dim RsGrid2 As ADODB.Recordset
Dim intTransaccion As Long, vrTotalTransaccion As Double
Dim strSQL As String, intCancel As Integer

Sub Totalizar_Transaccio()

End Sub

Private Sub chkDiferido_Click()

If chkDiferido Then
  txtFec_Diferido.Visible = True
  lbDiferido.Visible = True
Else
  txtFec_Diferido.Visible = False
  lbDiferido.Visible = False
  txtFec_Diferido.Text = ""
End If

End Sub

Private Sub CmdBack_Click()
    If gexGrid2.RowCount = 0 Then Exit Sub
    RsGrid1.AddNew
    RsGrid1.Fields("Correlativo").Value = gexGrid2.Value(gexGrid1.Columns("Correlativo").Index)
    RsGrid1.Fields("Numero").Value = gexGrid2.Value(gexGrid1.Columns("Numero").Index)
    RsGrid1.Fields("Fecha").Value = gexGrid2.Value(gexGrid1.Columns("Fecha").Index)
    RsGrid1.Fields("Tipo_Cambio").Value = gexGrid2.Value(gexGrid1.Columns("Tipo_Cambio").Index)
    RsGrid1.Fields("Moneda").Value = gexGrid2.Value(gexGrid1.Columns("Moneda").Index)
    RsGrid1.Fields("Monto_Origen1").Value = gexGrid2.Value(gexGrid1.Columns("Monto_Origen1").Index)
    RsGrid1.Fields("Monto_Origen").Value = gexGrid2.Value(gexGrid1.Columns("Monto_Origen").Index)
    RsGrid1.Fields("Monto_Aceptado").Value = gexGrid2.Value(gexGrid1.Columns("Monto_Aceptado").Index)
    RsGrid1.Fields("Cod_Cobranza").Value = gexGrid2.Value(gexGrid1.Columns("Cod_Cobranza").Index)
    RsGrid1.Fields("Debe_Haber").Value = gexGrid2.Value(gexGrid1.Columns("Debe_Haber").Index)
    RsGrid1.Fields("Observacion").Value = gexGrid2.Value(gexGrid1.Columns("Observacion").Index)
    RsGrid1.Fields("Tran_TipMonDoc").Value = gexGrid2.Value(gexGrid1.Columns("Tran_TipMonDoc").Index)
    RsGrid1.Fields("Doc_TipMonDoc").Value = gexGrid2.Value(gexGrid1.Columns("Doc_TipMonDoc").Index)
    RsGrid1.Fields("Otro_Tip_Cambio").Value = gexGrid2.Value(gexGrid1.Columns("Otro_Tip_Cambio").Index)

    
    RsGrid1.Update

    RsGrid2.MoveFirst
    Call BuscaCampo(RsGrid2, "Correlativo", gexGrid2.Value(gexGrid2.Columns("Correlativo").Index))

    RsGrid2.Delete
    gexGrid1.ClearFields
    Set gexGrid1.ADORecordset = RsGrid1
    ConfigurarGrid gexGrid1
    gexGrid2.ClearFields
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

Private Sub cmdDocumentos_Click()

  If DevuelveCampo("select count(*) from tg_moneda where cod_moneda='" & Trim(txtCod_Moneda.Text) & "'", cCONNECT) = 0 Then
    MsgBox "Seleccione una Moneda Valida", vbInformation, Me.Caption
    txtCod_Moneda.SetFocus
    Exit Sub
  End If
  
  If strCod_Anxo = "" Then
    MsgBox "Seleccione una Cliente", vbInformation, Me.Caption
    txtNum_Ruc.SetFocus
    Exit Sub
  End If
  
  If Me.WindowState <> 2 Then Me.Height = 7590
  
  frTransacciones.Visible = False
  frFacturas.Top = 0
  frFacturas.Visible = True
  
  fncBuscar.SetFocus
 
End Sub

Private Sub cmdNext_Click()

Dim Valor As Double, varCorrelativo As String

If gexGrid1.RowCount = 0 Then Exit Sub
    
If gexGrid1.EditMode = jgexEditModeOn Then
  MsgBox "Salga del Modo de Edicion de la Grilla" & vbCr & "Haga Click en la columna Numero", vbInformation, "IMPORTANTE"
  Exit Sub
End If
    
Valor = txt_ImpTotal_Doc_Cobra.Text - (gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) + CDbl(TxtMonto2))
If Valor < -10 And txt_ImpTotal_Doc_Cobra.Text <> 0 Then
   If MsgBox("Con este Documento el importe excederia en " & Valor & "  al importe del Documento de Cobranza de " & txt_ImpTotal_Doc_Cobra.Text, vbYesNo, "AVISO") = vbNo Then Exit Sub
End If

RsGrid2.AddNew
RsGrid2.Fields("Correlativo").Value = gexGrid1.Value(gexGrid1.Columns("Correlativo").Index)
RsGrid2.Fields("Numero").Value = gexGrid1.Value(gexGrid1.Columns("Numero").Index)
RsGrid2.Fields("Fecha").Value = gexGrid1.Value(gexGrid1.Columns("Fecha").Index)
RsGrid2.Fields("Moneda").Value = gexGrid1.Value(gexGrid1.Columns("Moneda").Index)
RsGrid2.Fields("Tipo_Cambio").Value = gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index)
RsGrid2.Fields("Monto_Origen1").Value = gexGrid1.Value(gexGrid1.Columns("Monto_Origen1").Index)
RsGrid2.Fields("Monto_Origen").Value = gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index)
RsGrid2.Fields("Monto_Aceptado").Value = gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index)
RsGrid2.Fields("Cod_Cobranza").Value = gexGrid1.Value(gexGrid1.Columns("Cod_Cobranza").Index)
RsGrid2.Fields("Debe_Haber").Value = gexGrid1.Value(gexGrid1.Columns("Debe_Haber").Index)
RsGrid2.Fields("Observacion").Value = gexGrid1.Value(gexGrid1.Columns("Observacion").Index)
RsGrid2.Fields("Tran_TipMonDoc").Value = gexGrid1.Value(gexGrid1.Columns("Tran_TipMonDoc").Index)
RsGrid2.Fields("Doc_TipMonDoc").Value = gexGrid1.Value(gexGrid1.Columns("Doc_TipMonDoc").Index)
RsGrid2.Fields("Otro_Tip_Cambio").Value = gexGrid1.Value(gexGrid1.Columns("Otro_Tip_Cambio").Index)

RsGrid2.Update

RsGrid1.MoveFirst
Call BuscaCampo(RsGrid1, "Correlativo", gexGrid1.Value(gexGrid1.Columns("Correlativo").Index))

RsGrid1.Delete
gexGrid1.ClearFields
Set gexGrid1.ADORecordset = RsGrid1
ConfigurarGrid gexGrid1
gexGrid2.ClearFields
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

Private Sub cmdObtieneComp_Click()

Dim rs As Object
Set rs = CreateObject("ADODB.Recordset")

Set rs = CargarRecordSetDesconectado("Select Cod_TipDoc,Ser_Docum,Num_Docum From Tg_Bancos_Cuentas Where Cod_Banco = '" & TxtCod_Banco & "' and Sec_Cuenta_Banco = '" & txtCuenta_Cod & "' and Fec_Transaccion_Cobranza = '" & txtFecha.Text & "' and Cod_Usuario = '" & vusu & "'", cCONNECT)

If Not (rs.BOF And rs.EOF) Then

  txtCod_TipDocCobra = rs!Cod_TipDoc
  txtSer_DocCobra = rs!Ser_Docum
  txtNum_DocCobra = rs!Num_Docum
  
  Set rs = Nothing
  
End If

End Sub

Private Sub fncBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
  Case "BUSCAR"
    CARGA_FACTURAS_PENDIENTES
  Case "CERRAR"
    frFacturas.Visible = False
    frTransacciones.Visible = True
    Resumen_Facturas
    Cambio_Apariencia
    If txtCod_TipCobra = "013" Then txt_ImpTotal_Doc_Cobra.Text = CDbl(txtTotDoc.Text)
    cmdDocumentos.SetFocus
End Select
End Sub
Sub Calcula_Monto_Letra()

Dim dbTotFac As Double, i As Integer, dbMontoLetra As Double, Nro_Letra As Integer, VrBookMark, dbTotLetras As Double

dbTotFac = CDbl(txtTotFact)
Nro_Letra = gexLetra.RowCount
txtTotLetras = 0

If dbTotFac = 0 Then Exit Sub
If Nro_Letra = 0 Then Exit Sub

gexLetra.MoveFirst
dbMontoLetra = 0

VrBookMark = gexLetra.Row

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

Private Sub Form_Load()
  txtFecha.Text = Date
  
  Set GridEX1.ADORecordset = CargarRecordSetDesconectado("Ventas_Muestra_Conceptos_Cobranzas", cCONNECT)
  
  GridEX1.Columns("Cod").Width = 900
  GridEX1.Columns("Descripcion").Width = 2475
  GridEX1.Columns("flg").Width = 345
  GridEX1.Columns("Imp_Debe").Width = 870
  GridEX1.Columns("Imp_Haber").Width = 915
  GridEX1.Columns("SEL").ColumnType = jgexCheckBox
  GridEX1.Columns("SEL").Visible = True
  GridEX1.Columns("SEL").EditType = jgexEditCheckBox
  GridEX1.Columns("SEL").Width = 660
  
  intTransaccion = 0
  
  Cambio_Apariencia
  DTFecha = Date
  DTFecha = Null
  intCancel = 1
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = intCancel
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim sMessage As Long
Dim strSQL As String

On Error GoTo dprError

Select Case ActionName
  Case "GRABAR"
  
    If MsgBox("Esta seguro de Generar una nueva Transaccion ", vbYesNo, "IMPORTANTE") = vbYes Then
      If lfSalvar_Datos Then
        lfSalvar = True
        With frmTransacciones
          .inpFec_Emi.Text = txtFecha.Text
          .txtCod_Origen = txtCod_Origen
          .txtDes_Origen = txtDes_Origen
        End With
        intCancel = 0
        Unload Me
      End If
    End If
    
  Case "CANCELAR"
    If MsgBox("Esta seguro de Cancelar esta Transaccion ", vbYesNo, "IMPORTANTE") = vbYes Then
      lfSalvar = False
      intCancel = 0
      Unload Me
    End If
End Select

Exit Sub

dprError:

errores err.Number

End Sub

Private Function lfSalvar_Datos() As Boolean

On Error GoTo hand

Dim SQL As String

If lfGenera_Cuadre Then
  SQL = "CN_VENTAS_GENERAR_NUEVA_TRANSACCION_COBRANZAS " & intTransaccion
  Call ExecuteCommandSQL(cCONNECT, SQL)
  lfSalvar_Datos = True
Else
  lfSalvar_Datos = False
End If

Exit Function

hand:

errores err.Number

lfSalvar_Datos = False

End Function

Private Function lfGenera_Cuadre() As Boolean

On Error GoTo hand

Dim rs As ADODB.Recordset
Dim SQL As String
Dim i As Integer
Dim strCod_Cobranza As String, strFlg_Debe_Haber As String

GridEX1.Redraw = False



SQL = "TM_VENTAS_TRANSACCIONES_COBRANZAS_INSERT '" & strOption & "'," & intTransaccion & ",'" & txtFecha.Text & "'," & 0 & ",'" _
      & txtCod_TipCobra & "','" & txtCod_TipAne & "','" & strCod_Anxo & "','" & TxtCod_Banco & "','" & txtCuenta_Cod & "','" _
      & txtCod_TipDocCobra & "','" & txtSer_DocCobra & "','" & txtNum_DocCobra & "','" & txtCod_Moneda & "','" _
      & Des_Apos(TxtObservacion) & "','" & vusu & "','" & ComputerName & "','" & txtCod_Origen & "','" & IIf(chkDiferido, "S", "N") & "'," _
      & IIf(txtFec_Diferido.Text <> "", "'" & txtFec_Diferido.Text & "'", "Null")
        
Set rs = CargarRecordSetDesconectado(SQL, cCONNECT)

If Not (rs.BOF Or rs.EOF) Then
  intTransaccion = rs!Num_Transaccion
  strCod_Cobranza = rs!Cod_Concepto_Cobranza
  strFlg_Debe_Haber = rs!Flg_Debe_Haber
End If

strOption = "I"

If txt_ImpTotal_Doc_Cobra.Text <> 0 Then
  SQL = "TM_Ventas_Transacciones_Cobranzas_DETALLE_MAN '" & strOption & "'," & intTransaccion & ",'" & txtFecha.Text & "'," _
        & 0 & "," & 0 & ",'" & strCod_Cobranza & "','" & strFlg_Debe_Haber & "',''," & txt_ImpTotal_Doc_Cobra.Text & ",''," & txtTipo_Cambio & ",'S'," & txtOtro_Tipo_Cambio
  Call ExecuteCommandSQL(cCONNECT, SQL)
End If

gexGrid2.MoveFirst
For i = 1 To gexGrid2.RowCount
  SQL = "TM_Ventas_Transacciones_Cobranzas_DETALLE_MAN '" & strOption & "'," & intTransaccion & ",'" & txtFecha.Text & "',0,0,'" _
        & gexGrid2.Value(gexGrid2.Columns("Cod_Cobranza").Index) & "','" & gexGrid2.Value(gexGrid2.Columns("Debe_Haber").Index) & "','" _
        & gexGrid2.Value(gexGrid2.Columns("Correlativo").Index) & "','" & gexGrid2.Value(gexGrid2.Columns("Monto_Origen").Index) & "','" _
        & gexGrid2.Value(gexGrid2.Columns("Observacion").Index) & "'," & gexGrid2.Value(gexGrid2.Columns("Tipo_Cambio").Index) & ",'S'," _
        & gexGrid2.Value(gexGrid2.Columns("Otro_Tip_Cambio").Index)
        
  Call ExecuteCommandSQL(cCONNECT, SQL)
  gexGrid2.MoveNext
Next

GridEX1.MoveFirst
For i = 1 To GridEX1.RowCount
  If GridEX1.Value(GridEX1.Columns("Sel").Index) Then
    SQL = "TM_Ventas_Transacciones_Cobranzas_DETALLE_MAN '" & strOption & "'," & intTransaccion & ",'" & txtFecha.Text & "',0,0,'" _
          & GridEX1.Value(GridEX1.Columns("Cod").Index) & "','" & GridEX1.Value(GridEX1.Columns("flg").Index) & "','" & "" & "','" _
          & IIf(GridEX1.Value(GridEX1.Columns("flg").Index) = "D", GridEX1.Value(GridEX1.Columns("Imp_Debe").Index), GridEX1.Value(GridEX1.Columns("Imp_Haber").Index)) & "','" _
          & GridEX1.Value(GridEX1.Columns("Observacion").Index) & "'," & txtTipo_Cambio & ",'S'," & txtOtro_Tipo_Cambio
    Call ExecuteCommandSQL(cCONNECT, SQL)
  End If
  GridEX1.MoveNext
Next

strOption = "U"

GridEX1.Redraw = True

With frmTransaccionesAddCuadre
  .strSQL = "TM_VENTAS_MUESTRA_CUADRE_COBRANZAS " & intTransaccion
  .intNum_Transaccion = intTransaccion
  .strCod_Anexo = strCod_Anxo
  .strCod_TipAnexo = txtCod_TipAne
  .strCod_Moneda = txtCod_Moneda
  .dFecha = txtFecha.Text
  .CARGA_GRID
  .Caption = "Detalle de Transaccion del Cliente " & txtDes_TipAne
  .Show 1
  lfGenera_Cuadre = .lfAceptar
End With

Exit Function
Resume
hand:

GridEX1.Redraw = True

errores err.Number

Set rs = Nothing

lfGenera_Cuadre = False

End Function


Private Sub gexGrid1_AfterColEdit(ByVal ColIndex As Integer)

Dim dbImporte As Double

  Select Case ColIndex
  
  Case Is = gexGrid1.Columns("Monto_Aceptado").Index
     
    If CDbl(gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index)) <= 0 Then
      MsgBox "El Monto del documento debe ser Mayor a Cero", vbInformation, "AVISO"
      gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) = gexGrid1.Value(gexGrid1.Columns("Monto_Origen1").Index)
      gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) = DevuelveCampo("select dbo.Convierte_Importe_Moneda_Destino('" & gexGrid1.Value(gexGrid1.Columns("Moneda").Index) & "'," & gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) & "," & gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index) & ",'" & txtCod_Moneda & "','" & txtFecha.Text & "'," & gexGrid1.Value(gexGrid1.Columns("Otro_Tip_Cambio").Index) & ")", cCONNECT)
      Exit Sub
    End If
    
    gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) = DevuelveCampo("select dbo.Convierte_Importe_Moneda_Destino('" & txtCod_Moneda & "'," & Val(gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index)) & "," & gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index) & ",'" & gexGrid1.Value(gexGrid1.Columns("Moneda").Index) & "','" & txtFecha.Text & "'," & gexGrid1.Value(gexGrid1.Columns("Otro_Tip_Cambio").Index) & ")", cCONNECT)
    
   If CDbl(gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index)) > CDbl(gexGrid1.Value(gexGrid1.Columns("Monto_Origen1").Index)) Then
      MsgBox "No se puede ingresar un monto Mayor al pendiente del Documento", vbInformation, "AVISO"
      gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) = gexGrid1.Value(gexGrid1.Columns("Monto_Origen1").Index)
      gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) = DevuelveCampo("select dbo.Convierte_Importe_Moneda_Destino('" & gexGrid1.Value(gexGrid1.Columns("Moneda").Index) & "'," & gexGrid1.Value(gexGrid1.Columns("Monto_Origen1").Index) & "," & gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index) & ",'" & txtCod_Moneda & "','" & txtFecha.Text & "'," & gexGrid1.Value(gexGrid1.Columns("Otro_Tip_Cambio").Index) & ")", cCONNECT)
   End If
   
    If Abs(gexGrid1.Value(gexGrid1.Columns("Monto_Origen1").Index) - gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index)) = 0.01 Then
        gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) = gexGrid1.Value(gexGrid1.Columns("Monto_Origen1").Index)
        gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) = DevuelveCampo("select dbo.Convierte_Importe_Moneda_Destino('" & gexGrid1.Value(gexGrid1.Columns("Moneda").Index) & "'," & gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) & "," & gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index) & ",'" & txtCod_Moneda & "','" & txtFecha.Text & "'," & gexGrid1.Value(gexGrid1.Columns("Otro_Tip_Cambio").Index) & ")", cCONNECT)
    End If
    
'    SendKeys "{ENTER}"
    
  Case Is = gexGrid1.Columns("Tipo_Cambio").Index
  
    If Trim(gexGrid1.Value(gexGrid1.Columns("Doc_TipMonDoc").Index)) <> "" Then
    
      If txtCod_Moneda <> gexGrid1.Value(gexGrid1.Columns("Moneda").Index) Then
        
        gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) = IIf(txtCod_Moneda = "SOL", gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) * gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index), Format(gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) / gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index), "###,###.00"))
      Else
        gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) = gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index)
      End If
      
    End If
    
 '   SendKeys "{ENTER}"
    
  Case Is = gexGrid1.Columns("Otro_Tip_Cambio").Index
    gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) = Format(gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) * gexGrid1.Value(gexGrid1.Columns("Otro_Tip_Cambio").Index), "###,###.000000")
    TxtMonto1 = CALCULA_MONTO_TOTAL(RsGrid1)
    
    SendKeys "{ENTER}"
    
  End Select
  
End Sub

Private Sub gexGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
  Select Case ColIndex
  Case Is = gexGrid1.Columns("Monto_Aceptado").Index
    Cancel = False
  Case Is = gexGrid1.Columns("Tipo_Cambio").Index
    Cancel = False
  Case Is = gexGrid1.Columns("Observacion").Index
    Cancel = False
  Case Is = gexGrid1.Columns("Otro_Tip_Cambio").Index
    'If Trim(gexGrid1.Value(gexGrid1.Columns("Doc_TipMonDoc").Index)) = "" Then Cancel = False Else Cancel = True
    Cancel = False
  Case Else
    Cancel = True
  End Select
End Sub

Private Sub gexGrid1_BeforeColUpdate(ByVal Row As Long, ByVal ColIndex As Integer, ByVal OldValue As String, ByVal Cancel As GridEX20.JSRetBoolean)
    'If gexGrid1.Columns("Monto_ORIGEN").Index = ColIndex Then
    '    If OldValue = gexGrid1.Value(gexGrid1.Columns("Monto_ORIGEN").Index) Then
    '        gexGrid1.Value(gexGrid1.Columns("Monto_ORIGEN").Index) = OldValue
    '    End If
    'End If
End Sub

Private Sub GridEX1_AfterColEdit(ByVal ColIndex As Integer)
  Select Case ColIndex
  Case Is = GridEX1.Columns("Sel").Index
    If Not GridEX1.Value(GridEX1.Columns("Sel").Index) Then
      GridEX1.Value(GridEX1.Columns("Imp_Debe").Index) = 0
      GridEX1.Value(GridEX1.Columns("Imp_Haber").Index) = 0
      GridEX1.Value(GridEX1.Columns("Observacion").Index) = ""
    End If
  End Select
End Sub

Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
  Select Case ColIndex
  Case Is = GridEX1.Columns("Sel").Index
    Cancel = False
  Case Is = GridEX1.Columns("Imp_Debe").Index
    Cancel = IIf(GridEX1.Value(GridEX1.Columns("Flg").Index) = "D" And GridEX1.Value(GridEX1.Columns("sel").Index), False, True)
  Case Is = GridEX1.Columns("Imp_Haber").Index
    Cancel = IIf(GridEX1.Value(GridEX1.Columns("Flg").Index) = "H" And GridEX1.Value(GridEX1.Columns("sel").Index), False, True)
  Case Is = GridEX1.Columns("Observacion").Index
    Cancel = IIf(GridEX1.Value(GridEX1.Columns("sel").Index), False, True)
  Case Else
    Cancel = True
  End Select
End Sub

Private Sub txt_ImpTotal_Doc_Cobra_Change()
 If txt_ImpTotal_Doc_Cobra = "" Then txt_ImpTotal_Doc_Cobra = 0
End Sub

Private Sub txt_ImpTotal_Doc_Cobra_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub txt_ImpTotal_Doc_Cobra_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
  SoloNumeros txt_ImpTotal_Doc_Cobra, KeyAscii, True, 2, 9
End Sub

Private Sub TxtCod_Banco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_Banco", "Nom_Banco", "Tg_Banco where ", TxtCod_Banco, TxtDes_Banco, 1, Me)
    txtCuenta_Cod = ""
    txtCuenta_Des = ""
    Check_Moneda_Cuenta
  End If
End Sub

Private Sub txtCod_Moneda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_Moneda", "Nom_Moneda", "TG_Moneda where ", txtCod_Moneda, txtDes_Moneda, 1, Me)
    Limpia_Doc
  End If
End Sub

Private Sub txtCod_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Origen", "Des_Origen", " Cn_Origen where ", txtCod_Origen, txtDes_Origen, 1, Me)
End Sub

Private Sub txtCod_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAne, txtDes_TipAnex, 1, Me)
    Limpia_Doc
  End If
End Sub

Private Sub txtCod_TipCobra_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion_Store("Cn_Ventas_Muestra_Tipos_Cobranza_Permitidos '" & vusu & "'", txtCod_TipCobra, txtDes_TipCobra, 1, Me)
    Cambio_Apariencia
  End If
End Sub

Private Sub txtCod_TipDocCobra_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipDoc", "Des_TipDoc", " CN_TiposDocum where Flg_Doc_Cobranza = '*' and ", txtCod_TipDocCobra, txtDes_DocCobra, 1, Me)
End Sub

Private Sub txtCuenta_Cod_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    txtCuenta_Cod = Format(txtCuenta_Cod, "000")
    Call Busca_Opcion("Sec_Cuenta_Banco", "cod_cuenta", "Tg_Bancos_Cuentas where Cod_Banco ='" & TxtCod_Banco & "' and ", txtCuenta_Cod, txtCuenta_Des, 1, Me)
    Check_Moneda_Cuenta
  End If
End Sub

Private Sub txtCuenta_Des_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Sec_Cuenta_Banco", "cod_cuenta", "Tg_Bancos_Cuentas where Cod_Banco ='" & TxtCod_Banco & "' and ", txtCuenta_Cod, txtCuenta_Des, 2, Me)
    Check_Moneda_Cuenta
  End If
End Sub

Private Sub TxtDes_Banco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_Banco", "Nom_Banco", "Tg_Banco where ", TxtCod_Banco, TxtDes_Banco, 2, Me)
    txtCuenta_Cod = ""
    txtCuenta_Des = ""
    Check_Moneda_Cuenta
  End If
End Sub

Private Sub txtDes_Moneda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_Moneda", "Nom_Moneda", "TG_Moneda where ", txtCod_Moneda, txtDes_Moneda, 2, Me)
    Limpia_Doc
  End If
End Sub


Private Sub txtDes_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Origen", "Des_Origen", " Cn_Origen where ", txtCod_Origen, txtDes_Origen, 2, Me)
End Sub

Private Sub txtDes_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 2, Me)
    SendKeys "{TAB}"
    Limpia_Doc
  End If
End Sub

Private Sub txtDes_TipCobra_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion_Store("Cn_Ventas_Muestra_Tipos_Cobranza_Permitidos '" & vusu & "'", txtCod_TipCobra, txtDes_TipCobra, 2, Me)
    Cambio_Apariencia
  End If
End Sub

Private Sub txtFec_Diferido_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    txtTipo_Cambio = DevuelveCampo("Select dbo.SM_OBTIENE_TIPO_CAMBIO('" & txtFecha.Text & "')", cCONNECT)
    txtOtro_Tipo_Cambio = DevuelveCampo("Select dbo.SM_OBTIENE_TIPO_CAMBIO_EUROS('" & txtFecha.Text & "')", cCONNECT)
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtNum_DocCobra_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 1, Me)
    Limpia_Doc
  End If
End Sub

Private Sub txtNumeroPendiente_Change()
  Call gexGrid1.Find(gexGrid1.Columns("Numero").Index, jgexContains, txtNumeroPendiente)
End Sub

Private Sub txtNumeroPendiente_GotFocus()
  SendKeys "{Home}+{End}"
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

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtOtro_Tipo_Cambio_LostFocus()
  If txtOtro_Tipo_Cambio = "" Then txtOtro_Tipo_Cambio = 0
End Sub

Private Sub txtSer_DocCobra_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Public Sub CARGA_FACTURAS_PENDIENTES()
On Error GoTo hand
Dim SQL As String

Set RsGrid1 = CreateObject("ADODB.Recordset")
RsGrid1.CursorLocation = adUseClient


SQL = "Ventas_Muestra_Docum_Pedientes_Cobranzas '" & txtCod_TipAne & "','" & strCod_Anxo & "','" & txtCod_Moneda & "','" & txtFecha.Text & "',0,'" & TxtCod_Banco & "','" & txtCuenta_Cod & "','" & IIf(chkmuestra_saldo_total.Value = "1", "S", "N") & "','" & Format(DTFecha.Value, "dd/mm/yyyy") & "'"
Set RsGrid1 = CargarRecordSetDesconectado(SQL, cCONNECT)

Set gexGrid1.ADORecordset = RsGrid1
ConfigurarGrid gexGrid1

If RsGrid1.RecordCount Then

    TxtMonto1 = CALCULA_MONTO_TOTAL(RsGrid1)

    Set RsGrid2 = CreateObject("ADODB.Recordset")
    RsGrid2.CursorLocation = adUseClient
    Set RsGrid2.ActiveConnection = Nothing

    RsGrid2.Fields.Append RsGrid1.Fields("Correlativo").Name, RsGrid1.Fields(0).Type, RsGrid1.Fields(0).DefinedSize
    RsGrid2.Fields.Append RsGrid1.Fields("Numero").Name, RsGrid1.Fields(1).Type, RsGrid1.Fields(1).DefinedSize
    RsGrid2.Fields.Append RsGrid1.Fields("Fecha").Name, adDate
    RsGrid2.Fields.Append RsGrid1.Fields("Moneda").Name, RsGrid1.Fields(3).Type, RsGrid1.Fields(3).DefinedSize
    RsGrid2.Fields.Append RsGrid1.Fields("Tipo_Cambio").Name, adDouble
    RsGrid2.Fields.Append RsGrid1.Fields("Monto_Origen1").Name, adDouble
    RsGrid2.Fields.Append RsGrid1.Fields("Monto_Origen").Name, adDouble
    RsGrid2.Fields.Append RsGrid1.Fields("Monto_Aceptado").Name, adDouble
    RsGrid2.Fields.Append RsGrid1.Fields("Cod_Cobranza").Name, RsGrid1.Fields(8).Type, RsGrid1.Fields(8).DefinedSize
    RsGrid2.Fields.Append RsGrid1.Fields("Debe_Haber").Name, RsGrid1.Fields(9).Type, RsGrid1.Fields(9).DefinedSize
    RsGrid2.Fields.Append RsGrid1.Fields("Otro_Tip_Cambio").Name, adDouble
    RsGrid2.Fields.Append RsGrid1.Fields("Observacion").Name, RsGrid1.Fields(11).Type, RsGrid1.Fields(11).DefinedSize
    RsGrid2.Fields.Append RsGrid1.Fields("Tran_TipMonDoc").Name, RsGrid1.Fields(12).Type, RsGrid1.Fields(12).DefinedSize
    RsGrid2.Fields.Append RsGrid1.Fields("Doc_TipMonDoc").Name, RsGrid1.Fields(13).Type, RsGrid1.Fields(13).DefinedSize
    
    
    RsGrid2.Open

    Set gexGrid2.ADORecordset = RsGrid2
    ConfigurarGrid gexGrid2
    
    TxtMonto2 = 0
End If

txtNumeroPendiente.SetFocus

Exit Sub
Resume
hand:
ErrorHandler err, "CARGA_FACTURAS_PENDIENTES"
End Sub

Sub ConfigurarGrid(mGridEx As GridEx)
    mGridEx.Columns(2).Width = 1305
    mGridEx.Columns(3).Width = 945
    mGridEx.Columns(4).Width = 720
    gexGrid1.Columns(5).Width = 1065
    gexGrid1.Columns(6).Width = 1500
    gexGrid1.Columns(7).Width = 1365
    mGridEx.Columns("Correlativo").Visible = False
    mGridEx.Columns("Monto_Origen1").Visible = False
    mGridEx.Columns("Cod_Cobranza").Visible = False
    mGridEx.Columns("Debe_Haber").Visible = False
    mGridEx.Columns("Tran_TipMonDoc").Visible = False
    mGridEx.Columns("Doc_TipMonDoc").Visible = False
    mGridEx.Columns("Monto_Origen").Format = "###,###.00"
    mGridEx.Columns("Monto_Aceptado").Format = "###,###.00"
    mGridEx.Columns("Otro_tip_Cambio").Format = "###,###.######"
    
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
CALCULA_MONTO_TOTAL = Format(Monto, "###,###.000000")
End Function

Sub Resumen_Facturas()

Dim Monto As Double

Dim i As Integer

    Monto = 0
    txtDocumentos = ""
    
    If gexGrid2.RowCount > 0 Then
    
      gexGrid2.MoveFirst
      For i = 1 To gexGrid2.RowCount
          Monto = Monto + gexGrid2.Value(gexGrid2.Columns("Monto_Aceptado").Index)
          txtDocumentos = txtDocumentos + " " + gexGrid2.Value(gexGrid2.Columns("Numero").Index)
          gexGrid2.MoveNext
      Next
      
    End If

txtTotDoc = Format(Monto, "###,###.00")

End Sub

Sub Limpia_Doc()
  Set gexGrid1.ADORecordset = Nothing
  Set gexGrid2.ADORecordset = Nothing
  txtDocumentos.Text = ""
  txtTotDoc.Text = 0
End Sub


Sub Cambio_Apariencia()
On Error Resume Next
  If DevuelveCampo("select Flg_Cobranza_Simple from Cn_Ventas_Tipos_Cobranza where Cod_Tipcobranza ='" & txtCod_TipCobra & "'", cCONNECT) = "N" Then
   If Me.WindowState <> 2 Then Me.Height = 7710
    FunctButt1.Top = 6940
    GridEX1.Visible = True
  Else
    If Me.WindowState <> 2 Then Me.Height = 5325
    FunctButt1.Top = 4560
    GridEX1.Visible = False
  End If
End Sub
Sub Check_Moneda_Cuenta()

 txtCod_Moneda = DevuelveCampo("select Cod_Moneda from tg_bancos_cuentas Where Cod_Banco = '" & TxtCod_Banco & "' and Sec_Cuenta_Banco = '" & txtCuenta_Cod & "'", cCONNECT)
 If txtCod_Moneda <> "" Then
    txtDes_Moneda = DevuelveCampo("select Nom_Moneda from tg_moneda Where Cod_Moneda = '" & txtCod_Moneda & "'", cCONNECT)
    txtCod_Moneda.Enabled = False
    txtDes_Moneda.Enabled = False
 Else
    txtDes_Moneda = ""
    txtCod_Moneda.Enabled = True
    txtDes_Moneda.Enabled = True
 End If
 
End Sub

Private Sub txtOtro_Tipo_Cambio_GotFocus()
  SendKeys "{HOME}+{END}"
End Sub

Private Sub txtOtro_Tipo_Cambio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
  SoloNumeros txtOtro_Tipo_Cambio, KeyAscii, True, 4, 3
End Sub

Private Sub txtTipo_Cambio_GotFocus()
  SendKeys "{HOME}+{END}"
End Sub

Private Sub TxtTipo_Cambio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
  SoloNumeros txtTipo_Cambio, KeyAscii, True, 4, 3
End Sub

Private Sub txtTipo_Cambio_LostFocus()
  If txtTipo_Cambio = "" Then txtTipo_Cambio = 0
End Sub
