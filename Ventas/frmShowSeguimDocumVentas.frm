VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmShowSeguimDocumVentas 
   Caption         =   "Control de Documentos de Ventas"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDrawBack 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Seguimiento Draw Back"
      Height          =   2175
      Left            =   1170
      TabIndex        =   64
      Top             =   4635
      Visible         =   0   'False
      Width           =   9660
      Begin VB.TextBox txtEstado 
         BackColor       =   &H80000014&
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   4470
         MaxLength       =   30
         TabIndex        =   73
         Top             =   360
         Width           =   4380
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   480
         Left            =   7740
         TabIndex        =   71
         Top             =   1080
         Width           =   1740
      End
      Begin VB.TextBox txtDescrip_Tipdoc_DB 
         BackColor       =   &H80000014&
         Height          =   330
         Left            =   1650
         MaxLength       =   30
         TabIndex        =   69
         Top             =   375
         Width           =   1980
      End
      Begin VB.CommandButton cmdRetornaraPendiente 
         Caption         =   "Retornar a Pendiente de Envío"
         Height          =   480
         Left            =   5910
         TabIndex        =   68
         Top             =   1065
         Width           =   1740
      End
      Begin VB.CommandButton cmdEnviaraEnTramite 
         Caption         =   "Enviar a Aduana (En Trámite)"
         Height          =   480
         Left            =   165
         TabIndex        =   67
         Top             =   1080
         Width           =   1740
      End
      Begin VB.CommandButton cmdCobrado 
         Caption         =   "Cobrado"
         Height          =   480
         Left            =   2055
         TabIndex        =   66
         Top             =   1080
         Width           =   1740
      End
      Begin VB.CommandButton cmdRetornaraEnTramite 
         Caption         =   "Retornar a ""En Trámite"""
         Height          =   480
         Left            =   4020
         TabIndex        =   65
         Top             =   1080
         Width           =   1740
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Estado: "
         Height          =   315
         Left            =   3810
         TabIndex        =   72
         Tag             =   "Document Type"
         Top             =   420
         Width           =   1020
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Documento:"
         Height          =   315
         Left            =   675
         TabIndex        =   70
         Tag             =   "Document Type"
         Top             =   450
         Width           =   1020
      End
   End
   Begin VB.Frame fra_origen 
      Height          =   1935
      Left            =   4185
      TabIndex        =   58
      Top             =   4680
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox txtOrigen 
         Height          =   285
         Left            =   840
         TabIndex        =   62
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   1440
         TabIndex        =   61
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton Cmd_Imprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   600
         TabIndex        =   60
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1920
         TabIndex        =   59
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Origen"
         Height          =   195
         Left            =   240
         TabIndex        =   63
         Top             =   600
         Width           =   465
      End
   End
   Begin VB.Frame fraDeuda 
      Height          =   2715
      Left            =   8625
      TabIndex        =   43
      Top             =   -45
      Width           =   3165
      Begin VB.TextBox Txt_Importe 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   360
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   1890
         Width           =   1455
      End
      Begin VB.TextBox TxtDDolares 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox TxtDsoles 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   675
         Width           =   1455
      End
      Begin VB.TextBox TxtDOtros 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   1485
         Width           =   1455
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Total $"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   375
         TabIndex        =   50
         Top             =   1980
         Width           =   735
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Deuda Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   930
         TabIndex        =   49
         Top             =   255
         Width           =   1305
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Soles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   375
         TabIndex        =   48
         Top             =   750
         Width           =   615
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Dolares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   375
         TabIndex        =   47
         Top             =   1125
         Width           =   840
      End
      Begin VB.Label Label14 
         Caption         =   "Otra Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   375
         TabIndex        =   46
         Top             =   1440
         Width           =   900
      End
   End
   Begin VB.Frame fraGeneral 
      Caption         =   "Argumentos de Búsqueda General"
      Height          =   1995
      Left            =   45
      TabIndex        =   26
      Top             =   660
      Width           =   8565
      Begin VB.OptionButton optDrawBack 
         Caption         =   "Draw Back Pendiente / En Trámite"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   1665
         Width           =   3090
      End
      Begin VB.Frame fraDocumGen 
         BorderStyle     =   0  'None
         Height          =   1125
         Left            =   375
         TabIndex        =   28
         Top             =   480
         Width           =   8115
         Begin VB.TextBox Txt_Descripcion 
            Height          =   285
            Left            =   1575
            TabIndex        =   55
            Top             =   45
            Width           =   2160
         End
         Begin VB.TextBox Txt_Origen 
            Height          =   285
            Left            =   930
            TabIndex        =   54
            Text            =   "N"
            Top             =   45
            Width           =   615
         End
         Begin VB.TextBox txtSer_Hasta 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   5250
            MaxLength       =   3
            TabIndex        =   38
            Top             =   840
            Width           =   540
         End
         Begin VB.TextBox txtNum_Hasta 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   6645
            MaxLength       =   15
            TabIndex        =   37
            Top             =   840
            Width           =   1440
         End
         Begin VB.TextBox txtDes_TipDoc2 
            BackColor       =   &H80000014&
            Height          =   315
            Left            =   1290
            MaxLength       =   30
            TabIndex        =   32
            Top             =   450
            Width           =   2445
         End
         Begin VB.TextBox txtCod_TipDoc2 
            BackColor       =   &H80000014&
            Height          =   315
            Left            =   930
            MaxLength       =   2
            TabIndex        =   31
            Text            =   "FA"
            Top             =   450
            Width           =   360
         End
         Begin VB.TextBox txtSer_Desde 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   5250
            MaxLength       =   3
            TabIndex        =   30
            Top             =   435
            Width           =   540
         End
         Begin VB.TextBox txtNum_Desde 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   6645
            MaxLength       =   15
            TabIndex        =   29
            Top             =   435
            Width           =   1440
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Origen"
            Height          =   195
            Left            =   75
            TabIndex        =   56
            Top             =   90
            Width           =   465
         End
         Begin VB.Label Label13 
            Caption         =   "Hasta:"
            Height          =   225
            Left            =   3795
            TabIndex        =   41
            Tag             =   "Number"
            Top             =   915
            Width           =   645
         End
         Begin VB.Label Label12 
            Caption         =   "Número :"
            Height          =   225
            Left            =   5895
            TabIndex        =   40
            Tag             =   "Number"
            Top             =   915
            Width           =   645
         End
         Begin VB.Label Label11 
            Caption         =   "Serie: "
            Height          =   240
            Left            =   4725
            TabIndex        =   39
            Top             =   915
            Width           =   525
         End
         Begin VB.Label Label10 
            Caption         =   "Desde :"
            Height          =   225
            Left            =   3795
            TabIndex        =   36
            Tag             =   "Number"
            Top             =   525
            Width           =   645
         End
         Begin VB.Label Label9 
            Caption         =   "Número :"
            Height          =   225
            Left            =   5895
            TabIndex        =   35
            Tag             =   "Number"
            Top             =   510
            Width           =   645
         End
         Begin VB.Label Label7 
            Caption         =   "Serie: "
            Height          =   210
            Left            =   4725
            TabIndex        =   34
            Top             =   510
            Width           =   510
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo Documento:"
            Height          =   390
            Left            =   60
            TabIndex        =   33
            Tag             =   "Document Type"
            Top             =   420
            Width           =   930
         End
      End
      Begin VB.OptionButton optRangoDocum 
         Caption         =   "Rango de Documentos"
         Height          =   255
         Left            =   105
         TabIndex        =   27
         Top             =   255
         Value           =   -1  'True
         Width           =   2010
      End
   End
   Begin VB.Frame Frame3 
      Height          =   705
      Left            =   45
      TabIndex        =   23
      Top             =   -45
      Width           =   8565
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   435
         Left            =   6885
         TabIndex        =   42
         Top             =   180
         Width           =   1545
      End
      Begin VB.OptionButton optEspecifica 
         Caption         =   "Busqueda Específica"
         Height          =   240
         Left            =   3510
         TabIndex        =   25
         Top             =   285
         Width           =   2025
      End
      Begin VB.OptionButton optGeneral 
         Caption         =   "Busqueda General"
         Height          =   240
         Left            =   1725
         TabIndex        =   24
         Top             =   285
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame fraEspecifico 
      Caption         =   "Argumentos de Búsqueda Específica"
      Height          =   1995
      Left            =   75
      TabIndex        =   0
      Top             =   7860
      Visible         =   0   'False
      Width           =   8565
      Begin VB.Frame frProveedor 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   495
         Left            =   60
         TabIndex        =   16
         Top             =   210
         Width           =   7335
         Begin VB.TextBox txtRuc 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   600
            MaxLength       =   11
            TabIndex        =   20
            Top             =   120
            Width           =   1200
         End
         Begin VB.TextBox txtCod_TipAne 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   2010
            MaxLength       =   2
            TabIndex        =   19
            Top             =   120
            Width           =   360
         End
         Begin VB.TextBox txtDes_Anexo 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   3015
            MaxLength       =   30
            TabIndex        =   18
            Top             =   120
            Width           =   4305
         End
         Begin VB.TextBox txtCod_Anexo 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   2385
            MaxLength       =   4
            TabIndex        =   17
            Top             =   120
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1875
            TabIndex        =   22
            Tag             =   "Anexo Type"
            Top             =   135
            Width           =   90
         End
         Begin VB.Label Label8 
            Caption         =   "Ruc :"
            Height          =   180
            Left            =   120
            TabIndex        =   21
            Tag             =   "Anexo Type"
            Top             =   135
            Width           =   435
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   1575
         TabIndex        =   9
         Top             =   795
         Width           =   5580
         Begin MSComCtl2.DTPicker dtpFecEmiIni 
            Height          =   315
            Left            =   1350
            TabIndex        =   10
            Top             =   120
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   94306305
            CurrentDate     =   37543
         End
         Begin MSComCtl2.DTPicker dtpFecEmiFin 
            Height          =   315
            Left            =   3510
            TabIndex        =   11
            Top             =   120
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   94306305
            CurrentDate     =   37543
         End
         Begin VB.Label Label1 
            Caption         =   "Rango Fecha de Emisión:"
            Height          =   525
            Left            =   0
            TabIndex        =   12
            Top             =   75
            Width           =   1365
         End
      End
      Begin VB.OptionButton opTodas 
         Caption         =   "Todas"
         Height          =   255
         Left            =   90
         TabIndex        =   8
         Top             =   1140
         Width           =   855
      End
      Begin VB.OptionButton oprCanceladas 
         Caption         =   "Canceladas"
         Height          =   255
         Left            =   90
         TabIndex        =   7
         Top             =   900
         Width           =   1215
      End
      Begin VB.OptionButton opPendiente 
         Caption         =   "Pendientes"
         Height          =   255
         Left            =   90
         TabIndex        =   6
         Top             =   660
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox txtDes_TipDoc 
         BackColor       =   &H80000014&
         Height          =   330
         Left            =   2805
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1425
         Width           =   1980
      End
      Begin VB.TextBox txtCod_TipDoc 
         BackColor       =   &H80000014&
         Height          =   330
         Left            =   2445
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "FA"
         Top             =   1425
         Width           =   360
      End
      Begin VB.TextBox txtSer_Docum 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   5475
         MaxLength       =   3
         TabIndex        =   3
         Top             =   1440
         Width           =   540
      End
      Begin VB.TextBox txtNum_Docum 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   6870
         MaxLength       =   15
         TabIndex        =   2
         Top             =   1440
         Width           =   1440
      End
      Begin VB.OptionButton optDocRef 
         Caption         =   "Documento Específico"
         Height          =   435
         Left            =   90
         TabIndex        =   1
         Top             =   1380
         Width           =   1260
      End
      Begin VB.Label Label5 
         Caption         =   "Número :"
         Height          =   225
         Left            =   6120
         TabIndex        =   15
         Tag             =   "Number"
         Top             =   1515
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Serie Docum.:"
         Height          =   375
         Left            =   4860
         TabIndex        =   14
         Top             =   1410
         Width           =   750
      End
      Begin VB.Label lblCod_TipOrdCom 
         Caption         =   "Tipo Documento:"
         Height          =   390
         Left            =   1575
         TabIndex        =   13
         Tag             =   "Document Type"
         Top             =   1395
         Width           =   930
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4110
      Left            =   45
      TabIndex        =   53
      Top             =   2715
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   7250
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmShowSeguimDocumVentas.frx":0000
      Column(2)       =   "frmShowSeguimDocumVentas.frx":00C8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmShowSeguimDocumVentas.frx":016C
      FormatStyle(2)  =   "frmShowSeguimDocumVentas.frx":02A4
      FormatStyle(3)  =   "frmShowSeguimDocumVentas.frx":0354
      FormatStyle(4)  =   "frmShowSeguimDocumVentas.frx":0408
      FormatStyle(5)  =   "frmShowSeguimDocumVentas.frx":04E0
      FormatStyle(6)  =   "frmShowSeguimDocumVentas.frx":0598
      FormatStyle(7)  =   "frmShowSeguimDocumVentas.frx":0678
      FormatStyle(8)  =   "frmShowSeguimDocumVentas.frx":0724
      ImageCount      =   0
      PrinterProperties=   "frmShowSeguimDocumVentas.frx":07D4
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   5340
      TabIndex        =   57
      Top             =   6825
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   900
      Custom          =   $"frmShowSeguimDocumVentas.frx":09AC
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   105
      Top             =   7440
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmShowSeguimDocumVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrEstus As String, strSQL As String
Public codigo, Descripcion As String, TipoAdd As String

Dim OP_Opcion As String
Dim SNum_Corre  As String
Public oGroup As GridEX20.JSGroup
Public oFormat As JSFormatStyle



Sub Reporte_Masivo()
On Error GoTo ERROR
Dim sSQL As String
Dim oo As Object
Dim Ruta As String
Dim Reg1 As ADODB.Recordset
Dim sRuta_Logo As String

sRuta_Logo = DevuelveCampo("SELECT Ruta_Logo=ISNULL(Ruta_Logo,'') FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA = '" & vemp & "'", cCONNECT)

sSQL = "Ventas_Muestra_Documentos_por_Cerrar '1','','','1','" & txtOrigen & "','','','','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "'"

Set gridex1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)

Set Reg1 = GetRecordset1(cCONNECT, sSQL)

Configurar

Ruta = vRuta & "\RptCuentasCorrientes_Clientes.xlt"
Set oo = CreateObject("excel.application")
oo.Workbooks.Open Ruta
oo.Visible = True
oo.displayalerts = False
        
oo.Run "Reporte", Reg1, sRuta_Logo
Set oo = Nothing

Exit Sub
ERROR:
    errores err.Number
End Sub




Private Sub Cmd_Cancelar_Click()
    fra_origen.Visible = False
End Sub

Private Sub Cmd_Imprimir_Click()
Reporte_Masivo
fra_origen.Visible = False
End Sub

Private Sub CmdImprimir_Click()
Reporte
End Sub
Sub Reporte()
On Error GoTo ERROR
Dim sSQL As String
Dim oo As Object
Dim Ruta As String
Dim sRuta_Logo As String

sRuta_Logo = DevuelveCampo("SELECT Ruta_Logo=ISNULL(Ruta_Logo,'') FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA = '" & vemp & "'", cCONNECT)


If gridex1.RowCount = 0 Then Exit Sub

If Not optGeneral.Value Then
    Ruta = vRuta & "\RptCuentasCorrientes_Clientes.xlt"
Else
    Ruta = vRuta & "\RptCuentasCorrientes_Clientes_Rango.xlt"
End If
Set oo = CreateObject("excel.application")
oo.Workbooks.Open Ruta
oo.Visible = True
oo.displayalerts = False
        
oo.Run "Reporte", gridex1.ADORecordset, sRuta_Logo
Set oo = Nothing

Exit Sub
ERROR:
    errores err.Number
End Sub


Private Sub cmdVerDetalle_Click()
Call GridEX1_DblClick
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Command3_Click()
    fra_origen.Visible = True
    Txt_Origen.SetFocus

End Sub

Private Sub cmdCobrado_Click()
    
    CambiarStatus "CN_VENTAS_CAMBIAR_STATUS_REGISTRO_COBRO_DRAWBACK"

End Sub

Private Sub cmdEnviaraEnTramite_Click()
    
    CambiarStatus "CN_VENTAS_CAMBIAR_STATUS_REGISTRO_ENVIO_DRAWBACK"
End Sub

Private Sub cmdRetornaraEnTramite_Click()
    
    CambiarStatus "CN_VENTAS_CAMBIAR_STATUS_REGISTRO_COBRO_A_ENVIO_DRAWBACK"

End Sub

Private Sub cmdRetornaraPendiente_Click()
    
    CambiarStatus "CN_VENTAS_CAMBIAR_STATUS_REGISTRO_PENDIENTE_DRAWBACK"
End Sub

Private Sub CmdSalir_Click()
    Me.fraDrawBack.Visible = False
End Sub

Private Sub dtpFecEmiFin_Validate(Cancel As Boolean)
If dtpFecEmiIni > dtpFecEmiFin Then
  MsgBox "Fecha Final no puede ser menor a la fecha Inicial", vbInformation, "AVISO"
  dtpFecEmiIni = dtpFecEmiFin
End If
End Sub

Private Sub dtpFecEmiIni_Change()
  gridex1.ClearFields
  dtpFecEmiFin.Value = Date
End Sub

Private Sub dtpFecEmiIni_Validate(Cancel As Boolean)
If dtpFecEmiIni > dtpFecEmiFin Then
  MsgBox "Fecha Inicial no puede ser mayor a la fecha final", vbInformation, "AVISO"
  dtpFecEmiIni = dtpFecEmiFin
End If
End Sub

Private Sub Form_Load()
  FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name) & "/SALIR"

  txtCod_TipAne = "C"
  dtpFecEmiIni.Value = Date
  dtpFecEmiFin.Value = Date
  
  OP_Opcion = "5"
  CargaValoresDefault
  
  BUSCA_ORIGEN 1
  txtCod_TipDoc2_KeyPress 13
End Sub

Private Sub cmdBuscar_Click()
  Buscar
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub Buscar()

Dim sSQL As String

If optEspecifica Then
    If (Trim(txtCod_Anexo) = "" And frProveedor.Visible) Then
    
      If optDocRef Then
            If txtCod_TipAne = "" Or txtCod_Anexo = "" Then
                Aviso "Debe Ingresar un Anexo Específico", 1
                Exit Sub
            End If
      
      
            If txtCod_TipDoc = "" Or txtNum_Docum = "" Then
                Aviso "Debe Ingresar un documento Específico", 1
                Exit Sub
            End If
      Else
    
          If (IsNull(dtpFecEmiIni) Or IsNull(dtpFecEmiIni)) Then
            MsgBox "Ingrese un Rango de Fechas", vbInformation, "AVISO"
            Exit Sub
          End If
        
          If (dtpFecEmiFin - dtpFecEmiIni) > 60 Then
            MsgBox "No puede Ingresar un Rango Mayor a 60 Dias", vbInformation, "AVISO"
            Exit Sub
          End If
      End If
    Else
    
      If (oprCanceladas Or opTodas) And (IsNull(dtpFecEmiIni) Or IsNull(dtpFecEmiIni)) Then
        MsgBox "Ingrese un Rango de Fechas", vbInformation, "AVISO"
        Exit Sub
      
      End If
        
    End If
Else
    If Trim(Txt_Origen.Text) = "" Then
        Aviso "Ingrese un Origen de Documento", 2
        Txt_Origen.SetFocus
        Exit Sub
    End If
    
    If Trim(txtCod_TipDoc2.Text) = "" Then
        Aviso "Ingrese un Tipo de Documento", 2
        txtCod_TipDoc2.SetFocus
        Exit Sub
    End If
    
    If Trim(txtSer_Desde.Text) = "" Then
        Aviso "Ingrese una Serie de Documento", 2
        txtSer_Desde.SetFocus
        Exit Sub
    End If
    
    If Trim(txtSer_Hasta.Text) = "" Then
        Aviso "Ingrese una Serie de Documento", 2
        txtSer_Hasta.SetFocus
        Exit Sub
    End If
    
    If Trim(txtNum_Desde.Text) = "" Then
        Aviso "Ingrese un Número de Documento", 2
        txtNum_Desde.SetFocus
        Exit Sub
    End If
    
    If Trim(txtNum_Hasta.Text) = "" Then
        Aviso "Ingrese un Número de Documento", 2
        txtNum_Hasta.SetFocus
        Exit Sub
    End If
    
    txtCod_TipDoc.Text = txtCod_TipDoc2.Text
    txtSer_Docum = txtSer_Desde
    txtNum_Docum = txtNum_Desde
End If



sSQL = "Ventas_Muestra_Documentos_por_Cerrar '" & OP_Opcion & "','" & txtCod_TipAne & "','" & txtCod_Anexo & "','2','" & Txt_Origen & "','" & txtCod_TipDoc & "','" & txtSer_Docum & "','" & txtNum_Docum & _
"','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','" & txtSer_Hasta & "','" & txtNum_Hasta & "'"

gridex1.ClearFields

gridex1.DefaultGroupMode = jgexDGMExpanded
Set gridex1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)


Configurar

End Sub
Sub Configurar()

If Not optGeneral.Value Then
    Set oGroup = gridex1.Groups.Add(gridex1.Columns("Cliente").Index, jgexSortAscending)
End If

gridex1.BackColorRowGroup = &H80000005

Txt_Importe = Format(gridex1.Value(gridex1.Columns("SALDO_TOTAl").Index), "##,##0.00")
TxtDDolares = Format(gridex1.Value(gridex1.Columns("SALDO_total_DOLARES").Index), "##,##0.00")
TxtDsoles = Format(gridex1.Value(gridex1.Columns("SALDO_total_SOLES").Index), "##,##0.00")
TxtDOtros = Format(gridex1.Value(gridex1.Columns("SALDO_total_OTROS").Index), "##,##0.00")

gridex1.Columns("Cod_Tipdoc").Caption = "Tipo"
gridex1.Columns("Cod_Tipdoc").Width = 600

gridex1.Columns("SALDO_TOTAl").Visible = False
gridex1.Columns("Ruc").Visible = False


gridex1.Columns("Num_Corre").Width = 0
gridex1.Columns("saldo_equivalente").Width = 0

If Not optGeneral.Value Then
    gridex1.Columns("Cliente").Width = 0
    gridex1.Columns("Anexo_Contable").Visible = False
Else
    gridex1.Columns("Cliente").Visible = True
    gridex1.Columns("Anexo_Contable").Visible = True
    gridex1.Columns("cliente").Width = 2000
    gridex1.Columns("Cod_TipDoc").Visible = False
End If

gridex1.Columns("SALDO_TOTAl").Visible = False
gridex1.Columns("SALDO_total_SOLES").Visible = False
gridex1.Columns("SALDO_total_DOLARES").Visible = False
gridex1.Columns("SALDO_total_otros").Visible = False

gridex1.Columns("Imp_Total").Width = 1200
gridex1.Columns("Imp_Total").Caption = "Imp Total"

gridex1.Columns("Saldo_Dolares").Width = 1200
gridex1.Columns("Saldo_Dolares").Caption = "Saldo Dolares"

gridex1.Columns("Saldo_Soles").Width = 1200
gridex1.Columns("Saldo_Soles").Caption = "Saldo Soles"

gridex1.Columns("Saldo_Otros").Width = 1200
gridex1.Columns("Saldo_Otros").Caption = "Saldo Otra Moneda"

gridex1.Columns("Importe_Cancelado").Width = 1300
gridex1.Columns("Importe_Cancelado").Caption = "Imp Cancelado"

gridex1.Columns("saldo_equivalente").Width = 1300
gridex1.Columns("saldo_equivalente").Caption = "saldo Equivalente"



gridex1.Columns("Fec_Emision").Width = 1125
'GridEX1.Columns("Fec_VenDoc").Width = 1080
gridex1.Columns("Num_Registro").Width = 1155
gridex1.Columns("Moneda").Width = 720

gridex1.Columns("Status_DrawBack").Visible = False
gridex1.Columns("Des_Status_DrawBack").Visible = False

gridex1.Columns("Flg_Status_DrawBack").Visible = False
gridex1.Columns("Des_Status").Caption = "Estado Draw Back"

If txtCod_TipAne = "" Then gridex1.DefaultGroupMode = jgexDGMCollapsed Else gridex1.DefaultGroupMode = jgexDGMExpanded

gridex1.ContinuousScroll = True

End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "DETALLE"
            cmdVerDetalle_Click
        Case "IMPRIMIRTODOS"
            fra_origen.Visible = True
            Txt_Origen.SetFocus
        Case "IMPRIMIR"
            Reporte
        Case "DRAWBACK"
            SNum_Corre = gridex1.Value(gridex1.Columns("NUM_CORRE").Index)
            txtDescrip_Tipdoc_DB = gridex1.Value(gridex1.Columns("COD_TIPDOC").Index) & " " & gridex1.Value(gridex1.Columns("DOCUMENTO").Index)
            txtEstado = gridex1.Value(gridex1.Columns("DES_STATUS").Index)
            Me.fraDrawBack.Visible = True
        Case "SALIR"
            Unload Me
    End Select
End Sub

Private Sub GridEX1_DblClick()
  If gridex1.RowCount = 0 Then Exit Sub
  Load frmMuestraGeneral
  frmMuestraGeneral.Caption = "Detalle Cliente " & gridex1.Value(gridex1.Columns("Cliente").Index) & " Documento : " & gridex1.Value(gridex1.Columns("Documento").Index)
  frmMuestraGeneral.strSQL = "Ventas_Muestra_Cobranzas_del_Documento '" & gridex1.Value(gridex1.Columns("NUM_CORRE").Index) & "'"
  frmMuestraGeneral.Buscar
  frmMuestraGeneral.Show vbModal
End Sub

Private Sub opPendiente_Click()
StrEstus = "P"
OP_Opcion = "1"

End Sub

Private Sub oprCanceladas_Click()
StrEstus = "C"
OP_Opcion = "2"
End Sub

Private Sub optDocumEspecifico_Click()
    StrEstus = "E"
    txtCod_TipDoc.SetFocus
End Sub

Private Sub optDocRef_Click()
    StrEstus = "R"
    OP_Opcion = "4"
    txtCod_TipDoc.SetFocus
End Sub



Private Sub Option1_Click()
Op: OP_Opcion = "1"
End Sub

Private Sub optDrawBack_Click()
    OP_Opcion = "6"
End Sub

Private Sub optEspecifica_Click()
    Me.fraGeneral.Visible = False
    Me.fraEspecifico.Top = Me.fraGeneral.Top
    Me.fraEspecifico.Left = Me.fraGeneral.Left
    Me.fraEspecifico.Visible = True
    Me.txtRuc.SetFocus
End Sub

Private Sub optGeneral_Click()
    CargaValoresDefault
    Me.fraGeneral.Visible = True
    Me.fraEspecifico.Visible = False
    Me.cmdBuscar.SetFocus
    
    If optRangoDocum.Value Then
        OP_Opcion = "5"
    Else
        OP_Opcion = "6"
    End If
    
End Sub

Private Sub opTodas_Click()
StrEstus = "T"
OP_Opcion = "3"
End Sub

Sub BUSCA_TIPO_ANEXO(Tipo As Integer, Ubic As Integer)
    Select Case Tipo
        Case 1:
                If Ubic = 1 Then
                    strSQL = "SELECT DES_TIPANEX FROM CN_TipoAnexoContable WHERE COD_TIPANEX = '" & txtCod_TipAne.Text & "'"
                    txtCod_Anexo.SetFocus
                Else
                End If
        Case 2:
                Dim oTipo As New frmBusqGeneral
                Dim RS As Object
                Set RS = CreateObject("ADODB.Recordset")
                Set oTipo.oParent = Me
                If Ubic = 1 Then
                    oTipo.SQuery = "SELECT COD_TIPANEX as Código, DES_TIPANEX as Descripción FROM CN_TipoAnexoContable "
                Else
                End If
                oTipo.CARGAR_DATOS
                oTipo.Show 1
                If codigo <> "" Then
                    If Ubic = 1 Then
                        txtCod_TipAne.Text = Trim(codigo)
                        txtCod_Anexo.SetFocus
                    Else
                    End If
                End If
                Set oTipo = Nothing
                Set RS = Nothing
                
    End Select
End Sub


Sub BUSCA_ANEXO(Tipo As Integer, Ubic As Integer)

Dim iLen As Integer
    Select Case Tipo
        Case 1:
                If Ubic = 1 Then
                    strSQL = "SELECT MIN(DATALENGTH(COD_ANXO)) FROM CN_AnexosContables"
                    iLen = Trim(DevuelveCampo(strSQL, cCONNECT))
                    
                    txtCod_Anexo.Text = Right(Repl("0", iLen) & txtCod_Anexo, iLen)
                    
                     
                     strSQL = "SELECT Des_Anexo FROM CN_AnexosContables WHERE Cod_TipAnEX = '" & txtCod_TipAne.Text & "' AND Cod_Anxo = '" & txtCod_Anexo.Text & "'"
                     txtDes_Anexo.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
                     SendKeys "{TAB}"
                     
                     Exit Sub
                     
                Else
                End If
        Case 2:
        
                Dim oTipo As New frmBusqGeneral
                Dim RS As Object
                Set RS = CreateObject("ADODB.Recordset")
                Set oTipo.oParent = Me
                If Ubic = 1 Then
                    oTipo.SQuery = "SELECT Cod_Anxo as Código, Des_Anexo as Descripción FROM CN_AnexosContables WHERE Cod_TipAnEX = '" & txtCod_TipAne.Text & "' AND Des_Anexo like '%" & Trim(txtDes_Anexo.Text) & "%'"
                Else
                End If
                oTipo.CARGAR_DATOS
                oTipo.Top = txtDes_Anexo.Top + txtDes_Anexo.Height
                oTipo.Left = txtDes_Anexo.Left
                oTipo.DGridLista.Columns(1).Width = 1000
                oTipo.Show 1
                If codigo <> "" Then
                    If Ubic = 1 Then
                        txtCod_Anexo.Text = Trim(codigo)
                        txtDes_Anexo.Text = Trim(Descripcion)
                        strSQL = "SELECT num_ruc FROM CN_AnexosContables WHERE Cod_TipAnEX = '" & txtCod_TipAne.Text & "' AND Cod_Anxo = '" & txtCod_Anexo.Text & "'"
                        txtRuc = Trim(DevuelveCampo(strSQL, cCONNECT))

                        SendKeys "{TAB}"
                    Else
                    End If
                End If
                Set oTipo = Nothing
                Set RS = Nothing
                
    End Select
    
End Sub


Private Sub optProveedor_Click()
  frProveedor.Visible = True

  LimpiaFr
  txtCod_TipAne = "P"
End Sub

Sub LimpiaFr()
  gridex1.ClearFields
  txtCod_Anexo = ""
  txtDes_Anexo = ""
  txtCod_TipAne = ""


  txtRuc = ""
End Sub


Private Sub optRangoDocum_Click()
    OP_Opcion = "5"
End Sub

Private Sub Txt_Descripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Trim(Me.Txt_Descripcion.Text) = "" Then
            Call Me.BUSCA_ORIGEN(3)
            
        Else
            Call Me.BUSCA_ORIGEN(1)
        End If
        
    End If
End Sub

Private Sub Txt_Origen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Trim(Me.Txt_Origen.Text) = "" Then
            Call Me.BUSCA_ORIGEN(3)
            
        Else
            Call Me.BUSCA_ORIGEN(1)
        End If
        Me.txtCod_TipDoc2.SetFocus
End If

End Sub

Private Sub txtCod_Anexo_KeyPress(KeyAscii As Integer)
    gridex1.ClearFields
    If KeyAscii = vbKeyReturn Then
        If Trim(txtCod_Anexo.Text) <> "" Then
            Call BUSCA_ANEXO(1, 1)
        End If
    End If
End Sub


Private Sub txtCod_TipAne_KeyPress(KeyAscii As Integer)
  gridex1.ClearFields
  If KeyAscii = vbKeyReturn Then
      If Trim(txtCod_TipAne.Text) <> "" Then
          Call BUSCA_TIPO_ANEXO(1, 1)
      Else
          Call BUSCA_TIPO_ANEXO(2, 1)
      End If
  End If
End Sub

Private Sub txtCod_TipDoc2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        CargaValoresDefault
        Set frmBusqGeneral.oParent = Me
            frmBusqGeneral.SQuery = "SELECT COD_TIPDOC AS CODIGO, DES_TIPDOC AS DESCRIPCION , DOC_SUNAT AS TIPO FROM CN_TIPOSDOCUM WHERE COD_TIPDOC LIKE '%" & Trim(txtCod_TipDoc2.Text) & "%'"
            frmBusqGeneral.CARGAR_DATOS
            If frmBusqGeneral.DGridLista.RowCount > 1 Then
                frmBusqGeneral.Show 1
            Else
                frmBusqGeneral.cmdAceptar_Click
            End If
        If codigo <> "" Then
            txtCod_TipDoc2.Text = codigo
            txtDes_TipDoc2.Text = Descripcion
            If Me.Visible Then
                txtNum_Desde.SetFocus
            End If
            CargaValoresDefault
        Else
            txtCod_TipDoc2.Text = ""
            txtDes_TipDoc2.Text = ""
        End If
        codigo = ""
        Descripcion = ""
        
    End If
End Sub

Private Sub txtDes_Anexo_KeyPress(KeyAscii As Integer)
    gridex1.ClearFields
    If KeyAscii = vbKeyReturn Then
        If Trim(txtDes_Anexo.Text) <> "" Then
            If Len(Trim(txtDes_Anexo)) > 2 Then
                Call BUSCA_ANEXO(2, 1)
            Else
                Aviso "Debe ingresar al menos 3 caracteres del Nombre requerido", 1
                Exit Sub
            End If
        Else
            Aviso "Debe ingresar al menos 3 caracteres del Nombre requerido", 1
            Exit Sub
        End If
    End If
End Sub

Private Sub BUSCARUC(opcion As Integer)

On Error GoTo Fin
Dim strSQL As String
Dim oTipo As New frmBusqGeneral

    strSQL = "SELECT num_ruc as Ruc,Des_Anexo Descripcion FROM CN_AnexosContables "
    txtRuc = Trim(txtRuc)
    
    strSQL = strSQL & " where num_ruc like '%" & txtRuc & "%' and Cod_TipAnex ='C'"
    
    txtRuc = ""
        
    Set oTipo.oParent = Me
    
    oTipo.SQuery = strSQL
    oTipo.CARGAR_DATOS
    oTipo.DGridLista.Columns(1).Width = 4350.047
    oTipo.Show 1
    If codigo <> "" Then
      txtRuc = Trim(codigo)
      txtDes_Anexo = Trim(Descripcion)
      
      strSQL = "SELECT Cod_TipAnEx FROM CN_AnexosContables WHERE num_ruc = '" & txtRuc.Text & "' and Cod_TipAnex ='C'"
      txtCod_TipAne.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
      strSQL = "SELECT Cod_Anxo FROM CN_AnexosContables WHERE num_ruc = '" & txtRuc.Text & "' and Cod_TipAnex ='C'"
      txtCod_Anexo.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
      
      SendKeys "{TAB}"
    End If
    Set oTipo = Nothing
    
Exit Sub
Resume
Fin:
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda (" & opcion & ")"
End Sub


Private Sub txtNum_Desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtNum_Hasta.SetFocus
    End If
End Sub

Private Sub txtNum_Desde_LostFocus()
    If txtNum_Desde <> "" Then
        txtNum_Desde = StrZero(txtNum_Desde, 8)
    End If
End Sub

Private Sub txtNum_Hasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdBuscar.SetFocus
    End If

End Sub


Private Sub txtNum_Hasta_LostFocus()
    If txtNum_Hasta <> "" Then
        txtNum_Hasta = StrZero(txtNum_Hasta, 8)
    End If

End Sub

Private Sub txtOrigen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Trim(Me.txtOrigen.Text) = "" Then
            Call Me.BUSCA_ORIGEN2(3)
            
        Else
            Call Me.BUSCA_ORIGEN2(1)
        End If
        Cmd_Imprimir.SetFocus
End If

End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    BUSCARUC 1
  End If
End Sub
Private Sub txtCod_TipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then

            Set frmBusqGeneral.oParent = Me
                frmBusqGeneral.SQuery = "SELECT COD_TIPDOC AS CODIGO, DES_TIPDOC AS DESCRIPCION , DOC_SUNAT AS TIPO FROM CN_TIPOSDOCUM WHERE COD_TIPDOC LIKE '%" & Trim(txtCod_TipDoc.Text) & "%'"
                frmBusqGeneral.CARGAR_DATOS
                frmBusqGeneral.Show 1
            If codigo <> "" Then
                txtCod_TipDoc.Text = codigo
                txtDes_TipDoc.Text = Descripcion
                txtSer_Docum.SetFocus
                
            Else
                txtCod_TipDoc.Text = ""
                txtDes_TipDoc.Text = ""
            End If
            codigo = ""
            Descripcion = ""

    End If
End Sub


Private Sub txtSer_Docum_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtSer_Docum = Format(txtSer_Docum, "000")
        txtNum_Docum.SetFocus
    End If
End Sub

Private Sub txtNum_Docum_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtNum_Docum = Format(txtNum_Docum, "00000000")
        cmdBuscar.SetFocus
    End If
End Sub


Public Sub BUSCA_ORIGEN(Tipo As Integer)
On Error GoTo hand
    Select Case Tipo
        Case 1:
                    strSQL = "SELECT Des_Origen as 'Descripción' FROM  cn_origen WHERE Origen = '" & Trim(Me.Txt_Origen.Text) & "'"
                    Me.Txt_Descripcion.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
                    CargaValoresDefault
                    
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim RS As Object
                    Set RS = CreateObject("ADODB.Recordset")
                    Set oTipo.oParent = Me
                    
                    
                    If Tipo = 2 Then
                        oTipo.SQuery = "SELECT Origen as 'Código', Des_Origen as 'Descripción' FROM cn_origen WHERE Des_Origen LIKE '%" & Trim(Me.Txt_Origen) & "%' ORDER BY Des_Origen"
                    Else
                        oTipo.SQuery = "SELECT ORIGEN as 'Código', Des_Origen AS 'Descripción' FROM Cn_Origen ORDER BY Des_Origen"
                    End If
                    
                    
                    oTipo.CARGAR_DATOS
                    If oTipo.DGridLista.RowCount > 1 Then
                        oTipo.Show 1
                    End If
                    If codigo <> "" Then
                        Me.Txt_Origen = Trim(codigo)
                        Me.Txt_Descripcion = Trim(Descripcion)
                        CargaValoresDefault
                    End If
                    Set oTipo = Nothing
                    Set RS = Nothing
    End Select
    codigo = ""
    Descripcion = ""
    
Exit Sub
hand:
ErrorHandler err, "BUSCA ORIGEN"
End Sub


Public Function GetRecordset1(ByVal Connect As String, ByVal SQL As String) As Object 'ADOR.Recordset
  On Error GoTo ehGetRecordset
  Dim objADORs As Object ' CreateObject("ADODB.Recordset") '
  Dim objAdoCn As Object ' New ADODB.Connection '
  
 
  Set objADORs = CreateObject("ADODB.Recordset") 'CreateObject("ADODB.Recordset") '
  Set objAdoCn = CreateObject("ADODB.Connection") ' New ADODB.Connection  '
  objAdoCn.CursorLocation = 3
  objAdoCn.Open Connect
  objAdoCn.CommandTimeout = 900
  objADORs.Open SQL, objAdoCn, 3, 4 ', 4  'adOpenStatic= 3 ,  adLockBatchOptimistic = 4  (orignal)  'cambio desde 24/07/2000 ' 1 adLockReadOnly , ' 4 adCmdStoredProc
  Set GetRecordset1 = objADORs
  Set GetRecordset1.ActiveConnection = objAdoCn
  Set objADORs.ActiveConnection = Nothing
  objAdoCn.Close
  Set objAdoCn = Nothing
 
Exit Function
ehGetRecordset:
  err.Raise err.Number, err.Source, err.Description
  MsgBox err.Description
End Function





Private Sub CargaValoresDefault()
On Error GoTo errx
Dim sSQL As String
Dim RS As ADODB.Recordset

sSQL = "CN_VENTAS_OBTIENE_Presentacion_Default '$','$'"
sSQL = VBsprintf(sSQL, Txt_Origen, txtCod_TipDoc2.Text)

Set RS = GetDataSet(cCONNECT, sSQL)

If Not RS Is Nothing Then
    If Not RS.EOF Then
        txtSer_Desde.Text = FixNulos(RS!SERIE_DOCUM, vbString)
        txtSer_Hasta.Text = FixNulos(RS!SERIE_DOCUM, vbString)
        txtNum_Desde.Text = FixNulos(RS!DESDE_DOCUM, vbString)
        txtNum_Hasta.Text = FixNulos(RS!HASTA_DOCUM, vbString)
        RS.Close
    End If
End If
Set RS = Nothing

Exit Sub
errx:
    errores err.Number
End Sub



Public Sub BUSCA_ORIGEN2(Tipo As Integer)
On Error GoTo hand
    Select Case Tipo
        Case 1:
                    strSQL = "SELECT Des_Origen as 'Descripción' FROM  cn_origen WHERE Origen = '" & Trim(Me.txtOrigen.Text) & "'"
                    Me.txtDescripcion.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
                    
                    
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim RS As Object
                    Set RS = CreateObject("ADODB.Recordset")
                    Set oTipo.oParent = Me
                    
                    
                    If Tipo = 2 Then
                        oTipo.SQuery = "SELECT Origen as 'Código', Des_Origen as 'Descripción' FROM cn_origen WHERE Des_Origen LIKE '%" & Trim(Me.txtOrigen) & "%' ORDER BY Des_Origen"
                    Else
                        oTipo.SQuery = "SELECT ORIGEN as 'Código', Des_Origen AS 'Descripción' FROM Cn_Origen ORDER BY Des_Origen"
                    End If
                    
                    
                    oTipo.CARGAR_DATOS
                    oTipo.Show 1
                    If codigo <> "" Then
                        Me.txtOrigen = Trim(codigo)
                        Me.txtDescripcion = Trim(Descripcion)
                        
                    End If
                    Set oTipo = Nothing
                    Set RS = Nothing
    End Select
    codigo = ""
    Descripcion = ""
    
Exit Sub
hand:
ErrorHandler err, "BUSCA ORIGEN"
End Sub


Private Sub CambiarStatus(sSQL As String)
On Error GoTo errx
Dim RS As ADODB.Recordset
Dim vResp As Variant

sSQL = sSQL & "'$','$','$'"

vResp = MsgBox("Desea Cambiar de Estado al Documento Indicado : " & txtDescrip_Tipdoc_DB & " ? ", vbOKCancel + vbQuestion, "Confirmación")

If vResp <> vbOK Then
    Exit Sub
End If


sSQL = VBsprintf(sSQL, SNum_Corre, vusu, ComputerName())

ExecuteCommandSQL cCONNECT, sSQL
Mensaje kMESSAGE_INF_PROCESS_SATISFACTO

fraDrawBack.Visible = False
Buscar

Exit Sub
errx:
    errores err.Number
End Sub
