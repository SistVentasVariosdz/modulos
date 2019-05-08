VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "numbox.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmShowTG_Embarque 
   Caption         =   "Embarques"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12555
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   12555
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDua1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "DUA"
      Height          =   1755
      Left            =   1800
      TabIndex        =   55
      Top             =   4320
      Visible         =   0   'False
      Width           =   7710
      Begin NumBoxProject.NumBox txtEntregaDrauBack1 
         Height          =   330
         Left            =   5850
         TabIndex        =   57
         Top             =   360
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         TypeVal         =   3
         Mask            =   "99/99/9999"
         Formato         =   "dd/MM/yyyy"
         AllowedMask     =   0
         MaskLen         =   20
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
      Begin NumBoxProject.NumBox txtFec_RecepcionDUA1 
         Height          =   285
         Left            =   1995
         TabIndex        =   56
         Top             =   420
         Width           =   1305
         _ExtentX        =   2302
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
      Begin FunctionsButtons.FunctButt FunctButt6 
         Height          =   516
         Left            =   2640
         TabIndex        =   58
         Top             =   1008
         Width           =   2352
         _ExtentX        =   4154
         _ExtentY        =   900
         Custom          =   $"frmShowTG_Embarque.frx":0000
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   0
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fecha de Recepción "
         Height          =   330
         Left            =   315
         TabIndex        =   60
         Tag             =   "FEC_RECEPCIONDUA"
         Top             =   450
         Width           =   1500
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fecha Entrega Tramite Drau Back"
         Height          =   405
         Left            =   4290
         TabIndex        =   59
         Top             =   390
         Width           =   1350
      End
   End
   Begin VB.Frame fraPenalidad 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Penalidades del Embarque"
      Height          =   2130
      Left            =   3960
      TabIndex        =   49
      Top             =   2220
      Visible         =   0   'False
      Width           =   3555
      Begin VB.TextBox txtImp_Flete_Aereo 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1560
         TabIndex        =   52
         Tag             =   "SET"
         Top             =   690
         Width           =   1260
      End
      Begin VB.CheckBox chkIncPenalidad 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Genera Penalidad"
         Height          =   240
         Left            =   240
         TabIndex        =   50
         Top             =   315
         Width           =   1995
      End
      Begin FunctionsButtons.FunctButt FunctButt5 
         Height          =   510
         Left            =   555
         TabIndex        =   53
         Top             =   1230
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   900
         Custom          =   $"frmShowTG_Embarque.frx":0097
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   0
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Imp. Flete Aereo"
         Height          =   270
         Left            =   225
         TabIndex        =   51
         Tag             =   "NUM_DUA"
         Top             =   735
         Width           =   1500
      End
   End
   Begin VB.Frame fraDua 
      BackColor       =   &H00C0FFFF&
      Caption         =   "DUA"
      Height          =   3045
      Left            =   1890
      TabIndex        =   40
      Top             =   4725
      Visible         =   0   'False
      Width           =   7710
      Begin VB.TextBox txtNum_BL 
         Height          =   300
         Left            =   5600
         TabIndex        =   38
         Top             =   1500
         Width           =   1980
      End
      Begin VB.TextBox txtDolares 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   5610
         TabIndex        =   37
         Top             =   1080
         Width           =   1305
      End
      Begin MSMask.MaskEdBox txtNum_Dua 
         Height          =   285
         Left            =   1695
         TabIndex        =   32
         Top             =   240
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   18
         Mask            =   "###-####-##-######"
         PromptChar      =   "_"
      End
      Begin NumBoxProject.NumBox txtEntregaDrauBack 
         Height          =   330
         Left            =   1710
         TabIndex        =   34
         Top             =   1050
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         TypeVal         =   3
         Mask            =   "99/99/9999"
         Formato         =   "dd/MM/yyyy"
         AllowedMask     =   0
         MaskLen         =   20
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
      Begin NumBoxProject.NumBox txtFec_NumeracionDua 
         Height          =   285
         Left            =   1710
         TabIndex        =   33
         Top             =   615
         Width           =   1305
         _ExtentX        =   2302
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
      Begin NumBoxProject.NumBox txtFec_RecepcionDUA 
         Height          =   285
         Left            =   5610
         TabIndex        =   36
         Top             =   615
         Width           =   1305
         _ExtentX        =   2302
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
      Begin NumBoxProject.NumBox txtFec_EmbarqueReal 
         Height          =   285
         Left            =   5610
         TabIndex        =   35
         Top             =   225
         Width           =   1305
         _ExtentX        =   2302
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
      Begin FunctionsButtons.FunctButt FunctButt4 
         Height          =   510
         Left            =   2625
         TabIndex        =   39
         Top             =   2160
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   900
         Custom          =   $"frmShowTG_Embarque.frx":012E
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   0
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Número Bill of Landing"
         Height          =   480
         Left            =   3900
         TabIndex        =   61
         Tag             =   "FEC_EMBARQUEREAL"
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Importe FOB $ DUA"
         Height          =   255
         Left            =   3960
         TabIndex        =   54
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fecha Entrega Tramite Drau Back"
         Height          =   405
         Left            =   150
         TabIndex        =   48
         Top             =   1080
         Width           =   1350
      End
      Begin VB.Label lblFec_RecepcionDua 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fecha de Recepción "
         Height          =   330
         Left            =   3930
         TabIndex        =   44
         Tag             =   "FEC_RECEPCIONDUA"
         Top             =   645
         Width           =   1500
      End
      Begin VB.Label lblNum_Dua 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Numero de Dua"
         Height          =   270
         Left            =   150
         TabIndex        =   43
         Tag             =   "NUM_DUA"
         Top             =   270
         Width           =   1500
      End
      Begin VB.Label lblFec_NumeracionDua 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fecha de Numeración"
         Height          =   420
         Left            =   135
         TabIndex        =   42
         Tag             =   "FEC_NUMERACIONDUA"
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label lblFec_EmbarqueReal 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fec Real de Embarque"
         Height          =   480
         Left            =   3930
         TabIndex        =   41
         Tag             =   "FEC_EMBARQUEREAL"
         Top             =   195
         Width           =   1065
      End
   End
   Begin VB.Frame fraAgenteCarga 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Envio Instrucciones al Agente de Carga"
      Height          =   2295
      Left            =   1785
      TabIndex        =   25
      Top             =   3615
      Visible         =   0   'False
      Width           =   7725
      Begin VB.TextBox txtObs_EnvioInstruccionesalAgenteCarga 
         Height          =   825
         Left            =   1710
         TabIndex        =   28
         Tag             =   "SET"
         Top             =   675
         Width           =   5880
      End
      Begin NumBoxProject.NumBox txtFec_EnvioInstruccionesalAgenteCarga 
         Height          =   285
         Left            =   1710
         TabIndex        =   26
         Top             =   270
         Width           =   1425
         _ExtentX        =   2514
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
      Begin FunctionsButtons.FunctButt FunctButt3 
         Height          =   510
         Left            =   2685
         TabIndex        =   29
         Top             =   1620
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   900
         Custom          =   $"frmShowTG_Embarque.frx":01C5
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   0
      End
      Begin VB.Label lblFec_EnvioInstruccionesalAgenteCarga 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fecha Envio"
         Height          =   405
         Left            =   165
         TabIndex        =   30
         Tag             =   "FEC_ENVIOINSTRUCCIONESALAGENTECARGA"
         Top             =   255
         Width           =   1500
      End
      Begin VB.Label lblObs_EnvioInstruccionesalAgenteCarga 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Observaciones"
         Height          =   480
         Left            =   165
         TabIndex        =   27
         Tag             =   "OBS_ENVIOINSTRUCCIONESALAGENTECARGA"
         Top             =   660
         Width           =   1500
      End
   End
   Begin VB.Frame fraAgenteAduana 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Envío Instrucciones Agente de Aduana"
      Height          =   3150
      Left            =   1680
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   7725
      Begin VB.CommandButton cmdAlmacenAduana 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7125
         TabIndex        =   45
         Top             =   1635
         Width           =   420
      End
      Begin VB.TextBox txtNom_AlmacenAduana 
         Height          =   300
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   21
         Tag             =   "SET"
         Top             =   1650
         Width           =   4815
      End
      Begin VB.TextBox txtCod_AlmacenAduana 
         Height          =   300
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   20
         Tag             =   "SET"
         Top             =   1650
         Width           =   450
      End
      Begin VB.TextBox txtObs_EnvioInstruccionesalAgenteAduanas 
         Height          =   915
         Left            =   1710
         TabIndex        =   19
         Tag             =   "SET"
         Top             =   630
         Width           =   5865
      End
      Begin NumBoxProject.NumBox txtFec_EnvioInstruccionesalAgenteAduana 
         Height          =   285
         Left            =   1710
         TabIndex        =   17
         Top             =   285
         Width           =   1425
         _ExtentX        =   2514
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
      Begin FunctionsButtons.FunctButt FunctOKCancel 
         Height          =   510
         Left            =   2685
         TabIndex        =   23
         Top             =   2460
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   900
         Custom          =   $"frmShowTG_Embarque.frx":025C
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   0
      End
      Begin NumBoxProject.NumBox txtFec_EnvioFacturaalAgenteAduanas 
         Height          =   285
         Left            =   1710
         TabIndex        =   22
         Top             =   2025
         Width           =   1425
         _ExtentX        =   2514
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
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Almacén de Aduana"
         Height          =   330
         Left            =   165
         TabIndex        =   46
         Tag             =   "COD_AGENTEADUANA"
         Top             =   1650
         Width           =   1500
      End
      Begin VB.Label lblFec_EnvioFacturaalAgenteAduanas 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fecha Envio Factura"
         Height          =   345
         Left            =   150
         TabIndex        =   31
         Tag             =   "FEC_ENVIOFACTURAALAGENTEADUANAS"
         Top             =   2025
         Width           =   1500
      End
      Begin VB.Label lblFec_EnvioInstruccionesalAgenteAduanas 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fecha Envio"
         Height          =   345
         Left            =   165
         TabIndex        =   24
         Tag             =   "FEC_ENVIOINSTRUCCIONESALAGENTEADUANAS"
         Top             =   255
         Width           =   1500
      End
      Begin VB.Label lblObs_EnvioInstruccionesalAgenteAduanas 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Observaciones"
         Height          =   480
         Left            =   120
         TabIndex        =   18
         Tag             =   "OBS_ENVIOINSTRUCCIONESALAGENTEADUANAS"
         Top             =   720
         Width           =   1500
      End
   End
   Begin VB.OptionButton optEstado 
      Caption         =   "Estado"
      Height          =   240
      Left            =   105
      TabIndex        =   13
      Top             =   1215
      Width           =   1290
   End
   Begin VB.OptionButton optReferencia 
      Caption         =   "Referencia"
      Height          =   240
      Left            =   105
      TabIndex        =   12
      Top             =   810
      Width           =   1290
   End
   Begin VB.OptionButton optPeriodo 
      Caption         =   "Período/Cliente"
      Height          =   240
      Left            =   105
      TabIndex        =   11
      Top             =   465
      Width           =   1515
   End
   Begin VB.TextBox txtAno 
      Height          =   300
      Left            =   1785
      TabIndex        =   1
      Top             =   435
      Width           =   675
   End
   Begin VB.TextBox txtMes 
      Height          =   300
      Left            =   2490
      TabIndex        =   2
      Top             =   435
      Width           =   390
   End
   Begin VB.TextBox txtDes_Status 
      Height          =   300
      Left            =   2415
      TabIndex        =   10
      Tag             =   "SET"
      Top             =   1130
      Width           =   4215
   End
   Begin VB.TextBox txtNom_cliente 
      Height          =   300
      Left            =   4260
      TabIndex        =   9
      Tag             =   "SET"
      Top             =   420
      Width           =   4215
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   435
      Left            =   11040
      TabIndex        =   6
      Top             =   240
      Width           =   1170
   End
   Begin VB.TextBox txtFlg_Status 
      Height          =   300
      Left            =   1770
      TabIndex        =   5
      Top             =   1130
      Width           =   585
   End
   Begin VB.TextBox txtRef_Embarque 
      Height          =   300
      Left            =   1770
      TabIndex        =   4
      Top             =   795
      Width           =   2025
   End
   Begin VB.TextBox txtAbr_Cliente 
      Height          =   300
      Left            =   3645
      TabIndex        =   3
      Top             =   420
      Width           =   555
   End
   Begin VB.TextBox txtNum_Embarque 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1770
      TabIndex        =   0
      Top             =   50
      Width           =   900
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   8055
      Left            =   10920
      TabIndex        =   8
      Top             =   840
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   14208
      Custom          =   $"frmShowTG_Embarque.frx":02F3
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1320
      ControlHeigth   =   490
      ControlSeparator=   90
   End
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   510
      Left            =   -120
      TabIndex        =   15
      Top             =   8280
      Width           =   10890
      _ExtentX        =   19050
      _ExtentY        =   900
      Custom          =   $"frmShowTG_Embarque.frx":084C
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1700
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   6528
      Left            =   156
      TabIndex        =   7
      Top             =   1452
      Width           =   9828
      _ExtentX        =   17330
      _ExtentY        =   11509
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmShowTG_Embarque.frx":0AA6
      Column(2)       =   "frmShowTG_Embarque.frx":0B6E
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmShowTG_Embarque.frx":0C12
      FormatStyle(2)  =   "frmShowTG_Embarque.frx":0D4A
      FormatStyle(3)  =   "frmShowTG_Embarque.frx":0DFA
      FormatStyle(4)  =   "frmShowTG_Embarque.frx":0EAE
      FormatStyle(5)  =   "frmShowTG_Embarque.frx":0F86
      FormatStyle(6)  =   "frmShowTG_Embarque.frx":103E
      ImageCount      =   0
      PrinterProperties=   "frmShowTG_Embarque.frx":111E
   End
   Begin VB.OptionButton optNumero 
      Caption         =   "Numero"
      Height          =   315
      Left            =   135
      TabIndex        =   47
      Top             =   75
      Value           =   -1  'True
      Width           =   1365
   End
   Begin VB.Label Label2 
      Caption         =   "Cliente"
      Height          =   330
      Left            =   2955
      TabIndex        =   14
      Tag             =   "COD_TIPANEX"
      Top             =   465
      Width           =   615
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   9480
      Top             =   480
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmShowTG_Embarque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public codigo As String
Public Descripcion As String
Public TipoAdd As String

Public oParent As Object
Dim sTipo As String
Dim lTemp_Embarque As String
Dim strSQL As String

Private Sub cmdAlmacenAduana_Click()
frmMantAlmacenAduana.Show 1
End Sub

Private Sub cmdBuscar_Click()
    BUSCAR
End Sub


Private Sub Form_Load()
Dim sSeguridad  As String
    sSeguridad = get_botones1(Me, vper, vemp, Me.Name)

    Me.FunctButt1.FunctionsUser = sSeguridad
    Me.FunctButt2.FunctionsUser = sSeguridad
    'Me.FunctButt1.FunctionsUser = "ADICIONAR/MODIFICAR/ELIMINAR/CONSULTAR/CAMBIODEESTADO/IMPRIMIR_AG_ADUANA/IMPRIMIR_AG_CARGA/ENVIOAGENTEADUANA/ENVIOAGENTECARGA/RECEPCIONDUA/IMPRIMIRLISTADO/PENALIDADES/FECHAS"
    sTipo = "1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub





Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo hand
Dim strSQL As String
    Select Case ActionName
        Case "ADICIONAR"
            Load frmTG_Embarque
            Set frmTG_Embarque.oParent = Me
            frmTG_Embarque.Saccion = "I"
            frmTG_Embarque.Show vbModal
            Set frmTG_Embarque = Nothing
        Case "MODIFICAR"
            If GridEX1.RowCount = 0 Then Exit Sub
            Load frmTG_Embarque
            CargarData
            Set frmTG_Embarque.oParent = Me
            frmTG_Embarque.Saccion = "U"
            frmTG_Embarque.Show vbModal
            Set frmTG_Embarque = Nothing
        Case "ELIMINAR"
            If GridEX1.RowCount = 0 Then Exit Sub
            Load frmTG_Embarque
            CargarData
            Set frmTG_Embarque.oParent = Me
            frmTG_Embarque.Saccion = "D"
            frmTG_Embarque.Show vbModal
            Set frmTG_Embarque = Nothing
        Case "DETALLE"
            If GridEX1.RowCount = 0 Then Exit Sub
            If GridEX1.Value(GridEX1.Columns("cod_origen").Index) = "7" Then
                Load frmShowTG_Embarque_Detalle
                frmShowTG_Embarque_Detalle.lNum_Embarque = GridEX1.Value(GridEX1.Columns("NUM_EMBARQUE").Index)
                frmShowTG_Embarque_Detalle.BUSCAR
                frmShowTG_Embarque_Detalle.Show vbModal
                Set frmShowTG_Embarque_Detalle = Nothing
            End If
        Case "DETALLETELAS"
            If GridEX1.RowCount = 0 Then Exit Sub
            If GridEX1.Value(GridEX1.Columns("cod_origen").Index) = "3" Then
                Load frmShowTG_Embarque_DetalleTelas
                frmShowTG_Embarque_DetalleTelas.Caption = "Detalle Embarque Telas Num. Embarque : " & GridEX1.Value(GridEX1.Columns("NUM_EMBARQUE").Index)
                frmShowTG_Embarque_DetalleTelas.lNum_Embarque = GridEX1.Value(GridEX1.Columns("NUM_EMBARQUE").Index)
                frmShowTG_Embarque_DetalleTelas.BUSCAR
                frmShowTG_Embarque_DetalleTelas.Show vbModal
            End If
            Set frmShowTG_Embarque_DetalleTelas = Nothing
            
        Case "IMPRIMIR_AG_ADUANA"
            If GridEX1.RowCount = 0 Then Exit Sub
            Imprimir_Ag_Aduana GridEX1.Value(GridEX1.Columns("Num_Embarque").Index)
        Case "IMPRIMIR_AG_CARGA"
            If GridEX1.RowCount = 0 Then Exit Sub
            Imprimir_Ag_Carga GridEX1.Value(GridEX1.Columns("Num_Embarque").Index)
        Case "IMPRIMIRLISTADO"
            If GridEX1.RowCount = 0 Then Exit Sub
            Imprimir_Listado GridEX1.ADORecordset
        Case "CAMBIODEESTADO"
            CAMBIODEESTADO
            
        Case "IMPR_AG_CARGATELAS"
            If GridEX1.RowCount = 0 Then Exit Sub
            Imprimir_Ag_CargaTelas GridEX1.Value(GridEX1.Columns("Num_Embarque").Index)
        
        Case "IMPR_AG_ADUANATELAS"
            If GridEX1.RowCount = 0 Then Exit Sub
            Imprimir_Ag_AduanaTelas GridEX1.Value(GridEX1.Columns("Num_Embarque").Index)
               
        Case "HOJCOSEXPO"
        If Me.GridEX1.RowCount = 0 Then Exit Sub
        
        
      Call ImprimirHojaCostoExpo
        
        
        Case "SALIR"
            Unload Me
    End Select
Exit Sub
hand:
    errores err.Number
End Sub


 Public Sub ImprimirHojaCostoExpo()
 On Error GoTo ErrorImpresion
 Dim oo As Object
 
Dim Adors1 As New ADODB.Recordset
Dim rutaLogo As String
rutaLogo = DevuelveCampo("select ruta_logo=isNUll(ruta_logo,'') from seguridad..seg_empresas where cod_empresa='" & vemp & "'", cCONNECT)
Dim vnum_embarque As String
vnum_embarque = GridEX1.Value(GridEX1.Columns("Num_Embarque").Index)

Set oo = CreateObject("Excel.Application")
    oo.workbooks.Open vRuta & "\RptHojaCostoExportacion.XLT"
    oo.Visible = True
    oo.displayalerts = False
    oo.run "Reporte", rutaLogo, vnum_embarque, cCONNECT
Set oo = Nothing

Exit Sub
ErrorImpresion:

   Screen.MousePointer = vbNormal
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte" & err.Description, vbCritical, "Impresion"
 End Sub





Public Function BUSCAR() As Boolean
On Error GoTo errores
Dim ssql As String
Dim vBookmark As Variant

ssql = "TG_Embarques_Muestra '$','$','$','$','$','$','$'"
ssql = VBsprintf(ssql, sTipo, txtNum_Embarque, txtRef_Embarque, txtAno, txtMes, txtAbr_Cliente.Tag, txtFlg_Status)

  
vBookmark = GridEX1.Row
GridEX1.ClearFields

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(ssql, cCONNECT)

GridEX1.Row = vBookmark


GridEX1.ContinuousScroll = True

GridEX1.FrozenColumns = 3
Exit Function

errores:
    errores err.Number
End Function


Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ENVIOAGENTEADUANA"
            If GridEX1.RowCount = 0 Then Exit Sub
            LoadAgenteAduana
        Case "ENVIOAGENTECARGA"
            If GridEX1.RowCount = 0 Then Exit Sub
            LoadAgenteCarga
        Case "RECEPCIONDUA"
            If GridEX1.RowCount = 0 Then Exit Sub
            LoadRecepcionDUA
        Case "PENALIDADES"
            If GridEX1.RowCount = 0 Then Exit Sub
            LoadPenalidades
        Case "FECHAS"
            If GridEX1.RowCount = 0 Then Exit Sub
            LoadRecepcionDUA1
      Case "GASTOS"
       If GridEX1.RowCount = 0 Then Exit Sub
       frmGastosAsociados.vnum_emb = GridEX1.Value(GridEX1.Columns("Num_embarque").Index)
        frmGastosAsociados.BUSCAR
        frmGastosAsociados.GridEX1.Columns("Des_Anexo").Width = 2300
       frmGastosAsociados.Show vbModal
        
    End Select
End Sub

Private Sub FunctButt4_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            GrabarDUA
        Case "CANCELAR"
            Me.fraDua.Visible = False
    End Select
End Sub

Private Sub FunctButt5_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            GrabarPenalidad
        Case "CANCELAR"
            Me.fraPenalidad.Visible = False
    End Select
End Sub

Private Sub FunctButt6_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            GrabarDUA1
        Case "CANCELAR"
            Me.fraDua1.Visible = False
    End Select
End Sub

Private Sub FunctOKCancel_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            GrabarAgenteAduana
        Case "CANCELAR"
            Me.fraAgenteAduana.Visible = False
    End Select
End Sub

Private Sub FunctButt3_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            GrabarAgenteCarga
        Case "CANCELAR"
            Me.fraAgenteCarga.Visible = False
    End Select
End Sub

Private Sub GridEX1_DblClick()
    Dim I As Integer
    For I = 1 To GridEX1.Columns.Count
        Debug.Print GridEX1.Name & ".Columns( & Chr(34) & GridEX1.Columns(i).Key & Chr(34) & ).width =  & CStr(GridEX1.Columns(i).Width)"
    Next
End Sub



Private Sub NumBox1_Change()

End Sub

Private Sub txtfec_BL_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub optEstado_Click()
    sTipo = "4"
    txtFlg_Status.SetFocus
End Sub

Private Sub optNumero_Click()
    sTipo = "1"
    txtNum_Embarque.SetFocus
End Sub

Private Sub optPeriodo_Click()
    sTipo = "2"
    txtAno.SetFocus
End Sub

Private Sub optReferencia_Click()
On Error GoTo errx
    sTipo = "3"
        txtRef_Embarque.SetFocus
    Exit Sub
errx:
    If err.Number = 5 Then
        Resume Next
    End If
End Sub

Private Sub CargarData()
    frmTG_Embarque.txtNum_Embarque = GridEX1.Value(GridEX1.Columns("NUM_EMBARQUE").Index)
    frmTG_Embarque.txtCod_Origen = GridEX1.Value(GridEX1.Columns("Cod_Origen").Index)
    frmTG_Embarque.txtDes_Origen = GridEX1.Value(GridEX1.Columns("des_Origen").Index)
    frmTG_Embarque.txtCod_TipAnex = GridEX1.Value(GridEX1.Columns("Cod_TipAnex").Index)
    frmTG_Embarque.txtCod_Anxo = GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index)
    frmTG_Embarque.txtDes_Anexo = GridEX1.Value(GridEX1.Columns("des_Anexo").Index)
    frmTG_Embarque.txtAbr_Cliente = GridEX1.Value(GridEX1.Columns("Abr_Cliente").Index)
    frmTG_Embarque.txtAbr_Cliente.Tag = GridEX1.Value(GridEX1.Columns("cod_Cliente").Index)
    frmTG_Embarque.txtNom_cliente = GridEX1.Value(GridEX1.Columns("nom_cliente").Index)
    
    frmTG_Embarque.txtCod_AgenteCarga = GridEX1.Value(GridEX1.Columns("Cod_AgenteCarga").Index)
    frmTG_Embarque.txtDes_AgenteCarga = GridEX1.Value(GridEX1.Columns("des_AgenteCarga").Index)
    frmTG_Embarque.txtCod_AgenteAduana = GridEX1.Value(GridEX1.Columns("cod_AgenteAduana").Index)
    frmTG_Embarque.txtNom_AgenteAduana = GridEX1.Value(GridEX1.Columns("nom_AgenteAduana").Index)
    frmTG_Embarque.txtCod_Ejecutivo_AgenteCarga = GridEX1.Value(GridEX1.Columns("Cod_Ejecutivo_AgenteCarga").Index)
    frmTG_Embarque.txtNom_Ejecutivo = GridEX1.Value(GridEX1.Columns("Nom_Ejecutivo").Index)
    frmTG_Embarque.txtRef_Embarque = GridEX1.Value(GridEX1.Columns("Ref_Embarque").Index)
    frmTG_Embarque.txtTip_Embarque = GridEX1.Value(GridEX1.Columns("Tip_Embarque").Index)
    frmTG_Embarque.txtDes_TipEmbarque = GridEX1.Value(GridEX1.Columns("Des_TipEmbarque").Index)
    frmTG_Embarque.txtObs_Embarque = GridEX1.Value(GridEX1.Columns("Obs_Embarque").Index)
    frmTG_Embarque.txtCod_Embarque = GridEX1.Value(GridEX1.Columns("Cod_Embarque").Index)
    frmTG_Embarque.txtDes_Embarque = GridEX1.Value(GridEX1.Columns("des_Embarque").Index)
    frmTG_Embarque.txtNom_Embarque = GridEX1.Value(GridEX1.Columns("nom_Embarque").Index)
    frmTG_Embarque.txtCod_Flete = GridEX1.Value(GridEX1.Columns("Cod_Flete").Index)
    frmTG_Embarque.txtDes_Flete_Ingles = GridEX1.Value(GridEX1.Columns("Des_Flete_ingles").Index)
    frmTG_Embarque.txtFlg_Status = GridEX1.Value(GridEX1.Columns("Flg_Status").Index)
    frmTG_Embarque.txtDes_Estado = GridEX1.Value(GridEX1.Columns("Des_status").Index)
    frmTG_Embarque.txtFec_EnvioInstruccionesalAgenteAduana.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Fec_EnvioInstruccionesalAgenteAduanas").Index), vbString)
    frmTG_Embarque.txtObs_EnvioInstruccionesalAgenteAduanas = GridEX1.Value(GridEX1.Columns("Obs_EnvioInstruccionesalAgenteAduanas").Index)
    frmTG_Embarque.txtFec_EnvioFacturaalAgenteAduanas.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Fec_EnvioFacturaalAgenteAduanas").Index), vbString)
    frmTG_Embarque.txtFec_RecepcionDUA.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Fec_RecepcionDua").Index), vbString)
    frmTG_Embarque.txtNum_Dua = GridEX1.Value(GridEX1.Columns("Num_Dua").Index)
    frmTG_Embarque.txtFec_NumeracionDua.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Fec_NumeracionDua").Index), vbString)
    
    frmTG_Embarque.txtEntregaDrauBack.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Fec_Entrega_Contabilidad_Tramites_DrawBack").Index), vbString)

      
    frmTG_Embarque.txtFec_EmbarqueReal.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Fec_EmbarqueReal").Index), vbString)
    frmTG_Embarque.txtFec_EnvioInstruccionesalAgenteCarga.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Fec_EnvioInstruccionesalAgenteCarga").Index), vbString)
    frmTG_Embarque.txtObs_EnvioInstruccionesalAgenteCarga = GridEX1.Value(GridEX1.Columns("Obs_EnvioInstruccionesalAgenteCarga").Index)
    frmTG_Embarque.txtAno = GridEX1.Value(GridEX1.Columns("Ano").Index)
    frmTG_Embarque.txtMes = GridEX1.Value(GridEX1.Columns("Mes").Index)
    frmTG_Embarque.txtCod_Termino_Venta = GridEX1.Value(GridEX1.Columns("cod_termino_venta").Index)
    frmTG_Embarque.txtDes_Termino_Venta = GridEX1.Value(GridEX1.Columns("des_termino_venta").Index)
    
    frmTG_Embarque.txtTip_Embarque = GridEX1.Value(GridEX1.Columns("TIP_EMBARQUE").Index)
    frmTG_Embarque.txtDes_TipEmbarque = GridEX1.Value(GridEX1.Columns("DES_TIPEMBARQUE").Index)
    frmTG_Embarque.txtCod_AlmacenAduana = GridEX1.Value(GridEX1.Columns("COD_ALMACENADUANA").Index)
    frmTG_Embarque.txtNom_AlmacenAduana = GridEX1.Value(GridEX1.Columns("NOM_ALMACENADUANA").Index)
    
    frmTG_Embarque.txtCod_Origen.Enabled = False
    frmTG_Embarque.txtDes_Origen.Enabled = False
    frmTG_Embarque.txtAno.Enabled = False
    frmTG_Embarque.txtAbr_Cliente.Enabled = False
    frmTG_Embarque.txtNom_cliente.Enabled = False
End Sub

Private Sub CAMBIODEESTADO()
On Error GoTo errores
Dim ssql As String
Dim vResp As Variant

vResp = MsgBox("Confirma cambio de Estado ? ", vbYesNo, "CAMBIO DE ESTADO EMBARQUES")

If vResp = vbNo Then Exit Sub

ssql = "TG_EMBARQUE_CAMBIOESTADO '$'"
ssql = VBsprintf(ssql, GridEX1.Value(GridEX1.Columns("NUM_EMBARQUE").Index))
  
ExecuteCommandSQL cCONNECT, ssql
Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
BUSCAR

Exit Sub

errores:
    errores err.Number
End Sub

Private Sub GrabarAgenteAduana()
On Error GoTo errores
Dim ssql As String

ssql = "TG_EMBARQUE_ENVIOAGENTE_ADUANA '$','$','$','$'"
ssql = VBsprintf(ssql, lTemp_Embarque, txtFec_EnvioInstruccionesalAgenteAduana.Text, txtObs_EnvioInstruccionesalAgenteAduanas, txtCod_AlmacenAduana)

If txtFec_EnvioFacturaalAgenteAduanas.Text = "" Then
    ssql = ssql & ",NULL"
Else
    ssql = ssql & ",'" & txtFec_EnvioFacturaalAgenteAduanas.Text & "'"
End If

ExecuteCommandSQL cCONNECT, ssql

Me.fraAgenteAduana.Visible = False
BUSCAR

Exit Sub

errores:
    errores err.Number
End Sub

Private Sub GrabarAgenteCarga()
On Error GoTo errores
Dim ssql As String

ssql = "TG_EMBARQUE_ENVIOAGENTE_CARGA '$','$','$'"
ssql = VBsprintf(ssql, lTemp_Embarque, txtFec_EnvioInstruccionesalAgenteCarga.Text, txtObs_EnvioInstruccionesalAgenteCarga)
  
ExecuteCommandSQL cCONNECT, ssql

Me.fraAgenteCarga.Visible = False
BUSCAR

Exit Sub

errores:
    errores err.Number
End Sub


Private Sub GrabarPenalidad()
On Error GoTo errores
Dim ssql As String
Dim sFlg_Pendalidad As String

If chkIncPenalidad.Value = "1" Then
    sFlg_Pendalidad = "S"
Else
    sFlg_Pendalidad = "N"
End If

ssql = "TG_EMBARQUE_DATOS_PENALIDAD '$','$','$'"
ssql = VBsprintf(ssql, lTemp_Embarque, sFlg_Pendalidad, txtImp_Flete_Aereo.Text)
  
ExecuteCommandSQL cCONNECT, ssql

Me.fraPenalidad.Visible = False
BUSCAR

Exit Sub

errores:
    errores err.Number
End Sub


Private Sub LoadAgenteAduana()
lTemp_Embarque = GridEX1.Value(GridEX1.Columns("NUM_EMBARQUE").Index)
fraAgenteAduana.Visible = True
txtFec_EnvioInstruccionesalAgenteAduana.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Fec_EnvioInstruccionesalAgenteAduanas").Index), vbString)
txtObs_EnvioInstruccionesalAgenteAduanas = GridEX1.Value(GridEX1.Columns("Obs_EnvioInstruccionesalAgenteAduanas").Index)
txtFec_EnvioFacturaalAgenteAduanas.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Fec_EnvioFacturaalAgenteAduanas").Index), vbString)
txtCod_AlmacenAduana = FixNulos(GridEX1.Value(GridEX1.Columns("Cod_AlmacenAduana").Index), vbString)
txtNom_AlmacenAduana = FixNulos(GridEX1.Value(GridEX1.Columns("Nom_AlmacenAduana").Index), vbString)
txtFec_EnvioInstruccionesalAgenteAduana.SetFocus

'Fec_RecepcionDua                                       Num_Dua                   Fec_NumeracionDua                                      Fec_EmbarqueReal

End Sub

Private Sub LoadAgenteCarga()
lTemp_Embarque = GridEX1.Value(GridEX1.Columns("NUM_EMBARQUE").Index)
fraAgenteCarga.Visible = True
txtFec_EnvioInstruccionesalAgenteCarga.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Fec_EnvioInstruccionesalAgenteCarga").Index), vbString)
txtObs_EnvioInstruccionesalAgenteCarga = GridEX1.Value(GridEX1.Columns("Obs_EnvioInstruccionesalAgenteCarga").Index)
txtFec_EnvioInstruccionesalAgenteCarga.SetFocus
End Sub

Private Sub LoadRecepcionDUA()
lTemp_Embarque = GridEX1.Value(GridEX1.Columns("NUM_EMBARQUE").Index)
fraDua.Visible = True

txtNum_Dua.Mask = ""
txtNum_Dua.Text = ""
txtNum_Dua.Mask = "###-####-##-######"

If RTrim(FixNulos(GridEX1.Value(GridEX1.Columns("NUM_DUA").Index), vbString)) <> "" Then
    txtNum_Dua = RTrim(FixNulos(GridEX1.Value(GridEX1.Columns("NUM_DUA").Index), vbString))
Else
    txtNum_Dua.Mask = ""
    txtNum_Dua.Text = ""
    txtNum_Dua.Mask = "###-####-##-######"
End If
txtFec_NumeracionDua.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Fec_NumeracionDUA").Index), vbString)

txtEntregaDrauBack.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Fec_Entrega_Contabilidad_Tramites_DrawBack").Index), vbString)


txtFec_EmbarqueReal.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Fec_EmbarqueReal").Index), vbString)
txtFec_RecepcionDUA.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Fec_RecepcionDUA").Index), vbString)
txtDolares.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Imp_FOB_Dol_Dua").Index), vbString)

If Val(txtDolares.Text) = 0 Then
    txtDolares.Text = DevuelveCampo("SELECT DBO.CN_VENTAS_PREDETERMINA_IMP_FOB_DOL_DUA('" & lTemp_Embarque & "')", cCONNECT)
End If

txtNum_BL = FixNulos(GridEX1.Value(GridEX1.Columns("Num_Bill_of_Landing").Index), vbString)

txtNum_Dua.SetFocus
End Sub

Private Sub LoadRecepcionDUA1()
'txtFec_RecepcionDUA1
'
'txtEntregaDrauBack1

lTemp_Embarque = GridEX1.Value(GridEX1.Columns("NUM_EMBARQUE").Index)
fraDua1.Visible = True

txtEntregaDrauBack1.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Fec_Entrega_Contabilidad_Tramites_DrawBack").Index), vbString)
txtFec_RecepcionDUA1.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Fec_RecepcionDUA").Index), vbString)
txtFec_RecepcionDUA1.SetFocus
End Sub

Private Sub LoadPenalidades()
lTemp_Embarque = GridEX1.Value(GridEX1.Columns("NUM_EMBARQUE").Index)
fraPenalidad.Visible = True
If GridEX1.Value(GridEX1.Columns("Flg_Penalidad").Index) = "S" Then
    chkIncPenalidad.Value = "1"
    txtImp_Flete_Aereo.Text = GridEX1.Value(GridEX1.Columns("Imp_Flete_Aereo").Index)
Else
    chkIncPenalidad.Value = "0"
    txtImp_Flete_Aereo.Text = 0
End If
chkIncPenalidad.SetFocus
End Sub

Private Sub GrabarDUA1()
On Error GoTo errores
Dim ssql As String

ssql = "TG_EMBARQUE_DATOS_DUA_FECHAS '$','$','$'"
ssql = VBsprintf(ssql, lTemp_Embarque, _
    txtFec_RecepcionDUA1.Text, txtEntregaDrauBack1.Text)


ExecuteCommandSQL cCONNECT, ssql

fraDua1.Visible = False
BUSCAR

Exit Sub

errores:
    errores err.Number
End Sub
Private Sub GrabarDUA()
On Error GoTo errores
Dim ssql As String

If RTrim(txtNum_Dua.Text) <> "" Then
    If Len(RTrim(txtNum_Dua.Text)) <> 18 Then
        Aviso "Formato de Número de DUA incorrecto. Revisar", 2
        Exit Sub
    End If
End If

ssql = "TG_EMBARQUE_DATOS_DUA '$','$','$','$'"
ssql = VBsprintf(ssql, lTemp_Embarque, _
    txtFec_RecepcionDUA.Text, txtNum_Dua.Text, txtFec_NumeracionDua.Text)

If txtFec_EmbarqueReal.Text = "" Then
    ssql = ssql & ",NULL,''"
Else
    ssql = ssql & ",'" & txtFec_EmbarqueReal.Text & "', '" & txtEntregaDrauBack.Text & "'"
    
End If


ssql = ssql & " ,'" & txtDolares.Text & "','"


ssql = ssql & txtNum_BL.Text & "'"


ExecuteCommandSQL cCONNECT, ssql

fraDua.Visible = False
BUSCAR

Exit Sub

errores:
    errores err.Number
End Sub

Private Sub txtAbr_Cliente_GotFocus()
    SelectionText txtAbr_Cliente
End Sub

Private Sub TxtAbr_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaCliente 1
        SendKeys "{TAB}"
    End If
End Sub
Public Sub BuscaCliente(opcion As String)
Dim rstAux As ADODB.Recordset

    strSQL = "SELECT Cod_Cliente, Abr_Cliente, Nom_Cliente FROM TG_CLIENTE WHERE "
    
    txtAbr_Cliente = Trim(txtAbr_Cliente)
    txtNom_cliente = Trim(txtNom_cliente)
    
    Select Case opcion
    Case 1: strSQL = strSQL & "Abr_Cliente LIKE '%" & txtAbr_Cliente & "%'"
    Case 2: strSQL = strSQL & "Nom_Cliente LIKE '%" & txtNom_cliente & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.sQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    
    
    frmBusqGeneral3.gexLista.Columns("Cod_Cliente").Visible = False
    frmBusqGeneral3.gexLista.Columns("Abr_Cliente").Width = 570
    frmBusqGeneral3.gexLista.Columns("Nom_Cliente").Width = 2370
    
    frmBusqGeneral3.gexLista.Columns("Abr_Cliente").Caption = "Abrev."
    frmBusqGeneral3.gexLista.Columns("Nom_Cliente").Caption = "Cliente"
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtAbr_Cliente.Tag = ""
    txtAbr_Cliente = ""
    txtNom_cliente = ""
    If codigo <> "" Then
        
        txtAbr_Cliente = Descripcion
        txtNom_cliente = TipoAdd
        txtAbr_Cliente.Tag = codigo
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
        
    codigo = ""
    Descripcion = ""
End Sub

Private Sub txtAno_GotFocus()
    SelectionText txtAno
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
    optPeriodo.Value = True
    If KeyAscii = vbKeyReturn Then
        If optPeriodo Then
            txtMes.SetFocus
        Else
            txtRef_Embarque.SetFocus
        End If
    End If
End Sub

Private Sub txtCod_AlmacenAduana_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
            If Trim(txtCod_AlmacenAduana.Text) = "" Then
                BUSCA_ALMACENREF (3)
            Else
                BUSCA_ALMACENREF (1)
            End If
            SendKeys "{TAB}"
    End If
End Sub

 

Private Sub txtDolares_GotFocus()
    SelectionText txtDolares
End Sub

Private Sub txtDolares_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtEntregaDrauBack_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtEntregaDrauBack1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtFec_EmbarqueReal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtFec_EnvioFacturaalAgenteAduanas_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtFec_EnvioInstruccionesalAgenteAduana_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtFec_EnvioInstruccionesalAgenteCarga_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtFec_NumeracionDua_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtFec_RecepcionDUA_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtFec_RecepcionDUA1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtFlg_Status_GotFocus()
    SelectionText txtFlg_Status
End Sub

Private Sub txtFlg_Status_KeyPress(KeyAscii As Integer)
    optEstado.Value = True
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaStatus 1
        SendKeys "{TAB}"
    End If
End Sub


Public Sub BuscaStatus(opcion As String)
Dim rstAux As ADODB.Recordset

    strSQL = "SELECT Flg_Status , Des_Status FROM TG_EMBARQUE_STATUS WHERE "
    
    txtFlg_Status = Trim(txtFlg_Status)
    txtDes_Status = Trim(txtDes_Status)
    
    Select Case opcion
    Case 1: strSQL = strSQL & "Flg_Status LIKE '%" & txtFlg_Status & "%'"
    Case 2: strSQL = strSQL & "Des_Status  LIKE '%" & txtDes_Status & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.sQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    

    frmBusqGeneral3.gexLista.Columns("Flg_status").Width = 570
    frmBusqGeneral3.gexLista.Columns("Des_status").Width = 2370
    
    frmBusqGeneral3.gexLista.Columns("Flg_status").Caption = "Status"
    frmBusqGeneral3.gexLista.Columns("Des_status").Caption = "Descripción"
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtFlg_Status = ""
    txtDes_Status = ""
    
    If codigo <> "" Then
        txtFlg_Status = codigo
        txtDes_Status = Descripcion
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
        
    codigo = ""
    Descripcion = ""
End Sub

Private Sub txtImp_Flete_Aereo_GotFocus()
    SelectionText txtImp_Flete_Aereo
End Sub

Private Sub txtMes_GotFocus()
    SelectionText txtMes
End Sub

Private Sub txtMes_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If optPeriodo And txtMes <> "" Then
            BUSCAR
        Else
            txtRef_Embarque.SetFocus
        End If
    End If
End Sub

Private Sub txtNom_AlmacenAduana_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtNom_cliente_GotFocus()
    SelectionText txtNom_cliente
End Sub

Private Sub TxtNom_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        BuscaCliente 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtNum_BL_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtNum_Dua_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtNum_Embarque_GotFocus()
    SelectionText txtNum_Embarque
End Sub

Private Sub txtNum_Embarque_KeyPress(KeyAscii As Integer)
    optNumero.Value = True
    If KeyAscii = vbKeyReturn Then
        If optNumero And txtNum_Embarque <> "" Then
            BUSCAR
        Else
            txtAno.SetFocus
        End If
    End If
End Sub

Private Sub txtObs_EnvioInstruccionesalAgenteAduanas_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtObs_EnvioInstruccionesalAgenteCarga_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtRef_Embarque_GotFocus()
    SelectionText txtRef_Embarque
End Sub

Private Sub txtRef_Embarque_KeyPress(KeyAscii As Integer)
    optReferencia_Click
    If KeyAscii = vbKeyReturn Then
        If optReferencia And txtRef_Embarque <> "" Then
            BUSCAR
        Else
            txtFlg_Status.SetFocus
        End If
    End If
End Sub

Private Sub Imprimir_Ag_Aduana(ByVal lNum_Embarque As Long)
On Error GoTo ErrorImpresion
Dim oo As Object
Dim sFormato_Invoice As String

    Set oo = CreateObject("excel.application")
    If vemp = "01" Then
        oo.workbooks.Open vRuta & "\InstrucAgenteAduanas.XLT"
        ElseIf vemp = "03" Then
        oo.workbooks.Open vRuta & "\InstrucAgenteAduanas_INKA.XLT"
    End If
    oo.Visible = True
    oo.displayalerts = False
    oo.run "reporte", cCONNECT, lNum_Embarque
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub
Private Sub Imprimir_Ag_AduanaTelas(ByVal lNum_Embarque As Long)
On Error GoTo ErrorImpresion
Dim oo As Object
Dim sFormato_Invoice As String

    Set oo = CreateObject("excel.application")
    If vemp = "01" Then
        oo.workbooks.Open vRuta & "\InstrucAgenteAduanasT.XLT"
    ElseIf vemp = "03" Then
        oo.workbooks.Open vRuta & "\InstrucAgenteAduanasT_INKA.XLT"
    End If
    oo.Visible = True
    oo.displayalerts = False
    oo.run "reporte", cCONNECT, lNum_Embarque
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub

Private Sub Imprimir_Ag_Carga(ByVal lNum_Embarque As Long)
On Error GoTo ErrorImpresion
Dim oo As Object
Dim sFormato_Invoice As String

    Set oo = CreateObject("excel.application")
    oo.workbooks.Open vRuta & "\InstrucAgenteCarga.xlt"
    oo.Visible = True
    oo.displayalerts = False
    oo.run "reporte", cCONNECT, lNum_Embarque, vemp
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub
Private Sub Imprimir_Ag_CargaTelas(ByVal lNum_Embarque As Long)
On Error GoTo ErrorImpresion
Dim oo As Object
Dim sFormato_Invoice As String

    Set oo = CreateObject("excel.application")
    oo.workbooks.Open vRuta & "\InstrucAgenteCargaT.xlt"
    oo.Visible = True
    oo.displayalerts = False
    oo.run "reporte", cCONNECT, lNum_Embarque
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub


Private Sub Imprimir_Listado(ByVal rsListado As ADODB.Recordset)
On Error GoTo ErrorImpresion
Dim oo As Object
Dim sRango As String
Dim strSQL As String
Dim sEmpresa As String

    strSQL = "SELECT DES_EMPRESA FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA='" & vemp & "'"
    sEmpresa = DevuelveCampo(strSQL, cCONNECT)


If optNumero Then
    sRango = "Embarque: " & GridEX1.Value(GridEX1.Columns("NUM_EMBARQUE").Index)
End If
If optPeriodo Then
    sRango = "Período: " & GridEX1.Value(GridEX1.Columns("ANO").Index) & IIf(txtMes.Text = "", "", "/" & txtMes.Text) & IIf(txtAbr_Cliente.Text = "", "", " Cliente:" & txtNom_cliente.Text)
End If
If optNumero Then
    sRango = "Ref.Embarque: " & GridEX1.Value(GridEX1.Columns("REF_EMBARQUE").Index)
End If
If optEstado Then
    sRango = "Estado: " & GridEX1.Value(GridEX1.Columns("DES_STATUS").Index)
End If

Set oo = CreateObject("excel.application")
oo.workbooks.Open vRuta & "\EmbarquesExportacion.XLT"
oo.Visible = True
oo.displayalerts = False
oo.run "reporte", cCONNECT, sRango, sTipo, Val(txtNum_Embarque.Text), txtRef_Embarque, txtAno, txtMes, txtAbr_Cliente.Tag, txtFlg_Status, sEmpresa
Set oo = Nothing
Exit Sub

ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub

Public Sub BUSCA_ALMACENREF(tipo As Integer)
    Select Case tipo
        Case 1:
                    strSQL = "SELECT nom_AlmacenAduana FROM CF_ALMACEN_ADUANA WHERE Cod_AlmacenAduana = '" & Trim(Me.txtCod_AlmacenAduana.Text) & "'"
                    Me.txtNom_AlmacenAduana.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
                    
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim Rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If tipo = 2 Then
                        oTipo.sQuery = "SELECT Cod_AlmacenAduana AS 'Código', nom_AlmacenAduana AS 'Descripción' FROM CF_ALMACEN_ADUANA where nom_AlmacenAduana like '%" & Trim(txtNom_AlmacenAduana.Text) & "%' order by Cod_AlmacenAduana"
                    Else
                        oTipo.sQuery = "SELECT Cod_AlmacenAduana AS 'Código', nom_AlmacenAduana AS 'Descripción' FROM CF_ALMACEN_ADUANA order by Cod_AlmacenAduana"
                    End If
                    
                    oTipo.CARGAR_DATOS
                    oTipo.Show 1
                    If codigo <> "" Then
                         Me.txtCod_AlmacenAduana.Text = Trim(codigo)
                         Me.txtNom_AlmacenAduana.Text = Trim(Descripcion)
                         
                         codigo = "": Descripcion = ""
                    End If
                    Set oTipo = Nothing
                    Set Rs = Nothing
    End Select
        
End Sub

