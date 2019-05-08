VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "numbox.ocx"
Begin VB.Form frmTG_Embarque 
   Caption         =   "Detalle de Embarque"
   ClientHeight    =   10350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8355
   LinkTopic       =   "Form2"
   ScaleHeight     =   10350
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCod_AlmacenAduana 
      Height          =   300
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   27
      Tag             =   "SET"
      Top             =   10920
      Width           =   450
   End
   Begin VB.TextBox txtNom_AlmacenAduana 
      Height          =   300
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   28
      Tag             =   "SET"
      Top             =   10920
      Width           =   5535
   End
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
      Left            =   7800
      TabIndex        =   74
      Top             =   10950
      Width           =   420
   End
   Begin VB.TextBox txtDes_Origen 
      Height          =   300
      Left            =   2190
      TabIndex        =   73
      Tag             =   "SET"
      Top             =   405
      Width           =   3570
   End
   Begin VB.TextBox txtCod_Termino_Venta 
      Height          =   315
      Left            =   4380
      TabIndex        =   10
      Top             =   1485
      Width           =   450
   End
   Begin VB.TextBox txtDes_Termino_Venta 
      Height          =   315
      Left            =   4860
      TabIndex        =   11
      Top             =   1485
      Width           =   2835
   End
   Begin VB.CommandButton cmdTipoFlete 
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
      Left            =   7770
      TabIndex        =   64
      Top             =   4470
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton cmdModoEmbarque 
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
      Left            =   7770
      TabIndex        =   63
      Top             =   3765
      Width           =   420
   End
   Begin VB.CommandButton cmdEjecutivoAgCarga 
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
      Left            =   7770
      TabIndex        =   62
      Top             =   2700
      Width           =   420
   End
   Begin VB.CommandButton cmdAgenteCarga 
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
      Left            =   7770
      TabIndex        =   61
      Top             =   1890
      Width           =   420
   End
   Begin VB.CommandButton cmdAgenteAduana 
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
      Left            =   7770
      TabIndex        =   60
      Top             =   2280
      Width           =   420
   End
   Begin VB.TextBox txtMes 
      Height          =   300
      Left            =   7335
      TabIndex        =   3
      Tag             =   "SET"
      Top             =   405
      Width           =   390
   End
   Begin VB.TextBox txtAno 
      Height          =   300
      Left            =   6615
      TabIndex        =   2
      Tag             =   "SET"
      Top             =   405
      Width           =   675
   End
   Begin VB.TextBox txtDes_Estado 
      Enabled         =   0   'False
      Height          =   300
      Left            =   4110
      TabIndex        =   58
      Tag             =   "SET"
      Top             =   60
      Width           =   3615
   End
   Begin VB.TextBox txtFlg_Status 
      Enabled         =   0   'False
      Height          =   300
      Left            =   3600
      TabIndex        =   57
      Tag             =   "SET"
      Top             =   60
      Width           =   435
   End
   Begin VB.TextBox txtNom_Cliente 
      Height          =   300
      Left            =   2265
      TabIndex        =   8
      Tag             =   "SET"
      Top             =   1125
      Width           =   5430
   End
   Begin VB.TextBox txtDes_Anexo 
      Height          =   300
      Left            =   2610
      TabIndex        =   6
      Tag             =   "SET"
      Top             =   765
      Width           =   5100
   End
   Begin VB.Frame fraAgenteCarga 
      Caption         =   "Envio Instrucciones al Agente de Carga"
      Enabled         =   0   'False
      Height          =   1620
      Left            =   60
      TabIndex        =   52
      Top             =   6450
      Width           =   7695
      Begin VB.TextBox txtObs_EnvioInstruccionesalAgenteCarga 
         Height          =   825
         Left            =   1710
         TabIndex        =   55
         Tag             =   "SET"
         Top             =   675
         Width           =   5880
      End
      Begin NumBoxProject.NumBox txtFec_EnvioInstruccionesalAgenteCarga 
         Height          =   285
         Left            =   1710
         TabIndex        =   67
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
      Begin VB.Label lblObs_EnvioInstruccionesalAgenteCarga 
         Caption         =   "Observaciones"
         Height          =   480
         Left            =   165
         TabIndex        =   54
         Tag             =   "OBS_ENVIOINSTRUCCIONESALAGENTECARGA"
         Top             =   660
         Width           =   1500
      End
      Begin VB.Label lblFec_EnvioInstruccionesalAgenteCarga 
         Caption         =   "Fecha Envio"
         Height          =   480
         Left            =   165
         TabIndex        =   53
         Tag             =   "FEC_ENVIOINSTRUCCIONESALAGENTECARGA"
         Top             =   255
         Width           =   1500
      End
   End
   Begin VB.Frame fraEnvioAgenteAduanas 
      Caption         =   "Envio Instrucciones al Agente Aduanas"
      Enabled         =   0   'False
      Height          =   1635
      Left            =   60
      TabIndex        =   47
      Top             =   4770
      Width           =   7695
      Begin VB.TextBox txtObs_EnvioInstruccionesalAgenteAduanas 
         Height          =   915
         Left            =   1710
         TabIndex        =   50
         Tag             =   "SET"
         Top             =   585
         Width           =   5865
      End
      Begin NumBoxProject.NumBox txtFec_EnvioInstruccionesalAgenteAduana 
         Height          =   285
         Left            =   1710
         TabIndex        =   66
         Top             =   240
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
      Begin NumBoxProject.NumBox txtFec_EnvioFacturaalAgenteAduanas 
         Height          =   285
         Left            =   5610
         TabIndex        =   70
         Top             =   225
         Width           =   1290
         _ExtentX        =   2275
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
      Begin VB.Label lblFec_EnvioFacturaalAgenteAduanas 
         Caption         =   "Fecha Envio Factura"
         Height          =   270
         Left            =   4020
         TabIndex        =   51
         Tag             =   "FEC_ENVIOFACTURAALAGENTEADUANAS"
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label lblObs_EnvioInstruccionesalAgenteAduanas 
         Caption         =   "Observaciones"
         Height          =   480
         Left            =   165
         TabIndex        =   49
         Tag             =   "OBS_ENVIOINSTRUCCIONESALAGENTEADUANAS"
         Top             =   585
         Width           =   1500
      End
      Begin VB.Label lblFec_EnvioInstruccionesalAgenteAduanas 
         Caption         =   "Fecha Envio"
         Height          =   480
         Left            =   165
         TabIndex        =   48
         Tag             =   "FEC_ENVIOINSTRUCCIONESALAGENTEADUANAS"
         Top             =   210
         Width           =   1500
      End
   End
   Begin VB.Frame fraDua 
      Caption         =   "DUA"
      Enabled         =   0   'False
      Height          =   1605
      Left            =   60
      TabIndex        =   42
      Top             =   8100
      Width           =   7680
      Begin NumBoxProject.NumBox txtEntregaDrauBack 
         Height          =   315
         Left            =   1695
         TabIndex        =   77
         Top             =   1035
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
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
      Begin VB.TextBox txtNum_Dua 
         Height          =   300
         Left            =   1710
         TabIndex        =   45
         Tag             =   "SET"
         Top             =   225
         Width           =   2025
      End
      Begin NumBoxProject.NumBox txtFec_NumeracionDua 
         Height          =   285
         Left            =   1710
         TabIndex        =   68
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
         TabIndex        =   69
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
         TabIndex        =   72
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
      Begin VB.Label Label4 
         Caption         =   "Fecha Entrega Tramite Drau Back"
         Height          =   390
         Left            =   165
         TabIndex        =   76
         Top             =   1005
         Width           =   1410
      End
      Begin VB.Label lblFec_EmbarqueReal 
         Caption         =   "Fec Real de Embarque"
         Height          =   480
         Left            =   3930
         TabIndex        =   71
         Tag             =   "FEC_EMBARQUEREAL"
         Top             =   195
         Width           =   1065
      End
      Begin VB.Label lblFec_NumeracionDua 
         Caption         =   "Fecha de Numeración"
         Height          =   390
         Left            =   150
         TabIndex        =   46
         Tag             =   "FEC_NUMERACIONDUA"
         Top             =   540
         Width           =   1500
      End
      Begin VB.Label lblNum_Dua 
         Caption         =   "Numero de Dua"
         Height          =   480
         Left            =   150
         TabIndex        =   44
         Tag             =   "NUM_DUA"
         Top             =   270
         Width           =   1500
      End
      Begin VB.Label lblFec_RecepcionDua 
         Caption         =   "Fecha de Recepción "
         Height          =   330
         Left            =   3930
         TabIndex        =   43
         Tag             =   "FEC_RECEPCIONDUA"
         Top             =   645
         Width           =   1500
      End
   End
   Begin VB.TextBox txtDes_Flete_Ingles 
      Height          =   300
      Left            =   2160
      TabIndex        =   25
      Tag             =   "SET"
      Top             =   4470
      Width           =   5535
   End
   Begin VB.TextBox txtDes_Embarque 
      Height          =   300
      Left            =   2160
      TabIndex        =   22
      Tag             =   "SET"
      Top             =   3750
      Width           =   5535
   End
   Begin VB.TextBox txtDes_TipEmbarque 
      Height          =   300
      Left            =   2145
      TabIndex        =   19
      Tag             =   "SET"
      Top             =   3030
      Width           =   5550
   End
   Begin VB.TextBox txtNom_Ejecutivo 
      Height          =   300
      Left            =   2145
      TabIndex        =   17
      Tag             =   "SET"
      Top             =   2670
      Width           =   5550
   End
   Begin VB.TextBox txtNom_AgenteAduana 
      Height          =   300
      Left            =   2160
      TabIndex        =   15
      Tag             =   "SET"
      Top             =   2265
      Width           =   5535
   End
   Begin VB.TextBox txtDes_AgenteCarga 
      Height          =   300
      Left            =   2160
      TabIndex        =   13
      Tag             =   "SET"
      Top             =   1875
      Width           =   5535
   End
   Begin VB.TextBox txtCod_Flete 
      Height          =   300
      Left            =   1650
      TabIndex        =   24
      Tag             =   "SET"
      Top             =   4485
      Width           =   465
   End
   Begin VB.TextBox txtNom_Embarque 
      Height          =   300
      Left            =   1650
      TabIndex        =   23
      Tag             =   "SET"
      Top             =   4110
      Width           =   2655
   End
   Begin VB.TextBox txtCod_Embarque 
      Height          =   300
      Left            =   1650
      TabIndex        =   21
      Tag             =   "SET"
      Top             =   3750
      Width           =   465
   End
   Begin VB.TextBox txtObs_Embarque 
      Height          =   300
      Left            =   1650
      TabIndex        =   20
      Tag             =   "SET"
      Top             =   3390
      Width           =   6045
   End
   Begin VB.TextBox txtTip_Embarque 
      Height          =   300
      Left            =   1650
      TabIndex        =   18
      Tag             =   "SET"
      Top             =   3030
      Width           =   450
   End
   Begin VB.TextBox txtCod_Ejecutivo_AgenteCarga 
      Height          =   300
      Left            =   1650
      TabIndex        =   16
      Tag             =   "SET"
      Top             =   2685
      Width           =   435
   End
   Begin VB.TextBox txtCod_AgenteAduana 
      Height          =   300
      Left            =   1650
      TabIndex        =   14
      Tag             =   "SET"
      Top             =   2265
      Width           =   450
   End
   Begin VB.TextBox txtCod_AgenteCarga 
      Height          =   300
      Left            =   1650
      TabIndex        =   12
      Tag             =   "SET"
      Top             =   1875
      Width           =   435
   End
   Begin VB.TextBox txtRef_Embarque 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1650
      TabIndex        =   9
      Tag             =   "SET"
      Top             =   1485
      Width           =   1590
   End
   Begin VB.TextBox txtAbr_Cliente 
      Height          =   300
      Left            =   1650
      TabIndex        =   7
      Tag             =   "SET"
      Top             =   1125
      Width           =   555
   End
   Begin VB.TextBox txtCod_Anxo 
      Height          =   300
      Left            =   1965
      TabIndex        =   5
      Tag             =   "SET"
      Top             =   765
      Width           =   585
   End
   Begin VB.TextBox txtCod_TipAnex 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1650
      TabIndex        =   4
      Tag             =   "SET"
      Top             =   765
      Width           =   240
   End
   Begin VB.TextBox txtCod_Origen 
      Height          =   300
      Left            =   1650
      TabIndex        =   1
      Tag             =   "SET"
      Top             =   405
      Width           =   495
   End
   Begin VB.TextBox txtNum_Embarque 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1650
      TabIndex        =   0
      Tag             =   "SET"
      Top             =   45
      Width           =   1200
   End
   Begin FunctionsButtons.FunctButt FunctOKCancel 
      Height          =   510
      Left            =   2835
      TabIndex        =   26
      Top             =   9735
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmTG_Embarque.frx":0000
      Orientacion     =   0
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Label Label3 
      Caption         =   "Almacén de Aduana"
      Height          =   330
      Left            =   105
      TabIndex        =   75
      Tag             =   "COD_AGENTEADUANA"
      Top             =   10980
      Width           =   1500
   End
   Begin VB.Label Label12 
      Caption         =   "Termino de Venta"
      Height          =   360
      Left            =   3360
      TabIndex        =   65
      Top             =   1455
      Width           =   1020
   End
   Begin VB.Label Label2 
      Caption         =   "Periodo"
      Height          =   330
      Left            =   5955
      TabIndex        =   59
      Tag             =   "COD_TIPANEX"
      Top             =   450
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Estado"
      Height          =   270
      Left            =   2925
      TabIndex        =   56
      Tag             =   "COD_AGENTECARGA"
      Top             =   90
      Width           =   825
   End
   Begin VB.Label lblCod_Flete 
      Caption         =   "Tipo de Flete"
      Height          =   480
      Left            =   105
      TabIndex        =   41
      Tag             =   "COD_FLETE"
      Top             =   4485
      Width           =   1500
   End
   Begin VB.Label lblNom_Embarque 
      Caption         =   "Nombre Embarque"
      Height          =   195
      Left            =   105
      TabIndex        =   40
      Tag             =   "NOM_EMBARQUE"
      Top             =   4155
      Width           =   1515
   End
   Begin VB.Label lblCod_Embarque 
      Caption         =   "Modo de Embarque"
      Height          =   315
      Left            =   90
      TabIndex        =   39
      Tag             =   "COD_EMBARQUE"
      Top             =   3780
      Width           =   1500
   End
   Begin VB.Label lblObs_Embarque 
      Caption         =   "Observaciones"
      Height          =   300
      Left            =   105
      TabIndex        =   38
      Tag             =   "OBS_EMBARQUE"
      Top             =   3435
      Width           =   1500
   End
   Begin VB.Label lblCod_LugEntrega 
      Caption         =   "Tipo Embarque"
      Height          =   345
      Left            =   105
      TabIndex        =   37
      Tag             =   "COD_LUGENTREGA"
      Top             =   3060
      Width           =   1500
   End
   Begin VB.Label lblCod_Ejecutivo_AgenteCarga 
      Caption         =   "Ejecutivo Ag.Carga"
      Height          =   240
      Left            =   105
      TabIndex        =   36
      Tag             =   "COD_EJECUTIVO_AGENTECARGA"
      Top             =   2700
      Width           =   1470
   End
   Begin VB.Label lblCod_AgenteAduana 
      Caption         =   "Agente de Aduana"
      Height          =   330
      Left            =   105
      TabIndex        =   35
      Tag             =   "COD_AGENTEADUANA"
      Top             =   2325
      Width           =   1500
   End
   Begin VB.Label lblCod_AgenteCarga 
      Caption         =   "Agente de Carga"
      Height          =   225
      Left            =   105
      TabIndex        =   34
      Tag             =   "COD_AGENTECARGA"
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label lblRef_Embarque 
      Caption         =   "Referencia"
      Height          =   315
      Left            =   105
      TabIndex        =   33
      Tag             =   "REF_EMBARQUE"
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label lblCod_Cliente 
      Caption         =   "Cliente"
      Height          =   285
      Left            =   105
      TabIndex        =   32
      Tag             =   "COD_CLIENTE"
      Top             =   1185
      Width           =   1500
   End
   Begin VB.Label lblCod_TipAnex 
      Caption         =   "Anexo"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Tag             =   "COD_TIPANEX"
      Top             =   810
      Width           =   1500
   End
   Begin VB.Label lblCod_Origen 
      Caption         =   "Origen"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Tag             =   "COD_ORIGEN"
      Top             =   420
      Width           =   1485
   End
   Begin VB.Label lblNum_Embarque 
      Caption         =   "Número"
      Height          =   255
      Left            =   105
      TabIndex        =   29
      Tag             =   "NUM_EMBARQUE"
      Top             =   90
      Width           =   1500
   End
End
Attribute VB_Name = "frmTG_Embarque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public Codigo As String
Public Descripcion As String
Public TipoAdd As String
Dim rstAux As ADODB.Recordset
Public oParent As Object
Public sAccion As String

Private Sub cmdLugEnt_Click()
    Load frmMantLugaresEntrega
    frmMantLugaresEntrega.sCod_Cliente = Me.txtAbr_Cliente.Tag
    frmMantLugaresEntrega.CARGA_GRID
    frmMantLugaresEntrega.Show vbModal
    Set frmMantLugaresEntrega = Nothing
End Sub

Public Sub BuscaCliente(opcion As String)
Dim rstAux As ADODB.Recordset

    strSQL = "SELECT Cod_Cliente, Abr_Cliente, Nom_Cliente FROM TG_CLIENTE WHERE "
    
    txtAbr_Cliente = Trim(txtAbr_Cliente)
    txtNom_Cliente = Trim(txtNom_Cliente)
    
    Select Case opcion
    Case 1: strSQL = strSQL & "Abr_Cliente LIKE '%" & txtAbr_Cliente & "%'"
    Case 2: strSQL = strSQL & "Nom_Cliente LIKE '%" & txtNom_Cliente & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
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
    txtNom_Cliente = ""
    If Codigo <> "" Then
        
        txtAbr_Cliente = Descripcion
        txtNom_Cliente = TipoAdd
        txtAbr_Cliente.Tag = Codigo
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
        
    Codigo = ""
    Descripcion = ""
End Sub

Private Sub cmdAgenteAduana_Click()
frmMantAgenteAduana.Show 1

End Sub

Private Sub cmdAgenteCarga_Click()
frmMantAgenteCarga.Show 1
End Sub

Private Sub cmdAlmacenAduana_Click()
frmMantAlmacenAduana.Show 1
End Sub

Private Sub cmdEjecutivoAgCarga_Click()
frmMantEjecutivoCarga.sCOD = txtCod_AgenteAduana.Text
frmMantEjecutivoCarga.sDES = txtNom_AgenteAduana.Text
frmMantEjecutivoCarga.Show 1
End Sub

Private Sub cmdModoEmbarque_Click()
frmMantModoEmbarque.Show 1

End Sub

Private Sub Form_Load()
    txtCod_TipAnex.Text = DevuelveCampo("SELECT Prefijo_Anexo_Ventas_Exportacion FROM cn_control ", cCONNECT)
    txtCod_Origen = "7"
    txtAno = Year(Date)
    txtmes = StrZero(Month(Date), 2)
End Sub

Private Sub FunctOKCancel_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            Grabar
        Case "CANCELAR"
            Unload Me
    End Select
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCod_AgenteAduana_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaAgenteAduana 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCod_AgenteCarga_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BUSCA_AGENTECARGA 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCod_AlmacenAduana_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub



Private Sub txtCod_Ejecutivo_AgenteCarga_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaEjecutivo 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCod_Embarque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaModoTransporte 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCod_Flete_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaFlete 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCod_Origen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaOrigen 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDes_AgenteCarga_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtDes_Embarque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtDes_Flete_Ingles_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtDes_Termino_Venta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtLinea1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtDes_TipEmbarque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtNom_AgenteAduana_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtNom_AlmacenAduana_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtNom_Ejecutivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtNom_embarque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub TxtAbr_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaCliente 1
        SendKeys "{TAB}"
    End If
End Sub


Private Sub TxtNom_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaCliente 2
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txttip_Embarque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        
        BuscaTipEmbarque 1
        SendKeys "{TAB}"
    End If
End Sub

Public Sub BuscaTipEmbarque(opcion As String)
Dim rstAux As ADODB.Recordset
    strSQL = "SELECT Tip_Embarque , Des_TipEmbarque  FROM TG_TIPOEMBARQUE " & _
             "WHERE "
    
    txtTip_Embarque = Trim(txtTip_Embarque)
    txtDes_TipEmbarque = Trim(txtDes_TipEmbarque)
    
    Select Case opcion
    Case 1: strSQL = strSQL & "Tip_Embarque  like '%" & txtTip_Embarque & "%'"
    Case 2: strSQL = strSQL & "RTRIM(Des_TipEmbarque) LIKE '%" & txtDes_TipEmbarque & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
        
    frmBusqGeneral3.gexLista.Columns("Tip_Embarque").Width = 570
    frmBusqGeneral3.gexLista.Columns("Des_TipEmbarque").Width = 2370
    
    frmBusqGeneral3.gexLista.Columns("Tip_Embarque").Caption = "Tipo Embarque"
    frmBusqGeneral3.gexLista.Columns("Des_TipEmbarque").Caption = "Descripcion"
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtTip_Embarque = ""
    txtDes_TipEmbarque = ""
    
    If Codigo <> "" Then
        txtTip_Embarque = Codigo
        txtDes_TipEmbarque = Descripcion
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    Codigo = ""
    Descripcion = ""
End Sub

Private Sub txtCod_Termino_Venta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaTerminoVent 1
        SendKeys "{TAB}"
    End If
End Sub

Public Sub BuscaTerminoVent(opcion As String)
Dim rstAux As ADODB.Recordset

    strSQL = "SELECT Cod_Termino_Venta, Des_Termino_Venta FROM CN_Termino_Venta WHERE "

    txtCod_Termino_Venta = Trim(txtCod_Termino_Venta)
    txtDes_Termino_Venta = Trim(txtDes_Termino_Venta)

    Select Case opcion
    Case 1: strSQL = strSQL & "Cod_Termino_Venta like '%" & txtCod_Termino_Venta & "%'"
    Case 2: strSQL = strSQL & "Des_Termino_Venta LIKE '%" & txtDes_Termino_Venta & "%'"
    End Select

    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset

    frmBusqGeneral3.gexLista.Columns("Cod_Termino_Venta").Width = 700
    frmBusqGeneral3.gexLista.Columns("Des_Termino_Venta").Width = 2000

    frmBusqGeneral3.gexLista.Columns("Cod_Termino_Venta").Caption = "Termino.Venta"
    frmBusqGeneral3.gexLista.Columns("Des_Termino_Venta").Caption = "Descrip."

    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If

    txtCod_Termino_Venta = ""
    txtDes_Termino_Venta = ""

    If Codigo <> "" Then
        txtCod_Termino_Venta = Codigo
        txtDes_Termino_Venta = Descripcion
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing

    Codigo = ""
    Descripcion = ""
End Sub




Private Sub txtCod_Emabarque_Venta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaModoTransporte 1
        SendKeys "{TAB}"
    End If
End Sub

Public Sub BuscaModoTransporte(opcion As String)
Dim rstAux As ADODB.Recordset

    strSQL = "SELECT Cod_Embarque, Des_Embarque FROM TG_TIPEMB WHERE "
    
    txtCod_Embarque = Trim(txtCod_Embarque)
    txtDes_Embarque = Trim(txtDes_Embarque)
    
    Select Case opcion
    Case 1: strSQL = strSQL & "Cod_Embarque like '%" & txtCod_Embarque & "%'"
    Case 2: strSQL = strSQL & "Des_Embarque LIKE '%" & txtDes_Embarque & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    
    frmBusqGeneral3.gexLista.Columns("Cod_Embarque").Width = 700
    frmBusqGeneral3.gexLista.Columns("Des_Embarque").Width = 2000
    
    frmBusqGeneral3.gexLista.Columns("Cod_Embarque").Caption = "Embarque"
    frmBusqGeneral3.gexLista.Columns("Des_Embarque").Caption = "Descrip."
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtCod_Embarque = ""
    txtDes_Embarque = ""
    
    If Codigo <> "" Then
        txtCod_Embarque = Codigo
        txtDes_Embarque = Descripcion
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    Codigo = ""
    Descripcion = ""
End Sub

Public Function CargaValores(ByRef ObjTemp As Object) As Boolean
    ObjTemp.txtAbr_Cliente.Text = txtAbr_Cliente.Text
    ObjTemp.txtAbr_Cliente.Tag = txtAbr_Cliente.Tag
    ObjTemp.TxtDes_Cliente.Text = txtNom_Cliente.Text
End Function

Private Sub BuscaFlete(opcion As String)
Dim rstAux As ADODB.Recordset

    strSQL = "SELECT Cod_Flete, Des_Flete_Ingles as Des_Flete_Ingles FROM TG_TIPOFLETE WHERE "
    
    txtCod_Flete = Trim(txtCod_Flete)
    txtDes_Flete_Ingles = Trim(txtDes_Flete_Ingles)
    
    Select Case opcion
    Case 1: strSQL = strSQL & "Cod_Flete like '%" & txtCod_Flete & "%'"
    Case 2: strSQL = strSQL & "Des_Flete_Ingles LIKE '%" & txtDes_Flete_Ingles & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    
    frmBusqGeneral3.gexLista.Columns("Cod_Flete").Width = 700
    frmBusqGeneral3.gexLista.Columns("Des_Flete_Ingles").Width = 2000
    
    frmBusqGeneral3.gexLista.Columns("Cod_Flete").Caption = "Embarque"
    frmBusqGeneral3.gexLista.Columns("Des_Flete_Ingles").Caption = "Descrip."
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtCod_Flete = ""
    txtDes_Flete_Ingles = ""
    
    If Codigo <> "" Then
        txtCod_Flete = Codigo
        txtDes_Flete_Ingles = Descripcion
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    Codigo = ""
    Descripcion = ""
End Sub

Public Sub BUSCA_AGENTECARGA(Tipo As Integer)
    strSQL = "SELECT Cod_AgenteCarga , Des_AgenteCarga FROM TG_AGENTECARGA WHERE "
    Select Case Tipo
        Case 1: strSQL = strSQL & "Cod_AgenteCarga  like '%" & txtCod_AgenteCarga & "%'"
        Case 2: strSQL = strSQL & "Des_AgenteCarga LIKE '%" & txtDes_AgenteCarga & "%'"
    End Select

    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
        
    frmBusqGeneral3.gexLista.Columns("Cod_AgenteCarga").Width = 570
    frmBusqGeneral3.gexLista.Columns("des_AgenteCarga").Width = 2370
    
    frmBusqGeneral3.gexLista.Columns("Cod_AgenteCarga").Caption = "Agente"
    frmBusqGeneral3.gexLista.Columns("des_AgenteCarga").Caption = "Nombre"
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtCod_AgenteCarga = ""
    txtDes_AgenteCarga = ""
    
    If Codigo <> "" Then
        txtCod_AgenteCarga = Codigo
        txtDes_AgenteCarga = Descripcion
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    Codigo = ""
    Descripcion = ""
        
End Sub

Public Sub BuscaAgenteAduana(Tipo As Integer)
    strSQL = "SELECT Cod_AgenteAduana , Nom_AgenteAduana FROM TG_AGENTEAduana WHERE "
    Select Case Tipo
        Case 1: strSQL = strSQL & "Cod_AgenteAduana  like '%" & txtCod_AgenteAduana & "%'"
        Case 2: strSQL = strSQL & "Nom_AgenteAduana LIKE '%" & txtNom_AgenteAduana & "%'"
    End Select

    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
        
    frmBusqGeneral3.gexLista.Columns("Cod_AgenteAduana").Width = 570
    frmBusqGeneral3.gexLista.Columns("Nom_AgenteAduana").Width = 2370
    
    frmBusqGeneral3.gexLista.Columns("Cod_AgenteAduana").Caption = "Agente"
    frmBusqGeneral3.gexLista.Columns("nom_AgenteAduana").Caption = "Nombre"
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtCod_AgenteAduana = ""
    txtNom_AgenteAduana = ""
    
    If Codigo <> "" Then
        txtCod_AgenteAduana = Codigo
        txtNom_AgenteAduana = Descripcion
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    Codigo = ""
    Descripcion = ""
        
End Sub

Public Sub BuscaEjecutivo(opcion As String)
Dim rstAux As ADODB.Recordset
    strSQL = "SELECT Cod_Ejecutivo, nom_ejecutivo FROM TG_AgenteCarga_Ejecutivo " & _
             "WHERE Cod_AgenteCarga = '" & txtCod_AgenteCarga & "' AND "
    
    txtCod_Ejecutivo_AgenteCarga = Trim(txtCod_Ejecutivo_AgenteCarga)
    txtNom_Ejecutivo = Trim(txtNom_Ejecutivo)
    
    Select Case opcion
    Case 1: strSQL = strSQL & "cod_ejecutivo like '%" & txtCod_Ejecutivo_AgenteCarga & "%'"
    Case 2: strSQL = strSQL & "RTRIM(Nom_Ejecutivo) LIKE '%" & txtNom_AgenteAduana & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
        
    frmBusqGeneral3.gexLista.Columns("Cod_ejecutivo").Width = 570
    frmBusqGeneral3.gexLista.Columns("Nom_ejecutivo").Width = 2370
    
    frmBusqGeneral3.gexLista.Columns("Cod_Ejecutivo").Caption = "Ejecutivo"
    frmBusqGeneral3.gexLista.Columns("Nom_Ejecutivo").Caption = "Nombre"
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtCod_Ejecutivo_AgenteCarga = ""
    txtNom_Ejecutivo = ""
    
    If Codigo <> "" Then
        txtCod_Ejecutivo_AgenteCarga = Codigo
        txtNom_Ejecutivo = Descripcion
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    Codigo = ""
    Descripcion = ""
End Sub


Private Sub txtCod_TipAnex_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Trim(txtCod_Anxo) <> "" Then BuscaAnexo 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDes_Anexo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaAnexo 2
        SendKeys "{TAB}"
    End If
End Sub

Private Sub BuscaAnexo(opcion As String)
    
    strSQL = "SELECT Cod_TipAnex, Cod_Anxo, Des_Anexo FROM CN_ANEXOSCONTABLES " & _
             "WHERE Cod_TipAnex = '" & txtCod_TipAnex & "' AND "
    
    txtCod_Anxo = Trim(txtCod_Anxo)
    txtDes_Anexo = Trim(txtDes_Anexo)
    
    Select Case opcion
    Case 1: strSQL = strSQL & "Cod_Anxo LIKE '%" & txtCod_Anxo & "%'"
    Case 2: strSQL = strSQL & "Des_Anexo LIKE '%" & txtDes_Anexo & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    'frmBusqGeneralJanus.Show vbModal
    
    frmBusqGeneral3.gexLista.Columns("Cod_TipAnex").Width = 400
    frmBusqGeneral3.gexLista.Columns("Cod_Anxo").Width = 570
    frmBusqGeneral3.gexLista.Columns("Des_Anexo").Width = 2370
    
    frmBusqGeneral3.gexLista.Columns("Cod_TipAnex").Caption = "Tipo"
    frmBusqGeneral3.gexLista.Columns("Cod_Anxo").Caption = "Codigo"
    frmBusqGeneral3.gexLista.Columns("Des_Anexo").Caption = "Anexo Contable"
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtCod_Anxo = ""
    txtDes_Anexo = ""
    
    If Codigo <> "" Then
        txtCod_TipAnex = Codigo
        txtCod_Anxo = Descripcion
        txtDes_Anexo = TipoAdd
    End If
    Codigo = ""
    Descripcion = ""
End Sub



Private Sub txtCod_Anxo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        'KeyAscii = 0
        'BuscaAnexo 1
        SendKeys "{TAB}"
    End If
End Sub




Private Sub Grabar()
On Error GoTo errx
Dim ssql As String
Dim mRs As ADODB.Recordset
ssql = "TG_EMBARQUE_MAN '$','$','$','$','$','$','$','$','$','$','$','$','$','$','$','$','$','$','$'"
  
ssql = VBsprintf(ssql, sAccion, txtNum_Embarque, txtCod_Origen, txtAno, txtmes, txtCod_TipAnex, txtCod_Anxo, _
  txtAbr_Cliente.Tag, txtCod_AgenteAduana, txtCod_AgenteCarga, txtCod_Ejecutivo_AgenteCarga, _
  txtTip_Embarque, txtObs_Embarque, txtCod_Embarque, txtNom_Embarque, txtCod_Flete, txtCod_Termino_Venta, _
  vusu, ComputerName())

Set mRs = GetDataSet(cCONNECT, ssql)

If Not mRs.EOF Then
    oParent.txtRef_Embarque.Text = mRs!Ref_Embarque
    oParent.optReferencia.Value = True
    oParent.BUSCAR
End If

Unload Me

Exit Sub
errx:
    errores err.Number
End Sub


Public Sub BuscaOrigen(opcion As String)
Dim rstAux As ADODB.Recordset

    strSQL = "SELECT Cod_Origen, Descripcion FROM CN_Ventas_Origen_Factura  WHERE "

    txtCod_Origen = Trim(txtCod_Origen)
    txtDes_Origen = Trim(txtDes_Origen)

    Select Case opcion
    Case 1: strSQL = strSQL & "Cod_Origen like '%" & txtCod_Origen & "%'"
    Case 2: strSQL = strSQL & "Descripcion LIKE '%" & txtDes_Origen & "%'"
    End Select

    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset

    frmBusqGeneral3.gexLista.Columns("Cod_Origen").Width = 700
    frmBusqGeneral3.gexLista.Columns("Descripcion").Width = 2000

    frmBusqGeneral3.gexLista.Columns("Cod_Origen").Caption = "Origen"
    frmBusqGeneral3.gexLista.Columns("Descripcion").Caption = "Descrip."

    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If

    txtCod_Origen = ""
    txtDes_Origen = ""

    If Codigo <> "" Then
        txtCod_Origen = Codigo
        txtDes_Origen = Descripcion
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing

    Codigo = ""
    Descripcion = ""
End Sub


Private Sub txtObs_Embarque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub
