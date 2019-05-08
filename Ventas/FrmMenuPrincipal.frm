VERSION 5.00
Begin VB.Form FrmMenuPrincipal 
   Caption         =   "Menu Principal"
   ClientHeight    =   10950
   ClientLeft      =   450
   ClientTop       =   345
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   11775
   Begin VB.CommandButton Command10 
      Caption         =   "Facturas Canceladas Segun Rango de Fechas"
      Height          =   735
      Left            =   9840
      TabIndex        =   75
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Liquidar Factura"
      Height          =   495
      Left            =   9840
      TabIndex        =   74
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ResumenDeuda"
      Height          =   615
      Left            =   9840
      TabIndex        =   73
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton Command75 
      Caption         =   "salir"
      Height          =   735
      Left            =   9840
      TabIndex        =   72
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton Command81 
      Caption         =   "Facturas Emitidas Segun Rango de Fechas"
      Height          =   735
      Left            =   9840
      TabIndex        =   71
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton Command82 
      Caption         =   "estadistica anual"
      Height          =   735
      Left            =   9840
      TabIndex        =   70
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command84 
      Caption         =   "Cta Clientes Factoring"
      Height          =   735
      Left            =   9840
      TabIndex        =   69
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton Command85 
      Caption         =   "Facturas Pendientes recuperacion DrawBack"
      Height          =   735
      Left            =   9840
      TabIndex        =   68
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton Command87 
      Caption         =   "Tipo De Cambio"
      Height          =   615
      Left            =   5880
      TabIndex        =   67
      Top             =   8040
      Width           =   1935
   End
   Begin VB.CommandButton Command80 
      Caption         =   "Motivo Notas"
      Height          =   675
      Left            =   7800
      TabIndex        =   66
      Top             =   8040
      Width           =   1935
   End
   Begin VB.CommandButton Command79 
      Caption         =   "Ventas de Hilo Comprado"
      Height          =   675
      Left            =   7800
      TabIndex        =   65
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton Command77 
      Caption         =   "Autorización de Pago de Documentos Saldos - Tela Cruda / Teñida"
      Height          =   675
      Left            =   7830
      TabIndex        =   64
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton Command76 
      Caption         =   "Emision de Venta por Tipo Venta"
      Height          =   675
      Left            =   7830
      TabIndex        =   63
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Penalidades de Venta"
      Height          =   675
      Left            =   7830
      TabIndex        =   62
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton Command73 
      Caption         =   "Emision de Ventas por Cliente"
      Height          =   675
      Left            =   5880
      TabIndex        =   61
      Top             =   7365
      Width           =   1935
   End
   Begin VB.CommandButton Command72 
      Caption         =   "VENTAS_FACTURAS_EXPO_SUJETAS_DRAW_BACK"
      Height          =   675
      Left            =   5880
      TabIndex        =   60
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CommandButton Command71 
      Caption         =   "Transmision Facturas Venta a INKA"
      Height          =   675
      Left            =   7830
      TabIndex        =   59
      Top             =   4710
      Width           =   1935
   End
   Begin VB.CommandButton Command70 
      Caption         =   "Emision Reporte Ventas por Grupo  [ECN]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7800
      TabIndex        =   58
      Top             =   4035
      Width           =   1935
   End
   Begin VB.CommandButton Command69 
      Caption         =   "Ranking de Ventas por Pais-Destino [ECN]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7830
      TabIndex        =   57
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Command68 
      Caption         =   "Ranking de Ventas [ECN]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7830
      TabIndex        =   56
      Top             =   2685
      Width           =   1935
   End
   Begin VB.CommandButton Command66 
      Caption         =   "Cierre de Cobranzas Diversas"
      Height          =   675
      Left            =   7830
      TabIndex        =   55
      Top             =   2055
      Width           =   1935
   End
   Begin VB.CommandButton Command64 
      Caption         =   "Mantenimiento Conceptos de Cobranza"
      Height          =   675
      Left            =   7830
      TabIndex        =   54
      Top             =   1380
      Width           =   1935
   End
   Begin VB.CommandButton Command65 
      Caption         =   "Mantenimiento de Tipos de Cobranza"
      Height          =   675
      Left            =   7830
      TabIndex        =   53
      Top             =   705
      Width           =   1935
   End
   Begin VB.CommandButton Command63 
      Caption         =   "Documentos de Venta"
      Height          =   675
      Left            =   7680
      TabIndex        =   52
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command62 
      Caption         =   "Estadistica de Ventas"
      Height          =   675
      Left            =   3960
      TabIndex        =   51
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CommandButton Command61 
      Caption         =   "Rpt Cancelaciones Incobra"
      Height          =   675
      Left            =   3960
      TabIndex        =   50
      Top             =   8040
      Width           =   1935
   End
   Begin VB.CommandButton Command60 
      Caption         =   "Rpt Facturas Canjes"
      Height          =   675
      Left            =   5880
      TabIndex        =   49
      Top             =   6690
      Width           =   1935
   End
   Begin VB.CommandButton Command59 
      Caption         =   "Rpt Cancelaciones Notas Abono"
      Height          =   675
      Left            =   5880
      TabIndex        =   48
      Top             =   6015
      Width           =   1935
   End
   Begin VB.CommandButton Command58 
      Caption         =   "Rpt Anticipos Canjes"
      Height          =   675
      Left            =   5880
      TabIndex        =   47
      Top             =   5340
      Width           =   1935
   End
   Begin VB.CommandButton Command57 
      Caption         =   "Reporte Anual Ventas"
      Height          =   675
      Left            =   5880
      TabIndex        =   46
      Top             =   4665
      Width           =   1935
   End
   Begin VB.CommandButton Command56 
      Caption         =   "Emision Ventas por Grupo de Ventas y Fechas"
      Height          =   675
      Left            =   5880
      TabIndex        =   45
      Top             =   3990
      Width           =   1935
   End
   Begin VB.CommandButton Command55 
      Caption         =   "Resumen Ventas SobrePartida Arancelaria "
      Height          =   675
      Left            =   5880
      TabIndex        =   44
      Top             =   3315
      Width           =   1935
   End
   Begin VB.CommandButton Command54 
      Caption         =   "Flujo de Cobranza"
      Height          =   675
      Left            =   5880
      TabIndex        =   43
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton Command53 
      Caption         =   "Control de Letras"
      Height          =   675
      Left            =   5880
      TabIndex        =   42
      Top             =   1965
      Width           =   1935
   End
   Begin VB.CommandButton Command52 
      Caption         =   "Reporte Cancelacion Fac Vs Letras"
      Height          =   675
      Left            =   5880
      TabIndex        =   41
      Top             =   1290
      Width           =   1935
   End
   Begin VB.CommandButton Command50 
      Caption         =   "Cobranza por Periodo"
      Height          =   675
      Left            =   5880
      TabIndex        =   40
      Top             =   660
      Width           =   1935
   End
   Begin VB.CommandButton Command48 
      Caption         =   "Reporte Año - Periodo"
      Height          =   675
      Left            =   5880
      TabIndex        =   39
      Top             =   30
      Width           =   1935
   End
   Begin VB.CommandButton Command47 
      Caption         =   "Mantenimiento Almacen Aduana"
      Height          =   675
      Left            =   3960
      TabIndex        =   38
      Top             =   7365
      Width           =   1935
   End
   Begin VB.CommandButton Command46 
      Caption         =   "Mantenimiento Modo Embarque"
      Height          =   675
      Left            =   3960
      TabIndex        =   37
      Top             =   6690
      Width           =   1935
   End
   Begin VB.CommandButton Command45 
      Caption         =   "Mantenimiento Ejecutivo Cargo"
      Height          =   675
      Left            =   3960
      TabIndex        =   36
      Top             =   6015
      Width           =   1935
   End
   Begin VB.CommandButton Command44 
      Caption         =   "Mantenimiento Agente Aduana"
      Height          =   675
      Left            =   3960
      TabIndex        =   35
      Top             =   5340
      Width           =   1935
   End
   Begin VB.CommandButton Command42 
      Caption         =   "Mantenimiento Agente Carga"
      Height          =   675
      Left            =   3960
      TabIndex        =   34
      Top             =   4710
      Width           =   1935
   End
   Begin VB.CommandButton Command41 
      Caption         =   "Deporte de Detraccion"
      Height          =   675
      Left            =   3960
      TabIndex        =   33
      Top             =   4035
      Width           =   1935
   End
   Begin VB.CommandButton Command40 
      Caption         =   "Cierre Ano  Mes"
      Height          =   675
      Left            =   3960
      TabIndex        =   32
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Command39 
      Caption         =   "Descargo de Letras en Descunto"
      Height          =   675
      Left            =   3960
      TabIndex        =   31
      Top             =   2685
      Width           =   1935
   End
   Begin VB.CommandButton Command38 
      Caption         =   "Descargo de Letras en Descunto"
      Height          =   675
      Left            =   3960
      TabIndex        =   30
      Top             =   2010
      Width           =   1935
   End
   Begin VB.CommandButton Command37 
      Caption         =   "Cancelaciones de Boletas"
      Height          =   675
      Left            =   3960
      TabIndex        =   29
      Top             =   1335
      Width           =   1935
   End
   Begin VB.CommandButton Command36 
      Caption         =   "Reporte Detalle Exportacion Agrupado"
      Height          =   675
      Left            =   3960
      TabIndex        =   28
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command34 
      Caption         =   "Reporte de Status de Letra"
      Height          =   675
      Left            =   3960
      TabIndex        =   27
      Top             =   30
      Width           =   1935
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Aceptacion PArte Cancelacion"
      Height          =   675
      Left            =   2025
      TabIndex        =   26
      Top             =   8760
      Width           =   1935
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Seguimiento Draw Back"
      Height          =   675
      Left            =   2025
      TabIndex        =   25
      Top             =   8085
      Width           =   1935
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Facturas Diferidas"
      Height          =   675
      Left            =   2025
      TabIndex        =   24
      Top             =   7410
      Width           =   1935
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Parte de Cobranza"
      Height          =   675
      Left            =   2025
      TabIndex        =   23
      Top             =   6735
      Width           =   1935
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Control Documentos Ventas"
      Height          =   675
      Left            =   2025
      TabIndex        =   22
      Top             =   6060
      Width           =   1935
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Liquidacion Diaria"
      Height          =   675
      Left            =   90
      TabIndex        =   21
      Top             =   8040
      Width           =   1935
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Revision Parte Cobranza}"
      Height          =   675
      Left            =   2025
      TabIndex        =   20
      Top             =   5385
      Width           =   1935
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Cta Clientes"
      Height          =   675
      Left            =   2025
      TabIndex        =   19
      Top             =   4710
      Width           =   1935
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Letras Pendientes de Pago"
      Height          =   675
      Left            =   2025
      TabIndex        =   18
      Top             =   4035
      Width           =   1935
   End
   Begin VB.CommandButton Command23 
      Caption         =   "GRAFICO"
      Height          =   675
      Left            =   2025
      TabIndex        =   17
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Cheques diferidos"
      Height          =   675
      Left            =   2025
      TabIndex        =   16
      Top             =   2685
      Width           =   1935
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Adelantos"
      Height          =   675
      Left            =   2025
      TabIndex        =   15
      Top             =   2010
      Width           =   1935
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Documento de Cobranza"
      Height          =   675
      Left            =   2025
      TabIndex        =   14
      Top             =   1335
      Width           =   1935
   End
   Begin VB.CommandButton Command18 
      Caption         =   "ANEXOS POR CLIENTE"
      Height          =   675
      Left            =   2025
      TabIndex        =   13
      Top             =   705
      Width           =   1935
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Facturas Regiustradas Exportacion"
      Height          =   675
      Left            =   2025
      TabIndex        =   12
      Top             =   30
      Width           =   1935
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Asociar Cliente - Anexo Contable"
      Height          =   675
      Left            =   90
      TabIndex        =   11
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Guias por Facturar -Prendas"
      Height          =   675
      Left            =   90
      TabIndex        =   10
      Top             =   6645
      Width           =   1935
   End
   Begin VB.CommandButton Command15 
      Caption         =   "REporte de Grupos"
      Height          =   675
      Left            =   90
      TabIndex        =   9
      Top             =   5295
      Width           =   1935
   End
   Begin VB.CommandButton Command14 
      Caption         =   "REporte de Grupos"
      Height          =   675
      Left            =   90
      TabIndex        =   8
      Top             =   4620
      Width           =   1935
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Letras"
      Height          =   675
      Left            =   90
      TabIndex        =   7
      Top             =   5970
      Width           =   1935
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Reporte de Exportacion"
      Height          =   675
      Left            =   90
      TabIndex        =   6
      Top             =   3945
      Width           =   1935
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Registro De Ventas"
      Height          =   675
      Left            =   90
      TabIndex        =   5
      Top             =   3270
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Reporte Ventas"
      Height          =   675
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Control Numeracion Facturas"
      Height          =   675
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1965
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Actualizar Precio O/C Tejeduria"
      Height          =   675
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1290
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Guias Por Facturar Tela Tenida"
      Height          =   675
      Left            =   90
      TabIndex        =   1
      Top             =   615
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Muestra Facturas REgiustradas"
      Height          =   675
      Left            =   90
      TabIndex        =   0
      Top             =   -60
      Width           =   1935
   End
End
Attribute VB_Name = "FrmMenuPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub Command1_Click()
Frm_Resumen_Ventas.Show 1
End Sub

Private Sub Command10_Click()
frmFacturasCanceladasRango.Show 1
End Sub

Private Sub Command11_Click()
Frm_Registro_Ventas.Show 1
End Sub



Private Sub Command13_Click()
  frmLetra.Show 1
End Sub

Private Sub Command14_Click()
FrmRptVentasxGrupo.Show 1
End Sub

Private Sub Command15_Click()
  FrmRptRelacionArticulos_x_Grupos.Show 1
End Sub

Private Sub Command16_Click()
    Load frmMantAnxCli
    frmMantAnxCli.Show vbModal
    Set frmMantAnxCli = Nothing
End Sub

Private Sub Command17_Click()
    Load frmShowFactVentasPrendasExportacion
    frmShowFactVentasPrendasExportacion.Show vbModal
    Set frmShowFactVentasPrendasExportacion = Nothing
End Sub

Private Sub Command18_Click()
    frmMantAnxCli.Show vbModal
    Set frmMantAnxCli = Nothing
End Sub





Private Sub Command2_Click()
    frmShowFactVentas_Liquidar.Show 1
End Sub

Private Sub Command20_Click()
  frmTransacciones.Show 1
End Sub

Private Sub Command21_Click()
  frmAdelantos.Show 1
End Sub

Private Sub Command22_Click()
    Load frmChequesDiferidos
    frmChequesDiferidos.Show vbModal
    Set frmChequesDiferidos = Nothing
End Sub



Private Sub Command24_Click()
  FrmRptLetrasPendientePago.Show 1
End Sub

Private Sub Command25_Click()
frmShowCtaCte.Show 1


End Sub

Private Sub Command26_Click()
frmShowPartexAutorizar.Show 1
End Sub

Private Sub Command27_Click()
  frmLiquidacionDiaria.Show 1
End Sub

Private Sub Command28_Click()
    frmShowSeguimDocumVentas.Show vbModal
    Set frmShowSeguimDocumVentas = Nothing
End Sub

Private Sub Command29_Click()
    frmShowPartesCobranzas.Show vbModal
    Set frmShowPartesCobranzas = Nothing
End Sub

Private Sub Command3_Click()
  frmShowFactVentas.Show 1
End Sub

Private Sub Command30_Click()
    frmFacturasDiferidas.Show vbModal
    Set frmFacturasDiferidas = Nothing
End Sub

Private Sub Command31_Click()
    frmShowSeguimDrawBack.Show vbModal
    Set frmShowSeguimDrawBack = Nothing
    
End Sub

Private Sub Command32_Click()
frmShowCanjeAutorizar.Show 1
End Sub



Private Sub Command34_Click()
  FrmRptLetrasStatus.Show 1
End Sub



Private Sub Command36_Click()
  FrmRptDetalleExport.Show 1
End Sub

Private Sub Command37_Click()
  FrmRptCancelaciones_Boletas.Show 1
End Sub

Private Sub Command38_Click()
  FrmRptLetrasDescuentos.Show 1
End Sub

Private Sub Command39_Click()
    frnRepCanjeLetras.Show vbModal
    Set frnRepCanjeLetras = Nothing
End Sub

Private Sub Command4_Click()
FrmPenalidadesVentas.Show 1
End Sub

Private Sub Command40_Click()
  frmCierreAnoMes.Show 1
End Sub

Private Sub Command41_Click()
  frmRptDetracciones.Show 1
End Sub

Private Sub Command42_Click()
frmMantAgenteCarga.Show 1
End Sub


Private Sub Command44_Click()
frmMantAgenteAduana.Show 1
End Sub

Private Sub Command45_Click()
frmMantEjecutivoCarga.Show 1

End Sub

Private Sub Command46_Click()
frmMantModoEmbarque.Show 1
End Sub

Private Sub Command47_Click()
frmMantAlmacenAduana.Show 1
End Sub

Private Sub Command48_Click()
frmReporteAnioPeriodo.Show 1
End Sub




Private Sub Command50_Click()
frmCobranzaXPeriodo.Show 1
End Sub



Private Sub Command52_Click()
FrmRptCancelaciones_x_Factura.Show 1
End Sub

Private Sub Command53_Click()
frmControl_Letras.Show 1
End Sub

Private Sub Command54_Click()
frmFlujoCobranza.Show 1
End Sub
Private Sub Command55_Click()
    frmResumenVentaSobrePartida.Show 1
End Sub
Private Sub Command56_Click()
    FrmRptVentasxGrupoxTipoVenta.Show 1
End Sub

Private Sub Command57_Click()
frmReporteResumenAnualVentas.Show 1
End Sub

Private Sub Command58_Click()
FrmRptAnticipos_Canjes.Show 1
End Sub

Private Sub Command59_Click()
FrmRptCancelaciones_NotasAbono.Show 1
End Sub

Private Sub Command6_Click()
frmShowGuiasxFact_TelaTenida.Show 1
End Sub

Private Sub Command60_Click()
FrmRptFacturas_Canjes.Show 1
End Sub

Private Sub Command61_Click()
FrmRptCancelaciones_Incobra.Show 1
End Sub

Private Sub Command62_Click()
frmEstadisticaVentas.Show 1
End Sub

Private Sub Command63_Click()
    frmShowFactVentas.Show 1
End Sub



Private Sub Command64_Click()
frmConceptoCobranza.Show 1
End Sub

Private Sub Command65_Click()
frmTiposCobranza.Show 1
End Sub

Private Sub Command66_Click()
frmShowCierreTipoDiario_Ventas.Show 1
End Sub


Private Sub Command68_Click()
    frmConVentasReq.Show 1
End Sub

Private Sub Command69_Click()
    frmRankingVentasPorPaisDestino.Show 1
End Sub

Private Sub Command7_Click()
FrmShowActPrecioOC.Show 1
End Sub

Private Sub Command70_Click()
    FrmRptVentasxGrupo.Show 1
End Sub

Private Sub Command71_Click()
frmTransFactVentas.Show 1
End Sub

Private Sub Command72_Click()
Frm_FactExpoSujetas.Show 1
End Sub

Private Sub Command73_Click()
FrmRptVentasxCliente.Show 1
End Sub

Private Sub Command74_Click()
frmShowGuiasxFact_Lavanderia.Show 1
End Sub

Private Sub Command75_Click()
Unload Me
End Sub

Private Sub Command76_Click()
FrmRptVentasxTipoVenta.Show 1
End Sub

Private Sub Command77_Click()
frmShowGuiasxFact_SaldosTelaTenida.Show 1
End Sub



Private Sub Command79_Click()
frmMuestraHiloComprado.Show 1
End Sub

Private Sub Command8_Click()
  frmShowControlNumeracion.Show 1
End Sub

Private Sub Command80_Click()
FrmMotivo_Notas.Show 1
End Sub

Private Sub Command81_Click()
frmFacturasEmiRanFecha.Show 1
End Sub

Private Sub Command82_Click()
FrmEstadisticaAnual.Show 1
End Sub



Private Sub Command84_Click()
frmCtaCteCliFacExt.Show 1
End Sub

Private Sub Command85_Click()
    FrmFacturas_Pendientes_Recuperacion_Draw.Show 1
End Sub


Private Sub Command87_Click()
frmShowTiposCambio.Show 1
End Sub

Private Sub Command9_Click()
  frmConVentasReq.Show 1
End Sub

Private Sub Form_Load()

cCONNECT = "Provider=SQLOLEDB.1;Password=soporte;Persist Security Info=True;User ID=soporte;Initial Catalog=textilesjoc;Data Source=192.168.1.10"
cSEGURIDAD = "Provider=SQLOLEDB.1;Password=soporte;Persist Security Info=True;User ID=soporte;Initial Catalog=Seguridad;Data Source=192.168.1.10"
vusu = "SISTEMAS"
vper = "0001"
vemp = "01"
'vRuta = App.Path
vRuta = "C:\Program Files (x86)\Sistema Produccion"
iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))


End Sub



