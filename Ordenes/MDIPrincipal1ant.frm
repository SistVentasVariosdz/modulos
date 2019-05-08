VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H00808080&
   Caption         =   "Menú Principal"
   ClientHeight    =   6675
   ClientLeft      =   165
   ClientTop       =   1020
   ClientWidth     =   11175
   Icon            =   "MDIPrincipal1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Tag             =   "Menu"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6300
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "EMPRESA :"
            TextSave        =   "EMPRESA :"
            Object.Tag             =   "COMPANY :"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "USUARIO :"
            TextSave        =   "USUARIO :"
            Object.Tag             =   "USER :"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7938
            MinWidth        =   7938
            Text            =   "CONEXION :"
            TextSave        =   "CONEXION :"
            Object.Tag             =   "CONNECTION :"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "EQUIPO :"
            TextSave        =   "EQUIPO :"
            Object.Tag             =   "PC:"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "16/05/2003"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal1.frx":0442
            Key             =   "mancli"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal1.frx":0894
            Key             =   "manfab"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal1.frx":0CE6
            Key             =   "manOrg"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal1.frx":1138
            Key             =   "mantra"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal1.frx":158A
            Key             =   "mancomisin"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal1.frx":19DC
            Key             =   "manBan"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal1.frx":1E2E
            Key             =   "mandestino"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal1.frx":2280
            Key             =   "mantippre"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1680
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   1140
      Top             =   4560
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Menu mnuTabla 
      Caption         =   "Maestros"
      Begin VB.Menu mancli 
         Caption         =   "Clients"
         Shortcut        =   ^L
      End
      Begin VB.Menu manFab 
         Caption         =   "Factory"
         Shortcut        =   ^F
      End
      Begin VB.Menu CarShip 
         Caption         =   "Shipment"
         Begin VB.Menu manPagemb 
            Caption         =   "Shipment Pay"
            Shortcut        =   ^P
         End
         Begin VB.Menu manTipEmb 
            Caption         =   "Shipment Type"
            Shortcut        =   ^T
         End
      End
      Begin VB.Menu mantippre 
         Caption         =   "Garment Type"
         Shortcut        =   ^G
      End
      Begin VB.Menu manunimed 
         Caption         =   "Unit Messure"
         Shortcut        =   ^U
      End
      Begin VB.Menu manMon 
         Caption         =   "Moneys"
         Shortcut        =   ^M
      End
      Begin VB.Menu manDestino 
         Caption         =   "Destinations"
         Shortcut        =   ^D
      End
      Begin VB.Menu mancomisin 
         Caption         =   "Agent"
         Shortcut        =   ^S
      End
      Begin VB.Menu manorg 
         Caption         =   "Organizations"
         Shortcut        =   ^O
      End
      Begin VB.Menu mantal 
         Caption         =   "Sizes"
         Shortcut        =   ^Z
      End
      Begin VB.Menu manban 
         Caption         =   "Banks"
         Shortcut        =   ^B
      End
      Begin VB.Menu manmotatr 
         Caption         =   "Delay Motive"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mandivpre 
         Caption         =   "Divisiones de Prenda"
         Shortcut        =   ^W
      End
      Begin VB.Menu mantCargos 
         Caption         =   "Charge Maintenance"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnuPurch 
      Caption         =   "Comercial"
      Begin VB.Menu manPOObs 
         Caption         =   "Consulta PO"
      End
      Begin VB.Menu mnuRegIn 
         Caption         =   "Registro PO"
      End
      Begin VB.Menu mnuConFa 
         Caption         =   "Consulta Facturas"
      End
      Begin VB.Menu mnuWizPO 
         Caption         =   "Wizard P.O."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuupdate 
         Caption         =   "Actualizaciones"
         Begin VB.Menu mnupoco 
            Caption         =   "PO/Colores/Estilos"
         End
         Begin VB.Menu mnuprecio 
            Caption         =   "Precio/Prendas"
         End
      End
      Begin VB.Menu manseg 
         Caption         =   "Seguimiento PO"
      End
      Begin VB.Menu DespPrenda 
         Caption         =   "Despacho de Prendas"
      End
      Begin VB.Menu ActEstCli 
         Caption         =   "Estilos Cliente"
      End
      Begin VB.Menu mnuproform 
         Caption         =   "Proforma"
      End
      Begin VB.Menu mnuCotiza 
         Caption         =   "Cotizaciones"
      End
      Begin VB.Menu mnubustel 
         Caption         =   "Busqueda de Telas"
      End
      Begin VB.Menu mnufacdet 
         Caption         =   "Detalle Facturación"
      End
      Begin VB.Menu mnucosgru 
         Caption         =   "Costos x Grupo"
      End
      Begin VB.Menu mnucosconf 
         Caption         =   "Costos Confeccion"
      End
      Begin VB.Menu mnuCosSem 
         Caption         =   "Costos Semanales"
      End
      Begin VB.Menu mnuConCot 
         Caption         =   "Consulta Cotizaciones"
      End
      Begin VB.Menu mnuGruPro 
         Caption         =   "Grupos de Producción"
      End
   End
   Begin VB.Menu mnuRepor 
      Caption         =   "Reportes"
      Begin VB.Menu RepTra 
         Caption         =   "Production Tracking"
      End
      Begin VB.Menu mnuResPr 
         Caption         =   "Resumen Production"
         Visible         =   0   'False
      End
      Begin VB.Menu RepDelDet 
         Caption         =   "Delivery Detail"
      End
      Begin VB.Menu mnuProCo 
         Caption         =   "Proyectado de Comisionistas"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuComRep 
         Caption         =   "Commisionist Report"
      End
      Begin VB.Menu mnuDeliv 
         Caption         =   "Despachos"
      End
      Begin VB.Menu mnuResDe 
         Caption         =   "Ventas Proyect.-Detalle"
      End
      Begin VB.Menu mnuCiCom 
         Caption         =   "Cierre de Comisiones"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDeCom 
         Caption         =   "Delivery Comisionistas (prendas e importes)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDelCP 
         Caption         =   "Delivery Contador Prendas (prendas, importes , comisiones)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConFF 
         Caption         =   "Control Facturacion Fabrica"
      End
      Begin VB.Menu mnuDeOpe 
         Caption         =   "Despachos de Operadora (contramuestras)"
      End
      Begin VB.Menu mnuforecast 
         Caption         =   "Proyeccion Comisiones"
      End
      Begin VB.Menu concierre 
         Caption         =   "Consulta Cierres"
      End
      Begin VB.Menu Cierre 
         Caption         =   "Cierre Operaciones"
      End
   End
   Begin VB.Menu mnuEst 
      Caption         =   "Estilos"
      Begin VB.Menu mnuestprop 
         Caption         =   "Estilos Propios"
      End
      Begin VB.Menu mnuTipComp 
         Caption         =   "Tipos Componente"
      End
      Begin VB.Menu mnuComp 
         Caption         =   "Componentes"
      End
      Begin VB.Menu mnutiprec 
         Caption         =   "Tipos Receta"
      End
      Begin VB.Menu mnugamas 
         Caption         =   "Gamas Colores"
      End
      Begin VB.Menu mnuintcol 
         Caption         =   "Intensidad Colores"
      End
      Begin VB.Menu mnusolid 
         Caption         =   "Colores"
      End
      Begin VB.Menu mnucarcol 
         Caption         =   "Carta de Colores"
      End
      Begin VB.Menu mnuproest 
         Caption         =   "Servicios/Procesos"
      End
      Begin VB.Menu mantGruTal 
         Caption         =   "Grupos Tallas"
      End
      Begin VB.Menu mnumot 
         Caption         =   "Mot.Pre.Prod."
      End
      Begin VB.Menu mnuPiezas 
         Caption         =   "Piezas"
      End
      Begin VB.Menu mnuimpmas 
         Caption         =   "Impresión Masiva Estilos"
      End
      Begin VB.Menu mnuPrePro 
         Caption         =   "Precios por Estilo Proceso"
      End
   End
   Begin VB.Menu mnuLogis 
      Caption         =   "Logistica"
      Begin VB.Menu mnumantp 
         Caption         =   "Proveedores"
      End
      Begin VB.Menu mnuclaitm 
         Caption         =   "Clase Item"
      End
      Begin VB.Menu mnuorig 
         Caption         =   "Origen/Proc."
      End
      Begin VB.Menu mnutit 
         Caption         =   "Titulos Hilados"
      End
      Begin VB.Menu mnugalgas 
         Caption         =   "Galgas"
      End
      Begin VB.Menu mnutipraya 
         Caption         =   "Tipos Listados"
      End
      Begin VB.Menu mantItem 
         Caption         =   "Items"
      End
      Begin VB.Menu MantTelas 
         Caption         =   "Telas"
      End
      Begin VB.Menu mantHil 
         Caption         =   "Hilados"
      End
      Begin VB.Menu mnumpr 
         Caption         =   "Materia Prima"
      End
      Begin VB.Menu frmGruposReq 
         Caption         =   "Grupos Explosion"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProceso 
         Caption         =   "Procesos Textíles"
      End
      Begin VB.Menu mnuraya 
         Caption         =   "-"
      End
      Begin VB.Menu TipOrdComp 
         Caption         =   "Tipo de Orden Compra"
      End
      Begin VB.Menu StaOrdComp 
         Caption         =   "Estado Orden de Compra"
      End
      Begin VB.Menu LugEntr 
         Caption         =   "Lugar Entrega"
      End
      Begin VB.Menu Dscto 
         Caption         =   "Descuento"
      End
      Begin VB.Menu CondVent 
         Caption         =   "Condición de Venta"
      End
      Begin VB.Menu mnuOrdComp 
         Caption         =   "Orden de Compra"
      End
      Begin VB.Menu mnuServTen 
         Caption         =   "Stocks. Serv. Teñido"
      End
      Begin VB.Menu mnurvsreal 
         Caption         =   "Requerimiento Vs Real Avios"
      End
      Begin VB.Menu mnuHorStk 
         Caption         =   "Horizonte de Stocks"
      End
      Begin VB.Menu mnuEntPen 
         Caption         =   "Seguimiento de Entregas Pendientes"
      End
   End
   Begin VB.Menu mnuConf 
      Caption         =   "Confecciones"
      Begin VB.Menu mnugrupo 
         Caption         =   "Grupos"
      End
      Begin VB.Menu mnuOrdPro 
         Caption         =   "Ordenes Produccion"
      End
      Begin VB.Menu mnuProdMen 
         Caption         =   "Producción Mensual"
      End
      Begin VB.Menu mnuultimos 
         Caption         =   "Ultimos Datos"
      End
      Begin VB.Menu mnuconsumo 
         Caption         =   "Consumos Unitarios"
      End
      Begin VB.Menu mnuPrgRea 
         Caption         =   "Programado Vs. Real"
      End
      Begin VB.Menu mnuordcort 
         Caption         =   "Ordenes de Corte"
      End
      Begin VB.Menu mnuposcort 
         Caption         =   "Posicion Corte"
      End
      Begin VB.Menu mnuFacPro 
         Caption         =   "Facturación Producción Semanal"
      End
      Begin VB.Menu mnuetiq 
         Caption         =   "Etiquetas"
      End
      Begin VB.Menu mnuOrdCnf 
         Caption         =   "Orden de Confección"
      End
      Begin VB.Menu mnutarifado 
         Caption         =   "Tarifado"
      End
      Begin VB.Menu mnuLecTic 
         Caption         =   "Lectura de Tickets"
      End
      Begin VB.Menu mnuActTic 
         Caption         =   "Actualización de Tickets"
      End
      Begin VB.Menu mnuResAct 
         Caption         =   "Resumen de Actualización"
      End
      Begin VB.Menu mnuEfic 
         Caption         =   "Eficiencia - Trabajador"
      End
      Begin VB.Menu mnuStocks 
         Caption         =   "Stocks en Proceso"
      End
      Begin VB.Menu frmImpSit 
         Caption         =   "Reporte Situacion O/P"
      End
      Begin VB.Menu mnuProHab 
         Caption         =   "Producción Habilitada"
      End
      Begin VB.Menu mnuSecCon 
         Caption         =   "Sectores de Confección"
      End
      Begin VB.Menu mnuLinPro 
         Caption         =   "Líneas de Producción"
      End
      Begin VB.Menu mnuProPro 
         Caption         =   "Producción en Proceso"
      End
      Begin VB.Menu mnuEstSem 
         Caption         =   "Minutos Estandares x Sem"
      End
      Begin VB.Menu mnuEstDia 
         Caption         =   "Minutos Estandares x Día"
      End
      Begin VB.Menu mnuStkCor 
         Caption         =   "Reporte Stocks en Corte"
      End
      Begin VB.Menu mnuStkCos 
         Caption         =   "Reporte Stocks en Costura"
      End
      Begin VB.Menu mnuESemCS 
         Caption         =   "Eficiencia Semanal por Linea Costura"
      End
      Begin VB.Menu mnuLecEsp 
         Caption         =   "Lectura de Tickets de Movim Confección"
      End
      Begin VB.Menu mnuPDCorte 
         Caption         =   "Produccion Diaria Corte"
      End
      Begin VB.Menu mnuSPProv 
         Caption         =   "Stocks por Proveedor"
      End
      Begin VB.Menu mnuAOpDia 
         Caption         =   "Avance Operaciones Diaria"
      End
      Begin VB.Menu mnuAvaAca 
         Caption         =   "Avance Diario Acabados"
      End
      Begin VB.Menu mnuAvaCn2 
         Caption         =   "Avance por Orden de Confección"
      End
      Begin VB.Menu mnuPakLis 
         Caption         =   "Packing List"
      End
      Begin VB.Menu mnuBiHorE 
         Caption         =   "Control Bi-Horario Eficiencia"
      End
      Begin VB.Menu mnuEfiDia 
         Caption         =   "Eficiencia Diaria"
      End
      Begin VB.Menu mnuEstClC 
         Caption         =   "Estilos Cliente Habilitados por Corte"
      End
      Begin VB.Menu mnuAvaGer 
         Caption         =   "Avance por Orden Confeccion (Gerencial)"
      End
      Begin VB.Menu mnuIndSem 
         Caption         =   "Indicadores Semanales"
      End
      Begin VB.Menu mnuFlujopd 
         Caption         =   "Flujo Producción Semanal"
      End
      Begin VB.Menu mnuCtPrMe 
         Caption         =   "Control de Producción Mensual"
      End
   End
   Begin VB.Menu mnugrupos 
      Caption         =   "Grupos"
      Begin VB.Menu mnuGrupoL 
         Caption         =   "Grupo Logístico"
      End
      Begin VB.Menu mnuGrupoT 
         Caption         =   "Grupo Textil"
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Administración"
      Begin VB.Menu manDocs 
         Caption         =   "Documentos"
      End
      Begin VB.Menu mnutipdoc 
         Caption         =   "Tipos Documentos"
      End
      Begin VB.Menu mnuanxcon 
         Caption         =   "Anexos Contables"
      End
      Begin VB.Menu mnutipanx 
         Caption         =   "Tipos de Anexos"
      End
      Begin VB.Menu mnugrupreg 
         Caption         =   "Grupos Reg"
      End
      Begin VB.Menu mnuTipCam 
         Caption         =   "Tipo de Cambio"
      End
      Begin VB.Menu mnutippro 
         Caption         =   "Tipo Producto"
      End
      Begin VB.Menu mnuRegCom 
         Caption         =   "Registro de Compras"
      End
      Begin VB.Menu mnuGenCoa 
         Caption         =   "Generación Archivo COA"
      End
      Begin VB.Menu mnuAutorPg 
         Caption         =   "Autorización de Documentos de Pago"
      End
      Begin VB.Menu mnuAutorLT 
         Caption         =   "Autorización de Letras"
      End
      Begin VB.Menu mnuCancDoc 
         Caption         =   "Cancelación de Documentos de Pago"
      End
      Begin VB.Menu mnuConsAut 
         Caption         =   "Consulta Autorizaciones de Documentos de Pago"
      End
      Begin VB.Menu mnuletra 
         Caption         =   "Canje de Letras"
      End
      Begin VB.Menu mnuConLet 
         Caption         =   "Consulta de Letras"
      End
      Begin VB.Menu mnuND 
         Caption         =   "Notas de Debito"
      End
      Begin VB.Menu mnuMovBan 
         Caption         =   "Movimientos Bancos"
      End
      Begin VB.Menu mnuTraPag 
         Caption         =   "Transferencia Docum Cancelados a Contabilidad"
      End
      Begin VB.Menu mnuRepDRB 
         Caption         =   "Impresión de Facturas DR_BACK"
      End
      Begin VB.Menu mnuKarMTA 
         Caption         =   "Kardex Mensual Tela Acabada"
      End
      Begin VB.Menu mnuNumDoc 
         Caption         =   "Configuración de Número de Documento"
      End
      Begin VB.Menu mnuTrans 
         Caption         =   "Transmisiones a Contabilidad"
      End
   End
   Begin VB.Menu mnumovi 
      Caption         =   "Movimientos"
      Begin VB.Menu mnukardex 
         Caption         =   "Kardex"
      End
      Begin VB.Menu mnumantalm 
         Caption         =   "Almacen"
      End
      Begin VB.Menu mnumovperm 
         Caption         =   "Movimientos Permitidos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnutipmov 
         Caption         =   "Tipos de Movimiento"
      End
      Begin VB.Menu mnumovalm 
         Caption         =   "Movimientos de Almacen"
      End
      Begin VB.Menu mnuguiaman 
         Caption         =   "Guía Manual"
      End
      Begin VB.Menu mnuclasoc 
         Caption         =   "Clases de Orden de Compra"
      End
      Begin VB.Menu mnumovsto 
         Caption         =   "Stocks por Almacen"
      End
      Begin VB.Menu mnutelaca 
         Caption         =   "Kardex - Tela Acabada"
      End
      Begin VB.Menu mnutelcru 
         Caption         =   "Kardex Tela Cruda"
      End
      Begin VB.Menu mnuhilten 
         Caption         =   "Kardex Hilo Teñido"
      End
      Begin VB.Menu mnuhilcru 
         Caption         =   "Kardex Hilo Crudo"
      End
      Begin VB.Menu mnustkfam 
         Caption         =   "Stock por Familia"
      End
      Begin VB.Menu mnumstock 
         Caption         =   "Movimientos Stocks"
      End
      Begin VB.Menu mnudatec 
         Caption         =   "Datos Tecn. de la Tela"
      End
      Begin VB.Menu mnuguias 
         Caption         =   "Control de Guias"
      End
      Begin VB.Menu mnuMovCnf 
         Caption         =   "Movimiento Confecciones"
      End
      Begin VB.Menu mnuKarTer 
         Caption         =   "Kardex Costura Terceros"
      End
      Begin VB.Menu mnumovsal 
         Caption         =   "Movimiento de Saldos"
      End
      Begin VB.Menu mnuconsal 
         Caption         =   "Consulta Prendas Saldos"
      End
      Begin VB.Menu mnuParTelB 
         Caption         =   "Partidas Tela Acabada"
      End
   End
   Begin VB.Menu mnuTextil 
      Caption         =   "Textil"
      Begin VB.Menu mnuPartida 
         Caption         =   "Partidas de Tintorería"
      End
      Begin VB.Menu mnuTelSer 
         Caption         =   "Telas en Servicio de Teñido"
      End
      Begin VB.Menu mnuMatReq 
         Caption         =   "Matriz de Requerimientos Textiles"
      End
   End
   Begin VB.Menu mnuasist 
      Caption         =   "Asistencia"
      Begin VB.Menu mnutrab 
         Caption         =   "Trabajadores"
      End
      Begin VB.Menu mnumaestro 
         Caption         =   "Maestros"
         Begin VB.Menu mnuconcep 
            Caption         =   "Conceptos"
         End
         Begin VB.Menu mnuhorario 
            Caption         =   "Horarios"
         End
      End
      Begin VB.Menu mnusubmar 
         Caption         =   "Subir Ranura"
      End
      Begin VB.Menu mnugenmar 
         Caption         =   "Genera Marcaciones"
      End
      Begin VB.Menu mnuactinf 
         Caption         =   "Actualiza Inf. Diara"
      End
      Begin VB.Menu mnureverr 
         Caption         =   "Revisa Errados"
      End
      Begin VB.Menu mnuautsob 
         Caption         =   "Autozacion Sobretiempo"
      End
      Begin VB.Menu mnuactinfs 
         Caption         =   "Actualiza Inf. Semanal"
      End
      Begin VB.Menu mnuRegDir 
         Caption         =   "Reg.Directo Asistencia"
      End
      Begin VB.Menu mnuextmar 
         Caption         =   "Extorno de Marcaciones"
      End
      Begin VB.Menu mnuCIPla 
         Caption         =   "Interfase Planilla"
      End
      Begin VB.Menu mnurepasis 
         Caption         =   "Reportes"
         Begin VB.Menu mnuentreg 
            Caption         =   "Entradas Registradas"
         End
         Begin VB.Menu mnusalnreg 
            Caption         =   "Salidas no Registradas"
         End
         Begin VB.Menu mnuhortra 
            Caption         =   "Horas Trabajadas"
         End
         Begin VB.Menu mnuRptIna 
            Caption         =   "Inasistencias"
         End
         Begin VB.Menu mnuResAsi 
            Caption         =   "Resumen de Asistencia"
         End
      End
   End
   Begin VB.Menu ISO0 
      Caption         =   "ISO"
      Index           =   0
      Begin VB.Menu ISO 
         Caption         =   "Sistema de Gestion de calidad"
         Index           =   4
         Begin VB.Menu ISO4 
            Caption         =   "Requisitos generales"
            Index           =   1
         End
         Begin VB.Menu ISO4 
            Caption         =   "Requisitos de la documentacion"
            Index           =   2
            Begin VB.Menu ISO42 
               Caption         =   "Generalidades"
               Index           =   1
            End
            Begin VB.Menu ISO42 
               Caption         =   "Manual de Calidad"
               Index           =   2
            End
            Begin VB.Menu ISO42 
               Caption         =   "Control de los Documentos"
               Index           =   3
            End
            Begin VB.Menu ISO42 
               Caption         =   "Control de Registros"
               Index           =   4
            End
         End
      End
      Begin VB.Menu ISO 
         Caption         =   "Responsabilidad de Dirección"
         Index           =   5
         Begin VB.Menu ISO5 
            Caption         =   "Compromiso de Direccion"
            Index           =   1
         End
         Begin VB.Menu ISO5 
            Caption         =   "Enfoque al Cliente"
            Index           =   2
         End
         Begin VB.Menu ISO5 
            Caption         =   "Politica de Calidad"
            Index           =   3
         End
         Begin VB.Menu ISO5 
            Caption         =   "Planificación"
            Index           =   4
            Begin VB.Menu ISO54 
               Caption         =   "Objetivos de Calidad"
               Index           =   1
            End
            Begin VB.Menu ISO54 
               Caption         =   "Planificacion del sistema de gestion de la calidad"
               Index           =   2
            End
         End
         Begin VB.Menu ISO5 
            Caption         =   "Responsabilidad,Autoridad y Comunicación"
            Index           =   5
            Begin VB.Menu ISO55 
               Caption         =   "Responsabilidad y Autoridad"
               Index           =   1
            End
            Begin VB.Menu ISO55 
               Caption         =   "Representante de la Direccion"
               Index           =   2
            End
            Begin VB.Menu ISO55 
               Caption         =   "Comunicacion Interna"
               Index           =   3
            End
         End
         Begin VB.Menu ISO5 
            Caption         =   "Revision de la Direccion"
            Index           =   6
            Begin VB.Menu ISO56 
               Caption         =   "Informacion para la revision"
               Index           =   1
            End
            Begin VB.Menu ISO56 
               Caption         =   "Informacion para la revision"
               Index           =   2
            End
            Begin VB.Menu ISO56 
               Caption         =   "Resultados de Revision"
               Index           =   3
            End
         End
      End
      Begin VB.Menu ISO 
         Caption         =   "Gestión de los Recursos"
         Index           =   6
         Begin VB.Menu ISO6 
            Caption         =   "Provisión de Recursos"
            Index           =   1
         End
         Begin VB.Menu ISO6 
            Caption         =   "Recursos Humanos"
            Index           =   2
            Begin VB.Menu ISO62 
               Caption         =   "Generalidades"
               Index           =   1
            End
            Begin VB.Menu ISO62 
               Caption         =   "Competencia, toma de conciencia y formación"
               Index           =   2
            End
         End
         Begin VB.Menu ISO6 
            Caption         =   "Infraestructura"
            Index           =   3
         End
         Begin VB.Menu ISO6 
            Caption         =   "Ambiente de Trabajo"
            Index           =   4
         End
      End
      Begin VB.Menu ISO 
         Caption         =   "Realización del Producto"
         Index           =   7
         Begin VB.Menu ISO7 
            Caption         =   "Planificación de la realización del producto"
            Index           =   1
         End
         Begin VB.Menu ISO7 
            Caption         =   "Procesos relacionados con el cliente"
            Index           =   2
            Begin VB.Menu ISO72 
               Caption         =   "Determinación de los requisitos relacionados con el producto"
               Index           =   1
            End
            Begin VB.Menu ISO72 
               Caption         =   "Revisión de los requisitos relacionados con el producto"
               Index           =   2
            End
            Begin VB.Menu ISO72 
               Caption         =   "Comunicación con el Cliente"
               Index           =   3
            End
         End
         Begin VB.Menu ISO7 
            Caption         =   "Diseño y Desarrollo"
            Index           =   3
            Begin VB.Menu ISO73 
               Caption         =   "Planificacion del diseño y desarrollo"
               Index           =   1
            End
            Begin VB.Menu ISO73 
               Caption         =   "Elementos de Entrada para el diseño y desarrollo"
               Index           =   2
            End
            Begin VB.Menu ISO73 
               Caption         =   "Resultados del diseño y desarrollo"
               Index           =   3
            End
            Begin VB.Menu ISO73 
               Caption         =   "Revisión del diseño y desarrollo"
               Index           =   4
            End
            Begin VB.Menu ISO73 
               Caption         =   "Verificación del diseño y desarrollo"
               Index           =   5
            End
            Begin VB.Menu ISO73 
               Caption         =   "Validación del diseño y desarrollo"
               Index           =   6
            End
            Begin VB.Menu ISO73 
               Caption         =   "Control de Cambios del diseño y desarrollo"
               Index           =   7
            End
         End
         Begin VB.Menu ISO7 
            Caption         =   "Compras"
            Index           =   4
            Begin VB.Menu ISO74 
               Caption         =   "Proceso de Compras"
               Index           =   1
            End
            Begin VB.Menu ISO74 
               Caption         =   "Información de las Compras"
               Index           =   2
            End
            Begin VB.Menu ISO74 
               Caption         =   "Verificación de los Productos Comprados"
               Index           =   3
            End
         End
         Begin VB.Menu ISO7 
            Caption         =   "Producto y Prestación del Servicio"
            Index           =   5
            Begin VB.Menu ISO75 
               Caption         =   "Control de la Producción y de la prestación del Servicio"
               Index           =   1
            End
            Begin VB.Menu ISO75 
               Caption         =   "Validacion de los procesos de la producción y la prestación"
               Index           =   2
            End
            Begin VB.Menu ISO75 
               Caption         =   "Identificación y Trazabilidad"
               Index           =   3
            End
            Begin VB.Menu ISO75 
               Caption         =   "Propiedad del Cliente"
               Index           =   4
            End
            Begin VB.Menu ISO75 
               Caption         =   "Preservación del Producto"
               Index           =   5
            End
         End
         Begin VB.Menu ISO7 
            Caption         =   "Control de los dispositivos de Seguimiento y Medición"
            Index           =   6
         End
      End
      Begin VB.Menu ISO 
         Caption         =   "Medición y Análisis"
         Index           =   8
         Begin VB.Menu ISO8 
            Caption         =   "Generalidades"
            Index           =   1
         End
         Begin VB.Menu ISO8 
            Caption         =   "Seguimiento y Medicion"
            Index           =   2
            Begin VB.Menu ISO82 
               Caption         =   "Satisfaccion del Cliente"
               Index           =   1
            End
            Begin VB.Menu ISO82 
               Caption         =   "Auditoria Interna"
               Index           =   2
            End
            Begin VB.Menu ISO82 
               Caption         =   "Seguimiento y Medición de los Procesos"
               Index           =   3
            End
            Begin VB.Menu ISO82 
               Caption         =   "Seguimiento y Medicion de los Productos"
               Index           =   4
            End
         End
         Begin VB.Menu ISO8 
            Caption         =   "Control del Producto No Conforme"
            Index           =   3
         End
         Begin VB.Menu ISO8 
            Caption         =   "Analisis de Datos"
            Index           =   4
         End
         Begin VB.Menu ISO8 
            Caption         =   "Mejora"
            Index           =   5
            Begin VB.Menu ISO85 
               Caption         =   "Mejora Continua"
               Index           =   1
            End
            Begin VB.Menu ISO85 
               Caption         =   "Accion Correctiva"
               Index           =   2
            End
            Begin VB.Menu ISO85 
               Caption         =   "Accion Preventiva"
               Index           =   3
            End
         End
      End
   End
   Begin VB.Menu mnuwin 
      Caption         =   "Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascada 
         Caption         =   "Cascada"
      End
      Begin VB.Menu mnuMosaico 
         Caption         =   "Mosaico"
      End
      Begin VB.Menu mnuExcel 
         Caption         =   "Microsoft Excel"
      End
      Begin VB.Menu mnuExplorer 
         Caption         =   "Explorador de Windows"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "Exit"
   End
   Begin VB.Menu mnuPopmenu 
      Caption         =   "Popmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAgregar 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mnuQuitar 
         Caption         =   "Quitar"
      End
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sOpcion As String

Sub BorrarTablas()
On Error Resume Next

Dim Reg As New ADODB.Recordset
Set Reg = Nothing
Set Reg = Nothing
Reg.CursorLocation = adUseClient
Reg.Open "drop table cf_clie", cCONNECT

Set Reg = Nothing
Reg.CursorLocation = adUseClient
Reg.Open "drop table CF_DES", cCONNECT

Set Reg = Nothing
Reg.CursorLocation = adUseClient
Reg.Open "drop table cf_pedd", cCONNECT

Set Reg = Nothing
Reg.CursorLocation = adUseClient
Reg.Open "drop table cf_pedi", cCONNECT

Set Reg = Nothing
Reg.CursorLocation = adUseClient
Reg.Open "drop table CF_PEDR", cCONNECT
Set Reg = Nothing
End Sub


Sub CambiaCaptionMenu()
On Error GoTo hand
Dim ctl As Control
Dim Reg As New ADODB.Recordset
Reg.CursorLocation = adUseClient
Reg.Open "select Cod_Opcion,Des_Opcion,Des_Opcion_Eng from seg_opciones order by 1", conn.ConnectionString

If Reg.RecordCount > 0 Then
    For Each ctl In MDIPrincipal.Controls
        'Debug.Print ctl.Name
        If TypeOf ctl Is Menu Then
            If Mid(ctl.Name, 1, 3) <> "ISO" Then
                If iLanguage = 1 Then
                    If DevuelveCampo("select Des_Opcion from seg_opciones where Cod_Opcion='" & ctl.Name & "'", sconnect) <> "" Then
                        ctl.Caption = DevuelveCampo("select Des_Opcion from seg_opciones where Cod_Opcion='" & ctl.Name & "'", sconnect)
                    End If
                Else
                    If DevuelveCampo("select Des_Opcion_Eng from seg_opciones where Cod_Opcion='" & ctl.Name & "'", sconnect) <> "" Then
                        ctl.Caption = DevuelveCampo("select Des_Opcion_Eng from seg_opciones where Cod_Opcion='" & ctl.Name & "'", sconnect)
                    End If
                End If
            End If
        End If
    Next
End If
Set Reg = Nothing
Exit Sub
hand:
ErrorHandler Err, "CambiaCaptionMenu"
Set Reg = Nothing
End Sub


Private Sub ActEstCli_Click()
EjecutaOpcionMenu "ActEstCli", Me.perfil, Me.pEmpresa
End Sub

Private Sub Cierre_Click()
EjecutaOpcionMenu "Cierre", Me.perfil, Me.pEmpresa
End Sub

Private Sub ConCierre_Click()
EjecutaOpcionMenu "concierre", Me.perfil, Me.pEmpresa
End Sub

Private Sub CondVent_Click()
EjecutaOpcionMenu "CondVent", Me.perfil, Me.pEmpresa
End Sub

Private Sub DespPrendas_Click()
EjecutaOpcionMenu "DespPrend", Me.perfil, Me.pEmpresa
End Sub

Private Sub Dscto_Click()
EjecutaOpcionMenu "Dscto", Me.perfil, Me.pEmpresa
End Sub

Private Sub frmGruposReq_Click()
EjecutaOpcionMenu "frmgruposreq", Me.perfil, Me.pEmpresa
End Sub

Private Sub frmImpSit_Click()
EjecutaOpcionMenu "frmImpSit", Me.perfil, Me.pEmpresa
End Sub

Private Sub ISO4_Click(Index As Integer)

sOpcion = "ISO4" & CStr(Index)
EjecutaOpcionMenu sOpcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO42_Click(Index As Integer)
sOpcion = "ISO42" & CStr(Index)
EjecutaOpcionMenu sOpcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO5_Click(Index As Integer)
sOpcion = "ISO5" & CStr(Index)
EjecutaOpcionMenu sOpcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO54_Click(Index As Integer)
sOpcion = "ISO54" & CStr(Index)
EjecutaOpcionMenu sOpcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO55_Click(Index As Integer)
sOpcion = "ISO55" & CStr(Index)
EjecutaOpcionMenu sOpcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO56_Click(Index As Integer)
sOpcion = "ISO56" & CStr(Index)
EjecutaOpcionMenu sOpcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO6_Click(Index As Integer)
sOpcion = "ISO6" & CStr(Index)
EjecutaOpcionMenu sOpcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO62_Click(Index As Integer)
sOpcion = "ISO62" & CStr(Index)
EjecutaOpcionMenu sOpcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO7_Click(Index As Integer)
sOpcion = "ISO7" & CStr(Index)
EjecutaOpcionMenu sOpcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO72_Click(Index As Integer)
sOpcion = "ISO72" & CStr(Index)
EjecutaOpcionMenu sOpcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO73_Click(Index As Integer)
sOpcion = "ISO73" & CStr(Index)
EjecutaOpcionMenu sOpcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO74_Click(Index As Integer)
sOpcion = "ISO74" & CStr(Index)
EjecutaOpcionMenu sOpcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO75_Click(Index As Integer)
sOpcion = "ISO75" & CStr(Index)
EjecutaOpcionMenu sOpcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO8_Click(Index As Integer)
sOpcion = "ISO8" & CStr(Index)
EjecutaOpcionMenu sOpcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO82_Click(Index As Integer)
sOpcion = "ISO82" & CStr(Index)
EjecutaOpcionMenu sOpcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO85_Click(Index As Integer)
sOpcion = "ISO85" & CStr(Index)
EjecutaOpcionMenu sOpcion, Me.perfil, Me.pEmpresa

End Sub

'Private Sub ISO41_Click()
'    EjecutaOpcionMenu "ISO41", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO42_Click()
'    EjecutaOpcionMenu "ISO42", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO5_Click()
'    EjecutaOpcionMenu "ISO5", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO61_Click()
'    EjecutaOpcionMenu "ISO61", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO621_Click()
'    EjecutaOpcionMenu "ISO621", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO622_Click()
'    EjecutaOpcionMenu "ISO622", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO63_Click()
'    EjecutaOpcionMenu "ISO63", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO64_Click()
'    EjecutaOpcionMenu "ISO64", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO731_Click()
'    EjecutaOpcionMenu "ISO731", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO732_Click()
'        EjecutaOpcionMenu "ISO732", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO733_Click()
'    EjecutaOpcionMenu "ISO733", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO734_Click()
'    EjecutaOpcionMenu "ISO734", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO735_Click()
'    EjecutaOpcionMenu "ISO735", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO736_Click()
'    EjecutaOpcionMenu "ISO736", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO737_Click()
'    EjecutaOpcionMenu "ISO737", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO8_Click()
'    EjecutaOpcionMenu "ISO8", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOComCli_Click()
'    EjecutaOpcionMenu "ISOComCli", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOConDis_Click()
'EjecutaOpcionMenu "ISOConDis", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOControl_Click()
'EjecutaOpcionMenu "ISOControl", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISODDP_Click()
'
'End Sub
'
'Private Sub ISODetReq_Click()
'EjecutaOpcionMenu "ISODetReq", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOIdeTra_Click()
'EjecutaOpcionMenu "ISOIdeTra", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOInfor_Click()
'EjecutaOpcionMenu "ISOInfor", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOPlaRea_Click()
'EjecutaOpcionMenu "ISOPlaRea", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOPrePro_Click()
'EjecutaOpcionMenu "ISOPrePro", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOProce_Click()
'EjecutaOpcionMenu "ISOProce", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOProCli_Click()
'EjecutaOpcionMenu "ISOProCli", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISORevReq_Click()
'EjecutaOpcionMenu "ISORevReq", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOValPro_Click()
'EjecutaOpcionMenu "ISOValPro", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOVerPro_Click()
'EjecutaOpcionMenu "ISOVerPro", Me.perfil, Me.pEmpresa
'End Sub

Private Sub LugEntr_Click()
EjecutaOpcionMenu "LugEntr", Me.perfil, Me.pEmpresa
End Sub

Private Sub manban_Click()
EjecutaOpcionMenu "manBan", Me.perfil, Me.pEmpresa
End Sub

Private Sub mancli_Click()
EjecutaOpcionMenu "MANCLI", Me.perfil, Me.pEmpresa
End Sub

Private Sub mancomisin_Click()
EjecutaOpcionMenu "manComisin", Me.perfil, Me.pEmpresa
End Sub

Private Sub manDestino_Click()
EjecutaOpcionMenu "manDestino", Me.perfil, Me.pEmpresa
End Sub

Private Sub mandivpre_Click()
EjecutaOpcionMenu "mandivpre", Me.perfil, Me.pEmpresa
End Sub

Private Sub manDocs_Click()
EjecutaOpcionMenu "MANDocs", Me.perfil, Me.pEmpresa
End Sub

Private Sub manFab_Click()
EjecutaOpcionMenu "MANfab", Me.perfil, Me.pEmpresa
End Sub

Private Sub manfun_Click()
EjecutaOpcionMenu "manfun", Me.perfil, Me.pEmpresa
End Sub

Private Sub manMon_Click()
EjecutaOpcionMenu "manMon", Me.perfil, Me.pEmpresa
End Sub

Private Sub manmotatr_Click()
EjecutaOpcionMenu "manmotatr", Me.perfil, Me.pEmpresa
End Sub

Private Sub manopc_Click()
EjecutaOpcionMenu "manopc", Me.perfil, Me.pEmpresa
End Sub

Private Sub manorg_Click()
EjecutaOpcionMenu "manOrg", Me.perfil, Me.pEmpresa
End Sub

Private Sub manPagemb_Click()
EjecutaOpcionMenu "manPagEmb", Me.perfil, Me.pEmpresa
End Sub

Private Sub manper_Click()
EjecutaOpcionMenu "manper", Me.perfil, Me.pEmpresa
End Sub

Private Sub manPOObs_Click()
EjecutaOpcionMenu "manPoObs", Me.perfil, Me.pEmpresa
End Sub

Private Sub manseg_Click()
EjecutaOpcionMenu "MANSEG", Me.perfil, Me.pEmpresa
End Sub

Private Sub mantal_Click()
EjecutaOpcionMenu "manTal", Me.perfil, Me.pEmpresa
End Sub

Private Sub mantCargos_Click()
EjecutaOpcionMenu "MANtcargos", Me.perfil, Me.pEmpresa

End Sub

Private Sub mantGruTal_Click()
EjecutaOpcionMenu "mantgrutal", Me.perfil, Me.pEmpresa
End Sub

Private Sub mantHil_Click()
EjecutaOpcionMenu "manthil", Me.perfil, Me.pEmpresa
End Sub

Private Sub manTipEmb_Click()
EjecutaOpcionMenu "manTipEmb", Me.perfil, Me.pEmpresa
End Sub

Private Sub mantippre_Click()
EjecutaOpcionMenu "manTipPre", Me.perfil, Me.pEmpresa
End Sub

Private Sub mantra_Click()
EjecutaOpcionMenu "manTra", Me.perfil, Me.pEmpresa
End Sub

Private Sub mantitm_Click()
EjecutaOpcionMenu "mantitm", Me.perfil, Me.pEmpresa
End Sub

Private Sub mantItem_Click()
EjecutaOpcionMenu "mantitem", Me.perfil, Me.pEmpresa
End Sub

Private Sub MantTelas_Click()
EjecutaOpcionMenu "manttelas", Me.perfil, Me.pEmpresa
End Sub

Private Sub manunimed_Click()
EjecutaOpcionMenu "manUniMed", Me.perfil, Me.pEmpresa
End Sub

Private Sub manusu_Click()
EjecutaOpcionMenu "manusu", Me.perfil, Me.pEmpresa
End Sub

'Option Explicit
Private Sub MDIForm_Load()
Dim f As Form
' iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))
IdiomaEtiquetas1 Me
 Set f = Me
 f.Caption = Caption & "-" & NEmpresa
 get_accesos3 pEmpresa, perfil, f
 get_favoritos pEmpresa, pUsuario, f, iLanguage
 set_barra (iLanguage)
CambiaCaptionMenu
 'InitMessages 'C.A.R.
'FrmMantEmpUsuPer.Show
'FrmMantopciones.Show
 End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    conn.Close
    Set conn = Nothing
End Sub

Private Sub mnuBanco_Click()
PopupMenu mnuPopmenu
End Sub

Private Sub mnuClien_Click()
EjecutaOpcionMenu "MANCLI", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuDesti_Click()
'frmMotivos.Show
End Sub

Private Sub mnuCieMe_Click()

End Sub

Private Sub mnu1_Click()
EjecutaOpcionMenu "mnu1", Me.perfil, Me.pEmpresa
End Sub


Private Sub mnuactinf_Click()
EjecutaOpcionMenu "mnuactinf", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuactinfs_Click()
EjecutaOpcionMenu "mnuactinfs", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuActTic_Click()
EjecutaOpcionMenu "mnuActTic", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuanxcon_Click()
EjecutaOpcionMenu "mnuanxcon", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuAOpDia_Click()
EjecutaOpcionMenu "mnuAOpDia", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuAutorLT_Click()
EjecutaOpcionMenu "mnuAutorLT", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuAutorPg_Click()
EjecutaOpcionMenu "mnuAutorPg", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuautsob_Click()
EjecutaOpcionMenu "mnuautsob", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuAvaAca_Click()
EjecutaOpcionMenu "mnuAvaAca", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuAvaCn2_Click()
EjecutaOpcionMenu "mnuAvaCn2", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuAvaGer_Click()
EjecutaOpcionMenu "mnuAvaGer", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuBiHorE_Click()
EjecutaOpcionMenu "mnuBiHorE", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnubustel_Click()
EjecutaOpcionMenu "mnubustel", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuCancDoc_Click()
    EjecutaOpcionMenu "mnuCancDoc", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnucarcol_Click()
    EjecutaOpcionMenu "mnucarcol", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuCascada_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuCIPla_Click()
EjecutaOpcionMenu "mnuCIPla", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuclaitm_Click()
EjecutaOpcionMenu "mnuclaitm", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuclasoc_Click()
EjecutaOpcionMenu "mnuclasoc", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuComp_Click()
EjecutaOpcionMenu "mnucomp", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuComRep_Click()
EjecutaOpcionMenu "mnuComRep", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuconcep_Click()
EjecutaOpcionMenu "mnuconcep", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuConCot_Click()
EjecutaOpcionMenu "mnuConCot", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuConFa_Click()
EjecutaOpcionMenu "ConsFact", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuConFF_Click()
EjecutaOpcionMenu "mnuConFF", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuCot_Click()

End Sub

Private Sub mnuConLet_Click()
EjecutaOpcionMenu "mnuConLet", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuconsal_Click()
EjecutaOpcionMenu "mnuconsal", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuConsAut_Click()
EjecutaOpcionMenu "mnuConsAut", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuconsumo_Click()
EjecutaOpcionMenu "mnuconsumo", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnucosconf_Click()
EjecutaOpcionMenu "mnucosconf", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnucosgru_Click()
EjecutaOpcionMenu "mnucosgru", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuCosSem_Click()
EjecutaOpcionMenu "mnuCosSem", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuCotiza_Click()
EjecutaOpcionMenu "mnuCotiza", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuCtPrMe_Click()
EjecutaOpcionMenu "mnuCtPrMe", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnudatec_Click()
EjecutaOpcionMenu "mnudatec", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuDeliv_Click()
GeneraReportes DeliverySummary
End Sub

Private Sub mnupocolest_Click()
EjecutaOpcionMenu "mnupocol", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuDeOpe_Click()
EjecutaOpcionMenu "MnuDeOpe", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuEfic_Click()
EjecutaOpcionMenu "mnuEfic", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuEfiDia_Click()
EjecutaOpcionMenu "mnuEfiDia", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuEntPen_Click()
EjecutaOpcionMenu "mnuEntPen", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuentreg_Click()
EjecutaOpcionMenu "mnuentreg", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuESemCS_Click()
EjecutaOpcionMenu "mnuESemCS", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuEstClC_Click()
EjecutaOpcionMenu "mnuEstClC", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuEstDia_Click()
EjecutaOpcionMenu "mnuEstDia", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuestprop_Click()
EjecutaOpcionMenu "mnuestprop", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuEstSem_Click()
EjecutaOpcionMenu "mnuEstSem", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuetiq_Click()
EjecutaOpcionMenu "mnuetiq", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuExcel_Click()
    Shell "C:\Archivos de programa\Microsoft Office\Office10\excel.EXE", vbNormalFocus
End Sub

Private Sub mnuExplorer_Click()
    Shell "explorer.exe", vbNormalFocus
End Sub

Private Sub mnuextmar_Click()
EjecutaOpcionMenu "mnuextmar", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnufacdet_Click()
EjecutaOpcionMenu "mnufacdet", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuFacPro_Click()
EjecutaOpcionMenu "mnuFacPro", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuFlujopd_Click()
EjecutaOpcionMenu "mnuFlujopd", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuforecast_Click()
    GeneraReportes Forecast
End Sub

Private Sub mnugalgas_Click()
EjecutaOpcionMenu "mnugalgas", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnugamas_Click()
EjecutaOpcionMenu "mnugamas", Me.perfil, Me.pEmpresa
End Sub



Private Sub mnuGenCoa_Click()
EjecutaOpcionMenu "mnuGenCoa", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnugenmar_Click()
EjecutaOpcionMenu "mnugenmar", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuGrupo_Click()
EjecutaOpcionMenu "mnuGrupo", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuGrupoL_Click()
EjecutaOpcionMenu "mnuGrupoL", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnugrupreg_Click()
EjecutaOpcionMenu "mnugrupreg", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuGrupoT_Click()
EjecutaOpcionMenu "mnuGrupoT", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuGruPro_Click()
EjecutaOpcionMenu "mnuGruPro", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuguiaman_Click()
EjecutaOpcionMenu "mnuguiaman", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuguias_Click()
EjecutaOpcionMenu "mnuguias", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuhilcru_Click()
EjecutaOpcionMenu "mnuhilcru", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuhilten_Click()
EjecutaOpcionMenu "mnuhilten", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuhorario_Click()
EjecutaOpcionMenu "mnuhorario", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuHorStk_Click()
EjecutaOpcionMenu "mnuHorStk", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuhortra_Click()
EjecutaOpcionMenu "mnuhortra", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuImpMas_Click()
EjecutaOpcionMenu "mnuimpmas", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuIndSem_Click()
EjecutaOpcionMenu "mnuIndSem", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuintcol_Click()
EjecutaOpcionMenu "mnuintcol", Me.perfil, Me.pEmpresa
End Sub


Private Sub mnulist_Click()
EjecutaOpcionMenu "mnulist", Me.perfil, Me.pEmpresa
End Sub



Private Sub mnukardex_Click()
EjecutaOpcionMenu "mnukardex", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuKarMTA_Click()
EjecutaOpcionMenu "mnuKarMTA", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuKarTer_Click()
EjecutaOpcionMenu "mnuKarTer", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuLecEsp_Click()
EjecutaOpcionMenu "mnuLecEsp", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuLecTic_Click()
EjecutaOpcionMenu "mnuLecTic", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuLetra_Click()
EjecutaOpcionMenu "mnuletra", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuLinPro_Click()
EjecutaOpcionMenu "mnuLinPro", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnumantalm_Click()
EjecutaOpcionMenu "mnumantalm", Me.perfil, Me.pEmpresa
End Sub


Private Sub mnumantp_Click()
EjecutaOpcionMenu "mnumantp", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuMatReq_Click()
EjecutaOpcionMenu "mnuMatReq", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuMosaico_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnumot_Click()
EjecutaOpcionMenu "mnumot", Me.perfil, Me.pEmpresa
End Sub


Private Sub mnumovalm_Click()
EjecutaOpcionMenu "mnumovalm", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuMovBan_Click()
EjecutaOpcionMenu "mnuMovBan", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuMovCnf_Click()
EjecutaOpcionMenu "mnuMovCnf", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnumovperm_Click()
EjecutaOpcionMenu "mnumovperm", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnumovsal_Click()
EjecutaOpcionMenu "mnumovsal", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnumovsto_Click()
EjecutaOpcionMenu "mnumovsto", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnumstock_Click()
    EjecutaOpcionMenu "mnumstock", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnumpr_Click()
EjecutaOpcionMenu "mnumpr", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuND_Click()
EjecutaOpcionMenu "mnuND", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuNumDoc_Click()
EjecutaOpcionMenu "mnuNumDoc", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuOrdCnf_Click()
EjecutaOpcionMenu "mnuOrdCnf", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuOrdComp_Click()
EjecutaOpcionMenu "mnuOrdComp", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuordcort_Click()
EjecutaOpcionMenu "mnuordcort", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuOrdPro_Click()
EjecutaOpcionMenu "mnuOrdPro", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuorig_Click()
EjecutaOpcionMenu "mnuorig", Me.perfil, Me.pEmpresa
End Sub


Private Sub mnuPakLis_Click()
EjecutaOpcionMenu "mnuPakLis", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuParTelB_Click()
EjecutaOpcionMenu "mnuParTelB", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuPartida_Click()
EjecutaOpcionMenu "mnuPartida", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuPDCorte_Click()
EjecutaOpcionMenu "mnuPDCorte", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuPiezas_Click()
EjecutaOpcionMenu "mnupiezas", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnupoco_Click()
EjecutaOpcionMenu "Mnupoco", Me.perfil, Me.pEmpresa
End Sub


Private Sub mnuposcort_Click()
EjecutaOpcionMenu "mnuposcort", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuprecio_Click()
EjecutaOpcionMenu "mnuprecio", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuPrePro_Click()
EjecutaOpcionMenu "mnuPrePro", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuPrgRea_Click()
EjecutaOpcionMenu "mnuPrgRea", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuProceso_Click()
EjecutaOpcionMenu "mnuProceso", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuProdMen_Click()
    EjecutaOpcionMenu "mnuProdMen", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuproest_Click()
EjecutaOpcionMenu "mnuproest", Me.perfil, Me.pEmpresa
End Sub


Private Sub mnuproform_Click()
EjecutaOpcionMenu "mnuproform", Me.perfil, Me.pEmpresa
End Sub


Private Sub mnuProHab_Click()
EjecutaOpcionMenu "mnuProHab", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuProPro_Click()
EjecutaOpcionMenu "mnuProPro", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuRegCom_Click()
EjecutaOpcionMenu "mnuRegCom", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuRegDir_Click()
EjecutaOpcionMenu "mnuRegDir", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuRegIn_Click()
    Dim frmShowTG_PurOrd1 As frmShowTG_PurOrd
    
    Set frmShowTG_PurOrd1 = New frmShowTG_PurOrd
    Load frmShowTG_PurOrd1
    Set frmShowTG_PurOrd1.oParent = Me
    frmShowTG_PurOrd1.Show
    
End Sub

Private Sub mnuRepDRB_Click()
EjecutaOpcionMenu "mnuRepDRB", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnurvsreal_Click()
EjecutaOpcionMenu "mnurvsreal", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuResAct_Click()
EjecutaOpcionMenu "mnuResAct", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuResAsi_Click()
EjecutaOpcionMenu "mnuResAsi", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuResDe_Click()
GeneraReportes TrackingReporteDetail
End Sub

Private Sub mnureverr_Click()
EjecutaOpcionMenu "mnureverr", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuRptIna_Click()
EjecutaOpcionMenu "mnuRptIna", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuSalir_Click()
End
Unload Me
End Sub


Private Sub mnuSeman_Click()

End Sub

Private Sub mnusalnreg_Click()
EjecutaOpcionMenu "mnusalnreg", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuSecCon_Click()
EjecutaOpcionMenu "mnuSecCon", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuServTen_Click()
EjecutaOpcionMenu "mnuServTen", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnusolid_Click()
EjecutaOpcionMenu "mnusolid", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuSPProv_Click()
EjecutaOpcionMenu "mnuSPProv", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuStkCor_Click()
EjecutaOpcionMenu "mnuStkCor", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuStkCos_Click()
EjecutaOpcionMenu "mnuStkCos", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnustkfam_Click()
EjecutaOpcionMenu "mnustkfam", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuStocks_Click()
EjecutaOpcionMenu "mnuStocks", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnusubmar_Click()
EjecutaOpcionMenu "mnusubmar", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutarifado_Click()
EjecutaOpcionMenu "mnutarifa", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutelaca_Click()
EjecutaOpcionMenu "mnutelaca", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutelcru_Click()
EjecutaOpcionMenu "mnutelcru", Me.perfil, Me.pEmpresa
End Sub




Private Sub mnuTelSer_Click()
EjecutaOpcionMenu "mnuTelSer", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutipanx_Click()
EjecutaOpcionMenu "mnutipanx", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuTipCam_Click()
EjecutaOpcionMenu "mnutipcam", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuTipComp_Click()
EjecutaOpcionMenu "mnutipcomp", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutipdoc_Click()
EjecutaOpcionMenu "mnutipdoc", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutipmov_Click()
EjecutaOpcionMenu "mnutipmov", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutippro_Click()
EjecutaOpcionMenu "mnutippro", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutipraya_Click()
EjecutaOpcionMenu "mnutipraya", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutiprec_Click()
EjecutaOpcionMenu "mnutiprec", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutit_Click()
EjecutaOpcionMenu "mnutit", Me.perfil, Me.pEmpresa
End Sub


Private Sub mnutrab_Click()
EjecutaOpcionMenu "mnutrab", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuTrans_Click()
EjecutaOpcionMenu "mnuTrans", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuTraPag_Click()
EjecutaOpcionMenu "mnuTraPag", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuultimos_Click()
On Error Resume Next
Screen.MousePointer = vbHourglass
DoEvents
'BorrarTablas
'EjecutaDBF2SQL
Dim Reg As New ADODB.Recordset
Reg.CursorLocation = adUseClient
Reg.Open "up_migracion", cCONNECT
'EjecutaMigracionSQLtoDBF2
EjecutaMigracionSQLtoDBF2
'BorrarTablas
Set Reg = Nothing
MsgBox "El proceso ha terminado", vbInformation
Screen.MousePointer = vbDefault
Exit Sub
hand:
Set Reg = Nothing
ErrorHandler Err, "mnuultimos_Click"
Screen.MousePointer = vbDefault
End Sub


Public Sub mnuWizPO_Click()
'    Dim frmNewWizard As frmWizard
'    Set frmNewWizard = New frmWizard
'    Load frmNewWizard
'    Set frmNewWizard.oParent = Me
'    frmNewWizard.Show
End Sub

Private Sub RepDelDet_Click()
EjecutaOpcionMenu "REPDELDET", Me.perfil, Me.pEmpresa
End Sub

Private Sub RepTra_Click()
EjecutaOpcionMenu "REPTRA", Me.perfil, Me.pEmpresa
End Sub

Private Sub StaOrdComp_Click()
EjecutaOpcionMenu "StaOrdComp", Me.perfil, Me.pEmpresa
End Sub

Private Sub TipOrdComp_Click()
EjecutaOpcionMenu "TipOrdComp", Me.perfil, Me.pEmpresa
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    ' PopupMenu mnuPopmenu

   ' Select Case Button.Key
   '     Case "PRINT"
   '         Me.ActiveForm.Imprimir
   '     Case "CLOSE"
   '         Me.ActiveForm.Cerrar
   '     Case "EXIT"
   '         Unload Me
   ' End Select
End Sub

Public Property Get pUsuario() As Variant
pUsuario = vusu
End Property

Public Property Let pUsuario(ByVal vnuevo As Variant)
vusu = vnuevo
End Property

Public Property Get pEmpresa() As Variant
pEmpresa = vemp
End Property

Public Property Let pEmpresa(ByVal vnuevo1 As Variant)
vemp = vnuevo1
End Property

Public Property Get PClave() As Variant
PClave = vpas
End Property

Public Property Let PClave(ByVal vnuevo2 As Variant)
vpas = vnuevo2
End Property
Public Property Get perfil() As Variant
perfil = vper
End Property

Public Property Let perfil(ByVal vnuevo3 As Variant)
vper = vnuevo3
End Property
Private Function get_accesos3(ByVal vcod_empresa As Variant, ByVal Vcod_perfil As Variant, ByVal f As Form)
On Error GoTo procesaerror
'on Error Resume Next
Dim RS1 As ADODB.Recordset
Dim RS2 As ADODB.Recordset
Dim sQuery As String
Dim j As Integer
Dim vCod_App As String

Set RS1 = New ADODB.Recordset
RS1.CursorLocation = adUseClient
sQuery = "SELECT * FROM SEG_ADMINISTRACION WHERE COD_PERFIL='" & Vcod_perfil & "'  AND COD_EMPRESA='" & vcod_empresa & "'"
'RS1.ActiveConnection = conn
RS1.Open sQuery, conn.ConnectionString

Set RS2 = New ADODB.Recordset
RS2.CursorLocation = adUseClient
'Opciones tipo Carpeta
'RS2.ActiveConnection = conn
If Not (RS1.BOF And RS1.EOF) Then
    For j = 1 To RS1.RecordCount
        vCod_App = RS1!COD_APLICACION
        RS2.Open "Sp_opciones2 '" & vCod_App & "','" & Vcod_perfil & "','" & vcod_empresa & "'", conn.ConnectionString
        If Not (RS2.BOF And RS2.EOF) Then
          RS2.MoveFirst
           While Not RS2.EOF
            mnu_invisible RS2!Cod_opcion, f
            RS2.MoveNext
           Wend
        End If
        RS2.Close
        RS1.MoveNext
    Next j
End If
RS1.Close
'Desactivar Aplicaciones no autorizadas
sQuery = "SELECT NOM_MENU FROM SEG_APLICACION WHERE COD_APLICACION NOT IN (SELECT distinct(cod_aplicacion) FROM SEG_ADMINISTRACION WHERE COD_PERFIL='" & Vcod_perfil & "'  AND COD_EMPRESA='" & vcod_empresa & "')"
RS1.Open sQuery
If Not (RS1.BOF And RS1.EOF) Then
    For j = 1 To RS1.RecordCount
        mnu_invisible RS1!nom_menu, f
    RS1.MoveNext
    Next j
End If
Set RS1 = Nothing
Set RS2 = Nothing

Exit Function

procesaerror:
ErrorHandler Err, "get_accesos3"

End Function
Private Sub mnu_invisible(ByVal sname As Variant, ByVal f As Form)
Dim ctl As Control, mnu As Menu
For Each ctl In f.Controls
        If TypeOf ctl Is Menu Then
            If LTrim(RTrim(UCase(sname))) = LTrim(RTrim(UCase(ctl.Name))) Then
                ctl.Visible = False
                Exit For
            End If
        End If
  Next ctl
End Sub
Private Sub mnu_OPCION(ByVal f As Form)
'Captura los name y caption del menu y los inserta en la tabla Tmp_Opcion
Dim ctl As Control, mnu As Menu
For Each ctl In f.Controls
        If TypeOf ctl Is Menu Then

                xname = ctl.Name
                xcaption = ctl.Caption
                sQuery = "insert into tmp_opcion (name,caption) values ('" & xname & "','" & xcaption & "')"
                conn.Execute sQuery
            'End If
        End If
  Next ctl
End Sub
Private Function get_favoritos(ByVal vcod_empresa As Variant, ByVal Vcod_usuario As Variant, ByVal f As Form, ByVal iLanguage As String)
Set RS1 = New ADODB.Recordset
sQuery = "SELECT A.COD_OPCION,A.ICONO,A.DES_OPCION,A.DES_OPCION_ENG  FROM SEG_OPCIONES A,SEG_FAVORITOS B WHERE A.COD_OPCION=B.COD_OPCION AND B.COD_USUARIO='" & Vcod_usuario & "'  AND B.COD_EMPRESA='" & vcod_empresa & "'"
RS1.ActiveConnection = conn
RS1.CursorType = adOpenStatic
RS1.Open sQuery
If Not (RS1.BOF And RS1.EOF) Then
  With Toolbar1
    For j = 1 To RS1.RecordCount
      xkey = LTrim(RTrim(RS1!Cod_opcion))
      ximg = LCase(RS1!icono)
      If iLanguage = "1" Then
      xtip = RS1!des_opcion
      Else
      xtip = RS1!des_opcion_eng
      End If
      .Buttons.Add j, xkey, "", , ximg
      .Buttons.Item(j).ToolTipText = xtip
      RS1.MoveNext
    Next j
  End With
End If
End Function
Private Sub mnuContext_Click()
   If mnuContext.Caption = "Agregar" Then
      mnuContext.Caption = "Quitar"
   Else
      mnuContext.Caption = "Agregar"
   End If
End Sub

Private Sub Toolbar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button = vbRightButton Then
 PopupMenu mnuPopmenu
 End If
End Sub

Public Property Get NEmpresa() As Variant
NEmpresa = vemp1
End Property

Public Property Let NEmpresa(ByVal vnuevo1 As Variant)
vemp1 = vnuevo1
End Property

Private Sub set_barra(iLanguage As String)
Dim Pan As Panel
 For Each Panel In StatusBar1.Panels
   If iLanguage = "2" Then
       Panel.Text = Panel.Tag
   End If
 Next Panel
 StatusBar1.Panels.Item(1).Text = StatusBar1.Panels.Item(1).Text & NEmpresa
 StatusBar1.Panels.Item(2).Text = StatusBar1.Panels.Item(2).Text & pUsuario
 StatusBar1.Panels.Item(4).Text = StatusBar1.Panels.Item(4).Text & ComputerName
 StatusBar1.Panels.Item(3).Text = StatusBar1.Panels.Item(3).Text & Fecha_Hora_Conexion
End Sub
