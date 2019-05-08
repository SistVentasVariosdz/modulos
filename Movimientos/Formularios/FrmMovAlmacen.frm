VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmMovAlmacen 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Movimientos de Almacen"
   ClientHeight    =   8715
   ClientLeft      =   2235
   ClientTop       =   465
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   13590
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   730
      Left            =   60
      TabIndex        =   14
      Top             =   0
      Width           =   13515
      Begin VB.CommandButton Command2 
         Caption         =   "&Buscar"
         Height          =   465
         Left            =   11880
         TabIndex        =   21
         Top             =   200
         Width           =   1515
      End
      Begin VB.ComboBox CmbAlmacen 
         Height          =   315
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Width           =   4020
      End
      Begin VB.TextBox TxtMov 
         BackColor       =   &H00FFFFFF&
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
         Left            =   5880
         MaxLength       =   6
         TabIndex        =   16
         Top             =   240
         Width           =   1155
      End
      Begin MSComCtl2.DTPicker DtFecha 
         Height          =   315
         Left            =   7680
         TabIndex        =   15
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   69140481
         CurrentDate     =   37270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Almacen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nro Mov :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   5070
         TabIndex        =   19
         Tag             =   "Hilado :"
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   7080
         TabIndex        =   18
         Tag             =   "Hilado :"
         Top             =   255
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8055
      Left            =   12000
      TabIndex        =   12
      Top             =   630
      Width           =   1575
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   7830
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   13811
         Custom          =   $"FrmMovAlmacen.frx":0000
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1400
         ControlHeigth   =   450
         ControlSeparator=   40
      End
   End
   Begin VB.Frame frCambioFecha 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cambio Fecha Movimiento"
      Height          =   1680
      Left            =   3120
      TabIndex        =   6
      Top             =   3960
      Visible         =   0   'False
      Width           =   4155
      Begin VB.CommandButton cmdEnd 
         Caption         =   "Salir"
         Height          =   525
         Left            =   2220
         TabIndex        =   10
         Top             =   975
         Width           =   1245
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   525
         Left            =   660
         TabIndex        =   9
         Top             =   975
         Width           =   1245
      End
      Begin MSComCtl2.DTPicker dtpNueFecMov 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   69140481
         CurrentDate     =   37270
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fecha:"
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
         Left            =   840
         TabIndex        =   8
         Tag             =   "Hilado :"
         Top             =   450
         Width           =   495
      End
   End
   Begin VB.Frame fraImpresion 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Modo de Impresión"
      Height          =   1680
      Left            =   3120
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   4155
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   525
         Left            =   660
         TabIndex        =   5
         Top             =   1095
         Width           =   1245
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   525
         Left            =   2280
         TabIndex        =   4
         Top             =   1095
         Width           =   1245
      End
      Begin VB.OptionButton optDetallado 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Detallado"
         Height          =   225
         Left            =   510
         TabIndex        =   3
         Top             =   690
         Width           =   2565
      End
      Begin VB.OptionButton optResumido 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Resumido"
         Height          =   225
         Left            =   510
         TabIndex        =   2
         Top             =   330
         Value           =   -1  'True
         Width           =   2565
      End
   End
   Begin VB.Frame Fralista 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   60
      TabIndex        =   0
      Tag             =   "List"
      Top             =   630
      Width           =   11895
      Begin GridEX20.GridEX DGridLista 
         Height          =   7740
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   13653
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         BackColorBkg    =   12648384
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "FrmMovAlmacen.frx":0569
         Column(2)       =   "FrmMovAlmacen.frx":0631
         FormatStylesCount=   6
         FormatStyle(1)  =   "FrmMovAlmacen.frx":06D5
         FormatStyle(2)  =   "FrmMovAlmacen.frx":080D
         FormatStyle(3)  =   "FrmMovAlmacen.frx":08BD
         FormatStyle(4)  =   "FrmMovAlmacen.frx":0971
         FormatStyle(5)  =   "FrmMovAlmacen.frx":0A49
         FormatStyle(6)  =   "FrmMovAlmacen.frx":0B01
         ImageCount      =   0
         PrinterProperties=   "FrmMovAlmacen.frx":0BE1
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   465
      Top             =   7005
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmMovAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Estado As String
Public Paso As String
Public CODIGO As String
Public DESCRIPCION As String
Public bOk  As Boolean
Public vEstImp As String
Public vMotivo As String

Public sCod_AlmacenOrigen As String
Public sNum_MovStkOrigen As String

Dim Tip_Accion
Dim Cod_TipOrdPro
Dim Cod_TipAnx
Dim Cod_ClaOrdComp
Dim Num_MovStk As String
Dim Cod_Fabrica
Public Almacen
Dim Flg_Rollo As String

Dim Cod_ClaMov  As String
Dim Cod_TipOrdTra

Dim Tip_item As String
Dim Tip_presentacion As String
Dim sguia As String



'Variable creada por AHSP
Dim sTipo As String
Dim varNum_Mov As String
Dim strSQL As String
Public varNum_SecOrd As String
Public varCod_Fabrica As String

Public varCod_TipOrdTra As String
Public varCod_Proveedor As String
Public varCod_OrdTra As String
Public varCod_color As String

Dim indicegrilla  As Long

Sub Reporte_Guia(sDoc As String)
On Error GoTo hand
Dim rs As New ADODB.Recordset
Dim vMessage As Variant
Dim vResp As String, sTit As String

'Doc = "Guia" ó Doc = "Parte"
sTit = "Guia de Remision"
If sDoc = "Parte" Then sTit = "Parte de Salida"

vMessage = (MsgBox("¿Es transportada por el mismo?", vbYesNo, sTit))
If vMessage = vbNo Then
    vResp = "N"
    rs.Open "select * from seguridad..seg_Empresas where cod_empresa='" & vemp1 & "'", cConnect, adOpenStatic
    If Not rs.EOF Then
        With frmDatosAdicionales
            .TxtTransportista.Text = rs!Des_Empresa
            .TxtDomicilio.Text = rs!Direccion
            .TxtRuc.Text = rs!Num_Ruc
        End With
    End If
    rs.Close

Else
    vResp = "S"
    rs.Open "select des_proveedor,dom_proveedor,num_ruc from lg_proveedor where cod_proveedor='" & DGridLista.Value(DGridLista.Columns("Cod_Proveedor").Index) & "'", cConnect, adOpenStatic
    If Not rs.EOF Then
        With frmDatosAdicionales
            .TxtTransportista = rs!Des_Proveedor
            .TxtDomicilio = rs!dom_proveedor
            .TxtRuc = rs!Num_Ruc
        End With
    End If
    rs.Close
End If

With frmDatosAdicionales
    .Caption = sTit
    .sDoc = sDoc
    .NumMovStk = DGridLista.Value(DGridLista.Columns("num_movstk").Index)
    .CodAlmacen = DGridLista.Value(DGridLista.Columns("cod_almacen").Index)
    .CodProveedor = Trim(DGridLista.Value(DGridLista.Columns("Cod_Proveedor").Index))
    .CodCenCost = DGridLista.Value(DGridLista.Columns("Cod_CenCost").Index)
    .Ser_OrdComp = DGridLista.Value(DGridLista.Columns("Ser_OrdComp").Index)
    .Cod_OrdComp = DGridLista.Value(DGridLista.Columns("Cod_OrdComp").Index)
    .vRespuesta = vResp
    '.vNumConosHilos = DGridLista.Value(DGridLista.Columns("NRO_CONOS_HILOS_COSER").Index)
    .varMoviStk_Guia = False
    .Show 1
End With

Set frmDatosAdicionales = Nothing

Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
End Sub

Sub Datos(Accion As String, EsAccion As Boolean)
On Error GoTo hand
bOk = False

strSQL = "UP_Lg_Movstk '" & UCase(Accion) & "','" & IIf(Accion = "V", Trim(Right(Me.CmbAlmacen, 2)), IIf(UCase(Accion) = "I", Trim(Right(Me.CmbAlmacen, 2)), Almacen)) & "','" & TxtMov & "','" & _
         DtFecha.Value & "','" & vusu & "','','','','','','','','','','','',0,'','','',''"

Set DGridLista.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)

If EsAccion = False Then
    DGridLista.Columns("Cod_TipMov").Visible = False
    DGridLista.Columns("Cod_Proveedor").Visible = False
    DGridLista.Columns("Cod_CenCost").Visible = False
    DGridLista.Columns("Cod_Cliente").Visible = False
    DGridLista.Columns("Num_MovStk").Visible = False
    
    DGridLista.Columns("Des_TipMov").Visible = False
    DGridLista.Columns("Des_Proveedor").Visible = False
    DGridLista.Columns("Cod_Ordpro").Visible = False
    DGridLista.Columns("Nom_Cliente").Visible = False
    DGridLista.Columns("Cod_TipOrdpRO").Visible = False
    DGridLista.Columns("Cod_Almacen").Visible = False
    
    DGridLista.Columns("Ser_OrdComp").Visible = False
    DGridLista.Columns("Cod_Almacen").Visible = False
    DGridLista.Columns("Cod_OrdComp").Visible = False
    DGridLista.Columns("Cod_Fabrica").Visible = False
    DGridLista.Columns("Cod_TipOrdTra").Visible = False
    DGridLista.Columns("Tip_PtMp").Visible = False
'    DGridLista.Columns("flg_adicionales").Visible = False
    DGridLista.Columns("Ser_Docum").Visible = False
    DGridLista.Columns("Num_Docum").Visible = False
    DGridLista.Columns("Usuario_Valorizo").Visible = False
    DGridLista.Columns("Cod_TipOrdTra1").Visible = False
    DGridLista.Columns("Cod_OrdTra1").Visible = False
    DGridLista.Columns("Fecha Creacion").Visible = False
    'DGridLista.Columns("Nombre_Solicitante").Width = 2500
    
    DGridLista.Columns("Num. Mov").Width = 800
    DGridLista.Columns("Fecha Mov").Width = 1100
    DGridLista.Columns("tipo mov").Width = 2700
    
    DGridLista.Columns("Num_MovStk_2da").Caption = "Num.2da"
    
    DGridLista.FrozenColumns = 4
    DGridLista.RowSelected(indicegrilla) = True
    
End If

bOk = True

Exit Sub
hand:
ErrorHandler err, "Datos"
'Set Reg = Nothing
End Sub


Sub ShowHiloCrudo()

Dim vAux As Variant
    
    FrmDetalleHilCru.Cod_OrdPro = DGridLista.Value(DGridLista.Columns("Cod_Ordpro").Index)
    'FrmDetalleHilCru.Cod_OrdTra = DevuelveCampo("select Cod_Ordtra from tx_ordtra where Cod_TipOrdTra='" & Reg("Cod_TipOrdTra") & "' and Cod_Proveedor='" & Reg("Cod_Proveedor") & "' and Cod_OrdProv='" & Reg("Lote Prov") & "'", cCONNECT)
    FrmDetalleHilCru.Cod_TipOrdTra = DGridLista.Value(DGridLista.Columns("Cod_TipOrdTra").Index)
    FrmDetalleHilCru.Fec_MOVsTK = DGridLista.Value(DGridLista.Columns("Fecha Mov").Index)
    FrmDetalleHilCru.Cod_TipOrdPro = DGridLista.Value(DGridLista.Columns("Cod_TipOrdPro").Index)
    FrmDetalleHilCru.Cod_Proveedor = DGridLista.Value(DGridLista.Columns("Cod_Proveedor").Index)
    'FrmDetalleHilCru.Sec_OrdComp = DevuelveCampo("select max(Sec_OrdComp) from lg_ordcompitem where Ser_OrdComp='" & Reg("Ser_OrdComp") & "' and Cod_OrdComp='" & Reg("Cod_OrdComp") & "'", cCONNECT)
    FrmDetalleHilCru.Cod_ClaMov = Cod_ClaMov
    FrmDetalleHilCru.Cod_TipMovi = DGridLista.Value(DGridLista.Columns("Cod_TipMov").Index)
    FrmDetalleHilCru.Cod_Almacen = Almacen
    FrmDetalleHilCru.Cod_ClaOrdComp = Cod_ClaOrdComp
    FrmDetalleHilCru.Num_MovStk = DGridLista.Value(DGridLista.Columns("Num_MovStk").Index)
    FrmDetalleHilCru.Cod_OrdComp = DGridLista.Value(DGridLista.Columns("Cod_OrdComp").Index)
    FrmDetalleHilCru.Ser_OrdComp = DGridLista.Value(DGridLista.Columns("Ser_OrdComp").Index)
    FrmDetalleHilCru.Tip_PtMp = DGridLista.Value(DGridLista.Columns("Tip_PtMp").Index)
    
    FrmDetalleHilCru.Cod_TipOrdTra1 = DGridLista.Value(DGridLista.Columns("Cod_TipOrdTra1").Index)
    FrmDetalleHilCru.Cod_OrdTra1 = DGridLista.Value(DGridLista.Columns("Cod_OrdTra1").Index)
    
    vAux = DevuelveCampo("select flg_partida_generada from lg_tiposmov where Cod_TipMov='" & DGridLista.Value(DGridLista.Columns("cod_tipmov").Index) & "'", cConnect)
    FrmDetalleHilCru.Flg_Partida_Generada = IIf(IsNull(vAux), "", vAux)
    
    strSQL = "select ISNULL(flg_ot_tejeduria_generada, 'N') from lg_tiposmov where Cod_TipMov='" & Trim(DGridLista.Value(DGridLista.Columns("cod_tipmov").Index)) & "'"
    FrmDetalleHilCru.sFlg_Ot_Tejeduria_Generada = DevuelveCampo(strSQL, cConnect)
    
    FrmDetalleHilCru.lblCod_OrdTra.Visible = (FrmDetalleHilCru.sFlg_Ot_Tejeduria_Generada = "S")
    FrmDetalleHilCru.txtCod_OrdTra.Visible = (FrmDetalleHilCru.sFlg_Ot_Tejeduria_Generada = "S")
    FrmDetalleHilCru.txtCod_OrdTra = DGridLista.Value(DGridLista.Columns("Cod_OrdTra1").Index)
    
    FrmDetalleHilCru.txtLote_Destino.Visible = (FrmDetalleHilCru.Cod_TipOrdPro = "TS")
    
    FrmDetalleHilCru.varValida_Factura = Valida_Factura
    
    FrmDetalleHilCru.Show 1
    
    Set FrmDetalleHilCru = Nothing
    
End Sub


Sub ShowHiloTenido()
Dim vAux As Variant
'    Almacen = DGridLista.Value(DGridLista.Columns("Cod_Almacen").Index)
    
    Cod_Fabrica = DGridLista.Value(DGridLista.Columns("Cod_Fabrica").Index)
    Cod_TipOrdTra = DGridLista.Value(DGridLista.Columns("Cod_TipOrdTra").Index)
    
    FrmDetalleHilTel.Cod_OrdPro = DGridLista.Value(DGridLista.Columns("Cod_Ordpro").Index)
    'FrmDetalleHilTel.Cod_OrdTra = DevuelveCampo("select Cod_Ordtra from tx_ordtra where Cod_TipOrdTra='" & Reg("Cod_TipOrdTra") & "' and Cod_Proveedor='" & Reg("Cod_Proveedor") & "' and Cod_OrdProv='" & Reg("Lote Prov") & "'", cCONNECT)
    FrmDetalleHilTel.Cod_TipOrdTra = DGridLista.Value(DGridLista.Columns("Cod_TipOrdTra").Index)
    FrmDetalleHilTel.Fec_MOVsTK = DGridLista.Value(DGridLista.Columns("Fecha Mov").Index)

    FrmDetalleHilTel.Cod_TipOrdPro = DGridLista.Value(DGridLista.Columns("Cod_TipOrdPro").Index)
    FrmDetalleHilTel.Cod_Proveedor = DGridLista.Value(DGridLista.Columns("Cod_Proveedor").Index)
    'FrmDetalleHilTel.Sec_OrdComp = DevuelveCampo("select max(Sec_OrdComp) from lg_ordcompitem where Ser_OrdComp='" & Reg("Ser_OrdComp") & "' and Cod_OrdComp='" & Reg("Cod_OrdComp") & "'", cCONNECT)
    FrmDetalleHilTel.Cod_ClaMov = Cod_ClaMov
    FrmDetalleHilTel.Cod_TipMovi = DGridLista.Value(DGridLista.Columns("Cod_TipMov").Index)
    FrmDetalleHilTel.Cod_Almacen = Almacen
    FrmDetalleHilTel.Cod_ClaOrdComp = Cod_ClaOrdComp
    FrmDetalleHilTel.Num_MovStk = DGridLista.Value(DGridLista.Columns("Num_MovStk").Index)
    FrmDetalleHilTel.Cod_OrdComp = DGridLista.Value(DGridLista.Columns("Cod_OrdComp").Index)
    FrmDetalleHilTel.Ser_OrdComp = DGridLista.Value(DGridLista.Columns("Ser_OrdComp").Index)
    FrmDetalleHilTel.Tip_PtMp = DGridLista.Value(DGridLista.Columns("Tip_PtMp").Index)
    
    vAux = DevuelveCampo("select isnull(flg_partida_generada,'') from lg_tiposmov where Cod_TipMov='" & DGridLista.Value(DGridLista.Columns("Cod_TipMov").Index) & "'", cConnect)
    FrmDetalleHilTel.Flg_Partida_Generada = vAux
    
    strSQL = "select ISNULL(flg_ot_tejeduria_generada, 'N') from lg_tiposmov where Cod_TipMov='" & DGridLista.Value(DGridLista.Columns("Cod_TipMov").Index) & "'"
    FrmDetalleHilTel.sFlg_Ot_Tejeduria_Generada = DevuelveCampo(strSQL, cConnect)
    
    FrmDetalleHilTel.lblCod_OrdTra.Visible = (FrmDetalleHilTel.sFlg_Ot_Tejeduria_Generada = "S")
    FrmDetalleHilTel.txtCod_OrdTra.Visible = (FrmDetalleHilTel.sFlg_Ot_Tejeduria_Generada = "S")
    FrmDetalleHilTel.txtCod_OrdTra = DGridLista.Value(DGridLista.Columns("Cod_OrdTra1").Index)
    
    FrmDetalleHilTel.varValida_Factura = Valida_Factura
    
    FrmDetalleHilTel.Show 1
    
    Set FrmDetalleHilTel = Nothing
    
End Sub

Sub ShowItem()
Dim dato As String
'    Almacen = DGridLista.Value(DGridLista.Columns("Cod_Almacen").Index)
    'CmbTipMov_DropDown
'    If sTipo = "I" Then
'        sTipo = ""
'    End If
    
    
    
    Cod_Fabrica = DGridLista.Value(DGridLista.Columns("Cod_Fabrica").Index)
    Cod_TipOrdTra = DGridLista.Value(DGridLista.Columns("Cod_TipOrdTra").Index)
    
'    FrmDetalleStock.sflg_adicionales = DGridLista.Value(DGridLista.Columns("flg_adicionales").Index)
    FrmDetalleStock.cod_tipmov = DGridLista.Value(DGridLista.Columns("Cod_TipMov").Index)
    FrmDetalleStock.Cod_Almacen = DGridLista.Value(DGridLista.Columns("cod_almacen").Index)
    FrmDetalleStock.Cod_ClaOrdComp = Cod_ClaOrdComp
    FrmDetalleStock.Num_MovStk = DGridLista.Value(DGridLista.Columns("Num_MovStk").Index)
    FrmDetalleStock.sFecmovstk = DGridLista.Value(DGridLista.Columns("FECHA MOV").Index)
    FrmDetalleStock.FLG_TRANSFERENCIA_EXTERNA = DevuelveCampo("select flg_transferencia_externa from lg_tiposmov where cod_tipmov = '" & DGridLista.Value(DGridLista.Columns("Cod_TipMov").Index) & "'", cConnect)
    FrmDetalleStock.num_guia = DGridLista.Value(DGridLista.Columns("NUM. GUIA").Index)
    
    FrmDetalleStock.Cod_OrdComp = DGridLista.Value(DGridLista.Columns("Cod_OrdComp").Index)
    FrmDetalleStock.Ser_OrdComp = DGridLista.Value(DGridLista.Columns("Ser_OrdComp").Index)
    FrmDetalleStock.Caption = FrmDetalleStock.Caption & Space(2) & DGridLista.Value(DGridLista.Columns("Num_MovStk").Index) & Space(2) & DGridLista.Value(DGridLista.Columns("Cod_TipMov").Index) & "-" & DGridLista.Value(DGridLista.Columns("Des_TipMov").Index)
    FrmDetalleStock.varValida_Factura = Valida_Factura
    FrmDetalleStock.vcod_cencost = Trim(DGridLista.Value(DGridLista.Columns("cod_cencost").Index))
    FrmDetalleStock.vFlg_Despacho_Acabado = DevuelveCampo("SELECT isnull(flg_despacho_acabado,'') FROM lg_tiposmov WHERE Cod_TipMov = '" & DGridLista.Value(DGridLista.Columns("cod_tipmov").Index) & "'", cConnect)
'    If IsNull(Cod_Fabrica) Then
'        FrmDetalleStock.varCod_Fabrica = ""
'    Else
    FrmDetalleStock.varCod_Fabrica = DGridLista.Value(DGridLista.Columns("Cod_Fabrica").Index)
'    End If
    FrmDetalleStock.varCod_OrdPro = DGridLista.Value(DGridLista.Columns("Cod_OrdPro").Index)
    
    FrmDetalleStock.varNum_SecOrd = Trim(DGridLista.Value(DGridLista.Columns("Num_SecOrd").Index))
    strSQL = "SELECT ISNULL(Flg_SecOrd,'') FROM lg_tiposmov WHERE Cod_TipMov = '" & DGridLista.Value(DGridLista.Columns("cod_tipmov").Index) & "'"
    FrmDetalleStock.varflg_secord = DevuelveCampo(strSQL, cConnect)
    
    FrmDetalleStock.vFlg_Almacen_Tejeduria = DevuelveCampo("select Flg_Almacen_Tejeduria from lg_almacen where cod_almacen ='" & DGridLista.Value(DGridLista.Columns("cod_almacen").Index) & "'", cConnect)
    FrmDetalleStock.vFLG_CREA_COMBINACION_ITEMS_TEJEDURIA = DevuelveCampo("SELECT isnull(FLG_CREA_COMBINACION_ITEMS_TEJEDURIA,'') FROM lg_tiposmov WHERE Cod_TipMov = '" & DGridLista.Value(DGridLista.Columns("cod_tipmov").Index) & "'", cConnect)
    strSQL = "SELECT cod_clamov FROM lg_tiposmov WHERE Cod_TipMov = '" & DGridLista.Value(DGridLista.Columns("cod_tipmov").Index) & "'"
    dato = DevuelveCampo(strSQL, cConnect)
    FrmDetalleStock.var_tipo = dato
    FrmDetalleStock.Datos "V", False
    FrmDetalleStock.Show 1

    Set FrmDetalleStock = Nothing
    

End Sub

Sub ShowTelaCruda()
Dim vAux As Variant
'    Almacen = DGridLista.Value(DGridLista.Columns("Cod_Almacen").Index)
    'CmbTipMov_DropDown
'    If sTipo = "I" Then
'        sTipo = ""
'    End If
    
    Cod_Fabrica = DGridLista.Value(DGridLista.Columns("Cod_Fabrica").Index)
    Cod_TipOrdTra = DGridLista.Value(DGridLista.Columns("Cod_TipOrdTra").Index)
    
    FrmDetalleTelCru.Cod_OrdPro = DGridLista.Value(DGridLista.Columns("Cod_Ordpro").Index)
'    FrmDetalleTelCru.Cod_OrdTra = DevuelveCampo("select Cod_Ordtra from tx_ordtra where Cod_TipOrdTra='" & Reg("Cod_TipOrdTra") & "' and Cod_Proveedor='" & Reg("Cod_Proveedor") & "' and Cod_OrdProv='" & Reg("Lote Prov") & "'", cCONNECT)
    FrmDetalleTelCru.Cod_TipOrdTra = DGridLista.Value(DGridLista.Columns("Cod_TipOrdTra").Index)
    FrmDetalleTelCru.Fec_MOVsTK = DGridLista.Value(DGridLista.Columns("Fecha Mov").Index)
    
    FrmDetalleTelCru.Cod_TipOrdPro = DGridLista.Value(DGridLista.Columns("Cod_TipOrdPro").Index)
    FrmDetalleTelCru.Cod_Proveedor = DGridLista.Value(DGridLista.Columns("Cod_Proveedor").Index)
    FrmDetalleTelCru.Sec_OrdComp = DevuelveCampo("select isnull(max(Sec_OrdComp),'') from lg_ordcompitem where Ser_OrdComp='" & DGridLista.Value(DGridLista.Columns("Ser_OrdComp").Index) & "' and Cod_OrdComp='" & DGridLista.Value(DGridLista.Columns("Cod_OrdComp").Index) & "'", cConnect)
    FrmDetalleTelCru.Cod_ClaMov = Cod_ClaMov
    FrmDetalleTelCru.Cod_TipMovi = DGridLista.Value(DGridLista.Columns("Cod_TipMov").Index)
    FrmDetalleTelCru.Cod_Almacen = Almacen
    FrmDetalleTelCru.Cod_ClaOrdComp = Cod_ClaOrdComp
    FrmDetalleTelCru.Num_MovStk = DGridLista.Value(DGridLista.Columns("Num_MovStk").Index)
    FrmDetalleTelCru.Cod_OrdComp = DGridLista.Value(DGridLista.Columns("Cod_OrdComp").Index)
    FrmDetalleTelCru.Ser_OrdComp = DGridLista.Value(DGridLista.Columns("Ser_OrdComp").Index)
    FrmDetalleTelCru.Tip_PtMp = DGridLista.Value(DGridLista.Columns("Tip_PtMp").Index)
    FrmDetalleTelCru.Cod_TipOrdTra1 = DGridLista.Value(DGridLista.Columns("Cod_TipOrdTra1").Index)
    FrmDetalleTelCru.Cod_OrdTra1 = DGridLista.Value(DGridLista.Columns("Cod_OrdTra1").Index)
    FrmDetalleTelCru.Cod_Calidad = DevuelveCampo("select Cod_Calidad from lg_tiposmov where Cod_TipMov='" & DGridLista.Value(DGridLista.Columns("cod_TipMov").Index) & "'", cConnect)
    vAux = DevuelveCampo("select flg_partida_generada from lg_tiposmov where Cod_TipMov='" & DGridLista.Value(DGridLista.Columns("cod_TipMov").Index) & "'", cConnect)
    FrmDetalleTelCru.Flg_Partida_Generada = IIf(IsNull(vAux), "", vAux)
    vAux = DevuelveCampo("select flg_partidas_tinto from lg_tiposmov where Cod_TipMov='" & DGridLista.Value(DGridLista.Columns("cod_TipMov").Index) & "'", cConnect)
    FrmDetalleTelCru.Flg_Partidas_Tinto = IIf(IsNull(vAux), "", vAux)
    
    If DevuelveCampo("SELECT RTRIM(COD_TIPMOVREL) FROM LG_TIPOSMOV WHERE COD_TIPMOV = '" & DGridLista.Value(DGridLista.Columns("cod_TipMov").Index) & "'", cConnect) <> "" Then
        FrmDetalleTelCru.CmdTransferir.Enabled = True
    Else
        FrmDetalleTelCru.bElijeDatos = True
    End If
    
    strSQL = "select ISNULL(flg_ot_tejeduria_generada, 'N') from lg_tiposmov where Cod_TipMov='" & DGridLista.Value(DGridLista.Columns("cod_TipMov").Index) & "'"
    FrmDetalleTelCru.sFlg_Ot_Tejeduria_Generada = DevuelveCampo(strSQL, cConnect)
    FrmDetalleTelCru.lote.Enabled = (FrmDetalleTelCru.sFlg_Ot_Tejeduria_Generada <> "S")
    FrmDetalleTelCru.varValida_Factura = Valida_Factura
    
    FrmDetalleTelCru.Show 1
    
    Set FrmDetalleTelCru = Nothing

End Sub

Sub ShowTelaTenida()
'    Almacen = DGridLista.Value(DGridLista.Columns("Cod_Almacen").Index)
    'CmbTipMov_DropDown
'    If sTipo = "I" Then
'        sTipo = ""
'    End If
    
    Cod_Fabrica = DGridLista.Value(DGridLista.Columns("Cod_Fabrica").Index)
    Cod_TipOrdTra = DGridLista.Value(DGridLista.Columns("Cod_TipOrdTra").Index)
    
    FrmDetalleTelaCa.Cod_Proveedor = DGridLista.Value(DGridLista.Columns("Cod_Proveedor").Index)
    FrmDetalleTelaCa.Cod_ClaMov = Cod_ClaMov
    FrmDetalleTelaCa.Cod_TipMovi = DGridLista.Value(DGridLista.Columns("Cod_TipMov").Index)
    FrmDetalleTelaCa.Cod_Almacen = Almacen
    FrmDetalleTelaCa.Cod_ClaOrdComp = Cod_ClaOrdComp
    FrmDetalleTelaCa.Num_MovStk = DGridLista.Value(DGridLista.Columns("Num_MovStk").Index)
    FrmDetalleTelaCa.Cod_OrdComp = DGridLista.Value(DGridLista.Columns("Cod_OrdComp").Index)
    FrmDetalleTelaCa.Ser_OrdComp = DGridLista.Value(DGridLista.Columns("Ser_OrdComp").Index)
    FrmDetalleTelaCa.Tip_PtMp = DGridLista.Value(DGridLista.Columns("tip_ptmp").Index)
    FrmDetalleTelaCa.Cod_Calidad = DevuelveCampo("select Cod_Calidad from lg_tiposmov where Cod_TipMov='" & DGridLista.Value(DGridLista.Columns("cod_tipmov").Index) & "'", cConnect)

    FrmDetalleTelaCa.varValida_Factura = Valida_Factura

    FrmDetalleTelaCa.varCod_OrdPro = DGridLista.Value(DGridLista.Columns("Cod_Ordpro").Index)

    strSQL = "SELECT COUNT(*) FROM LG_TIPOSMOV WHERE Cod_TipOrdPro = 'CF' AND Tip_Item = 'T' AND Cod_TipMov = '" & DGridLista.Value(DGridLista.Columns("Cod_TipMov").Index) & "'"
    If DevuelveCampo(strSQL, cConnect) = 0 Then
        FrmDetalleTelaCa.fraOP.Visible = False
    End If
    
    FrmDetalleTelaCa.Flg_Rollo = DevuelveCampo("SELECT ISNULL(FLG_ROLLO,'') FROM LG_TIPOSMOV WHERE Cod_TipMov = '" & DGridLista.Value(DGridLista.Columns("Cod_TipMov").Index) & "'", cConnect)
    Flg_Rollo = FrmDetalleTelaCa.Flg_Rollo
    
    If FrmDetalleTelaCa.Flg_Rollo = "*" Then
        FrmDetalleTelaCa.cmdRollos.Visible = True
        FrmDetalleTelaCa.cmdCapturarPeso.Visible = True
        FrmDetalleTelaCa.txtBultos.Text = "1"
    Else
        FrmDetalleTelaCa.cmdRollos.Visible = False
        FrmDetalleTelaCa.cmdCapturarPeso.Value = False
    End If
    
    FrmDetalleTelaCa.Show 1
    
    Set FrmDetalleTelaCa = Nothing
End Sub

Function ValidaCierre() As Boolean
Dim Ano
Dim mes

ValidaCierre = True

Ano = DevuelveCampo("select Ano_Cierre from lg_almacen where cod_almacen='" & Trim(Right(CmbAlmacen, 3)) & "'", cConnect)
mes = DevuelveCampo("select Mes_Cierre from lg_almacen where cod_almacen='" & Trim(Right(CmbAlmacen, 3)) & "'", cConnect)

If Ano < Me.DtFecha.Year And mes < DtFecha.Month Then
    MsgBox "El mes y año de cierre del almacen no deben ser menores al seleccionado", vbInformation
    ValidaCierre = False
End If
End Function

Function ValidaFlag() As Boolean
ValidaFlag = True

If DevuelveCampo("select Flg_StatusVAL from Lg_MoviStk where Cod_Almacen='" & DGridLista.Value(DGridLista.Columns("Cod_Almacen").Index) & "' and Num_MovStk='" & DGridLista.Value(DGridLista.Columns("num_movstk").Index) & "'", cConnect) = "S" Then
    MsgBox "Este registro no puede ser eliminado", vbInformation
   ValidaFlag = False
End If
End Function

Function ValidaItem() As Boolean
ValidaItem = True

If DevuelveCampo("select count(*) from Lg_MoviStkitem where Cod_Almacen='" & DGridLista.Value(DGridLista.Columns("Cod_Almacen").Index) & "' and Num_MovStk='" & DGridLista.Value(DGridLista.Columns("num_movstk").Index) & "'", cConnect) > 0 Then
    MsgBox "Este registro no puede ser eliminado, tiene items asociado", vbInformation
    ValidaItem = False
End If
End Function

Private Sub CmbAlmacen_Click()
Tip_presentacion = DevuelveCampo("select Tip_Presentacion from lg_almacen where cod_almacen='" & Right(CmbAlmacen, 2) & "'", cConnect)
Tip_item = DevuelveCampo("select Tip_Item from lg_almacen where cod_almacen='" & Right(CmbAlmacen, 2) & "'", cConnect)
sguia = DevuelveCampo("select Flg_Guia from lg_almacen where cod_almacen='" & Right(CmbAlmacen, 2) & "'", cConnect)

End Sub

Private Sub cmdAceptar_Click()
  Screen.MousePointer = vbHourglass
  strSQL = "costos_arregla_fecha_movs '" & DGridLista.Value(DGridLista.Columns("cod_almacen").Index) & "','" & DGridLista.Value(DGridLista.Columns("Num_MovStk").Index) & "','" & dtpNueFecMov & "'"
  ExecuteSQL cConnect, strSQL
  frCambioFecha.Visible = False
  Datos "V", False
  Screen.MousePointer = vbDefault
End Sub

Private Sub cmdEnd_Click()
  frCambioFecha.Visible = False
End Sub

Private Sub CmdImprimir_Click()
    If optResumido.Value Then
        Reporte2
    Else
        Reporte3
    End If
End Sub

Private Sub cmdSalir_Click()
    Me.fraImpresion.Visible = False
End Sub

Private Sub Command2_Click()
sTipo = ""
indicegrilla = 1
If CmbAlmacen <> "" Then
    If ValidaCierre Then Datos "V", False
Else
    Datos "V", False
End If
'Deshabilita


End Sub

Sub CARGA_DATOS(ByVal sAccion As String)
Dim dato As String
On Error GoTo hand
    Load FrmAddMovimAlm
    FrmAddMovimAlm.Accion = sAccion
    If sAccion = "A" Then
        FrmAddMovimAlm.Estado = "MODIFICAR"
        FrmAddMovimAlm.Fradetalle.Enabled = True
        FrmAddMovimAlm.Caption = FrmAddMovimAlm.Caption & "  " & DGridLista.Value(DGridLista.Columns("Num_MovStk").Index)
        FrmAddMovimAlm.DtFechaMov.Enabled = False
        FrmAddMovimAlm.TxtCod_TipMov.Enabled = False
        FrmAddMovimAlm.TxtDes_TipMov.Enabled = False
        'FrmAddMovimAlm.Txtproveedor.Enabled = False
        'FrmAddMovimAlm.TxtDetalle.Enabled = False
        FrmAddMovimAlm.Command1.Enabled = False
        If Trim(DGridLista.Value(DGridLista.Columns("cod_tipmov").Index)) = "S21" Then
            FrmAddMovimAlm.Txtproveedor.Enabled = True
            FrmAddMovimAlm.TxtDetalle.Enabled = True
        Else
            FrmAddMovimAlm.Txtproveedor.Enabled = False
            FrmAddMovimAlm.TxtDetalle.Enabled = False
        End If
    Else
        FrmAddMovimAlm.Estado = "ELIMINAR"
        FrmAddMovimAlm.Fradetalle.Enabled = False
    End If
    FrmAddMovimAlm.Habilita
    strSQL = "SELECT cod_clamov FROM lg_tiposmov WHERE Cod_TipMov = '" & DGridLista.Value(DGridLista.Columns("cod_tipmov").Index) & "'"
    dato = UCase(DevuelveCampo(strSQL, cConnect))
    
    If FrmAddMovimAlm.vFlg_Almacen_Tejeduria = "S" And dato = "S" Then
        FrmAddMovimAlm.FraSolicitante.Visible = True
        FrmAddMovimAlm.TxtTip_Trabajador.Text = Mid(DGridLista.Value(DGridLista.Columns("codigo_solicitante").Index), 1, 1)
        FrmAddMovimAlm.TxtCod_Trabajador.Text = Mid(DGridLista.Value(DGridLista.Columns("codigo_solicitante").Index), 2, 4)
        FrmAddMovimAlm.TxtNom_Trabajador.Text = DGridLista.Value(DGridLista.Columns("nombre_solicitante").Index)
    Else
        FrmAddMovimAlm.FraSolicitante.Visible = False
    End If
     
    FrmAddMovimAlm.TxtOrdPro.Enabled = False
    FrmAddMovimAlm.txtNum_SecOrd.Enabled = False
    FrmAddMovimAlm.Tip_presentacion = DevuelveCampo("select Tip_Presentacion from lg_almacen where cod_almacen='" & DGridLista.Value(DGridLista.Columns("cod_almacen").Index) & "'", cConnect)
    FrmAddMovimAlm.Tip_item = DevuelveCampo("select Tip_Item from lg_almacen where cod_almacen='" & DGridLista.Value(DGridLista.Columns("cod_almacen").Index) & "'", cConnect)
    FrmAddMovimAlm.vCod_Almacen = DGridLista.Value(DGridLista.Columns("Cod_Almacen").Index)
    FrmAddMovimAlm.TxtCod_TipMov = DGridLista.Value(DGridLista.Columns("Cod_TipMov").Index)
    FrmAddMovimAlm.TxtDes_TipMov = DGridLista.Value(DGridLista.Columns("Des_TipMov").Index)
    FrmAddMovimAlm.DtFechaMov.Value = DGridLista.Value(DGridLista.Columns("Fecha Mov").Index)
    FrmAddMovimAlm.Txtproveedor = Trim(DGridLista.Value(DGridLista.Columns("Cod_Proveedor").Index))
    FrmAddMovimAlm.TxtDetalle = Trim(DGridLista.Value(DGridLista.Columns("Des_Proveedor").Index))
    FrmAddMovimAlm.TxtOrdPro = DGridLista.Value(DGridLista.Columns("Cod_ordpro").Index)
    FrmAddMovimAlm.txtNum_SecOrd = DGridLista.Value(DGridLista.Columns("Num_SecOrd").Index)
    FrmAddMovimAlm.vCod_Cliente = Trim(DGridLista.Value(DGridLista.Columns("cod_cliente").Index))
    FrmAddMovimAlm.txtCod_Cliente = Trim(DevuelveCampo("select abr_cliente from tg_cliente where cod_cliente = '" & DGridLista.Value(DGridLista.Columns("cod_cliente").Index) & "'", cConnect))
    FrmAddMovimAlm.TxtNom_Cliente = Trim(DGridLista.Value(DGridLista.Columns("nom_cliente").Index))
    FrmAddMovimAlm.TxtCod_CenCosto = Trim(DGridLista.Value(DGridLista.Columns("cod_cencost").Index))
    FrmAddMovimAlm.TxtDes_CenCosto = Trim(DGridLista.Value(DGridLista.Columns("Centro Costo").Index))
'   FrmAddMovimAlm.Num_Conos
' FrmAddMovimAlm.txtNumConosHilosCoser.Text = Trim(DGridLista.Value(DGridLista.Columns("NRO_CONOS_HILOS_COSER").Index))
    
    Call FrmAddMovimAlm.CARGA_ORDCOMP
    BuscaCombo DGridLista.Value(DGridLista.Columns("Ser_OrdComp").Index) & "-" & DGridLista.Value(DGridLista.Columns("Cod_OrdComp").Index), 1, FrmAddMovimAlm.CmbOrdComp

    FrmAddMovimAlm.TxtObservaciones = Trim(DGridLista.Value(DGridLista.Columns("Observaciones").Index))

    FrmAddMovimAlm.Num_MovStk = DGridLista.Value(DGridLista.Columns("Num_MovStk").Index)
    
    FrmAddMovimAlm.TxtGuia.Text = Trim(DGridLista.Value(DGridLista.Columns("Num. Guia").Index))
'    FrmAddMovimAlm.txtParteSalida.Text = Trim(DGridLista.Value(DGridLista.Columns("ser_parte_salida").Index)) & " " & Trim(DGridLista.Value(DGridLista.Columns("Numero_parte_salida").Index))
    FrmAddMovimAlm.varCod_Fabrica = Trim(DGridLista.Value(DGridLista.Columns("Cod_Fabrica").Index))

    FrmAddMovimAlm.Cod_TipOrdTra = DGridLista.Value(DGridLista.Columns("Cod_TipOrdTra").Index)
   
    FrmAddMovimAlm.txtCod_OrdTra.Text = DGridLista.Value(DGridLista.Columns("Cod_Ordtra1").Index)
    FrmAddMovimAlm.txtCod_TipOrdTra.Text = DGridLista.Value(DGridLista.Columns("Cod_TipOrdTra1").Index)

    strSQL = "SELECT isnull(Cod_color,'') FROM TX_ORDTRA WHERE Cod_TipOrdTra = '" & Trim(FrmAddMovimAlm.txtCod_TipOrdTra.Text) & "' AND Cod_Ordtra = '" & Trim(FrmAddMovimAlm.txtCod_OrdTra.Text) & "'"
    strSQL = "SELECT Des_Color FROM LB_COLOR WHERE Cod_Color = '" & DevuelveCampo(strSQL, cConnect) & "'"
    FrmAddMovimAlm.txtDes_Color = Trim(DevuelveCampo(strSQL, cConnect))
'    FrmAddMovimAlm.txtGlosa_Hilado.Text = Trim(DGridLista.Value(DGridLista.Columns("Glosa_Hilado").Index))
    FrmAddMovimAlm.Num_MovStk = DGridLista.Value(DGridLista.Columns("Num_movstk").Index)
    FrmAddMovimAlm.Show vbModal
    If FrmAddMovimAlm.vOk = True Then
        Datos "V", False
    End If
    Set FrmAddMovimAlm = Nothing
Exit Sub
hand:
ErrorHandler err, "DGridLista_RowColChange"
End Sub


 
Private Sub DGridLista_Click()
indicegrilla = DGridLista.Row
End Sub


Private Sub Form_Load()
Me.DtFecha.Value = Date
dtpNueFecMov.Value = Date

'If vemp1 <> "09" Then
    FunctButt2.FunctionsUser = get_botones1(Me, vper, vemp1, Me.Name)
'End If

LlenarCombos
Datos "V", False
indicegrilla = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub AnularGuia()
On Error GoTo Fin
Dim sTit As String, sNum_MovStk As String
    sTit = "Anular Guia"
    
    If DGridLista.RowCount = 0 Then Exit Sub
    
    If MsgBox("Anular Guia?", vbQuestion + vbYesNo, sTit) = vbNo Then Exit Sub
    
    sNum_MovStk = DGridLista.Value(DGridLista.Columns("Num_MovStk").Index)
    
    strSQL = "EXEC UP_ANUL_IMPRESION_GUIA '" & DGridLista.Value(DGridLista.Columns("Cod_Almacen").Index) & _
    "','" & DGridLista.Value(DGridLista.Columns("Num_MovStk").Index) & "', '" & DGridLista.Value(DGridLista.Columns("Num. Guia").Index) & _
    "','" & vusu & "'"
    
    ExecuteSQL cConnect, strSQL
    
    Command2_Click
    
    'If DGridLista.RowCount = 0 Then Exit Sub
    
'    strSQL = "Num_MovStk = '" & sNum_Movstk & "'"
'    Reg.MoveFirst
'    Reg.Find strSQL
'    If Reg.EOF Then Reg.MoveFirst
    
Exit Sub
Fin:
    MsgBox err.Description, vbCritical + vbOKOnly, ""
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim varTemporal As Boolean
Dim varBusqueda As String, sNum_MovStk As String
Dim sImprimeDetalelRollos As String
Dim sAlmacen1  As String

If DGridLista.RowCount Then
    Almacen = DGridLista.Value(DGridLista.Columns("Cod_Almacen").Index)
    Cod_ClaOrdComp = DevuelveCampo("select rtrim(Cod_ClaOrdComp) from lg_tiposmov where Cod_TipMov='" & DGridLista.Value(DGridLista.Columns("cod_tipmov").Index) & "'", cConnect)
    Cod_ClaMov = DevuelveCampo("select cod_clamov from lg_tiposmov where Cod_TipMov='" & DGridLista.Value(DGridLista.Columns("cod_tipmov").Index) & "'", cConnect)
End If

Select Case ActionName
Case "ADICIONAR"
    If CmbAlmacen.ListIndex = -1 Then
        MsgBox "Debe seleccionar un almacen", vbCritical
        CmbAlmacen.SetFocus
        Exit Sub
    End If
    Load FrmAddMovimAlm
    FrmAddMovimAlm.Accion = "I"
    FrmAddMovimAlm.sCod_AlmacenOrigen = ""
    FrmAddMovimAlm.sNum_MovStkOrigen = ""
    FrmAddMovimAlm.vCod_Almacen = Right(CmbAlmacen.Text, 2)
    FrmAddMovimAlm.Estado = "NUEVO"
    FrmAddMovimAlm.DtFechaMov.Value = Me.DtFecha.Value
    FrmAddMovimAlm.vFlg_Almacen_Tejeduria = DevuelveCampo("select Flg_Almacen_Tejeduria from lg_almacen where cod_almacen ='" & Right(CmbAlmacen.Text, 2) & "'", cConnect)
    FrmAddMovimAlm.Show vbModal
    If FrmAddMovimAlm.vOk = True Then
        Datos "V", False
        FunctButt2_ActionClick 0, 0, "DETALLE"
    End If
    Set FrmAddMovimAlm = Nothing
Case "MODIFICAR"
    If DGridLista.RowCount = 0 Then Exit Sub
    If Valida_Factura = False Then
        MsgBox "El registro no se puede modificar por que posee una factura asociada.", vbInformation, "Mensaje"
        Exit Sub
    End If
    FrmAddMovimAlm.vFlg_Almacen_Tejeduria = DevuelveCampo("select Flg_Almacen_Tejeduria from lg_almacen where cod_almacen ='" & Right(CmbAlmacen.Text, 2) & "'", cConnect)
    'varBusqueda = DGridLista.Value(DGridLista.Columns("Num_MovStk").Index)
    CARGA_DATOS ("A")
    'Call BuscaCampo(DGridLista.ADORecordset, "Num_MovStk", varBusqueda)
Case "ELIMINAR"
    If DGridLista.RowCount = 0 Then Exit Sub
    If Valida_Factura = False Then
        MsgBox "El registro no puede ser eliminado por que posee una factura asociada.", vbInformation, "Mensaje"
        Exit Sub
    End If
    FrmAddMovimAlm.vFlg_Almacen_Tejeduria = DevuelveCampo("select Flg_Almacen_Tejeduria from lg_almacen where cod_almacen ='" & Right(CmbAlmacen.Text, 2) & "'", cConnect)
    If ValidaFlag And ValidaItem Then
        CARGA_DATOS ("B")
        'Datos "b", True
        'varNum_Mov = ""
    End If
    sTipo = ""
Case "DETALLE"
    If DGridLista.RowCount = 0 Then Exit Sub
    If Tip_item = "P" And Tip_presentacion = "T" Then
        ShowCorte
    ElseIf Tip_item = "H" And Tip_presentacion = "C" Then
        ShowHiloCrudo
    ElseIf Tip_item = "H" And Tip_presentacion = "T" Then
        ShowHiloTenido
    ElseIf Tip_item = "T" And Tip_presentacion = "C" Then
        ShowTelaCruda
    ElseIf Tip_item = "T" And Tip_presentacion = "T" Then
        ShowTelaTenida
    Else
       ShowItem
    
    End If
Case "VOUCHER"
    If DGridLista.RowCount = 0 Then Exit Sub
    Flg_Rollo = DevuelveCampo("SELECT ISNULL(FLG_ROLLO,'') FROM LG_TIPOSMOV WHERE Cod_TipMov = '" & DGridLista.Value(DGridLista.Columns("Cod_TipMov").Index) & "'", cConnect)
    If Flg_Rollo = "*" Then
        Me.fraImpresion.Visible = True
    Else
        If DevuelveCampo("select flg_almacen_tejeduria from lg_almacen where cod_almacen ='" & DGridLista.Value(DGridLista.Columns("Cod_Almacen").Index) & "'", cConnect) = "S" Then
            Reporte_Tej
        Else
            Reporte2
        End If
    End If
Case "GUIA"
    If DGridLista.RowCount = 0 Then Exit Sub
    If DevuelveCampo("select cod_clamov from lg_movistk a, lg_tiposmov b where a.cod_Tipmov=b.cod_Tipmov and a.cod_almacen='" & DGridLista.Value(DGridLista.Columns("Cod_Almacen").Index) & "' and a.num_movstk='" & DGridLista.Value(DGridLista.Columns("Num_MovStk").Index) & "'", cConnect) = "S" Then
        If Trim(DGridLista.Value(DGridLista.Columns("Num. Guia").Index)) <> "" Then
            MsgBox "Guia ya fue Impresa", vbInformation + vbOKOnly, "Imprimir Guia"
            Exit Sub
        End If
        Reporte_Guia "Guia"
        sNum_MovStk = DGridLista.Value(DGridLista.Columns("Num_MovStk").Index)
        
        Command2_Click
'        strSQL = "Num_MovStk = '" & sNum_Movstk & "'"
'        Reg.MoveFirst
'        Reg.Find strSQL
'        If Reg.EOF Then Reg.MoveFirst
    Else
        MsgBox "El tipo de movimiento no es 'Salida'", vbInformation, Me.Caption
    End If
Case "FACTURAR"
    If DGridLista.RowCount = 0 Then Exit Sub
    Load frmFacturar
    frmFacturar.varCOD_ALMACEN = DGridLista.Value(DGridLista.Columns("Cod_Almacen").Index)
    frmFacturar.varNUM_MOVSTK = DGridLista.Value(DGridLista.Columns("num_movstk").Index)
    
    frmFacturar.txtSer_Docum = DGridLista.Value(DGridLista.Columns("Ser_Docum").Index)
    frmFacturar.txtNum_Docum = DGridLista.Value(DGridLista.Columns("Num_Docum").Index)
    frmFacturar.txtUsuario_Valorizo = DGridLista.Value(DGridLista.Columns("Usuario_Valorizo").Index)
    frmFacturar.Show 1
    
    Set frmFacturar = Nothing
    'varNum_Mov = DGridLista.Value(DGridLista.Columns("num_movstk").Index)
    'Aqui refrescaremos la data que se muestra en la grilla
    Call Command2_Click
Case "TEMPORAL"
    If DGridLista.RowCount = 0 Then Exit Sub
    If Trim(DGridLista.Value(DGridLista.Columns("cod_tipordtra1").Index)) = "" And Trim(DGridLista.Value(DGridLista.Columns("cod_ordtra1").Index)) = "" Then
    Else
        MsgBox "Este registro ya tiene asignada una partida. Sirvase elegir otra", vbInformation, "Mensaje"
        Exit Sub
    End If

    varTemporal = False
    strSQL = "SELECT COUNT(*) FROM LG_ALMACEN WHERE Cod_Almacen = '" & Trim(DGridLista.Value(DGridLista.Columns("cod_almacen").Index)) & "' AND Tip_Item = 'T' AND Tip_Presentacion = 'C'"
    If DevuelveCampo(strSQL, cConnect) Then
        strSQL = "SELECT Cod_ClaOrdComp FROM LG_TIPOSMOV WHERE Cod_TipMov = '" & DGridLista.Value(DGridLista.Columns("cod_tipmov").Index) & "' AND Tip_Item = 'T' AND Flg_Partidas_Tinto = 'S'"
        strSQL = DevuelveCampo(strSQL, cConnect)
        If Trim(strSQL) <> "" Then
            strSQL = "SELECT COUNT(*) FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp = '" & strSQL & "' AND Tip_Item = 'T' AND Tip_Presentacion = 'T' AND Flg_Requerimiento = 'S' AND Cod_Protex IS NOT NULL"
            If DevuelveCampo(strSQL, cConnect) > 0 Then
                varTemporal = True
            End If
        Else
        End If
    End If
    
    If varTemporal Then
    
        strSQL = "SELECT Cod_Grupo FROM LG_ORDCOMP WHERE Ser_OrdComp= '" & DGridLista.Value(DGridLista.Columns("Ser_OrdComp").Index) & "' AND Cod_OrdComp = '" & DGridLista.Value(DGridLista.Columns("Cod_OrdComp").Index) & "'"
        'varBusqueda = DGridLista.Value(DGridLista.Columns("Num_MovStk").Index)
    
        Load frmMovAlmacenAnexoTemp
        frmMovAlmacenAnexoTemp.varCOD_ALMACEN = DGridLista.Value(DGridLista.Columns("Cod_Almacen").Index)
        frmMovAlmacenAnexoTemp.varNUM_MOVSTK = DGridLista.Value(DGridLista.Columns("Num_MovStk").Index)
        frmMovAlmacenAnexoTemp.varCod_Grupo = DevuelveCampo(strSQL, cConnect)
        frmMovAlmacenAnexoTemp.txtCod_GrupoTex = DevuelveCampo(strSQL, cConnect)
        Call frmMovAlmacenAnexoTemp.BUSCA_GRUPO(1)
        frmMovAlmacenAnexoTemp.varCod_ClaOrdComp = DevuelveCampo("select rtrim(Cod_ClaOrdComp) from lg_tiposmov where Cod_TipMov='" & DGridLista.Value(DGridLista.Columns("cod_tipmov").Index) & "'", cConnect)
        frmMovAlmacenAnexoTemp.varCod_Fabrica = DevuelveCampo("select rtrim(Cod_Fabrica ) from lg_tiposmov where Cod_TipMov='" & DGridLista.Value(DGridLista.Columns("Cod_TipMov").Index) & "'", cConnect)
        frmMovAlmacenAnexoTemp.varCod_Clamov = DevuelveCampo("select rtrim(Cod_ClaMov) from lg_tiposmov where Cod_TipMov='" & DGridLista.Value(DGridLista.Columns("Cod_TipMov").Index) & "'", cConnect)
        
        frmMovAlmacenAnexoTemp.varSer_OrdComp = DGridLista.Value(DGridLista.Columns("Ser_OrdComp").Index)
        frmMovAlmacenAnexoTemp.varCod_OrdComp = DGridLista.Value(DGridLista.Columns("Cod_OrdComp").Index)
        Call frmMovAlmacenAnexoTemp.CARGA_GRID
        Set frmMovAlmacenAnexoTemp.oParent = Me
        frmMovAlmacenAnexoTemp.Show 1
        Call Command2_Click
        'Call BuscaCampo(DGridLista.ADORecordset, "Num_MovStk", varBusqueda)
        Set frmMovAlmacenAnexoTemp = Nothing
    Else
        MsgBox "No se puede acceder a esta opción. Sirvase verificar", vbInformation, "Mensaje"
        Exit Sub
    End If
Case "IMPRESION"
    'If DGridLista.RowCount = 0 Then Exit Sub
    'Call ReporteAvios
    Call ReporteDetalleRollo
    
Case "ANULGUIA"
    If DGridLista.RowCount = 0 Then Exit Sub
    AnularGuia
Case "PARTESALIDA"
    If DGridLista.RowCount = 0 Then Exit Sub
    If DevuelveCampo("select cod_clamov from lg_movistk a, lg_tiposmov b where a.cod_Tipmov=b.cod_Tipmov and a.cod_almacen='" & DGridLista.Value(DGridLista.Columns("Cod_Almacen").Index) & "' and a.num_movstk='" & DGridLista.Value(DGridLista.Columns("Num_MovStk").Index) & "'", cConnect) = "S" Then
        Reporte_Guia "Parte"
    Else
        MsgBox "El tipo de movimiento no es 'Salida'", vbInformation, Me.Caption
    End If
Case "CAMBIOCOLOR"
    If DGridLista.RowCount = 0 Then Exit Sub
    Load FrmCambiosColor
    FrmCambiosColor.Cod_Almacen = DGridLista.Value(DGridLista.Columns("Cod_Almacen").Index)
    FrmCambiosColor.Num_MovStk = DGridLista.Value(DGridLista.Columns("Num_MovStk").Index)
    FrmCambiosColor.Show 1
    Set FrmCambiosColor = Nothing
Case "CAPTURADESPACHOS"
    If CmbAlmacen.ListIndex = -1 Then
        MsgBox "Debe seleccionar algun almacen", vbCritical, "Captura Despachos"
        Exit Sub
    End If
    sAlmacen1 = DevuelveCampo("SELECT COD_ALMACEN FROM LG_ALMACEN WHERE TIP_ITEM='T' AND TIP_PRESENTACION='C' AND COD_ALMACEN = '" & Right(CmbAlmacen, 2) & "'", cConnect)
    If sAlmacen1 <> "" Then
        Load frmCapturaDespachosTejeduria
        frmCapturaDespachosTejeduria.sCod_AlmacenDestino = Right(CmbAlmacen.Text, 2)
        frmCapturaDespachosTejeduria.BUSCAR
        frmCapturaDespachosTejeduria.Show vbModal
        Set frmCapturaDespachosTejeduria = Nothing
        Command2_Click
        If DGridLista.RowCount = 0 Then Exit Sub
    End If
Case Is = "BUSCAGUIA"
    If CmbAlmacen.ListIndex = -1 Then
        MsgBox "Debe seleccionar algun almacen", vbCritical, "Captura Despachos"
        Exit Sub
    End If
    frmBuscaMovGuia.sCod_Almacen = Trim(Right(CmbAlmacen, 2))
    frmBuscaMovGuia.txtNro_Guia = "***"
    frmBuscaMovGuia.fnbBuscar_ActionClick 0, 0, ""
    frmBuscaMovGuia.txtNro_Guia = ""
    frmBuscaMovGuia.Show vbModal
Case Is = "CAMBIOFECHA"
    If DGridLista.RowCount = 0 Then Exit Sub
    frCambioFecha.Visible = True
    dtpNueFecMov = DGridLista.Value(DGridLista.Columns("Fecha Mov").Index)
    dtpNueFecMov.SetFocus
Case Is = "IMPRIMIRCONOSENV"
    Dim frmEnvCoProv As New frmConosHilosCoserEnvProv
    frmEnvCoProv.Show 1
Case Is = "SALIR"
    Unload Me
End Select
End Sub





Private Sub TxtMov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtMov = Trim(DevuelveCampo("Select dbo.uf_devuelvecodigo(6," & IIf(Trim(TxtMov) = "", 0, TxtMov) & ")", cConnect))
    Command2_Click
Else
    Call SoloNumeros(TxtMov, KeyAscii, False)
End If
End Sub

Public Sub Reporte3()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String
Dim varReporte As Boolean

    If Tip_item = "T" And Tip_presentacion = "T" Then
        Ruta = vRuta & "\impvoucherTELTENROL.xlt"
    ElseIf Tip_item = "T" And Tip_presentacion = "C" Then
        Ruta = vRuta & "\ImpVoucherTelCru.XLT"
    ElseIf Tip_item = "H" And Tip_presentacion = "T" Then
        Ruta = vRuta & "\ImpVoucherHilTen.xlt"
    ElseIf Tip_item = "H" And Tip_presentacion = "C" Then
        Ruta = vRuta & "\ImpVoucherHilCru.xlt"
    ElseIf Tip_item = "P" And Tip_presentacion = "T" Then
        Ruta = vRuta & "\impvoucherTELCOR.xlt"
    Else
        Ruta = vRuta & "\impvoucher.xlt"
    End If
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    If Tip_item = "T" Or Tip_item = "H" Then
        oo.Run "reporte", Left(Trim(CmbAlmacen.Text), 50), Trim(DGridLista.Value(DGridLista.Columns("Tipo Mov").Index)), DGridLista.Value(DGridLista.Columns("cod_proveedor").Index) & "-" & DGridLista.Value(DGridLista.Columns("des_proveedor").Index), DGridLista.Value(DGridLista.Columns("Cliente").Index), DGridLista.Value(DGridLista.Columns("cod_ordpro").Index), DGridLista.Value(DGridLista.Columns("Fecha Mov").Index), DGridLista.Value(DGridLista.Columns("Num. Guia").Index), DGridLista.Value(DGridLista.Columns("Ser_OrdComp").Index) & "-" & DGridLista.Value(DGridLista.Columns("Cod_OrdComp").Index), DGridLista.Value(DGridLista.Columns("Centro Costo").Index), CStr(Almacen), CStr(Num_MovStk), DGridLista.Value(DGridLista.Columns("Observaciones").Index), DGridLista.Value(DGridLista.Columns("Num_MovStk_2da").Index), cConnect
        oo.DisplayAlerts = False
        oo.Workbooks.Close
    Else
        oo.Visible = True
        oo.Run "reporte", Left(Trim(CmbAlmacen.Text), 50), Trim(DGridLista.Value(DGridLista.Columns("Tipo Mov").Index)), DGridLista.Value(DGridLista.Columns("cod_proveedor").Index) & "-" & DGridLista.Value(DGridLista.Columns("des_proveedor").Index), DGridLista.Value(DGridLista.Columns("Cliente").Index), DGridLista.Value(DGridLista.Columns("cod_ordpro").Index), DGridLista.Value(DGridLista.Columns("Fecha Mov").Index), DGridLista.Value(DGridLista.Columns("Num. Guia").Index), DGridLista.Value(DGridLista.Columns("Ser_OrdComp").Index) & "-" & DGridLista.Value(DGridLista.Columns("Cod_OrdComp").Index), DGridLista.Value(DGridLista.Columns("Centro Costo").Index), CStr(Almacen), CStr(Num_MovStk), DGridLista.Value(DGridLista.Columns("Observaciones").Index), DGridLista.Value(DGridLista.Columns("Num_MovStk_2da").Index), cConnect
        oo.Run "reporte", Left(Trim(CmbAlmacen.Text), 50), Trim(DGridLista.Value(DGridLista.Columns("Tipo Mov").Index)), DGridLista.Value(DGridLista.Columns("cod_proveedor").Index) & "-" & DGridLista.Value(DGridLista.Columns("des_proveedor").Index), DGridLista.Value(DGridLista.Columns("Cliente").Index), DGridLista.Value(DGridLista.Columns("cod_ordpro").Index), DGridLista.Value(DGridLista.Columns("Fecha Mov").Index), DGridLista.Value(DGridLista.Columns("Num. Guia").Index), DGridLista.Value(DGridLista.Columns("Ser_OrdComp").Index) & "-" & DGridLista.Value(DGridLista.Columns("Cod_OrdComp").Index), DGridLista.Value(DGridLista.Columns("Centro Costo").Index), CStr(Almacen), CStr(Num_MovStk), DGridLista.Value(DGridLista.Columns("Observaciones").Index), DGridLista.Value(DGridLista.Columns("Num_MovStk_2da").Index), cConnect
        oo.DisplayAlerts = False
    End If
        
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub

Private Sub ReporteDetalleRollo()
On Error GoTo hand
    Dim oo As Object, sRuta As String
    Dim rs As New Recordset
    strSQL = "EXEC TJ_SM_MUESTRA_MOV_TELA_CRUDA_ROLLOS_REPORTE '" & Left(Trim(CmbAlmacen), 2) & "', '" & Trim(DGridLista.Value(DGridLista.Columns("num_movstk").Index)) & "',''"
    
    Set rs = CargarRecordSetDesconectado(strSQL, cConnect)
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\rptDetalleRollos.xlt"
    oo.Visible = True
    
    oo.Run "Reporte", rs
Exit Sub
hand:
ErrorHandler err, Me.Caption

End Sub
Sub ReporteAvios()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String
Dim alm As String
Dim Ordcomp As String
Dim hora As Date
Dim dato As String

    strSQL = "SELECT cod_clamov FROM lg_tiposmov WHERE Cod_TipMov = '" & DGridLista.Value(DGridLista.Columns("cod_tipmov").Index) & "'"
    dato = DevuelveCampo(strSQL, cConnect)
    If UCase(dato) <> "E" Then
        MsgBox "Registro debe ser de entrada"
        Exit Sub
    Else
        alm = DGridLista.Value(DGridLista.Columns("cod_almacen").Index)
        Ordcomp = DGridLista.Value(DGridLista.Columns("Ord.Compra").Index)
        
        Ruta = vRuta & "\InspeccionAvios.xlt"
        
        Set oo = CreateObject("excel.application")
        oo.Workbooks.Open Ruta
        oo.Visible = True
        oo.DisplayAlerts = False
        oo.Run "Reporte", alm, Left(Trim(CmbAlmacen), 20), DGridLista.Value(DGridLista.Columns("Num_movstk").Index), DGridLista.Value(DGridLista.Columns("Cliente").Index), Ordcomp, DGridLista.Value(DGridLista.Columns("Des_Proveedor").Index), DGridLista.Value(DGridLista.Columns("Num. Guia").Index), DGridLista.Value(DGridLista.Columns("Fecha Mov").Index), cConnect
        Set oo = Nothing
    End If
    Exit Sub
hand:
    ErrorHandler err, "ReporteAvios"
    Set oo = Nothing
End Sub

Public Function Valida_Factura() As Boolean
Valida_Factura = True
If DGridLista.RowCount > 0 Then
    If Trim(DGridLista.Value(DGridLista.Columns("Factura").Index)) <> "-" And Trim(DGridLista.Value(DGridLista.Columns("Factura").Index)) <> "" Then
        Valida_Factura = False
        Exit Function
    End If
End If
End Function
    
Sub LlenarCombos()
LlenaCombo CmbAlmacen, "Select a.Cod_Almacen + space(2) + a.Nom_Almacen+space(100)+ a.Cod_Almacen from lg_almacen a, lg_segalm b  where a.cod_almacen=b.cod_almacen and b.cod_usuario='" & vusu & "' order by 1", cConnect
End Sub

Public Sub Reporte2()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String
Dim varReporte As Boolean
Dim TipMovRel As String
    
    TipMovRel = RTrim(IIf(IsNull(DevuelveCampo("SELECT COD_TIPMOVREL FROM LG_TIPOSMOV WHERE COD_TIPMOV='" & Trim(DGridLista.Value(DGridLista.Columns("cod_tipmov").Index)) & "'", cConnect)), "", DevuelveCampo("SELECT COD_TIPMOVREL FROM LG_TIPOSMOV WHERE COD_TIPMOV='" & Trim(DGridLista.Value(DGridLista.Columns("cod_tipmov").Index)) & "'", cConnect)))

    If Tip_item = "T" And Tip_presentacion = "T" Then
        Ruta = vRuta & "\impvoucherTELTEN.xlt"
    ElseIf Tip_item = "T" And Tip_presentacion = "C" Then
        Ruta = vRuta & "\ImpVoucherTelCru.XLT"
    ElseIf Tip_item = "H" And Tip_presentacion = "T" Then
        Ruta = vRuta & "\ImpVoucherHilTen.xlt"
    ElseIf Tip_item = "H" And Tip_presentacion = "C" Then
        Ruta = vRuta & "\ImpVoucherHilCru.xlt"
    ElseIf Tip_item = "P" And Tip_presentacion = "T" Then
        Ruta = vRuta & "\impvoucherTELCOR.xlt"
    Else
        If TipMovRel = "" Then
            Ruta = vRuta & "\impvoucher.xlt"
        Else
            Ruta = vRuta & "\impvoucherRel.xlt"
        End If
    End If
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    
    If Tip_item = "T" Or Tip_item = "H" Then
        oo.Run "reporte", DevuelveCampo("select nom_almacen from lg_almacen where cod_almacen = '" & DGridLista.Value(DGridLista.Columns("cod_almacen").Index) & "'", cConnect), _
                Trim(DGridLista.Value(DGridLista.Columns("cod_tipmov").Index)), DGridLista.Value(DGridLista.Columns("Cod_Proveedor").Index) & "-" & DGridLista.Value(DGridLista.Columns("Des_Proveedor").Index), DGridLista.Value(DGridLista.Columns("nom_cliente").Index), DGridLista.Value(DGridLista.Columns("cod_ordpro").Index), DGridLista.Value(DGridLista.Columns("fecha mov").Index), DGridLista.Value(DGridLista.Columns("num. guia").Index), IIf(IsNull(DGridLista.Value(DGridLista.Columns("Ser_OrdComp").Index)), "", CStr(DGridLista.Value(DGridLista.Columns("Ser_OrdComp").Index)) & "-") & IIf(IsNull(DGridLista.Value(DGridLista.Columns("Cod_OrdComp").Index)), "", CStr(DGridLista.Value(DGridLista.Columns("Cod_OrdComp").Index))), DGridLista.Value(DGridLista.Columns("Centro Costo").Index), DGridLista.Value(DGridLista.Columns("cod_almacen").Index), _
                Trim(DGridLista.Value(DGridLista.Columns("num_movstk").Index)), DGridLista.Value(DGridLista.Columns("Observaciones").Index), DGridLista.Value(DGridLista.Columns("Num_MovStk_2da").Index), cConnect
        oo.DisplayAlerts = False
        'oo.Workbooks.Close
    Else
        
        oo.Run "reporte", DevuelveCampo("select nom_almacen from lg_almacen where cod_almacen = '" & DGridLista.Value(DGridLista.Columns("cod_almacen").Index) & "'", cConnect), _
                Trim(DGridLista.Value(DGridLista.Columns("cod_tipmov").Index)) & " " & Trim(DGridLista.Value(DGridLista.Columns("des_tipmov").Index)), _
                DGridLista.Value(DGridLista.Columns("Cod_Proveedor").Index) & "-" & DGridLista.Value(DGridLista.Columns("Des_Proveedor").Index), DGridLista.Value(DGridLista.Columns("nom_cliente").Index), DGridLista.Value(DGridLista.Columns("cod_ordpro").Index), DGridLista.Value(DGridLista.Columns("fecha mov").Index), DGridLista.Value(DGridLista.Columns("num. guia").Index), DGridLista.Value(DGridLista.Columns("Ser_OrdComp").Index) & "-" & DGridLista.Value(DGridLista.Columns("Cod_OrdComp").Index), DGridLista.Value(DGridLista.Columns("Centro Costo").Index), DGridLista.Value(DGridLista.Columns("cod_almacen").Index), DGridLista.Value(DGridLista.Columns("Num_MovStk").Index), DGridLista.Value(DGridLista.Columns("Observaciones").Index), _
                DGridLista.Value(DGridLista.Columns("Num_MovStk_2da").Index), cConnect, ""
        oo.DisplayAlerts = False
    End If
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "Reporte"
End Sub

Public Sub Reporte_Tej()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String
Dim varReporte As Boolean
    
    Ruta = vRuta & "\impvoucherTej.xlt"
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    
    oo.Visible = True
    oo.Run "reporte", DevuelveCampo("select nom_almacen from lg_almacen where cod_almacen = '" & DGridLista.Value(DGridLista.Columns("cod_almacen").Index) & "'", cConnect), _
            Trim(DGridLista.Value(DGridLista.Columns("cod_tipmov").Index)) & " " & Trim(DGridLista.Value(DGridLista.Columns("des_tipmov").Index)), _
            DGridLista.Value(DGridLista.Columns("Cod_Proveedor").Index) & "-" & DGridLista.Value(DGridLista.Columns("Des_Proveedor").Index), DGridLista.Value(DGridLista.Columns("nom_cliente").Index), DGridLista.Value(DGridLista.Columns("cod_ordpro").Index), DGridLista.Value(DGridLista.Columns("fecha mov").Index), DGridLista.Value(DGridLista.Columns("num. guia").Index), DGridLista.Value(DGridLista.Columns("Ser_OrdComp").Index) & "-" & DGridLista.Value(DGridLista.Columns("Cod_OrdComp").Index), DGridLista.Value(DGridLista.Columns("Centro Costo").Index), DGridLista.Value(DGridLista.Columns("cod_almacen").Index), DGridLista.Value(DGridLista.Columns("Num_MovStk").Index), DGridLista.Value(DGridLista.Columns("Observaciones").Index), _
            DGridLista.Value(DGridLista.Columns("Num_MovStk_2da").Index), cConnect
    oo.DisplayAlerts = False

    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "Reporte"
End Sub






Sub ShowCorte()
    frmDetalleCorte.sCod_TipMov = DGridLista.Value(DGridLista.Columns("Cod_TipMov").Index)
    frmDetalleCorte.sCod_Almacen = Almacen
    frmDetalleCorte.sCod_ClaOrdComp = Cod_ClaOrdComp
    frmDetalleCorte.sNum_MovStk = DGridLista.Value(DGridLista.Columns("Num_MovStk").Index)
    frmDetalleCorte.CARGA_GRID
    frmDetalleCorte.Show 1
    Set frmDetalleCorte = Nothing
End Sub
