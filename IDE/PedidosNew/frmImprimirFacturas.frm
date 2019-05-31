VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmImprimirFacturas 
   Caption         =   "Facturas"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   9525
   StartUpPosition =   3  'Windows Default
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   4200
      TabIndex        =   0
      Top             =   4680
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   900
      Custom          =   "0~0~SALIR~Verdadero~Verdadero~&Salir~0~0~1~~0~Falso~Falso~&Salir~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX grdPoAsociadas 
      Height          =   4575
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8070
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmImprimirFacturas.frx":0000
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmImprimirFacturas.frx":0352
      Column(2)       =   "frmImprimirFacturas.frx":041A
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmImprimirFacturas.frx":04BE
      FormatStyle(2)  =   "frmImprimirFacturas.frx":05F6
      FormatStyle(3)  =   "frmImprimirFacturas.frx":06A6
      FormatStyle(4)  =   "frmImprimirFacturas.frx":075A
      FormatStyle(5)  =   "frmImprimirFacturas.frx":0832
      FormatStyle(6)  =   "frmImprimirFacturas.frx":08EA
      FormatStyle(7)  =   "frmImprimirFacturas.frx":09CA
      FormatStyle(8)  =   "frmImprimirFacturas.frx":0E82
      ImageCount      =   1
      ImagePicture(1) =   "frmImprimirFacturas.frx":12CE
      PrinterProperties=   "frmImprimirFacturas.frx":1620
   End
End
Attribute VB_Name = "frmImprimirFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    If Me.ssgrdDatos2.Rows > 0 Then
        Select Case ActionName
            Case "ASIGNANRODESPACHO"
                Load frmAsignaNroDespacho
                frmAsignaNroDespacho.sCod_Cliente = Me.sCod_Cliente
                frmAsignaNroDespacho.sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
                frmAsignaNroDespacho.sCod_LotPurOrd = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                frmAsignaNroDespacho.sCod_EstCli = ssgrdDatos2.Columns("Cod_EstCli").Text
                frmAsignaNroDespacho.CargaNroDespachoActual
                frmAsignaNroDespacho.Show vbModal
                Set frmAsignaNroDespacho = Nothing
            Case "PO"
                Load frmTG_PurOrdDestinos
                Set frmTG_PurOrdDestinos.oParent = Me
                frmTG_PurOrdDestinos.sFlgOpcion_Nueva = "S"
                frmTG_PurOrdDestinos.sAccionName = "MODIFICAR"
                frmTG_PurOrdDestinos.sModoWizard = "   ESTDAT"
                frmTG_PurOrdDestinos.sCod_Cliente = Me.sCod_Cliente
                frmTG_PurOrdDestinos.sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
                frmTG_PurOrdDestinos.sCod_LotPurOrd = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                frmTG_PurOrdDestinos.sCod_EstCli = ssgrdDatos2.Columns("Cod_EstCli").Text
                frmTG_PurOrdDestinos.sCod_TemCli = ssgrdDatos.Columns("Cod_TemCli").Text
                frmTG_PurOrdDestinos.BUSCAR
                frmTG_PurOrdDestinos.Show vbModal
                Set frmTG_PurOrdDestinos = Nothing
             Case "DATOS ADICIONALES"
                Load frmDatosAdicionales
                Set frmDatosAdicionales.oParent = Me
                frmDatosAdicionales.sFlgOpcion_Nueva = "S"
                frmDatosAdicionales.sCod_Cliente = Me.sCod_Cliente
                frmDatosAdicionales.sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
                frmDatosAdicionales.sCod_LotPurOrd = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                frmDatosAdicionales.sCod_EstCli = ssgrdDatos2.Columns("Cod_EstCli").Text
                frmDatosAdicionales.sCod_TemCli = ssgrdDatos.Columns("Cod_TemCli").Text
                frmDatosAdicionales.BUSCAR
                frmDatosAdicionales.Show vbModal
                Set frmDatosAdicionales = Nothing
             Case "COMENTARIO"
                'Load FrmComentario
                Set FrmComentario.oParent = Me
                'FrmComentario.sFlgOpcion_Nueva = "S"
                FrmComentario.sCod_Cliente = Me.sCod_Cliente
                FrmComentario.sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
                FrmComentario.sCod_LotPurOrd = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                FrmComentario.sCod_EstCli = ssgrdDatos2.Columns("Cod_EstCli").Text
                'FrmComentario.Buscar
                FrmComentario.Cargar_Data
                FrmComentario.Show vbModal
                Set FrmComentario = Nothing
             Case "POASOCIADAS"
               Load frmPoAsociadas
               'frmPoAsociadas.Caption = "PO Asociadas " + "Cliente:" + Me.ssgrdDatos.Columns("Nom_cliente").Text + " PO:" + Me.ssgrdDatos.Columns("Cod_PurOrd").Text + " Lote:" + ssgrdDatos2.Columns("Cod_LotPurOrd").Text + " Numero:" + ssgrdDatos2.Columns("Cod_EstCli").Text
               frmPoAsociadas.Caption = "PO Asociadas " + "Cliente:" + Me.TxtNom_Cliente.Text + "  PO:" + Me.ssgrdDatos.Columns("Cod_PurOrd").Text + " Lote:" + ssgrdDatos2.Columns("Cod_LotPurOrd").Text + " Numero:" + ssgrdDatos2.Columns("Cod_EstCli").Text
               frmPoAsociadas.COD_CLIENTE = Me.ssgrdDatos.Columns("Cod_cliente").Text
               frmPoAsociadas.cod_purord = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
               frmPoAsociadas.cod_lotpurord = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
               frmPoAsociadas.cod_estcli = ssgrdDatos2.Columns("Cod_EstCli").Text
               frmPoAsociadas.BUSCAR
               frmPoAsociadas.Show vbModal
               Set frmPoAsociadas = Nothing
             Case "IMPRIMIRPOASO"
               mReporte
            Case "CAMBFECFINPROD"
                If ssgrdDatos.Rows > 0 Then
                    Load frmCambioFecFinProduccion
                    frmCambioFecFinProduccion.COD_CLIENTE = sCod_Cliente
                    frmCambioFecFinProduccion.cod_purord = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
                    frmCambioFecFinProduccion.cod_lotpurord = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                    frmCambioFecFinProduccion.cod_estcli = ssgrdDatos2.Columns("Cod_EstCli").Text
                    
                    frmCambioFecFinProduccion.txtCliente.Text = Trim(Me.txtAbr_Cliente.Text) & " - " & Trim(Me.TxtNom_Cliente.Text)
                    frmCambioFecFinProduccion.txtPO.Text = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
                    frmCambioFecFinProduccion.txtEstilo.Text = Me.ssgrdDatos2.Columns("Cod_EstCli").Text
                    frmCambioFecFinProduccion.dtpFecFinProd.value = ssgrdDatos2.Columns("Fec_DespachoAct").Text
                    frmCambioFecFinProduccion.Show 1
                    Set frmCambioFecFinProduccion = Nothing
                    Call BUSCAR
                End If
             Case "DATOSFINAN"
                Dim svpo As String
                Dim svcommit As String
                Dim strSql As String
                Load frmDatosFinanzas
                frmDatosFinanzas.varCod_Cliente = sCod_Cliente
                frmDatosFinanzas.varCod_EstCli = ssgrdDatos2.Columns("Cod_EstCli").Text
                frmDatosFinanzas.varCod_LotPurOrd = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                frmDatosFinanzas.varCod_TemCli = ssgrdDatos.Columns("Cod_TemCli").Text
                Set frmDatosFinanzas.oParent = Me
                        
'                Strsql = "SELECT FLG_IFINANCIERA_PO FROM TG_PURORD WHERE COD_CLIENTE = '" & sCod_Cliente & "' AND COD_PURORD = '" & Me.ssgrdDatos.Columns("Cod_PurOrd").Text & "'"
'                svpo = DevuelveCampo(Strsql, cCONNECT)
'
'                If svpo = "N" Then
'                    frmDatosFinanzas.chkop.value = 0
'                Else
'                    frmDatosFinanzas.chkop.value = 1
'                End If
'
'                Strsql = "SELECT FLG_IFINANCIERA_COMMIT FROM TG_PURORD WHERE COD_CLIENTE = '" & sCod_Cliente & "' AND COD_PURORD = '" & Me.ssgrdDatos.Columns("Cod_PurOrd").Text & "'"
'                svcommit = DevuelveCampo(Strsql, cCONNECT)
'
'                If svcommit = "N" Then
'                    frmDatosFinanzas.chkcommit.value = 0
'                Else
'                    frmDatosFinanzas.chkcommit.value = 1
'                End If


                If ssgrdDatos.Columns(47).Text = "N" Then
                    frmDatosFinanzas.chkcommit.value = 0
                Else
                    frmDatosFinanzas.chkcommit.value = 1
                End If
                
                If ssgrdDatos.Columns(46).Text = "N" Then
                    frmDatosFinanzas.chkop.value = 0
                Else
                    frmDatosFinanzas.chkop.value = 1
                End If

                frmDatosFinanzas.txtPO = ssgrdDatos.Columns("Cod_PurOrd").Text
                frmDatosFinanzas.txtCliente = Trim(Me.txtAbr_Cliente.Text) & " - " & Trim(Me.TxtNom_Cliente.Text)
                frmDatosFinanzas.txtEstilo = ssgrdDatos2.Columns("Cod_EstCli").Text
        
                
                frmDatosFinanzas.Show 1
                Set frmDatosFinanzas = Nothing
                
                Case "EXFACTORY"
                If ssgrdDatos.Rows > 0 Then
                    Load FrmExFactory
                    FrmExFactory.COD_CLIENTE = sCod_Cliente
                    FrmExFactory.cod_purord = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
                    FrmExFactory.cod_lotpurord = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                    FrmExFactory.cod_estcli = ssgrdDatos2.Columns("Cod_EstCli").Text
                    FrmExFactory.dtpFec.value = ssgrdDatos2.Columns("Fec_DespachoOri_Reprogramada").Text
                    FrmExFactory.Show 1
                    Set FrmExFactory = Nothing
                    Call BUSCAR
                End If
                
            Case "EFAFAC"
                If ssgrdDatos.Rows = 0 Then Exit Sub
                Load frmExFactoryFacturacion
                With frmExFactoryFacturacion
                    .COD_CLIENTE = sCod_Cliente
                    .PO = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
                    .LOTE_PO = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                    .ESTILO_CLIENTE = ssgrdDatos2.Columns("Cod_EstCli").Text
                    .dtpFec.value = ssgrdDatos2.Columns("Fec_ExFactory_Ajustada_Facturacion").Text
                    .Show 1
                End With
                Set frmExFactoryFacturacion = Nothing
                Call BUSCAR
         End Select
    End If
End Sub
