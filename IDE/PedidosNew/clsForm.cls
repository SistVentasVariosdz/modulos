VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function PutFocus _
                Lib "user32" _
                Alias "SetFocus" (ByVal hwnd As Long) As Long

Private mConnect As String

Public Property Get Parent() As Object
    Set Parent = oMDIParent
End Property

Public Property Let Parent(ByVal vnewvalue As Object)
    Set oMDIParent = vnewvalue
End Property

Public Property Get Language() As Integer
    Language = iLanguage
End Property

Public Property Let Language(ByVal vnewvalue As Integer)
    iLanguage = vnewvalue
End Property

Public Property Let Rutas(ByVal mRutas As Variant)
    vRuta = mRutas
End Property

Public Property Get Rutas()
    Rutas = vRuta
End Property

Public Property Let Cod_Empresa(ByVal mCod_Empresa As String)
    vemp = mCod_Empresa
    bCargaConexion = True
End Property

Public Property Let Cod_Perfil(ByVal mCod_Perfil As String)
    vper = mCod_Perfil
End Property

Public Property Let Cod_Opcion(ByVal mCod_Opcion As String)
    'gsCod_Opcion = mCod_Opcion
End Property

Public Property Let Key_Node(ByVal mKey_Node As String)
    'gsKey_Node = mKey_Node
End Property

Public Property Let Permisos(ByVal mPermisos As String)
    'gsPermisos = mPermisos
End Property

Public Property Let Nom_Empresa(ByVal vnewvalue As String)
    'gsNomEmpresa = vnewvalue
End Property

Public Property Let Nom_Aplicacion(ByVal vnewvalue As String)
    'gsNomAplicacion = vnewvalue
End Property

Public Property Let Nom_Opcion(ByVal vnewvalue As String)
    vNomFor = vnewvalue
End Property

Public Property Let UserName(ByVal vnewvalue As String)
    vusu = vnewvalue
End Property

Public Property Get UserName() As String
    UserName = vusu
End Property

Public Property Let ConnectEmpresa(ByVal vnewvalue As String)
    mConnect = vnewvalue
    cCONNECT = mConnect
End Property

Public Property Let ConnectSeguridad(ByVal vnewvalue As String)
    cSEGURIDAD = vnewvalue
End Property

Public Function GetForm(sFormName, ParamArray Parametros()) As Object

    Dim prms() As Variant

    prms = Parametros(0)(0)

    sFormName = Trim(UCase(sFormName))

    Select Case sFormName

        Case "FRMSHOWTG_PURORD"
            Set GetForm = New frmShowTG_PurOrd

        Case "FRMGENERAROP"
            Set GetForm = New frmGenerarOP
            GetForm.sCod_Cliente = CStr(prms(0))
            GetForm.sCod_PurOrd = CStr(prms(1))

        Case "FRMCHANGEPO"
            Set GetForm = New frmChangePO
            GetForm.varCod_Cliente = CStr(prms(0))
            GetForm.bNivelPO = CBool(prms(1))
            GetForm.varCod_EstCli = CStr(prms(2))
            GetForm.varCod_LotPurOrd = CStr(prms(3))
            GetForm.varCod_TemCli = CStr(prms(4))
            GetForm.txtPO.Text = CStr(prms(5))
            GetForm.txtCliente.Text = CStr(prms(6))
            GetForm.txtEstilo.Visible = CBool(prms(7))
            GetForm.txtEstilo.Text = CStr(prms(8))

        Case "FRMCAMBIACLASEPO"
            Set GetForm = New frmCambiaClasePO
            GetForm.varCod_Cliente = CStr(prms(0))
            GetForm.varNom_Cliente = CStr(prms(1))
            GetForm.varCod_Purord = CStr(prms(2))
            GetForm.varCod_Clapurord = CStr(prms(3))

        Case "FRMAVANENCAJADO"
            Set GetForm = New frmAvanEncajado
            GetForm.vCod_Cliente = CStr(prms(0))
            GetForm.vCod_PurOrd = CStr(prms(1))
            GetForm.vcod_lotpurord = CStr(prms(2))
            GetForm.vcod_estcli = CStr(prms(3))
            GetForm.Label1.Caption = CStr(prms(4))
            GetForm.Label2.Caption = CStr(prms(5))
            GetForm.CARGA_GRID

        Case "FRMMANTPURORDTAL"
            Set GetForm = New frmMantPurOrdTal
            GetForm.Cliente = CStr(prms(0))
            GetForm.txtNom_Cliente.Text = CStr(prms(1))
            GetForm.PO = CStr(prms(2))
            GetForm.txtCod_PurOrd.Text = CStr(prms(2))
            GetForm.Estilo = CStr(prms(3))
            GetForm.txtCod_EstCli.Text = CStr(prms(3))
            GetForm.Cargar_Data

        Case "FRMMANTCOLOR"
            Set GetForm = New frmMantColor
            GetForm.sCod_Cliente = CStr(prms(0))
            GetForm.sCod_PurOrd = CStr(prms(1))
            GetForm.sCod_EstCli = CStr(prms(2))
            GetForm.Inicializar
            GetForm.Cargar_Datos
            GetForm.txtabrecli.Text = CStr(prms(3))

        Case "FRMUPDATEDATGENLOTEST"
            Set GetForm = New frmUpdateDatGenLotEst
            GetForm.sCod_Cliente = CStr(prms(0))
            GetForm.sCod_PurOrd = CStr(prms(1))
            GetForm.sCod_LotPurOrd = CStr(prms(2))
            GetForm.sCod_EstCli = CStr(prms(3))
            GetForm.txtCod_DestinoLOT.Text = CStr(prms(4))
            GetForm.txtDes_DestinoLOT.Text = CStr(prms(5))

            If CStr(prms(13)) = "S" Then
                GetForm.optComisionEnPorcentaje.value = True
                GetForm.optComisionEnImporte.value = False
            Else
                GetForm.optComisionEnImporte.value = True
                GetForm.optComisionEnPorcentaje.value = False
            End If

            'GetForm.dtpFec_DespachoActLOT.value = CStr(prms(6))
            GetForm.txtPor_ComisionLOT.Text = CStr(prms(7))
            GetForm.dtpFec_DespachoOriLOT.value = CStr(prms(8))
            GetForm.txtDes_General.Text = CStr(prms(9))
            GetForm.txtPrecioLOT.Text = CStr(prms(10))
            GetForm.txtCod_DivPreLOT.Text = CStr(prms(11))
            GetForm.sFlg_Regular = CStr(prms(12))

            GetForm.txtImp_Comision = CStr(prms(14))

            If GetForm.sFlg_Regular = "S" Then
                GetForm.fraNORegular.Visible = False
            Else
                GetForm.dtpFec_RecCliLOT.value = CStr(prms(15))
                GetForm.txtPrecio_RecCliLOT.Text = CStr(prms(16))
                GetForm.fraNORegular.Visible = True
            End If

            'Pre Pack
            If CStr(prms(17)) = "S" Then
                GetForm.optPrePackSi = True
            Else
                GetForm.optPrePackNo = True
            End If

        Case "FRMVIEWOPS"
            Set GetForm = New frmViewOPs
            GetForm.sCod_Cliente = CStr(prms(0))
            GetForm.sCod_PurOrd = CStr(prms(1))
            GetForm.sCod_LotPurOrd = CStr(prms(2))
            GetForm.sCod_EstCli = CStr(prms(3))
            GetForm.BuscarOps

        Case "FRMVERIFICAMATRIZDETALLE"
            Set GetForm = New frmVerificaMatrizDetalle
            GetForm.sCod_Cliente = CStr(prms(0))
            GetForm.sCod_PurOrd = CStr(prms(1))
            GetForm.sCod_LotPurOrd = CStr(prms(2))
            GetForm.sCod_TemCli = CStr(prms(3))
            Set GetForm.rsData = prms(4)
            GetForm.BUSCAR

        Case "FRMCAMBIOSINGLEPREPACK"
            Set GetForm = New frmCambioSinglePrePack

            If CStr(prms(0)) = "S" Then GetForm.optPrePackSi = True Else: GetForm.optPrePackNo = True
            GetForm.sCodigoCliente = CStr(prms(1))
            GetForm.sCodigoPurOrd = CStr(prms(2))
            GetForm.sCodigoLotePurOrd = CStr(prms(3))
            GetForm.sCodigoEstadoCliente = CStr(prms(4))

        Case "FRMASIGNANRODESPACHO"
            Set GetForm = New frmAsignaNroDespacho
            GetForm.sCod_Cliente = CStr(prms(0))
            GetForm.sCod_PurOrd = CStr(prms(1))
            GetForm.sCod_LotPurOrd = CStr(prms(2))
            GetForm.sCod_EstCli = CStr(prms(3))
            GetForm.CargaNroDespachoActual

        Case "FRMTG_PURORDDESTINOS"
            Set GetForm = New frmTG_PurOrdDestinos
            
            GetForm.sFlgOpcion_Nueva = CStr(prms(0))
            GetForm.sAccionName = CStr(prms(1))
            GetForm.sModoWizard = CStr(prms(2))
            GetForm.sCod_Cliente = CStr(prms(3))
            GetForm.sCod_PurOrd = CStr(prms(4))
            GetForm.sCod_LotPurOrd = CStr(prms(5))
            GetForm.sCod_EstCli = CStr(prms(6))
            GetForm.sCod_TemCli = CStr(prms(7))
            GetForm.BUSCAR

        Case "FRMDATOSADICIONALES"
            Set GetForm = New frmDatosAdicionales
            
            GetForm.sFlgOpcion_Nueva = CStr(prms(0))
            GetForm.sCod_Cliente = CStr(prms(1))
            GetForm.sCod_PurOrd = CStr(prms(2))
            GetForm.sCod_LotPurOrd = CStr(prms(3))
            GetForm.sCod_EstCli = CStr(prms(4))
            GetForm.sCod_TemCli = CStr(prms(5))
            GetForm.BUSCAR

        Case "FRMDATOSADICIONALES"
            Set GetForm = New frmDatosAdicionales
            
            GetForm.sFlgOpcion_Nueva = CStr(prms(0))
            GetForm.sCod_Cliente = CStr(prms(1))
            GetForm.sCod_PurOrd = CStr(prms(2))
            GetForm.sCod_LotPurOrd = CStr(prms(3))
            GetForm.sCod_EstCli = CStr(prms(4))
            GetForm.Cargar_Data

        Case "FRMPOASOCIADAS"
            Set GetForm = New frmPoAsociadas

            GetForm.Caption = "PO Asociadas " + "Cliente:" + CStr(prms(0)) + " PO:" + CStr(prms(2)) + " Lote:" + CStr(prms(3)) + " Numero:" + CStr(prms(4))
            GetForm.COD_CLIENTE = CStr(prms(1))
            GetForm.cod_purord = CStr(prms(2))
            GetForm.cod_lotpurord = CStr(prms(3))
            GetForm.cod_estcli = CStr(prms(4))
            GetForm.BUSCAR

        Case "FRMDATOSFINANZAS"
            Set GetForm = New frmDatosFinanzas
            
            GetForm.varCod_Cliente = CStr(prms(0))
            GetForm.varCod_EstCli = CStr(prms(1))
            GetForm.varCod_LotPurOrd = CStr(prms(2))
            GetForm.varCod_TemCli = CStr(prms(3))
                    
            If CStr(prms(4)) = "N" Then
                GetForm.chkcommit.value = 0
            Else
                GetForm.chkcommit.value = 1
            End If
            
            If CStr(prms(5)) = "N" Then
                GetForm.chkop.value = 0
            Else
                GetForm.chkop.value = 1
            End If

            GetForm.txtPO = CStr(prms(6))
            GetForm.txtCliente = Trim(CStr(prms(7))) & " - " & Trim(CStr(prms(8)))
            GetForm.txtEstilo = CStr(prms(1))

        Case "FRMADDRECORD"
            Set GetForm = New frmAddRecord
            GetForm.bEnabledCodigo = CBool(prms(0))
            GetForm.sFrame = CStr(prms(1))
            GetForm.LoadFrame2
            GetForm.txtCod_Destino.Text = CStr(prms(2))
            GetForm.txtCod_DivCli.Text = CStr(prms(2))
            GetForm.txtCod_TemCli.Text = CStr(prms(2))
            GetForm.txtCod_PagEmb.Text = CStr(prms(2))
            GetForm.txtCod_Embarque.Text = CStr(prms(2))
            GetForm.txtCod_Banco.Text = CStr(prms(2))
            GetForm.sCod_Cliente = CStr(prms(3))

        Case "FRMADDGRUPOPRO"
            Set GetForm = New frmAddGrupoPro
            GetForm.txtAbr_Cliente = CStr(prms(0))
            GetForm.txtNom_Cliente = CStr(prms(1))
            GetForm.txtCod_GrupoPro = Trim(CStr(prms(0))) & "-" & Format(CStr(CStr(prms(2)) + 1), "00#")

        Case "FRMCAMBIOFECFINPRODUCCION"
            Set GetForm = New frmCambioFecFinProduccion
            'GetForm.txtAbr_Cliente = CStr(prms(0))
            'GetForm.txtNom_Cliente = CStr(prms(1))
            GetForm.COD_CLIENTE = CStr(prms(0))
            GetForm.cod_purord = CStr(prms(1))
            GetForm.cod_lotpurord = CStr(prms(2))
            GetForm.cod_estcli = CStr(prms(3))
            
            GetForm.txtCliente.Text = CStr(prms(4)) & " - " & CStr(prms(5))
            GetForm.txtPO.Text = CStr(prms(6))
            GetForm.txtEstilo.Text = CStr(prms(7))
            GetForm.dtpFecFinProd.value = CStr(prms(8))

        Case "FRMCAMBIAPRECIOCOLOR"
            Set GetForm = New frmCambiaPrecioColor
            'Set frmCambiaPrecioColor.oParent = Me
            GetForm.sCod_Cliente = CStr(prms(0))
            GetForm.sCod_PurOrd = CStr(prms(1))
            GetForm.sCod_LotPurOrd = CStr(prms(2))
            GetForm.sCod_EstCli = CStr(prms(3))
            GetForm.sCod_TemCli = CStr(prms(4))
            GetForm.BUSCAR

        Case "FRMEXFACTORY"
            Set GetForm = New FrmExFactory
            GetForm.COD_CLIENTE = CStr(prms(0))
            GetForm.cod_purord = CStr(prms(1))
            GetForm.cod_lotpurord = CStr(prms(2))
            GetForm.cod_estcli = CStr(prms(3))
            GetForm.dtpFec.value = CStr(prms(4))

        Case "FRMEXFACTORYFACTURACION"
            Set GetForm = New frmExFactoryFacturacion
            GetForm.COD_CLIENTE = CStr(prms(0))
            GetForm.PO = CStr(prms(1))
            GetForm.LOTE_PO = CStr(prms(2))
            GetForm.ESTILO_CLIENTE = CStr(prms(3))
            GetForm.dtpFec.value = CStr(prms(4))

        Case "FRMNEWPO"
            Set GetForm = New frmNewPO
            GetForm.txtIdCliente.Text = CStr(prms(0))
            GetForm.txtNomCliente.Text = CStr(prms(1))

        Case "FRM_ACTUALIZA_FECHAPO"
            Set GetForm = New Frm_Actualiza_FechaPO
            GetForm.sCod_Cliente = CStr(prms(0))
            GetForm.sCod_PurOrd = CStr(prms(1))
            GetForm.DTPicker1.value = CStr(prms(2))

        Case "FRMCOLORDETAIL"
            Set GetForm = New frmColorDetail
            GetForm.sCod_Cliente = CStr(prms(0))
            GetForm.sCod_PurOrd = CStr(prms(1))
            GetForm.sCod_LotPurOrd = CStr(prms(2))
            GetForm.sCod_EstCli = CStr(prms(3))
            GetForm.BUSCAR

        Case "FRMCOMENTARIO"
            Set GetForm = New FrmComentario
            GetForm.sCod_Cliente = CStr(prms(0))
            GetForm.sCod_PurOrd = CStr(prms(1))
            GetForm.sCod_LotPurOrd = CStr(prms(2))
            GetForm.sCod_EstCli = CStr(prms(3))
            GetForm.Cargar_Data
    End Select

End Function

Public Function OpenForm(ByVal strFormName As String, _
                         ByVal lngParentHwnd As Long, _
                         ParamArray Parametros()) As Long

    Dim objForm As Form

    On Error GoTo eH

    'We don't want to allow creating multiple instances of the same form
    'so we return the existing form handle if it is in the Forms collection

    'Set objForm = GetFormVB6(strFormName)
    Set objForm = GetForm(strFormName, Parametros)

    If objForm Is Nothing Then

        'Creates a new instance of the form
        Set objForm = VB.Forms.Add(strFormName)
        OpenForm = objForm.hwnd

    Else

        'If it is already open then return the handle
        OpenForm = objForm.hwnd

    End If

    Exit Function

eH:

    Err.Clear
    OpenForm = 0    'Return Error

End Function

Public Function GetWindowBorderStyle(ByVal lngHandle As Long) As Long

    Dim objForm As Form

    Set objForm = GetFormByHandle(lngHandle)

    If objForm Is Nothing Then
        GetWindowBorderStyle = -1
    Else
        GetWindowBorderStyle = objForm.BorderStyle
    End If

End Function

Public Sub SetWindowBorderStyle(ByVal lngHandle As Long, ByVal lngBorderStyle As Long)

    Dim objForm As Form

    Set objForm = GetFormByHandle(lngHandle)

    If Not objForm Is Nothing Then
        ChangeFormBorder objForm, lngBorderStyle
    End If

End Sub

Public Function GetWindowTitle(ByVal lngHandle As Long) As String

    Dim objForm As Form

    Set objForm = GetFormByHandle(lngHandle)

    If objForm Is Nothing Then
        GetWindowTitle = ""
    Else
        GetWindowTitle = objForm.Caption
    End If

End Function

Public Function GetWindowWidth(ByVal lngHandle As Long) As Long

    Dim objForm As Form

    Set objForm = GetFormByHandle(lngHandle)

    If objForm Is Nothing Then
        GetWindowWidth = 0
    Else
        GetWindowWidth = objForm.ScaleX(objForm.ScaleWidth, vbTwips, vbPixels)
    End If

End Function

Public Function GetWindowHeight(ByVal lngHandle As Long) As Long

    Dim objForm As Form

    Set objForm = GetFormByHandle(lngHandle)

    If objForm Is Nothing Then
        GetWindowHeight = 0
    Else
        GetWindowHeight = objForm.ScaleX(objForm.ScaleHeight, vbTwips, vbPixels)
    End If

End Function

Public Function GetWindowState(ByVal lngHandle As Long) As Long

    Dim objForm As Form

    Set objForm = GetFormByHandle(lngHandle)

    If objForm Is Nothing Then
        GetWindowState = 0
    Else
        GetWindowState = objForm.WindowState
    End If

End Function

Public Function GetStartUpPosition(ByVal lngHandle As Long) As Long

    Dim objForm As Form

    Set objForm = GetFormByHandle(lngHandle)

    If objForm Is Nothing Then
        GetStartUpPosition = 0
    Else
        GetStartUpPosition = objForm.StartUpPosition
    End If

End Function

Public Function ResizeForm(ByVal lngHandle As Long, _
                           ByVal lngAncho As Long, _
                           ByVal lngAlto As Long)

    Dim objForm As Form

    Set objForm = GetFormByHandle(lngHandle)

    If objForm Is Nothing Then Exit Function

    On Error Resume Next

    objForm.Move 0, 0, objForm.ScaleX(lngAncho, vbPixels, vbTwips), objForm.ScaleY(lngAlto, vbPixels, vbTwips)

End Function

Public Function IsFormOpen(ByVal lngHandle As Long) As Boolean

    Dim objForm As Form

    Set objForm = GetFormByHandle(lngHandle)

    If objForm Is Nothing Then
        IsFormOpen = False
    Else
        IsFormOpen = True
    End If

End Function

Public Function OpenFormCount() As Long

    OpenFormCount = VB.Forms.count

End Function

Public Sub CloseForm(ByVal lngHandle As Long)

    Dim objForm As Form

    Set objForm = GetFormByHandle(lngHandle)

    If objForm Is Nothing Then Exit Sub

    Unload objForm

End Sub

Public Function GetMessage(ByVal lngHandle As Long, _
                           ByRef strParameters As String) As String

    strParameters = gstrParameters
    GetMessage = gstrMessage

End Function

Public Sub DeleteMessage()

    gstrMessage = ""
    gstrParameters = ""

End Sub

'------------------------------------------------------------------------------
' Métodos privados
'------------------------------------------------------------------------------

Private Function GetFormVB6(ByVal strFormName As String) As Form

    Dim objForm As Form

    For Each objForm In VB.Forms

        If objForm.Name = strFormName Then
            Set GetFormVB6 = objForm

            Exit Function

        End If

    Next

    Set GetFormVB6 = Nothing

End Function

Private Function GetFormByHandle(ByVal lngHandle As Long) As Form

    Dim objForm As Form

    For Each objForm In VB.Forms

        If objForm.hwnd = lngHandle Then
            Set GetFormByHandle = objForm

            Exit Function

        End If

    Next

    Set GetFormByHandle = Nothing

End Function

'Public Function Showform(ByVal sform As String)
'    Main2 GetForm(sform)
'End Function

Public Sub PrintPreview(ByVal strFormName As String)

    Dim frm      As Form

    Dim oGrx     As GridEx

    Dim oControl As Control

    Set frm = GetFormVB6(strFormName)

    For Each oControl In frm.Controls

        If TypeOf oControl Is GridEX20.GridEx Then

            With oControl.PrinterProperties
                .RepeatHeaders = True
                .TranslateColors = True
                '.FooterString(jgexHFLeft) = Now
                'Right footer is set in the BeforePrintPage to indicate page number
            End With

            Load frmPreview
            frmPreview.Move frm.Left, frm.Top, frm.Width, frm.Height
            oControl.PrintPreview frmPreview.grPrev, False
            frmPreview.Show 1
            Set frmPreview = Nothing

            Exit For

        End If

    Next

End Sub

Public Sub GetControlFocus(ByVal lngHandle As Long)

    On Error GoTo errx

    Dim Success   As Long

    Dim SavedWnd  As Long

    Dim moveLines As Long

    Dim objForm   As Form

    Set objForm = GetFormByHandle(lngHandle)

    'save the window handle of the control that currently has focus
    'If Not objForm.ActiveControl Is Nothing Then
    SavedWnd = objForm.ActiveControl.hwnd
    'End If

    'Restore the focus to the original control
    Call PutFocus(lngHandle)
    'If Not objForm.ActiveControl Is Nothing Then
    Call PutFocus(SavedWnd)
    'End If
errx:
    'err.Raise 12555, "Configuración", "Configurando Ventana"
    MsgBox "Configurando Ventana", vbInformation, "Configuración"
    'MsgBox ""
End Sub

Public Property Get AnoPeriodoContableVigente() As String
    AnoPeriodoContableVigente = sAnoPeriodoContableVigente
End Property

Public Property Let AnoPeriodoContableVigente(ByVal vnewvalue As String)
    sAnoPeriodoContableVigente = vnewvalue
    sAnoContableVigente = Mid(sAnoPeriodoContableVigente, 1, 4)
    sPeriodoContableVigente = Mid(sAnoPeriodoContableVigente, 5, 2)
End Property
