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
Private mConnect As String

'Private Declare Function PutFocus Lib "user32" _
 '  Alias "SetFocus" _
  '(ByVal hwnd As Long) As Long

Public Property Get Language() As Integer
    Language = iLanguage
End Property

Public Property Let Language(ByVal vNewValue As Integer)
    iLanguage = vNewValue
End Property

Public Property Get Parent() As Object
    Set Parent = oParent
End Property

Public Property Let Parent(ByVal vNewValue As Object)
    Set oParent = vNewValue
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
Public Property Let Nom_Empresa(ByVal vNewValue As String)
    'gsNomEmpresa = vnewvalue
End Property
Public Property Let Nom_Aplicacion(ByVal vNewValue As String)
    'gsNomAplicacion = vnewvalue
End Property
Public Property Let Nom_Opcion(ByVal vNewValue As String)
    vNomFor = vNewValue
End Property
Public Property Let UserName(ByVal vNewValue As String)
    vusu = vNewValue
End Property
Public Property Get UserName() As String
    UserName = vusu
End Property
Public Property Let ConnectEmpresa(ByVal vNewValue As String)
    mConnect = vNewValue
    cCONNECT = mConnect
    
    
    
End Property
Public Property Let ConnectSeguridad(ByVal vNewValue As String)
    cSEGURIDAD = vNewValue
End Property
Public Function GetForm(sFormName) As Object

sFormName = Trim(UCase(sFormName))

Select Case sFormName
    Case "FRMSHOWFACTVENTAS"
        Set GetForm = New frmShowFactVentas
    Case "FRMSHOWGUIASXFACT_TELATENIDA"
        Set GetForm = New frmShowGuiasxFact_TelaTenida
    Case "FRMSHOWACTPRECIOOC"
        Set GetForm = New FrmShowActPrecioOC
    Case "FRMSHOWCONTROLNUMERACION"
        Set GetForm = New frmShowControlNumeracion
    Case "FRMCONVENTASREQ"
        Set GetForm = New frmConVentasReq

    Case "FRM_REGISTRO_VENTAS"
        Set GetForm = New Frm_Registro_Ventas

    Case "FRMRPTVENTASXGRUPO"
        Set GetForm = New FrmRptVentasxGrupo
    Case "FRMRPTRELACIONARTICULOS_X_GRUPOS"
        Set GetForm = New FrmRptRelacionArticulos_x_Grupos
    Case "FRMLETRA"
        Set GetForm = New frmLetra
 
    Case "FRMSHOWFACTVENTASPRENDASEXPORTACION"
        Set GetForm = New frmShowFactVentasPrendasExportacion
    Case "FRMMANTANXCLI"
        Set GetForm = New frmMantAnxCli
 
    Case "FRMTRANSACCIONES"
        Set GetForm = New frmTransacciones
    Case "FRMLIQUIDACIONDIARIA"
        Set GetForm = New frmLiquidacionDiaria
    Case "FRMCHEQUESDIFERIDOS"
        Set GetForm = New frmChequesDiferidos
    Case "FRMRPTCANCELACIONES_X_FACTURA"
        Set GetForm = New FrmRptCancelaciones_x_Factura
    Case "FRMRPTCANCELACIONES_X_FACTURA"
        Set GetForm = New FrmRptCancelaciones_x_Factura
  
    Case "FRMSHOWCTACTE"
             Set GetForm = New frmShowCtaCte
    Case "FRMADELANTOS"
             Set GetForm = New frmAdelantos
    Case "FRMRPTLETRASPENDIENTEPAGO"
          Set GetForm = New FrmRptLetrasPendientePago
    Case "FRMSHOWPARTEXAUTORIZAR"
          Set GetForm = New frmShowPartexAutorizar
    Case "FRMSHOWPARTESCOBRANZAS"
          Set GetForm = New frmShowPartesCobranzas
    Case "FRMSHOWSEGUIMDOCUMVENTAS"
          Set GetForm = New frmShowSeguimDocumVentas
    Case "FRMFACTURASDIFERIDAS"
          Set GetForm = New frmFacturasDiferidas
    Case "FRMRPTFACTURAS_CANJES"
        Set GetForm = New FrmRptFacturas_Canjes
    Case "FRMRPTCANCELACIONES_NOTASABONO"
        Set GetForm = New FrmRptCancelaciones_NotasAbono
    Case "FRMSHOWSEGUIMDRAWBACK"
        Set GetForm = New frmShowSeguimDrawBack
    Case "FRMSHOWCANJEAUTORIZAR"
        Set GetForm = New frmShowCanjeAutorizar

    Case "FRMRPTANEXOSCONTABLES"
        Set GetForm = New FrmRptAnexosContables
    Case "FRMRPTANTICIPOS_CANJES"
        Set GetForm = New FrmRptAnticipos_Canjes
    Case "FRMRPTLETRASSTATUS"
        Set GetForm = New FrmRptLetrasStatus

    Case "FRMRPTDETALLEEXPORT"
        Set GetForm = New FrmRptDetalleExport
    Case "FRMRPTCANCELACIONES_BOLETAS"
        Set GetForm = New FrmRptCancelaciones_Boletas
    Case "FRMRPTCANCELACIONES_INCOBRA"
        Set GetForm = New FrmRptCancelaciones_Incobra
    Case "FRMRPTLETRASDESCUENTOS"
        Set GetForm = New FrmRptLetrasDescuentos
    Case "FRNREPCANJELETRAS"
        Set GetForm = New frnRepCanjeLetras
    Case "FRMCIERREANOMES"
        Set GetForm = New frmCierreAnoMes
    Case "FRMRPTDETRACCIONES"
        Set GetForm = New frmRptDetracciones
    Case "FRMRPTVENTASXCLIENTE"
        Set GetForm = New FrmRptVentasxCliente
    Case "FRMRPTVENTASXGRUPOXTIPOVENTA"
        Set GetForm = New FrmRptVentasxGrupoxTipoVenta

    Case "FRMREPORTEANIOPERIODO"
        Set GetForm = New frmReporteAnioPeriodo

    Case "FRMCOBRANZAXPERIODO"
        Set GetForm = New frmCobranzaXPeriodo

    Case "FRMCONTROL_LETRAS"
        Set GetForm = New frmControl_Letras
    Case "FRMFLUJOCOBRANZA"
        Set GetForm = New frmFlujoCobranza
    Case "FRMRESUMENVENTASOBREPARTIDA"
        Set GetForm = New frmResumenVentaSobrePartida
    Case "FRMREPORTERESUMENANUALVENTAS"
        Set GetForm = New frmReporteResumenAnualVentas
    Case "FRMESTADISTICAVENTAS"
        Set GetForm = New frmEstadisticaVentas
    Case "FRMTIPOSCOBRANZA"
        Set GetForm = New frmTiposCobranza
    Case "FRMCONCEPTOCOBRANZA"
        Set GetForm = New frmConceptoCobranza
    Case "FRMSHOWCIERRETIPODIARIO_VENTAS"
        Set GetForm = New frmShowCierreTipoDiario_Ventas
    Case "FRMRANKINGVENTASPORPAISDESTINO"
        Set GetForm = New frmRankingVentasPorPaisDestino
    Case "FRMTRANSFACTVENTAS"
        Set GetForm = New frmTransFactVentas
    Case "FRM_FACTEXPOSUJETAS"
        Set GetForm = New Frm_FactExpoSujetas
    Case "FRMSHOWGUIASXFACT_LAVANDERIA"
        Set GetForm = New frmShowGuiasxFact_Lavanderia
    Case "FRMPENALIDADESVENTAS"
        Set GetForm = New FrmPenalidadesVentas
    Case "FRMRPTVENTASXTIPOVENTA"
        Set GetForm = New FrmRptVentasxTipoVenta

    Case "FRMSHOWGUIASXFACT_SALDOSTELATENIDA"
        Set GetForm = New frmShowGuiasxFact_SaldosTelaTenida
    Case "FRMMUESTRAHILOCOMPRADO"
        Set GetForm = New frmMuestraHiloComprado
    Case "FRMMOTIVO_NOTAS"
        Set GetForm = New FrmMotivo_Notas
     Case "FRMFACTURASEMIRANFECHA"
        Set GetForm = New frmFacturasEmiRanFecha
    Case "FRMESTADISTICAANUAL"
        Set GetForm = New FrmEstadisticaAnual
       
    Case "FRMCTACTECLIFACEXT"
        Set GetForm = New frmCtaCteCliFacExt
        
    
    Case "FRMFACTURAS_PENDIENTES_RECUPERACION_DRAW"
        Set GetForm = New FrmFacturas_Pendientes_Recuperacion_Draw
        

    Case "FRMSHOWTIPOSCAMBIO"
        Set GetForm = New frmShowTiposCambio
    Case "FRM_RESUMEN_VENTAS"
        Set GetForm = New Frm_Resumen_Ventas
    Case "FRMSHOWFACTVENTAS_LIQUIDAR"
        Set GetForm = New frmShowFactVentas_Liquidar
    Case "FRMFACTURASCANCELADASRANGO"
          Set GetForm = New frmFacturasCanceladasRango
End Select

End Function

Public Function OpenForm(ByVal strFormName As String, ByVal lngParentHwnd As Long) As Long

    Dim objForm As Form
    
    On Error GoTo eH

    'We don't want to allow creating multiple instances of the same form
    'so we return the existing form handle if it is in the Forms collection

    'Set objForm = GetFormVB6(strFormName)
    Set objForm = GetForm(strFormName)

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

    err.Clear
    OpenForm = 0 'Return Error
    
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

Public Function ResizeForm(ByVal lngHandle As Long, ByVal lngAncho As Long, ByVal lngAlto As Long)

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

    OpenFormCount = VB.Forms.Count

End Function

Public Sub CloseForm(ByVal lngHandle As Long)
    Dim objForm As Form
    Set objForm = GetFormByHandle(lngHandle)
    
    If objForm Is Nothing Then Exit Sub
    
    Dim oform As Form
    Set oform = GetFormVB6("Frm_Toolbar")
    
    If Not oform Is Nothing Then
        Unload oform
    End If
    Unload objForm
End Sub


Public Function GetMessage(ByVal lngHandle As Long, ByRef strParameters As String) As String

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


Public Sub PrintPreview(ByVal strFormName As String)
    Dim frm As Form
    Dim oGrx As GridEx
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





Public Property Get AnoPeriodoContableVigente() As String
    AnoPeriodoContableVigente = sAnoPeriodoContableVigente
End Property

Public Property Let AnoPeriodoContableVigente(ByVal vNewValue As String)
    sAnoPeriodoContableVigente = vNewValue
    sAnoContableVigente = Mid(sAnoPeriodoContableVigente, 1, 4)
    sPeriodoContableVigente = Mid(sAnoPeriodoContableVigente, 5, 2)
End Property


