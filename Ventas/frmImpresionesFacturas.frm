VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmImpresionesFacturas 
   Caption         =   "Imprimir Facturas"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3780
   LinkTopic       =   "frmConfirmacionDespacho"
   ScaleHeight     =   1665
   ScaleWidth      =   3780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   1245
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   2196
         Custom          =   $"frmImpresionesFacturas.frx":0000
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   3300
         ControlHeigth   =   580
         ControlSeparator=   60
      End
   End
End
Attribute VB_Name = "frmImpresionesFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public scoddestino As String, sdesdestino As String, scodembarque As String, sdesembarque As String
 Public SNum_Corre As String, SImp_Total As Double, SCod_TipDoc As String
     
Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo HandlerError

    Dim Msg As Variant
    
    Select Case ActionName
    
    Case "IMPRIMIR1"
        Imprimir SNum_Corre, SImp_Total, False, SCod_TipDoc
    Case "FIJO"
        Imprimir_Fijo SNum_Corre, SImp_Total, False, SCod_TipDoc
    Case "IMPRIMIR2"
        Imprimir_Exportacion_Prendas SNum_Corre, SCod_TipDoc
        'Carga_Data
        'Imprimir_Exp SNum_Corre, SCod_TipDoc, SImp_Total
    Case "IMPRIMIR3"
        Carga_Data
        Imprimir_Sunat SNum_Corre, SCod_TipDoc, SImp_Total
    Case "DEVANLAY"
        Carga_Data
        Imprimir_Devanlay SNum_Corre, SCod_TipDoc, SImp_Total
    Case "SALIR"
          Unload Me
    End Select
Exit Sub
Resume
HandlerError:
    MsgBox err.Description, vbCritical, "Mensaje del Sistema"
End Sub

Sub Carga_Data()

Dim RS As ADODB.Recordset

Set RS = CargarRecordSetDesconectado("Ventas_Up_Man 'V','" & SNum_Corre & "'", cCONNECT)

With RS
  If Not (.BOF Or .EOF) Then
   ' With frmAdicionaDocumVentas
        
      scodembarque = RS!Tip_Embarque
      sdesembarque = RS!Des_TipEmbarque

      scoddestino = RS!Cod_Destino
      sdesdestino = RS!Des_Destino
    
   ' End With
  End If
End With

End Sub

Private Sub Imprimir_Exp(ByVal SNum_Corre As String, ByVal SCod_TipDoc As String, dbImp_Total As Double)
On Error GoTo ErrorImpresion
Dim oo As Object
Dim sFormato_Invoice As String

   ' sFormato_Invoice = DevuelveCampo("SELECT FORMATO_INVOICE FROM TG_CLIENTE WHERE COD_CLIENTE = '" & GridEX1.Value(GridEX1.Columns("COD_CLIENTE").Index) & "'", cCONNECT)
    Set oo = CreateObject("excel.application")
   ' Select Case sCod_Tipdoc
    '    Case "FA"
            oo.Workbooks.Open vRuta & "\Invoice03.XLT" ' & sFormato_Invoice & ".XLT"
   ' End Select
    oo.Visible = True
    oo.displayalerts = False
    oo.Run "reporte", cCONNECT, SNum_Corre, UCase(EnLetras(Trim(CStr(dbImp_Total)))), scoddestino, sdesdestino, scodembarque, sdesembarque
    Set oo = Nothing
       
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub

Private Sub Imprimir_Sunat(ByVal SNum_Corre As String, ByVal SCod_TipDoc As String, dbImp_Total As Double)
On Error GoTo ErrorImpresion
Dim oo As Object
Dim sFormato_Invoice As String

   ' sFormato_Invoice = DevuelveCampo("SELECT FORMATO_INVOICE FROM TG_CLIENTE WHERE COD_CLIENTE = '" & GridEX1.Value(GridEX1.Columns("COD_CLIENTE").Index) & "'", cCONNECT)
    Set oo = CreateObject("excel.application")
   ' Select Case sCod_Tipdoc
    '    Case "FA"
            oo.Workbooks.Open vRuta & "\Invoice04.XLT" ' & sFormato_Invoice & ".XLT"
   ' End Select
    oo.Visible = True
    oo.displayalerts = False
    oo.Run "reporte", cCONNECT, SNum_Corre, UCase(EnLetras(Trim(CStr(dbImp_Total)))), scoddestino, sdesdestino, scodembarque, sdesembarque
    Set oo = Nothing
       
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub


Private Sub Imprimir_Devanlay(ByVal SNum_Corre As String, ByVal SCod_TipDoc As String, dbImp_Total As Double)
On Error GoTo ErrorImpresion
Dim oo As Object
Dim sFormato_Invoice As String

    Set oo = CreateObject("excel.application")

    oo.Workbooks.Open vRuta & "\Invoice06.XLT"

    oo.Visible = True
    oo.displayalerts = False
    oo.Run "reporte", cCONNECT, SNum_Corre, UCase(EnLetras(Trim(CStr(dbImp_Total)))), scoddestino, sdesdestino, scodembarque, sdesembarque
    Set oo = Nothing
       
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub

Private Sub Imprimir_Exportacion_Prendas(ByVal SNum_Corre As String, ByVal SCod_TipDoc As String)
On Error GoTo ErrorImpresion
Dim oo As Object
Dim sFormato_Invoice As String
Dim strSQL As String
Dim sRutaLogo As String
Dim sCod_Cliente As String

    strSQL = "SELECT Ruta_Logo = ISNULL(Ruta_Logo, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
    sRutaLogo = DevuelveCampo(strSQL, cCONNECT)
    
    Dim sEmpresa As String
    strSQL = "SELECT Des_Empresa = ISNULL(Des_Empresa, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
    sEmpresa = DevuelveCampo(strSQL, cCONNECT)
        
    sCod_Cliente = DevuelveCampo("select cod_cliente from cn_ventas where num_corre ='" & SNum_Corre & "'", cCONNECT)

    sFormato_Invoice = DevuelveCampo("SELECT FORMATO_INVOICE FROM TG_CLIENTE WHERE COD_CLIENTE = '" & sCod_Cliente & "'", cCONNECT)
    Set oo = CreateObject("excel.application")
    Select Case SCod_TipDoc
        Case "FA"
            oo.Workbooks.Open vRuta & "\Invoice" & sFormato_Invoice & ".XLT"
    End Select
    oo.Visible = True
    oo.displayalerts = False
    
    If sFormato_Invoice = "01" Then
        oo.Run "reporte", cCONNECT, SNum_Corre, sEmpresa, sRutaLogo
    Else
        oo.Run "reporte", cCONNECT, SNum_Corre
    End If
    
    Set oo = Nothing
       
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub

