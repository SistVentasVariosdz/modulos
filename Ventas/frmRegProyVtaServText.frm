VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmRegProyVtaServText 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro Proyeccion Ventas-Servicios Textiles"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin GridEX20.GridEX grxData 
      Height          =   4365
      Left            =   90
      TabIndex        =   6
      Top             =   930
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   7699
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmRegProyVtaServText.frx":0000
      Column(2)       =   "frmRegProyVtaServText.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmRegProyVtaServText.frx":016C
      FormatStyle(2)  =   "frmRegProyVtaServText.frx":02A4
      FormatStyle(3)  =   "frmRegProyVtaServText.frx":0354
      FormatStyle(4)  =   "frmRegProyVtaServText.frx":0408
      FormatStyle(5)  =   "frmRegProyVtaServText.frx":04E0
      FormatStyle(6)  =   "frmRegProyVtaServText.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmRegProyVtaServText.frx":0678
   End
   Begin VB.Frame fraBuscar 
      Height          =   915
      Left            =   60
      TabIndex        =   0
      Top             =   -30
      Width           =   10095
      Begin VB.TextBox txtStatusDes 
         Height          =   285
         Left            =   1980
         TabIndex        =   8
         Text            =   "PENDIENTE"
         Top             =   225
         Width           =   1425
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "Status"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   270
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton optNroProyeccion 
         Caption         =   "Nro Proyeccion"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1395
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   435
         Left            =   8760
         TabIndex        =   3
         Top             =   330
         Width           =   1215
      End
      Begin VB.TextBox txtNroProyecccion 
         Height          =   285
         Left            =   1590
         TabIndex        =   2
         Top             =   555
         Width           =   1815
      End
      Begin VB.TextBox txtStatus 
         Height          =   285
         Left            =   1590
         TabIndex        =   1
         Text            =   "P"
         Top             =   225
         Width           =   375
      End
   End
   Begin FunctionsButtons.FunctButt fnbOperacion 
      Height          =   510
      Left            =   1973
      TabIndex        =   7
      Top             =   5340
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   900
      Custom          =   $"frmRegProyVtaServText.frx":0850
      Orientacion     =   0
      Style           =   0
      Language        =   1
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   3540
      Top             =   5970
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmRegProyVtaServText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Codigo As String
Public Descripcion  As String

Private Sub cmdBuscar_Click()
Dim adoRs As ADODB.Recordset
Dim strSQL As String
Dim strOpcion As String
strOpcion = IIf(optStatus.Value, "2", "1")
strSQL = "EXEC ventas_muestra_proyeccion_textil_status '" & strOpcion & "','" & txtNroProyecccion.Text & "','" & txtStatus.Text & "'"
Set adoRs = CargarRecordSetDesconectado(strSQL, cCONNECT)
Set grxData.ADORecordset = adoRs
Call CONFIGURAR_GRID

End Sub

Public Sub CONFIGURAR_GRID()
    grxData.Columns("Id_Proyeccion").Caption = "Id Proyeccion"
    grxData.Columns("Id_Proyeccion").Width = "700"
    grxData.Columns("Cod_Tipo_Venta").Width = "0"
    grxData.Columns("Nombre_Venta").Caption = "Nom.Venta"
    grxData.Columns("Nombre_Venta").Width = "1600"
    grxData.Columns("Nom_Cliente").Caption = "Nom.Cliente"
    grxData.Columns("Nom_Cliente").Width = "1600"
    grxData.Columns("Fec_Creacion").Caption = "Fec.Creacion"
    grxData.Columns("Fec_Creacion").Width = "1200"
    grxData.Columns("Status").Width = "1200"
    grxData.Columns("Kgs_Requeridos").Caption = "Kgs.Requeridos"
    grxData.Columns("Kgs_Requeridos").Width = "1000"
    grxData.Columns("Fec_Requerimiento").Caption = "Fec.Requerimiento"
    grxData.Columns("Fec_Requerimiento").Width = "1200"
    grxData.Columns("Cod_Hilado").Caption = "Cod.Hilado"
    grxData.Columns("Cod_Hilado").Width = "1200"
    grxData.Columns("Cod_Tela").Caption = "Cod.Tela"
    grxData.Columns("Cod_Tela").Width = "1200"
    grxData.Columns("Nombre").Caption = "Nombre"
    grxData.Columns("Nombre").Width = "2000"
    grxData.Columns("Observaciones").Width = "1500"
    grxData.Columns("cod_cliente").Width = "0"
End Sub

Private Sub fnbOperacion_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ADICIONAR"
    frmMan_RegProyVtaServText.sOpcion = "I"
    frmMan_RegProyVtaServText.Sid_proyeccion = 0
    frmMan_RegProyVtaServText.Show 1
    cmdBuscar_Click
Case "MODIFICAR"
    frmMan_RegProyVtaServText.sOpcion = "U"
    frmMan_RegProyVtaServText.Sid_proyeccion = grxData.Value(grxData.Columns("Id_Proyeccion").Index)
    frmMan_RegProyVtaServText.TxtCod_TipoVenta = grxData.Value(grxData.Columns("Cod_Tipo_Venta").Index)
    frmMan_RegProyVtaServText.TxtDes_TipoVenta = grxData.Value(grxData.Columns("Nombre_Venta").Index)

    frmMan_RegProyVtaServText.txtAbr_Cliente = DevuelveCampo("SELECT Abr_Cliente FROM tx_cliente WHERE Cod_Cliente_Tex = '" & grxData.Value(grxData.Columns("cod_cliente").Index) & "'", cCONNECT)
    frmMan_RegProyVtaServText.txtNom_cliente = grxData.Value(grxData.Columns("Nom_Cliente").Index)
    
    frmMan_RegProyVtaServText.DTPInicio = grxData.Value(grxData.Columns("Fec_Requerimiento").Index)
    frmMan_RegProyVtaServText.txtkilos = grxData.Value(grxData.Columns("Kgs_Requeridos").Index)

    frmMan_RegProyVtaServText.txtOPCION_COD = grxData.Value(grxData.Columns("Cod_Hilado").Index)
    frmMan_RegProyVtaServText.txtHILADO_DES = DevuelveCampo("SELECT DESCRIPCION FROM HI_HILADOS WHERE COD_HILADO = '" & grxData.Value(grxData.Columns("Cod_Hilado").Index) & "'", cCONNECT)
    frmMan_RegProyVtaServText.txtcod_tela = grxData.Value(grxData.Columns("Cod_Tela").Index)
    frmMan_RegProyVtaServText.txtdes_tela = DevuelveCampo("SELECT Des_Tela FROM TX_TELA WHERE Cod_Tela = '" & grxData.Value(grxData.Columns("Cod_Tela").Index) & "'", cCONNECT)
    If DevuelveCampo("SELECT cod_grupo_ventas FROM CN_TIPOS_VENTA WHERE cod_grupo_ventas in ('1','2','3') AND Cod_Tipo_Venta= '" & grxData.Value(grxData.Columns("Cod_Tipo_Venta").Index) & "'", cCONNECT) = "1" Then
        frmMan_RegProyVtaServText.Frame2.Visible = True
        frmMan_RegProyVtaServText.Frame3.Visible = False
    Else
        frmMan_RegProyVtaServText.Frame3.Visible = True
        frmMan_RegProyVtaServText.Frame2.Visible = False
    End If
    frmMan_RegProyVtaServText.txtobservacion = grxData.Value(grxData.Columns("Observaciones").Index)
    frmMan_RegProyVtaServText.Show 1
    cmdBuscar_Click
Case "ELIMINAR"
            ELIMINAR = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Combinación-Detalle")
            If ELIMINAR = vbYes Then
                sTipo = "D"
                Call Eliminar_Datos
                cmdBuscar_Click
            End If
Case "IMPRIMIR"
    Call Reporte
Case "SALIR"
    Unload Me
End Select
End Sub

Sub Eliminar_Datos()
    Dim strSQL As String
    On Error GoTo Salvar_DatosErr

 
    strSQL = "EXEC ventas_up_act_proyeccion_textil_status 'D','" & Trim(grxData.Value(grxData.Columns("id_proyeccion").Index)) & "','','','" & Date & "',0,'','','',''"
      
    ExecuteCommandSQL cCONNECT, strSQL

    'Dim amensaje As New clsMessages
    'amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
    'Informa "", amensaje
    MsgBox "Registro eliminado satisfactoriamente......", vbInformation, Me.Caption
    
    Exit Sub
Salvar_DatosErr:
    ErrorHandler err, "Salvar_Datos"
End Sub

Sub Reporte()
On Error GoTo ErrorImpresion
Dim oo As Object
Dim adoRs As ADODB.Recordset
Dim strSQL As String

strSQL = "ventas_proyeccion_textil_cuadro_general "

Set adoRs = CargarRecordSetDesconectado(strSQL, cCONNECT)

If adoRs.RecordCount = 0 Then
    MsgBox "No hay datos para mostrar...verificar", vbInformation, "Mensaje del sistema"
    Exit Sub
End If

Set oo = CreateObject("excel.application")
oo.Workbooks.Open vRuta & "\rptCuadroProyeccionVentaTextiles.XLT"
oo.Visible = True
oo.DisplayAlerts = False

oo.Run "reporte", adoRs
Set oo = Nothing
Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub


Private Sub optNroProyeccion_Click()
If optNroProyeccion.Value Then
    txtNroProyecccion.Enabled = True
Else
    txtNroProyecccion.Enabled = False
End If

End Sub

Private Sub optStatus_Click()
If optStatus.Value Then
    txtStatus.Enabled = True
Else
    txtStatus.Enabled = False
End If
End Sub

Private Sub txtNroProyecccion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmdBuscar.SetFocus
End If

End Sub

Private Sub txtStatus_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call BuscarStatus("1")
End If
End Sub

Private Sub BuscarStatus(ByVal Opcion As String)
Dim adoRs As ADODB.Recordset
Dim strSQL As String
On Error GoTo lblError

    strSQL = "SELECT Flg_Status,Descripcion FROM Ventas_Proyeccion_Textil_Status "

    txtStatus.Text = Trim(txtStatus.Text)
    txtStatusDes.Text = Trim(txtStatusDes)
    
    Select Case Opcion
        Case 1: strSQL = strSQL & " WHERE Flg_Status " & " LIKE '%" & txtStatus.Text & "%'"
        Case 1: strSQL = strSQL & " WHERE Descripcion " & " LIKE '%" & txtStatusDes.Text & "%'"
    End Select
    
    txtStatus.Text = ""
    txtStatusDes.Text = ""
    
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = strSQL
        .CARGAR_DATOS
        
        Codigo = ".."
        Set adoRs = .DGridLista.ADORecordset
        If adoRs.RecordCount > 1 Then .Show vbModal
        
        If Codigo <> "" And adoRs.RecordCount > 0 Then
            txtStatus = Trim(adoRs!Flg_Status)
            txtStatusDes = Trim(adoRs!Descripcion)
            Select Case Opcion
                Case 1: SendKeys "{TAB}": SendKeys "{TAB}"
                Case 2: SendKeys "{TAB}"
            End Select
        Else
            SendKeys "{TAB}"
        End If
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    adoRs.Close
    Set adoRs = Nothing
Exit Sub
lblError:
    Set frmBusqGeneral = Nothing
    adoRs.Close
    Set adoRs = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, "Mensaje del Sistema"
End Sub


Private Sub txtStatusDes_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call BuscarStatus("2")
End If

End Sub
