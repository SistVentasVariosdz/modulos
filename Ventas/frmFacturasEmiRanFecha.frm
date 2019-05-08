VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmFacturasEmiRanFecha 
   Caption         =   "Facturas Emitidas Segun Rango de Fechas"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "Resumen"
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   16
      Top             =   2760
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Detalle"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   15
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.TextBox txtNum_Ruc 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   960
         MaxLength       =   11
         TabIndex        =   13
         Top             =   1440
         Width           =   1200
      End
      Begin VB.TextBox txtDes_Anexo 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   12
         Top             =   1440
         Width           =   4410
      End
      Begin VB.TextBox Txt_Cod_Usuario 
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Txt_DesUsuario 
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Top             =   1080
         Width           =   4410
      End
      Begin VB.OptionButton optCanceladas 
         Caption         =   "Canceladas"
         Height          =   195
         Left            =   4800
         TabIndex        =   8
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton OptPendientesPago 
         Caption         =   "Pendientes de Pago"
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   2160
         Width           =   1935
      End
      Begin VB.OptionButton OptTodas 
         Caption         =   "Todas"
         Height          =   195
         Left            =   720
         TabIndex        =   6
         Top             =   2160
         Value           =   -1  'True
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   69861379
         CurrentDate     =   39018
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   285
         Left            =   4200
         TabIndex        =   3
         Top             =   480
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   69861379
         CurrentDate     =   39018
      End
      Begin VB.Label Label4 
         Caption         =   "Nro Ruc:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta:"
         Height          =   210
         Left            =   3480
         TabIndex        =   4
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Desde:"
         Height          =   210
         Left            =   600
         TabIndex        =   2
         Top             =   480
         Width           =   540
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   615
      Left            =   1800
      TabIndex        =   5
      Top             =   3240
      Width           =   3240
      _ExtentX        =   5556
      _ExtentY        =   1111
      Custom          =   $"frmFacturasEmiRanFecha.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1500
      ControlHeigth   =   600
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   360
      Top             =   2160
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmFacturasEmiRanFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strSQL As String
Public vopcion As String
Public sNomOpcion As String
Public codigo As String
Public Descripcion As String
Public stipo_reporte  As Integer
Public scodclienteAne As String
Public strCod_Anxo As String
Private Sub Form_Load()
dtpInicio.Value = Date
dtpFin.Value = Date
stipo_reporte = 1
End Sub


 Public Sub ImprimirReporte()
 On Error GoTo ErrorImpresion
 Dim oo As Object
  
Dim Adors1 As Object
Set Adors1 = CreateObject("ADODB.Recordset")
Dim rutaLogo As String
rutaLogo = DevuelveCampo("select ruta_logo=isNUll(ruta_logo,'') from seguridad..seg_empresas where cod_empresa='" & vemp & "'", cCONNECT)


If OptTodas.Value = True Then
vopcion = "1"
sNomOpcion = OptTodas.Caption
ElseIf OptPendientesPago.Value = True Then
vopcion = "2"
sNomOpcion = OptPendientesPago.Caption
Else
vopcion = "3"
sNomOpcion = optCanceladas.Caption
End If


strSQL = " EXEC cn_ventas_muestra_facturas_segun_estatus '" & dtpInicio.Value & "','" & dtpFin.Value & "','" & vopcion & "','" & Left(Txt_Cod_Usuario, 1) & "','" & Right(Txt_Cod_Usuario, 4) & "','C','" & scodclienteAne & "'"

Set Adors1 = CargarRecordSetDesconectado(strSQL, cCONNECT)

Set oo = CreateObject("Excel.Application")
    oo.Workbooks.Open vRuta & "\Rpt_Facturas_Emitidas.XLT"
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "Reporte", rutaLogo, Adors1, dtpInicio.Value, dtpFin.Value, sNomOpcion
Set oo = Nothing
 

Exit Sub
ErrorImpresion:

   Screen.MousePointer = vbNormal
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte" & err.Description, vbCritical, "Impresion"
End Sub
 Public Sub ImprimirReporte_detalle()
 On Error GoTo ErrorImpresion
 Dim oo As Object
  
Dim Adors1 As Object
Set Adors1 = CreateObject("ADODB.Recordset")
Dim rutaLogo As String
rutaLogo = DevuelveCampo("select ruta_logo=isNUll(ruta_logo,'') from seguridad..seg_empresas where cod_empresa='" & vemp & "'", cCONNECT)


If OptTodas.Value = True Then
vopcion = "1"
sNomOpcion = OptTodas.Caption
ElseIf OptPendientesPago.Value = True Then
vopcion = "2"
sNomOpcion = OptPendientesPago.Caption
Else
vopcion = "3"
sNomOpcion = optCanceladas.Caption
End If


strSQL = " EXEC cn_ventas_muestra_facturas_segun_estatus_detalle '" & dtpInicio.Value & "','" & dtpFin.Value & "','" & vopcion & "','" & Left(Txt_Cod_Usuario, 1) & "','" & Right(Txt_Cod_Usuario, 4) & "','C','" & scodclienteAne & "'"

Set Adors1 = CargarRecordSetDesconectado(strSQL, cCONNECT)

Set oo = CreateObject("Excel.Application")
    oo.Workbooks.Open vRuta & "\Rpt_Facturas_Emitidas_detalle.XLT"
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "Reporte", rutaLogo, Adors1, dtpInicio.Value, dtpFin.Value, sNomOpcion
Set oo = Nothing
 

Exit Sub
ErrorImpresion:

   Screen.MousePointer = vbNormal
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte" & err.Description, vbCritical, "Impresion"
End Sub
Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "IMPRIMIR"
    
        If Trim(txtNum_Ruc.Text) = "" Or Trim(txtDes_Anexo.Text = "") Then scodclienteAne = ""
        
        If stipo_reporte = 1 Then
        Call ImprimirReporte
        
        Else
        Call ImprimirReporte_detalle
        End If
        
    Case "SALIR"
        Unload Me
End Select
End Sub


Public Sub Busca_Trabajador()
On Error GoTo Fin
Dim iCol As Long
Dim rstAux As New ADODB.Recordset
Dim Opcion As String
Dim strSQL As String
strSQL = "Tg_Sm_Muestra_Operario_Caracteristica '001'"
    With frmBusqGeneralOperario
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("Codigo").Caption = "Codigo"
        .DGridLista.Columns("Codigo").Width = 900
        .DGridLista.Columns("Apellido_Paterno").Caption = "Ape Paterno"
        .DGridLista.Columns("Apellido_Paterno").Width = 1500
        .DGridLista.Columns("Apellido_Materno").Caption = "Ape Materno"
        .DGridLista.Columns("Apellido_Materno").Width = 1500
        .DGridLista.Columns("Nombre_Trabajador").Caption = "Nombres"
        .DGridLista.Columns("Nombre_Trabajador").Width = 1500
        
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If codigo <> "" And rstAux.RecordCount > 0 Then
            Txt_Cod_Usuario = Trim(rstAux!codigo)
            Txt_Cod_Usuario.Tag = Left(Trim(rstAux!codigo), 1)
            Txt_DesUsuario = Trim(rstAux!Apellido_Paterno) + " " + Trim(rstAux!Apellido_Materno) + " " + Trim(rstAux!Nombre_Trabajador)
            Txt_DesUsuario.Tag = Right(Trim(rstAux!codigo), 4)
            'stip_Trabajador = Left(rstAux!codigo, 1)
            'scod_trabajador = Right(rstAux!codigo, 4)
        End If
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Color (" & Opcion & ")"
End Sub


Private Sub Option1_Click(Index As Integer)
stipo_reporte = Index + 1
End Sub

Private Sub Txt_Cod_Usuario_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Busca_Trabajador
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Txt_DesUsuario_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Busca_Trabajador
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  Call Busca_Opcion_Anexo("Num_Ruc", "Des_Anexo", " Cn_AnexosContables Where cod_tipanex ='C' AND TIP_TRABAJADOR+'-'+COD_TRABAJADOR = '" & Txt_Cod_Usuario.Text & "' and ", txtNum_Ruc, txtDes_Anexo, 1, Me)
  SendKeys "{TAB}"
 End If
End Sub


Private Sub txtDes_Anexo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion_Anexo("Num_Ruc", "Des_Anexo", " Cn_AnexosContables Where cod_tipanex ='C' AND TIP_TRABAJADOR+'-'+COD_TRABAJADOR = '" & Txt_Cod_Usuario.Text & "' and ", txtNum_Ruc, txtDes_Anexo, 2, Me)
End Sub

Sub Busca_Opcion_Anexo(strCampo1 As String, strCampo2 As String, StrTabla As String, txtCod As TextBox, txtDes As TextBox, Opcion As Integer, frmME As Form)

On Error GoTo Fin

Dim rstAux As Object, strSQL As String
Set rstAux = CreateObject("ADODB.Recordset")
    strSQL = "select Cod_Anxo as Cod,Des_Anexo as Nombre,Num_Ruc as Ruc from " & StrTabla
    
    'StrSql = "Select " & strCampo1 & " AS Cod," & strCampo2 & " as Descripcion from " & StrTabla

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    Select Case Opcion
    Case 1: strSQL = strSQL & strCampo1 & " like '%" & txtCod & "%'"
    Case 2: strSQL = strSQL & strCampo2 & " like '%" & txtDes & "%'"
    End Select
    txtCod = ""
    txtDes = ""
    scodclienteAne = ""
    frmME.strCod_Anxo = ""
    With frmBusqGeneral
        Set .oParent = frmME
        .sQuery = strSQL
        .Cargar_Datos
        
        codigo = ""
        .DGridLista.Columns("Cod").Visible = False
        .DGridLista.Columns("Nombre").Width = 4575
        .DGridLista.Columns("RUC").Width = 1695
        Set rstAux = .DGridLista.ADORecordset
        
        If rstAux.RecordCount > 1 Then
          .Show vbModal
        Else
          frmME.codigo = ".."
        End If
        If frmME.codigo <> "" And rstAux.RecordCount > 0 Then
            frmME.strCod_Anxo = Trim(rstAux!Cod)
            txtDes = Trim(rstAux!Nombre)
            txtCod = Trim(rstAux!Ruc)
            scodclienteAne = rstAux!Cod
            Select Case Opcion
            Case 1: SendKeys "{TAB}"
            Case 2: SendKeys "{TAB}"
            End Select
        Else
            SendKeys "{TAB}"
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & Opcion & ")"
End Sub


