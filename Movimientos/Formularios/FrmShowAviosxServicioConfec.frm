VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmShowAviosxServicioConfec 
   Caption         =   "Avios por Servicio Confecciones"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   12345
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   645
      Left            =   4237
      TabIndex        =   1
      Top             =   5640
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   1138
      Custom          =   $"FrmShowAviosxServicioConfec.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1200
      ControlHeigth   =   620
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX gexList 
      Height          =   5565
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   9816
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      TabKeyBehavior  =   1
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "FrmShowAviosxServicioConfec.frx":0103
      FormatStyle(2)  =   "FrmShowAviosxServicioConfec.frx":023B
      FormatStyle(3)  =   "FrmShowAviosxServicioConfec.frx":02EB
      FormatStyle(4)  =   "FrmShowAviosxServicioConfec.frx":039F
      FormatStyle(5)  =   "FrmShowAviosxServicioConfec.frx":0477
      FormatStyle(6)  =   "FrmShowAviosxServicioConfec.frx":052F
      FormatStyle(7)  =   "FrmShowAviosxServicioConfec.frx":060F
      ImageCount      =   0
      PrinterProperties=   "FrmShowAviosxServicioConfec.frx":062F
   End
End
Attribute VB_Name = "FrmShowAviosxServicioConfec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, sSeguridad As String
Public vCod_Fabrica As String, vCod_OrdPro As String


Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
    Call Reporte
Case "ACTUALIZAR"
    Call Actualizar
Case "SALIR"
    Unload Me
End Select
End Sub

Sub CARGA_GRID()
strSQL = "lg_muestra_envios_avios_por_np_servicio '" & vCod_Fabrica & "','" & vCod_OrdPro & "'"
Set Me.gexList.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)

gexList.Columns("des_proveedor").Width = 2000
gexList.Columns("cod_item").Width = 900
gexList.Columns("cod_unimed").Width = 700
gexList.Columns("des_comb").Width = 1800
gexList.Columns("estilo_cliente").Width = 1700
gexList.Columns("medida").Width = 1500
gexList.Columns("color").Width = 2000
gexList.Columns("cantidad_enviada").Width = 1400

gexList.Columns("COD_PROVEEDOR").Width = 0
gexList.Columns("COD_COMB").Width = 0
gexList.Columns("cod_color").Width = 0
gexList.Columns("COD_TALLA").Width = 0
gexList.Columns("cod_destino").Width = 0
gexList.Columns("Estilo_Cliente").Width = 0

gexList.Columns("Can_Devuelta").Width = 1000
gexList.Columns("Fec_Devolucion").Width = 1000
gexList.Columns("Observaciones").Width = 1000

gexList.Columns("des_proveedor").Caption = "Proveedor"
gexList.Columns("cod_item").Caption = "Item"
gexList.Columns("cod_unimed").Caption = "Uni.Med."
gexList.Columns("des_comb").Caption = "Comb."
gexList.Columns("estilo_cliente").Caption = "Estilo Cliente"
gexList.Columns("cantidad_enviada").Caption = "Cant. Enviada"

End Sub

Private Sub Reporte()
On Error GoTo hand
Dim oo As Object, Ruta As String

    Ruta = vRuta & "\RptAviosxServConfec.xlt"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "reporte", strSQL, vCod_OrdPro, cConnect
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "Reporte"
    Set oo = Nothing
End Sub

Private Sub Actualizar()
    Load FrmDetalleAvios
    'SCOD_PROVEEDOR, sCOD_ITEM,sCOD_COMB,scod_color,sCOD_TALLA,sCod_Descuento,sEstilo_Cliente
    FrmDetalleAvios.sCod_Fabrica = vCod_Fabrica
    FrmDetalleAvios.sCod_OrdPro = vCod_OrdPro
    FrmDetalleAvios.SCOD_PROVEEDOR = gexList.Value(gexList.Columns("COD_PROVEEDOR").Index)
    FrmDetalleAvios.sCOD_ITEM = Trim(gexList.Value(gexList.Columns("COD_ITEM").Index))
    FrmDetalleAvios.sCOD_COMB = Trim(gexList.Value(gexList.Columns("COD_COMB").Index))
    FrmDetalleAvios.scod_color = Trim(gexList.Value(gexList.Columns("cod_color").Index))
    FrmDetalleAvios.sCOD_TALLA = Trim(gexList.Value(gexList.Columns("COD_TALLA").Index))
    FrmDetalleAvios.sCod_Descuento = Mid(gexList.Value(gexList.Columns("cod_destino").Index), 1, 6)
    FrmDetalleAvios.sEstilo_Cliente = gexList.Value(gexList.Columns("Estilo_Cliente").Index)
    FrmDetalleAvios.sCantidad_Enviada = gexList.Value(gexList.Columns("Cantidad_Enviada").Index)
    
    FrmDetalleAvios.TxtCantidad.Text = gexList.Value(gexList.Columns("Can_Devuelta").Index)
    If IsNull(gexList.Value(gexList.Columns("Fec_Devolucion").Index)) Then
        FrmDetalleAvios.DTPFecha = Date
    Else
        FrmDetalleAvios.DTPFecha = gexList.Value(gexList.Columns("Fec_Devolucion").Index)
    End If
    FrmDetalleAvios.txtobservacion.Text = gexList.Value(gexList.Columns("Observaciones").Index)
    
    FrmDetalleAvios.Show vbModal
    Set FrmDetalleAvios = Nothing
End Sub

