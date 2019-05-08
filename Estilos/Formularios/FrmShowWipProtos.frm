VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmShowWipProtos 
   Caption         =   "Seguimiento Wip Protos"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   12165
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   540
      Left            =   3360
      TabIndex        =   1
      Top             =   4440
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   953
      Custom          =   $"FrmShowWipProtos.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1300
      ControlHeigth   =   520
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   7646
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
      Column(1)       =   "FrmShowWipProtos.frx":015C
      Column(2)       =   "FrmShowWipProtos.frx":0224
      FormatStylesCount=   6
      FormatStyle(1)  =   "FrmShowWipProtos.frx":02C8
      FormatStyle(2)  =   "FrmShowWipProtos.frx":0400
      FormatStyle(3)  =   "FrmShowWipProtos.frx":04B0
      FormatStyle(4)  =   "FrmShowWipProtos.frx":0564
      FormatStyle(5)  =   "FrmShowWipProtos.frx":063C
      FormatStyle(6)  =   "FrmShowWipProtos.frx":06F4
      ImageCount      =   0
      PrinterProperties=   "FrmShowWipProtos.frx":07D4
   End
End
Attribute VB_Name = "FrmShowWipProtos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public vCod_Cliente As String, vcod_TemCli As String, vCod_EstCli As String, vCod_estPro As String
Public Cliente As String, Temporada As String
Public varNumCot As Integer, varObs As String, varDes_Cliente As String, varNom_TemCli As String

Sub CARGA_GRID()

strSQL = "Tg_Muestra_EstCliEst_Protos '" & vCod_Cliente & "','" & vcod_TemCli & "','" & vCod_EstCli & "','" & vCod_estPro & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

GridEX1.Columns("Usuario_Comercial").Width = 1300
GridEX1.Columns("Modelista").Width = 1250
GridEX1.Columns("Fec_Recepcion_Consumos").Width = 1900
GridEX1.Columns("Fec_Recepcion_moldes").Width = 1900
GridEX1.Columns("cod_ordproproto").Width = 1000
GridEX1.Columns("Fec_Recepcion_Consumos").Caption = "Recepcion Consumos(Moldes)"

GridEX1.Columns("cod_ordproproto").Caption = "Proto"

GridEX1.FrozenColumns = 3
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
    If GridEX1.RowCount = 0 Then Exit Sub
    Call Reporte
Case "ESTILO"
    If GridEX1.RowCount = 0 Then Exit Sub
    Call GeneraReportes
Case "ELIMINAR"
    If GridEX1.RowCount = 0 Then Exit Sub
    Call Del_Iteracion
    
Case "SALIR"
    Unload Me
End Select
End Sub

Public Sub Reporte()
On Error GoTo ErrorImpresion
    strSQL = "Tg_Muestra_EstCliEst_Protos '" & vCod_Cliente & "','" & vcod_TemCli & "','" & vCod_EstCli & "','" & vCod_estPro & "'"
    Dim oo As Object
    Set oo = CreateObject("excel.application")
    oo.workbooks.Open vRuta & "\RptImpWipProtos.xlt"
    oo.Visible = True
    oo.run "REPORTE", strSQL, Cliente, Temporada, vCod_EstCli, vCod_estPro, cCONNECT
    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte WipProtos " & Err.Description, vbCritical, "Impresion"
End Sub

Sub GeneraReportes()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String
Dim strSQL As String
Dim vNum_Iteracion As Integer

    Ruta = vRuta & "\PROTOTIPOD_Estilo.xlt"

    'strSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(varAbr_Cliente) & "'"
    'vCod_Cliente = CStr(DevuelveCampo(strSQL, cCONNECT))
    
    'vCod_Version = DevuelveCampo("select cod_version_costeo from tg_estcliest where cod_cliente ='" & vCod_Cliente & "' and cod_temcli='" & varCod_TemCli & "' and cod_estcli='" & vCod_EstCli & "' and cod_estpro='" & varCod_EstPro & "'", cCONNECT)
    'vNum_Iteracion = DevuelveCampo("select num_iteracion from tg_estcliest where cod_cliente ='" & vCod_Cliente & "' and cod_temcli='" & varCod_TemCli & "' and cod_estcli='" & vCod_EstCli & "' and cod_estpro='" & varCod_EstPro & "'", cCONNECT)
    
    Set oo = CreateObject("excel.application")
    oo.workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.run "Reporte", vCod_Cliente, vcod_TemCli, varNumCot, cCONNECT, vemp, varDes_Cliente, varNom_TemCli, varObs, vusu, vCod_EstCli, vCod_estPro, GridEX1.Value(GridEX1.Columns("Cod_Version").Index), Val(GridEX1.Value(GridEX1.Columns("num_iteracion").Index))
    Set oo = Nothing
Exit Sub
Resume
hand:
    ErrorHandler Err, "GeneraReportes"
    Set oo = Nothing
End Sub

Sub Del_Iteracion()
On Error GoTo errIteracion

strSQL = "Es_Actualiza_Version_Costeo_Estilo '" & Me.vCod_Cliente & "','" & Me.vcod_TemCli & "','" & Me.vCod_EstCli & "','" & Me.vCod_estPro & "','" & GridEX1.Value(GridEX1.Columns("Cod_Version").Index) & "','D'," & Val(GridEX1.Value(GridEX1.Columns("num_iteracion").Index))
Call ExecuteSQL(cCONNECT, strSQL)

Call CARGA_GRID

Exit Sub
errIteracion:
    MsgBox Err.Description, vbCritical, "Eliminar Iteracion"
End Sub
