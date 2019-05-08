VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRankingVentasPorPaisDestino 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ranking de Ventas por Pais Destino"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15015
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   15015
   StartUpPosition =   3  'Windows Default
   Begin GridEX20.GridEX grxListado 
      Height          =   8700
      Left            =   1905
      TabIndex        =   0
      Top             =   0
      Width           =   13110
      _ExtentX        =   23125
      _ExtentY        =   15346
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      BorderStyle     =   2
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      HeaderFontSize  =   8.25
      FontSize        =   8.25
      GridLines       =   2
      BackColorBkg    =   15531775
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmRankingVentasPorPaisDestino.frx":0000
      Column(2)       =   "frmRankingVentasPorPaisDestino.frx":00C8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmRankingVentasPorPaisDestino.frx":016C
      FormatStyle(2)  =   "frmRankingVentasPorPaisDestino.frx":02A4
      FormatStyle(3)  =   "frmRankingVentasPorPaisDestino.frx":0354
      FormatStyle(4)  =   "frmRankingVentasPorPaisDestino.frx":0408
      FormatStyle(5)  =   "frmRankingVentasPorPaisDestino.frx":04E0
      FormatStyle(6)  =   "frmRankingVentasPorPaisDestino.frx":0598
      FormatStyle(7)  =   "frmRankingVentasPorPaisDestino.frx":0678
      FormatStyle(8)  =   "frmRankingVentasPorPaisDestino.frx":0724
      ImageCount      =   0
      PrinterProperties=   "frmRankingVentasPorPaisDestino.frx":07D4
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8730
      Left            =   50
      TabIndex        =   1
      Top             =   -75
      Width           =   1815
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   39591
      End
      Begin MSComCtl2.DTPicker dtpFecInicio 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   39591
      End
      Begin FunctionsButtons.FunctButt fbOperaciones 
         Height          =   3480
         Left            =   120
         TabIndex        =   6
         Top             =   5160
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   6138
         Custom          =   $"frmRankingVentasPorPaisDestino.frx":09AC
         Orientacion     =   1
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1500
         ControlHeigth   =   650
         ControlSeparator=   50
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fin"
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   885
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   1680
      Top             =   4560
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmRankingVentasPorPaisDestino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strSQL As String

Private Sub Form_Load()
    dtpFecInicio.Value = Date
    dtpFecFin.Value = Date
End Sub

Private Sub fbOperaciones_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "BUSCAR": Call BUSCAR
        Case "VER": Call VER
        Case "IMPRIMIR": Call Imprimir
        Case "IMPVEN": Call ImprimirVentas
        Case "SALIR": Unload Me
    End Select
End Sub

'****************************************************************************************************************************************************************************************************************
'==> PROCEDIMIENTOS LOCALES DE USUARIOS
'****************************************************************************************************************************************************************************************************************
Private Sub BUSCAR()
On Error GoTo SALTO_ERROR
    
    strSQL = "EXECUTE CN_VENTAS_RANKING_PAIS_DESTINO_EXPORTACION '" & dtpFecInicio.Value & "', '" & dtpFecFin.Value & "', '1', '', '', '', ''"
    Screen.MousePointer = vbHourglass
    With grxListado
        Set .ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
        .Columns("TIPO").Width = 0
        .Columns("COD_PAIS").Width = 0
        .Columns("DES_PAIS").Caption = "Pais Destino Embarque"
        .Columns("DES_PAIS").Width = 2200
        .Columns("CANTIDAD").Caption = "Cantidad"
        .Columns("CANTIDAD").Width = 900
        .Columns("IMPORTE_SOLES").Caption = "FOB Soles [S/.]"
        .Columns("IMPORTE_SOLES").Width = 1500
        .Columns("IMPORTE_DOLARES").Caption = "FOB Dólares [US$]"
        .Columns("IMPORTE_DOLARES").Width = 1700
        .Columns("FLETE").Caption = "Flete [US$]"
        .Columns("FLETE").Width = 1100
        .Columns("DESADUANAJE").Caption = "DesAdua. [US$]"
        .Columns("DESADUANAJE").Width = 1500
        .Columns("TRANSP_PAIS_DESTINO").Caption = "Tran. Pais Dest. [US$]"
        .Columns("TRANSP_PAIS_DESTINO").Width = 2000
        .Columns("TOTALDOLARES").Caption = "Total [US$]"
        .Columns("TOTALDOLARES").Width = 1300
        .Columns("PORCENTAJE").Caption = "[%]"
        .Columns("PORCENTAJE").Width = 600
        
        
        .Columns("CANTIDAD").Format = "### ###,###"
        .Columns("IMPORTE_SOLES").Format = "### ###,###.00"
        .Columns("IMPORTE_DOLARES").Format = "### ###,###.00"
        .Columns("FLETE").Format = "### ###,###.00"
        .Columns("DESADUANAJE").Format = "### ###,###.00"
        .Columns("TRANSP_PAIS_DESTINO").Format = "### ###,###.00"
        .Columns("TOTALDOLARES").Format = "### ###,###.00"
        .Columns("PORCENTAJE").Format = "###.00"
        
    End With
    
    grxListado.SetFocus
    Screen.MousePointer = vbCustom
    Exit Sub
    
SALTO_ERROR:
    Screen.MousePointer = vbCustom
    MsgBox err.Description, vbCritical, "[VENTAS] : Ranking por Pais-Destino"
End Sub

Private Sub VER()
On Error GoTo SALTO_ERROR
    
    Dim oRs As New Recordset
    Dim sCodPais As String
        
    sCodPais = Trim(grxListado.Value(grxListado.Columns("COD_PAIS").Index))
    strSQL = "EXECUTE CN_VENTAS_RANKING_PAIS_DESTINO_EXPORTACION '" & dtpFecInicio.Value & "', '" & dtpFecFin.Value & "', '2', '" & sCodPais & "', '', '', ''"
    Screen.MousePointer = vbHourglass
    Set oRs = CargarRecordSetDesconectado(strSQL, cCONNECT)
    Screen.MousePointer = vbCustom
    If oRs.RecordCount > 0 Then
        With frmRankingVentasPorPaisDestino_Detalle_2
            .FECHA_INICIO = dtpFecInicio.Value
            .FECHA_FIN = dtpFecFin.Value
            .opcion = "2"
            .COD_PAIS = sCodPais
            .DES_PAIS = Trim(grxListado.Value(grxListado.Columns("DES_PAIS").Index))
            With .grxListado
                Set .ADORecordset = oRs
                .Columns("ANEXO").Caption = "Anexo"
                .Columns("ANEXO").Width = 700
                .Columns("DES_ANEXO").Caption = "Descripción del Anexo"
                .Columns("DES_ANEXO").Width = 2500
                .Columns("CANTIDAD").Caption = "Cantidad"
                .Columns("CANTIDAD").Width = 900
                .Columns("IMPORTE_SOLES").Caption = "FOB Soles [S/.]"
                .Columns("IMPORTE_SOLES").Width = 1500
                .Columns("IMPORTE_DOLARES").Caption = "FOB Dólares [US$]"
                .Columns("IMPORTE_DOLARES").Width = 1500
                .Columns("FLETE").Caption = "Flete [US$]"
                .Columns("FLETE").Width = 900
                .Columns("DESADUANAJE").Caption = "DesAdua. [US$]"
                .Columns("DESADUANAJE").Width = 1300
                .Columns("TRANSP_PAIS_DESTINO").Caption = "Tran. Pais Dest. [US$]"
                .Columns("TRANSP_PAIS_DESTINO").Width = 1800
                .Columns("TOTALDOLARES").Caption = "Total [US$]"
                .Columns("TOTALDOLARES").Width = 1100
               
                
                
                .Columns("CANTIDAD").Format = "### ###,###"
                .Columns("IMPORTE_SOLES").Format = "### ###,###.00"
                .Columns("IMPORTE_DOLARES").Format = "### ###,###.00"
                .Columns("FLETE").Format = "### ###,###.00"
                .Columns("DESADUANAJE").Format = "### ###,###.00"
                .Columns("TRANSP_PAIS_DESTINO").Format = "### ###,###.00"
                .Columns("TOTALDOLARES").Format = "### ###,###.00"
        
                
            End With
            .Show 1
        End With
    Else
        MsgBox "No se han encontrados en la consulta de detalle....", vbInformation
    End If
    Exit Sub
    
SALTO_ERROR:
    Screen.MousePointer = vbCustom
    MsgBox err.Description, vbCritical, "[VENTAS] : Ranking por Pais-Destino"
End Sub

Private Sub Imprimir()
On Error GoTo ERROR
    
    If grxListado.ADORecordset.RecordCount > 0 Then
        Dim oo As Object, vRutaLogo As Variant
        Dim sRutaLogo As String
        Dim sTitulo As String, Ruta As String
        
        strSQL = "SELECT Ruta_Logo From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
        sTitulo = CStr(dtpFecInicio.Value) & "-" & CStr(dtpFecFin.Value)
        
        If MsgBox("Imprimir usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir") = vbYes Then
            Set oo = CreateObject("excel.application")
            oo.Workbooks.Open vRuta & "\RankingVentasPorPaisDestino_01.XLT"
            oo.DisplayAlerts = False
            oo.Visible = True
            
            sRutaLogo = DevuelveCampo(strSQL, cCONNECT)
            sRutaLogo = CStr(IIf(IsNull(sRutaLogo), "", sRutaLogo))
            
            oo.Run "REPORTE", CStr(sRutaLogo), grxListado.ADORecordset, sTitulo
        Else
            Ruta = vRuta & "\RankingVentasPorPaisDestino_01.OTS"
            Set oo = CreateObject("ooBusiness.Calc")
            oo.OfficeTemplateSheet = Ruta
            oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
            oo.MacroLibraryName = "Library1"
            oo.MacroModuleName = "Module1"
            oo.MacroName = "Reporte"
            
            strSQL = "SELECT Des_Empresa From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
            sRutaLogo = DevuelveCampo(strSQL, cCONNECT)
            sRutaLogo = CStr(IIf(IsNull(sRutaLogo), "", sRutaLogo))
            
            oo.Run CStr(sRutaLogo), grxListado.ADORecordset.Source, sTitulo, cCONNECT
        End If
        Set oo = Nothing
    Else
         MsgBox "No se han encontrado datos para imprirmir....", vbInformation
    End If
    
Exit Sub
ERROR:
    ErrorHandler err, "[VENTAS] : Ranking de Ventas por Pais-Destino"
End Sub

Private Sub ImprimirVentas()
Dim frm As New frmRankingVentasXEstilos
frm.f1 = dtpFecInicio.Value
frm.f2 = dtpFecFin.Value
frm.Show 1
'   On Error GoTo ERROR
'
'   'If grxListado.ADORecordset.RecordCount > 0 Then
'        Dim oo As Object, vRutaLogo As Variant
'        Dim sRutaLogo As String
'        Dim sTitulo As String
'
'        strSql = "SELECT des_empresa From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
'        sRutaLogo = DevuelveCampo(strSql, cCONNECT)
'
'        sTitulo = CStr(dtpFecInicio.Value) & "-" & CStr(dtpFecFin.Value)
'        Set oo = CreateObject("excel.application")
'        oo.Workbooks.Open vRuta & "\RPTVentasxEstilo.XLT"
'        oo.DisplayAlerts = False
'        oo.Visible = True
'
'        oo.Run "REPORTE", dtpFecInicio.Value, dtpFecFin.Value, cCONNECT, sRutaLogo
'        Set oo = Nothing
''   Else
''        MsgBox "No se han encontrado datos para imprirmir....", vbInformation
''   End If
'   Exit Sub
'
'ERROR:
'    ErrorHandler err, "[VENTAS] : RPTVentasxEstilo"
End Sub


