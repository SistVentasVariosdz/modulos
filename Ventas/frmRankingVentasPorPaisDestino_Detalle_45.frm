VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmRankingVentasPorPaisDestino_Detalle_45 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   15270
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optVerDetalle 
      Caption         =   "Ver Todo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2580
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   3045
      Begin VB.CommandButton cmdDocumentos 
         Caption         =   "Ver Documento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CommandButton cmdVerDetalle 
         Caption         =   "Consultar Detalle"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   765
         Width           =   2775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   1890
         Width           =   2775
      End
   End
   Begin GridEX20.GridEX grxListado 
      Height          =   7980
      Left            =   3360
      TabIndex        =   0
      Top             =   0
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   14076
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
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      BackColorBkg    =   15531775
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmRankingVentasPorPaisDestino_Detalle_45.frx":0000
      Column(2)       =   "frmRankingVentasPorPaisDestino_Detalle_45.frx":00C8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmRankingVentasPorPaisDestino_Detalle_45.frx":016C
      FormatStyle(2)  =   "frmRankingVentasPorPaisDestino_Detalle_45.frx":02A4
      FormatStyle(3)  =   "frmRankingVentasPorPaisDestino_Detalle_45.frx":0354
      FormatStyle(4)  =   "frmRankingVentasPorPaisDestino_Detalle_45.frx":0408
      FormatStyle(5)  =   "frmRankingVentasPorPaisDestino_Detalle_45.frx":04E0
      FormatStyle(6)  =   "frmRankingVentasPorPaisDestino_Detalle_45.frx":0598
      FormatStyle(7)  =   "frmRankingVentasPorPaisDestino_Detalle_45.frx":0678
      FormatStyle(8)  =   "frmRankingVentasPorPaisDestino_Detalle_45.frx":0724
      ImageCount      =   0
      PrinterProperties=   "frmRankingVentasPorPaisDestino_Detalle_45.frx":07D4
   End
   Begin VB.OptionButton optVerDetalle 
      Caption         =   "Por Registro seleccionado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   3200
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3200
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label lblANEXO 
      AutoSize        =   -1  'True
      Caption         =   "[DESCRIPCION]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Anexo"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   540
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3200
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label lblPAIS 
      AutoSize        =   -1  'True
      Caption         =   "[NOMBRE DEL PAIS]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   1545
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pais"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Fin"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Cli. Comercial"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   945
   End
   Begin VB.Label lblCLIENTE 
      AutoSize        =   -1  'True
      Caption         =   "[COD CLI COMER]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblFecFIN 
      AutoSize        =   -1  'True
      Caption         =   "[01/01/2008]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   1320
      TabIndex        =   7
      Top             =   480
      Width           =   1380
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   2700
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Inicio"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   465
   End
   Begin VB.Label lblFecINI 
      AutoSize        =   -1  'True
      Caption         =   "[01/01/2008]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   1380
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   2700
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "frmRankingVentasPorPaisDestino_Detalle_45"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DETALLE As Recordset

'--+------------------------------------------+--
'==> VARIABLES PARA EL SP DE LA CONSULTA
'--+------------------------------------------+--
Public PADRE  As Object

Public FECHA_INICIO As String
Public FECHA_FIN As String


Public opcion As String
Public COD_PAIS As String
Public DES_PAIS As String
Public TIPO_ANEXO As String
Public COD_ANEXO As String
Public DES_ANEXO As String
Public COD_CLIENTE_COMERCIAL As String
Public DES_CLIENTE_COMERCIAL As String
'--+------------------------------------------+--

Private strSQL As String
Private blSW As Boolean

Private Sub cmdImprimir_Click()
    blSW = True
    Call Imprimir
End Sub

Private Sub cmdVerDetalle_Click()
    Call VER_DETALLE
End Sub

Private Sub cmdDocumentos_Click()
     With frmMuestraDetalleDocumVentas
        .Caption = "Documento Nª" & grxListado.Value(grxListado.Columns("NUM_CORRE").Index)
        .strSQL = "Ventas_Muestra_Detalle_Factura_Items '" & grxListado.Value(grxListado.Columns("NUM_CORRE").Index) & "'"
        .Num_Corre = grxListado.Value(grxListado.Columns("NUM_CORRE").Index)
        .cmdImprimir.Visible = True
        .sDOCUMENTO = grxListado.Value(grxListado.Columns("DOCUMENTO").Index)
        .FunctButt1.Visible = False
        .BUSCAR
        .Show 1
      End With
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblFecINI = "[" & FECHA_FIN & "]"
    lblFecFIN = "[" & FECHA_FIN & "]"
    lblPAIS = "[" & COD_PAIS & "] " & DES_PAIS
    lblANEXO = "[" & TIPO_ANEXO & COD_ANEXO & "] " & DES_ANEXO
    lblCLIENTE = "[" & COD_CLIENTE_COMERCIAL & "] " & DES_CLIENTE_COMERCIAL
    cmdImprimir.Enabled = True
    blSW = False
End Sub


'****************************************************************************************************************************************************************************************************************
'==> PROCEDIMIENTOS LOCALES DE USUARIOS
'****************************************************************************************************************************************************************************************************************
Private Sub VER_DETALLE()
On Error GoTo SALTO_ERROR
           
    If optVerDetalle(0).Value = True Then opcion = "6"
    If optVerDetalle(1).Value = True Then
        Select Case Len(Trim(COD_CLIENTE_COMERCIAL))
            Case 0: opcion = "4"
            Case Else: opcion = "5"
        End Select
    End If
    
    strSQL = "EXECUTE CN_VENTAS_RANKING_PAIS_DESTINO_EXPORTACION '" & FECHA_INICIO & "', '" & _
                                                                      FECHA_FIN & "', '" & _
                                                                      opcion & "', '" & _
                                                                      COD_PAIS & "', '" & _
                                                                      TIPO_ANEXO & "', '" & _
                                                                      COD_ANEXO & "', '" & _
                                                                      COD_CLIENTE_COMERCIAL & "'"
    Dim oRs As New Recordset
    
    Screen.MousePointer = vbHourglass
    Set oRs = CargarRecordSetDesconectado(strSQL, cCONNECT)
    Screen.MousePointer = vbCustom
    With grxListado
        Set .ADORecordset = oRs
        .Columns("DOCUMENTO").Caption = "Documento"
        .Columns("DOCUMENTO").Width = 1500
        .Columns("COD_CLIENTE").Width = 0
        .Columns("NOM_CLIENTE").Caption = "Cliente"
        .Columns("NOM_CLIENTE").Width = 2400
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
    cmdImprimir.Enabled = False
    If oRs.RecordCount = 0 Then MsgBox "No se han encontrados en la consulta de detalle....", vbInformation _
    Else cmdImprimir.Enabled = True
    Exit Sub
    
SALTO_ERROR:
    Screen.MousePointer = vbCustom
    MsgBox err.Description, vbCritical, "[VENTAS] : Ranking por Pais-Destino"
End Sub


Private Sub Imprimir()
On Error GoTo ERROR
    
    If grxListado.ADORecordset.RecordCount > 0 Then
        Dim oo As Object, vRutaLogo As Variant
        Dim sRutaLogo As String, sTitulo As String, Ruta As String
        
        strSQL = "SELECT Ruta_Logo From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
        sRutaLogo = DevuelveCampo(strSQL, cCONNECT)
        sRutaLogo = CStr(IIf(IsNull(sRutaLogo), "", sRutaLogo))
        
        sTitulo = CStr(FECHA_INICIO) & "-" & CStr(FECHA_FIN)
        
        If MsgBox("Imprimir usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir") = vbYes Then
            Set oo = CreateObject("excel.application")
            If opcion = "4" Or opcion = "5" Then oo.Workbooks.Open vRuta & "\RankingVentasPorPaisDestino_0405.XLT"
            If opcion = "6" Then oo.Workbooks.Open vRuta & "\RankingVentasPorPaisDestino_06.XLT"
            oo.DisplayAlerts = False
            oo.Visible = True
            
            If opcion = "4" Or opcion = "5" Then
                oo.Run "REPORTE", CStr(sRutaLogo), grxListado.ADORecordset, sTitulo, lblPAIS, lblANEXO, lblCLIENTE
            End If
            If opcion = "6" Then
                oo.Run "REPORTE", CStr(sRutaLogo), grxListado.ADORecordset, sTitulo, lblPAIS, lblANEXO
            End If
        Else
            If opcion = "4" Or opcion = "5" Then Ruta = vRuta & "\RankingVentasPorPaisDestino_0405.OTS"
            If opcion = "6" Then Ruta = vRuta & "\RankingVentasPorPaisDestino_06.OTS"
            Set oo = CreateObject("ooBusiness.Calc")
            oo.OfficeTemplateSheet = Ruta
            oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
            oo.MacroLibraryName = "Library1"
            oo.MacroModuleName = "Module1"
            oo.MacroName = "REPORTE"
            
            strSQL = "SELECT Des_Empresa From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
            sRutaLogo = DevuelveCampo(strSQL, cCONNECT)
            sRutaLogo = CStr(IIf(IsNull(sRutaLogo), "", sRutaLogo))
            
            If opcion = "4" Or opcion = "5" Then
                oo.Run CStr(sRutaLogo), grxListado.ADORecordset.Source, sTitulo, lblPAIS, lblANEXO, lblCLIENTE, cCONNECT
            End If
            If opcion = "6" Then
                oo.Run CStr(sRutaLogo), grxListado.ADORecordset.Source, sTitulo, lblPAIS, lblANEXO, cCONNECT
            End If
        End If
        Set oo = Nothing
   Else
        MsgBox "No se han encontrado datos para imprirmir....", vbInformation
   End If
   Exit Sub
   
ERROR:
    ErrorHandler err, "[VENTAS] : Ranking de Ventas por Pais-Destino"
End Sub

Private Sub optVerDetalle_Click(Index As Integer)
    If blSW = False Then Exit Sub
    cmdImprimir.Enabled = False
End Sub




