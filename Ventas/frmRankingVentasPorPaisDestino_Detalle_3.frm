VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmRankingVentasPorPaisDestino_Detalle_3 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14310
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   14310
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctDetalle 
      BackColor       =   &H80000010&
      Height          =   1070
      Left            =   3255
      ScaleHeight     =   1005
      ScaleWidth      =   6975
      TabIndex        =   8
      Top             =   2025
      Visible         =   0   'False
      Width           =   7035
      Begin VB.CommandButton cmdConsultar 
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
         Height          =   615
         Left            =   3720
         TabIndex        =   11
         Top             =   200
         Width           =   1935
      End
      Begin VB.OptionButton optVerDetalle 
         BackColor       =   &H80000010&
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
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5760
         TabIndex        =   9
         Top             =   200
         Width           =   1095
      End
      Begin VB.OptionButton optVerDetalle 
         BackColor       =   &H80000010&
         Caption         =   "Ver por registro seleccionado"
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
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   975
         Left            =   15
         Top             =   15
         Width           =   6950
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2220
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   3135
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
         Left            =   240
         TabIndex        =   3
         Top             =   1530
         Width           =   2655
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
         Left            =   240
         TabIndex        =   2
         Top             =   885
         Width           =   2655
      End
      Begin VB.CommandButton cmdVerDetalle 
         Caption         =   "Ver Detalle"
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
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
   End
   Begin GridEX20.GridEX grxListado 
      Height          =   6060
      Left            =   3360
      TabIndex        =   13
      Top             =   0
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   10689
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
      HeaderFontName  =   "MS Sans Serif"
      FontName        =   "MS Sans Serif"
      BackColorBkg    =   15531775
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmRankingVentasPorPaisDestino_Detalle_3.frx":0000
      Column(2)       =   "frmRankingVentasPorPaisDestino_Detalle_3.frx":00C8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmRankingVentasPorPaisDestino_Detalle_3.frx":016C
      FormatStyle(2)  =   "frmRankingVentasPorPaisDestino_Detalle_3.frx":02A4
      FormatStyle(3)  =   "frmRankingVentasPorPaisDestino_Detalle_3.frx":0354
      FormatStyle(4)  =   "frmRankingVentasPorPaisDestino_Detalle_3.frx":0408
      FormatStyle(5)  =   "frmRankingVentasPorPaisDestino_Detalle_3.frx":04E0
      FormatStyle(6)  =   "frmRankingVentasPorPaisDestino_Detalle_3.frx":0598
      FormatStyle(7)  =   "frmRankingVentasPorPaisDestino_Detalle_3.frx":0678
      FormatStyle(8)  =   "frmRankingVentasPorPaisDestino_Detalle_3.frx":0724
      ImageCount      =   0
      PrinterProperties=   "frmRankingVentasPorPaisDestino_Detalle_3.frx":07D4
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pais"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   345
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
      TabIndex        =   16
      Top             =   1545
      Width           =   1560
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3200
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Anexo"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   540
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
      TabIndex        =   14
      Top             =   2280
      Width           =   1200
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3200
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   2700
      Y1              =   360
      Y2              =   360
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
      TabIndex        =   7
      Top             =   120
      Width           =   1380
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
   Begin VB.Line Line4 
      X1              =   120
      X2              =   2700
      Y1              =   720
      Y2              =   720
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
      TabIndex        =   5
      Top             =   480
      Width           =   1380
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Fin"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   240
   End
End
Attribute VB_Name = "frmRankingVentasPorPaisDestino_Detalle_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'--+------------------------------------------+--
'==> VARIABLES PARA EL SP DE LA CONSULTA
'--+------------------------------------------+--
Public FECHA_INICIO As String
Public FECHA_FIN As String

Public opcion As String
Public COD_PAIS As String
Public DES_PAIS As String
Public TIPO_ANEXO As String
Public COD_ANEXO As String
Public DES_ANEXO As String
'--+------------------------------------------+--

Private strSQL As String


Private Sub cmdImprimir_Click()
    Call Imprimir
End Sub

Private Sub Form_Load()
    lblFecINI = "[" & FECHA_INICIO & "]"
    lblFecFIN = "[" & FECHA_FIN & "]"
    lblPAIS = "[" & COD_PAIS & "] " & DES_PAIS
End Sub

Private Sub cmdVerDetalle_Click()
    Call VER_DETALLE
End Sub



Private Sub Command2_Click()
    Unload Me
End Sub


'****************************************************************************************************************************************************************************************************************
'==> PROCEDIMIENTOS LOCALES DE USUARIOS
'****************************************************************************************************************************************************************************************************************
Private Sub VER_DETALLE()
On Error GoTo SALTO_ERROR
    
    Dim oRs As New Recordset
    
    TIPO_ANEXO = Left(Trim(grxListado.Value(grxListado.Columns("ANEXO").Index)), 1)
    COD_ANEXO = Right(Trim(grxListado.Value(grxListado.Columns("ANEXO").Index)), 4)
    strSQL = "EXECUTE CN_VENTAS_RANKING_PAIS_DESTINO_EXPORTACION '" & FECHA_INICIO & "', '" & _
                                                                      FECHA_FIN & "', '3', '" & _
                                                                      COD_PAIS & "', '" & _
                                                                      TIPO_ANEXO & "', '" & _
                                                                      COD_ANEXO & "', ''" '"', '" & _
                                                                      'COD_CLIENTE_COMERCIAL & "'"
    Screen.MousePointer = vbHourglass
    Set oRs = CargarRecordSetDesconectado(strSQL, cCONNECT)
    Screen.MousePointer = vbCustom
    If oRs.RecordCount > 0 Then
        With frmRankingVentasPorPaisDestino_Detalle_3
            .FECHA_INICIO = FECHA_INICIO
            .FECHA_FIN = FECHA_FIN
            .opcion = opcion
            .COD_PAIS = COD_PAIS
            .DES_PAIS = DES_PAIS
            .TIPO_ANEXO = TIPO_ANEXO
            .COD_ANEXO = COD_ANEXO
            .DES_ANEXO = Trim(grxListado.Value(grxListado.Columns("DES_ANEXO").Index))
            With .grxListado
                Set .ADORecordset = oRs
                .Columns("ANEXO").Caption = "Anexo"
                .Columns("ANEXO").Width = 600
                .Columns("DES_ANEXO").Caption = "Descripción del Anexo"
                .Columns("DES_ANEXO").Width = 2500
                .Columns("COD_CLIENTE").Width = 0
                .Columns("NOM_CLIENTE").Caption = "Cliente"
                .Columns("NOM_CLIENTE").Width = 2500
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
        Dim sRutaLogo As String, Ruta As String, sTitulo As String
        
        sTitulo = CStr(FECHA_INICIO) & "-" & CStr(FECHA_FIN)
        
        If MsgBox("Imprimir usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir x Estilo") = vbYes Then
            Set oo = CreateObject("excel.application")
            oo.Workbooks.Open vRuta & "\RankingVentasPorPaisDestino_02.XLT"
            oo.DisplayAlerts = False
            oo.Visible = True
            
            strSQL = "SELECT Ruta_Logo From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
            sRutaLogo = DevuelveCampo(strSQL, cCONNECT)
            sRutaLogo = CStr(IIf(IsNull(sRutaLogo), "", sRutaLogo))
            
            oo.Run "REPORTE", CStr(sRutaLogo), grxListado.ADORecordset, sTitulo, "[" & COD_PAIS & "] " & DES_PAIS
        Else
            Ruta = vRuta & "\RankingVentasPorPaisDestino_02.OTS"
            Set oo = CreateObject("ooBusiness.Calc")
            oo.OfficeTemplateSheet = Ruta
            oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
            oo.MacroLibraryName = "Library1"
            oo.MacroModuleName = "Module1"
            oo.MacroName = "Reporte"
            
            strSQL = "SELECT Des_Empresa From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
            sRutaLogo = DevuelveCampo(strSQL, cCONNECT)
            sRutaLogo = CStr(IIf(IsNull(sRutaLogo), "", sRutaLogo))
            
            oo.Run CStr(sRutaLogo), grxListado.ADORecordset.Source, sTitulo, "[" & COD_PAIS & "] " & DES_PAIS, cCONNECT
        End If
        Set oo = Nothing
   Else
        MsgBox "No se han encontrado datos para imprirmir....", vbInformation
   End If
   Exit Sub
   
ERROR:
    ErrorHandler err, "[PLANEAMIENTO] : Ranking de Ventas por Pais-Destino"
End Sub



