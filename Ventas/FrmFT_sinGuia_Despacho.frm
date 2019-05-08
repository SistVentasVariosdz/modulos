VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form FrmFT_sinGuia_Despacho 
   Caption         =   "Factura Manufactura sin Guia de Despacho"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar Opción a Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3615
      Begin VB.OptionButton OptFT 
         Caption         =   "Por Facturar"
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton OptMov 
         Caption         =   "Por Movimiento "
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   480
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   720
      TabIndex        =   1
      Top             =   2640
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmFT_sinGuia_Despacho.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   0
      Top             =   0
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Factura Manufactura sin Guia de Despacho"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "FrmFT_sinGuia_Despacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "IMPRIMIR"
            Reporte
        Case "SALIR"
            Unload Me
    End Select
End Sub


Private Sub Reporte()
   On Error GoTo SALTO_ERROR
    Dim oRs As New Recordset
    Dim strSQL As String
    Dim opcion As String
    
    If OptMov.Value = True Then
        strSQL = "ventas_muestra_movimientos_Por_facturar_despachos_apt '1'"
        opcion = OptMov.Caption
    Else
        strSQL = "ventas_muestra_movimientos_Por_facturar_despachos_apt '2'"
        opcion = OptFT.Caption
    End If
    
    
    
    Set oRs = CargarRecordSetDesconectado(strSQL, cCONNECT)
    If oRs.RecordCount = 0 Then
        MsgBox "No se han encontrado datos para la impresión.....", vbExclamation
        Exit Sub
    End If
    
    Dim oo As Object
    Dim sRutaLogo As String, sTitulo As String
    
    Set oo = CreateObject("excel.application")
    strSQL = "SELECT Ruta_Logo = ISNULL(Ruta_Logo, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
    sRutaLogo = DevuelveCampo(strSQL, cCONNECT)
    
    If OptMov.Value = True Then
        oo.Workbooks.Open vRuta & "\rptFT_sinGuia_Despacho.XLT"
    Else
        oo.Workbooks.Open vRuta & "\rptFT_sinGuia_Despacho_Por_Facturar.XLT"
    End If
    
    oo.Visible = True
    oo.displayalerts = False
    
    oo.Run "reporte", sRutaLogo, oRs, opcion
    
    Set oo = Nothing
    Exit Sub

SALTO_ERROR:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
