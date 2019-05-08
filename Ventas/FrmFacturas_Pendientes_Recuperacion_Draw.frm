VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Begin VB.Form FrmFacturas_Pendientes_Recuperacion_Draw 
   Caption         =   "Imprimir"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   4365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Ordenar por"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   4095
      Begin VB.OptionButton OptFactura 
         Caption         =   "Número de Factura"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton OptFec_Embarque 
         Caption         =   "Fecha de Embarque"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtrar por"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4095
      Begin VB.OptionButton OptDraw 
         Caption         =   "Sin Expediente DrawBack"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1080
         Width           =   2295
      End
      Begin VB.OptionButton OptDua 
         Caption         =   "Sin F/Numeración DUA"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.OptionButton OptTodas 
         Caption         =   "Todas"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
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
      Caption         =   "Facturas Pendientes Recuperación DrawBack"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "FrmFacturas_Pendientes_Recuperacion_Draw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdImprimir_Click()
    Reporte
End Sub

Private Sub Reporte()
   On Error GoTo SALTO_ERROR
    Dim oRs As New Recordset
    Dim strSQL As String
    Dim opcion As String
    Dim orden As String
    Dim Filtro As String
    Dim Nom_Ordenado As String
    
    If OptTodas.Value Then
        opcion = "1"
        Filtro = OptTodas.Caption
    ElseIf OptDua.Value Then
        opcion = "2"
        Filtro = OptDua.Caption
    Else
        opcion = "3"
        Filtro = OptDraw.Caption
    End If
    
    If OptFec_Embarque.Value Then
        orden = "1"
        Nom_Ordenado = OptFec_Embarque.Caption
    Else
        orden = "2"
        Nom_Ordenado = OptFactura.Caption
    End If
    
    strSQL = "VENTAS_MUESTRA_FACTURAS_PENDIENTES_RECUPERACION_DRAWBACK '" & opcion & "','" & orden & "'"

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
    oo.Workbooks.Open vRuta & "\rptFACTURAS_PENDIENTES_RECUPERACION_DRAWBACK.XLT"
    oo.Visible = True
    oo.displayalerts = False
    
    oo.Run "reporte", sRutaLogo, oRs, Filtro & "  Ordenado por :  " & Nom_Ordenado
    
    Set oo = Nothing
    Exit Sub

SALTO_ERROR:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub
