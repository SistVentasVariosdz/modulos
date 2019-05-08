VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPenalidadesVentas 
   Caption         =   "Penalidades de Venta"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4440
   LinkTopic       =   "frmConfirmacionDespacho"
   ScaleHeight     =   1695
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   413
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin MSComCtl2.DTPicker dtpAnoMes 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM yyyy"
         Format          =   16777219
         CurrentDate     =   37887
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes - Año: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   990
      TabIndex        =   3
      Top             =   1080
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmPenalidadesVentas.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   240
      Top             =   1080
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmPenalidadesVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSQL As String

Private Sub Form_Load()
dtpAnoMes.Value = Date
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "IMPRIMIR"
         
                    Call Reporte
            
        Case "SALIR"
            Unload Me
    End Select
End Sub


Public Sub Reporte()
On Error GoTo ErrorImpresion
Dim oo As Object
Dim sRutaLogo As String, Ruta As String
    
    VB.Screen.MousePointer = vbHourglass
    
    strSQL = "CN_VENTAS_OBTIENE_PENALIDADES_MENSUALES  '" & Format(dtpAnoMes.Value, "YYYY") & "','" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2) & "'"
    
    If MsgBox("Desea imprimir usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir") = vbYes Then
        Set oo = CreateObject("excel.application")
        oo.Workbooks.Open vRuta & "\RptPenalidadesVentas.XLT"
        
        sRutaLogo = "SELECT Ruta_Logo = ISNULL(Ruta_Logo, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
        sRutaLogo = DevuelveCampo(sRutaLogo, cCONNECT)
        
        oo.Visible = True
        oo.Run "REPORTE", strSQL, Format(dtpAnoMes.Value, "YYYY") & "-" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2), cCONNECT, sRutaLogo
    Else
        Ruta = vRuta & "\RptPenalidadesVentas.OTS"
        Set oo = CreateObject("ooBusiness.Calc")
        oo.OfficeTemplateSheet = Ruta
        oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
        oo.MacroLibraryName = "Library1"
        oo.MacroModuleName = "Module1"
        oo.MacroName = "Reporte"
        
        sRutaLogo = "SELECT Des_Empresa = ISNULL(Des_Empresa, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
        sRutaLogo = DevuelveCampo(sRutaLogo, cCONNECT)
        
        oo.Run strSQL, Format(dtpAnoMes.Value, "YYYY") & "-" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2), cCONNECT, sRutaLogo
        
    End If
    Screen.MousePointer = vbNormal
    Set oo = Nothing
Exit Sub
ErrorImpresion:
    Screen.MousePointer = vbNormal
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte" & err.Description, vbCritical, "Impresion"
End Sub



