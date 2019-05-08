VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmRankingVentasXEstilos 
   Caption         =   "Ranking de Ventas por Estilos"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   12405
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   4740
      TabIndex        =   1
      Top             =   8040
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   900
      Custom          =   $"frmRankingVentasXEstilos.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   1
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX grxListado 
      Height          =   7920
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   12330
      _ExtentX        =   21749
      _ExtentY        =   13970
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
      GridLines       =   2
      BackColorBkg    =   15531775
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmRankingVentasXEstilos.frx":00E1
      Column(2)       =   "frmRankingVentasXEstilos.frx":01A9
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmRankingVentasXEstilos.frx":024D
      FormatStyle(2)  =   "frmRankingVentasXEstilos.frx":0385
      FormatStyle(3)  =   "frmRankingVentasXEstilos.frx":0435
      FormatStyle(4)  =   "frmRankingVentasXEstilos.frx":04E9
      FormatStyle(5)  =   "frmRankingVentasXEstilos.frx":05C1
      FormatStyle(6)  =   "frmRankingVentasXEstilos.frx":0679
      FormatStyle(7)  =   "frmRankingVentasXEstilos.frx":0759
      FormatStyle(8)  =   "frmRankingVentasXEstilos.frx":0805
      ImageCount      =   0
      PrinterProperties=   "frmRankingVentasXEstilos.frx":08B5
   End
End
Attribute VB_Name = "frmRankingVentasXEstilos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public f1 As Date
Public f2 As Date
Dim strSQL As String

Private Sub Form_Load()
    strSQL = "EXECUTE CN_VENTAS_RANKING_PAIS_DESTINO_EXPORTACION_ESTILO '" & f1 & "', '" & f2 & "', '7', '', '', '', ''"
    Screen.MousePointer = vbHourglass
    Set grxListado.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
    Screen.MousePointer = vbDefault
End Sub

Private Sub ImprimirVentas()
On Error GoTo ERROR
'If grxListado.ADORecordset.RecordCount > 0 Then
Dim oo As Object, vRutaLogo As Variant
Dim sRutaLogo As String, sTitulo As String, _
    Ruta As String
    
    strSQL = "SELECT des_empresa From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
    sRutaLogo = DevuelveCampo(strSQL, cCONNECT)
    
    sTitulo = CStr(f1) & "-" & CStr(f2)
    
    If MsgBox("Imprimir usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir x Estilo") = vbYes Then
        Set oo = CreateObject("excel.application")
        oo.Workbooks.Open vRuta & "\RPTVentasxEstilo.XLT"
        oo.DisplayAlerts = False
        oo.Visible = True
        
        oo.Run "REPORTE", CStr(f1), CStr(f2), cCONNECT, sRutaLogo
    Else
        Ruta = vRuta & "\RPTVentasxEstilo.OTS"
        Set oo = CreateObject("ooBusiness.Calc")
        oo.OfficeTemplateSheet = Ruta
        oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
        oo.MacroLibraryName = "Library1"
        oo.MacroModuleName = "Module1"
        oo.MacroName = "Reporte"
        
        oo.Run CStr(f1), CStr(f2), cCONNECT, sRutaLogo
    End If
    Set oo = Nothing
'   Else
'        MsgBox "No se han encontrado datos para imprirmir....", vbInformation
'   End If
   Exit Sub
   
ERROR:
    ErrorHandler err, "[VENTAS] : RPTVentasxEstilo"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "VERDETALLE": Call VER
        Case "IMPRIMIR": Call ImprimirVentas
        Case "SALIR": Unload Me
    End Select
End Sub

Private Sub VER()
Dim frm As New frmRankingVentasEstiloDetalle
frm.f1 = f1
frm.f2 = f2
frm.Cod_EstCli = Trim(grxListado.Value(grxListado.Columns("Cod_EstCli").Index))
frm.Cod_OrdPro = Trim(grxListado.Value(grxListado.Columns("NP").Index))
frm.Show 1
End Sub


