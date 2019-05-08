VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmCotizacionesAlternativas 
   Caption         =   "Cotizaciones Alternativas"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin GridEX20.GridEX GridEX1 
      Height          =   2475
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   4366
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmCotizacionesAlternativas.frx":0000
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   4
      Column(1)       =   "frmCotizacionesAlternativas.frx":0352
      Column(2)       =   "frmCotizacionesAlternativas.frx":041A
      Column(3)       =   "frmCotizacionesAlternativas.frx":04BE
      Column(4)       =   "frmCotizacionesAlternativas.frx":0562
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmCotizacionesAlternativas.frx":063E
      FormatStyle(2)  =   "frmCotizacionesAlternativas.frx":0776
      FormatStyle(3)  =   "frmCotizacionesAlternativas.frx":0826
      FormatStyle(4)  =   "frmCotizacionesAlternativas.frx":08DA
      FormatStyle(5)  =   "frmCotizacionesAlternativas.frx":09B2
      FormatStyle(6)  =   "frmCotizacionesAlternativas.frx":0A6A
      FormatStyle(7)  =   "frmCotizacionesAlternativas.frx":0B4A
      FormatStyle(8)  =   "frmCotizacionesAlternativas.frx":1002
      ImageCount      =   1
      ImagePicture(1) =   "frmCotizacionesAlternativas.frx":144E
      PrinterProperties=   "frmCotizacionesAlternativas.frx":17A0
   End
   Begin FunctionsButtons.FunctButt acbOkCancel 
      Height          =   480
      Left            =   4110
      TabIndex        =   1
      Top             =   2940
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   847
      Custom          =   $"frmCotizacionesAlternativas.frx":1978
      Orientacion     =   0
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1000
      ControlHeigth   =   450
      ControlSeparator=   80
   End
End
Attribute VB_Name = "frmCotizacionesAlternativas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_cliente As String
Public sCod_EstCli As String
Public vNum_cotizacion As Long
Public vCod_Estpro_Cotizacion As String
Public vCod_Version_Cotizacion As String
Public bOk As Boolean

Private Sub acbOkCancel_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            bOk = True
            vNum_cotizacion = GridEX1.value(GridEX1.Columns("Num_Solicitud_Costeo_Asignada").Index)
            vCod_Estpro_Cotizacion = GridEX1.value(GridEX1.Columns("Cod_EstPro_Asignada").Index)
            vCod_Version_Cotizacion = GridEX1.value(GridEX1.Columns("Cod_Version_Asignada").Index)
            Unload Me
        Case "VERCUADRO"
            Reporte
        Case "SALIR"
            Unload Me
    End Select
End Sub

Private Sub Reporte()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim RutaLogo As String
Dim Usu As String
Dim strSQL As String
Dim Rs As ADODB.Recordset

Dim vNumCotizacion As Long
Dim sCod_Fabrica  As String
Dim sCod_OrdPro As String

    vNumCotizacion = GridEX1.value(GridEX1.Columns("Num_Solicitud_Costeo_Asignada").Index)
    sCod_Fabrica = GridEX1.value(GridEX1.Columns("COD_FABRICA").Index)
    sCod_OrdPro = GridEX1.value(GridEX1.Columns("COD_ORDPRO").Index)
    
    strSQL = "SELECT Ruta_Logo FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA = '" & vemp & "'"
    RutaLogo = DevuelveCampo(strSQL, cCONNECT)
    
    
    Set Rs = GetDataSet(cCONNECT, "Exec SM_GENERA_MATRIZ_COTIZACION " & vNumCotizacion & ",'" & sCod_Fabrica & "','" & sCod_OrdPro & "'")
    
    Ruta = App.Path & "\cotizacion.xlt"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    
    oo.Run "reporte", Rs, RutaLogo, vNumCotizacion, cCONNECT
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler Err, "GeneraReportes"
    Set oo = Nothing
End Sub

    


Public Function Buscar() As Boolean
On Error GoTo errores
Dim sSQL As String
Dim vBookmark As Variant

sSQL = "SM_MUESTRA_COTIZACIONES_SIMILARES '$' ,'$' "
sSQL = VBsprintf(sSQL, sCod_cliente, sCod_EstCli)

vBookmark = GridEX1.Row
GridEX1.ClearFields

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)

GridEX1.Row = vBookmark


GridEX1.ContinuousScroll = True

GridEX1.FrozenColumns = 2
Exit Function

errores:
    errores Err.Number
End Function


Private Sub GridEX1_DblClick()
    Dim i As Integer
    For i = 1 To GridEX1.Columns.Count
        Debug.Print GridEX1.Name & ".Columns(" & Chr(34) & GridEX1.Columns(i).Caption & Chr(34) & ").width = " & CStr(GridEX1.Columns(i).Width)
    Next
End Sub

