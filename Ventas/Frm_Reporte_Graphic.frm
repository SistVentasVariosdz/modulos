VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{6D42EF51-35C6-4857-8449-90A05BE76B89}#1.5#0"; "csXGraphTrial.ocx"
Begin VB.Form Frm_Reporte_Graphic 
   Caption         =   "Posicion  De La Cartera Morosa"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   5.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin csXGraphTrial.Draw Draw1 
      Height          =   4695
      Left            =   0
      TabIndex        =   10
      Top             =   960
      Width           =   7935
      GraphType       =   0
      CenterX         =   210
      CenterY         =   150
      PieDia          =   200
      Offset          =   10
      BGColor         =   16777215
      LegendX         =   390
      LegendY         =   10
      Square          =   8
      Padding         =   5
      GraphPen        =   0
      ShowLabel       =   -1  'True
      ShowPercent     =   0   'False
      ShowNumbers     =   -1  'True
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowLegend      =   -1  'True
      ShowLegendBox   =   -1  'True
      BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleX          =   0
      TitleY          =   0
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title           =   ""
      BeginProperty AxisTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XAxisText       =   ""
      YAxisText       =   ""
      Decimals        =   2
      OriginX         =   50
      OriginY         =   250
      MaxX            =   250
      MaxY            =   200
      BarWidth        =   0
      BarGap          =   0
      YGrad           =   0
      YTop            =   0
      RNDColor        =   0
      LabelVertical   =   0   'False
      ShowGrid        =   0   'False
      GridStyle       =   1
      GridColor       =   0
      ShowBarTotal    =   0   'False
      TransColor      =   16777215
      XTop            =   0
      XGrad           =   0
      XValuesVertical =   -1  'True
      LineWidth       =   1
      PointStyle      =   0
      PointSize       =   2
      XOffset         =   0
      YOffset         =   0
      UseXAxisLabels  =   0   'False
      UseYAxisLabels  =   0   'False
      XMarkSize       =   4
      YMarkSize       =   4
      XAxisNegative   =   0
      YAxisNegative   =   0
      LabelBGColor    =   16777215
      TitleBGColor    =   16777215
      LegendBGColor   =   16777215
      AxisTextBGColor =   16777215
      Prefix          =   ""
      Suffix          =   ""
      PlotAreaColor   =   16777215
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowLine        =   -1  'True
      Transparent     =   0   'False
      ShowSeparator   =   -1  'True
      JpegQuality     =   100
      DoubleBuffered  =   0   'False
      Enabled         =   -1  'True
      Object.Visible         =   -1  'True
      Cursor          =   0
      HelpType        =   0
      HelpKeyword     =   ""
      Object.Height          =   313
      Object.Width           =   529
      UseRNDColor     =   0   'False
      TextBGColor     =   16777215
      LabelTransparent=   0   'False
      TitleTransparent=   0   'False
      AxisTextTransparent=   0   'False
      TextTransparent =   -1  'True
      ShowTotalIfZero =   -1  'True
      LabelColor      =   0
      LegendColor     =   255
      TitleColor      =   0
      AxisTextColor   =   0
      TextColor       =   0
      BeginProperty LineGraphTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LineGraphTextColor=   0
      LineGraphTextBGColor=   16777215
      LineGraphTextTransparent=   -1  'True
      LineGraphTextBorder=   0   'False
      LineGraphTextLeader=   0   'False
      LineGraphBorderColor=   0
      LineGraphLeaderColor=   0
      LineGraphTextAlign=   0
      LineGraphTextX  =   0
      LineGraphTextY  =   0
      ShowPlotBorder  =   0   'False
      TitleTextAlign  =   0
      HideHGrid       =   0   'False
      HideVGrid       =   0   'False
      UseXAxisDates   =   0   'False
      UseYAxisDates   =   0   'False
      DateTimeFormat  =   1
      UseLZW          =   -1  'True
      ShowStackedValue=   0   'False
      BeginProperty StackedTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StackedTextBGColor=   16777215
      StackedTextTransparent=   -1  'True
      StackedTextAlign=   1
      StackedTextColor=   0
      LegendVertical  =   -1  'True
      LegendAlign     =   0
      DateFormatString=   ""
      TimeFormatString=   ""
      PrefixX         =   ""
      PrefixY         =   ""
      SuffixX         =   ""
      SuffixY         =   ""
      ShowSeparatorX  =   0   'False
      ShowSeparatorY  =   0   'False
      LegendHideEmptyNames=   0   'False
      LegendInvert    =   0   'False
      StartAngle      =   0
      ShowTrendLine   =   0   'False
      TrendLineColor  =   0
      TrendLineName   =   ""
      TrendLineWidth  =   1
      BarTotalVertical=   0   'False
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Deuda Por Fecha Venc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   9
      Top             =   1680
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Cartera Morosa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   7
      Top             =   960
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9975
      Begin VB.TextBox Txt_Porcentaje 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4560
         TabIndex        =   0
         Text            =   "5"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Opt_Exportacion 
         Caption         =   "Cliente Exportacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton opt_local 
         Caption         =   "Cliente Local"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "% Minimo Para Agrupar Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   2235
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   1110
      Left            =   8040
      TabIndex        =   2
      Top             =   4560
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1958
      Custom          =   $"Frm_Reporte_Graphic.frx":0000
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   2895
      Left            =   0
      TabIndex        =   1
      Top             =   5760
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   5106
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      GroupByBoxVisible=   0   'False
      HeaderFontName  =   "MS Sans Serif"
      HeaderFontSize  =   8.25
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      ColumnHeaderHeight=   285
      IntProp1        =   0
      ColumnsCount    =   2
      Column(1)       =   "Frm_Reporte_Graphic.frx":0090
      Column(2)       =   "Frm_Reporte_Graphic.frx":0158
      FormatStylesCount=   6
      FormatStyle(1)  =   "Frm_Reporte_Graphic.frx":01FC
      FormatStyle(2)  =   "Frm_Reporte_Graphic.frx":0334
      FormatStyle(3)  =   "Frm_Reporte_Graphic.frx":03E4
      FormatStyle(4)  =   "Frm_Reporte_Graphic.frx":0498
      FormatStyle(5)  =   "Frm_Reporte_Graphic.frx":0570
      FormatStyle(6)  =   "Frm_Reporte_Graphic.frx":0628
      ImageCount      =   0
      PrinterProperties=   "Frm_Reporte_Graphic.frx":0708
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Deuda Por Tipo Doc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   8
      Top             =   1320
      Width           =   1935
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   10440
      Top             =   3840
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "Frm_Reporte_Graphic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public strSQL As String
Dim Tip_cliente As String
Public mRs As Object
Public nEsperado As Double
Public nLogrado As Double
Dim Colours

Private Sub Command1_Click()

mostrar
DrawBarChart
End Sub

Private Sub Form_Load()
Tip_cliente = "N"
Set mRs = CreateObject("ADODB.Recordset")

End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
    Reporte
Case "SALIR"
    Unload Me

End Select
End Sub

Public Sub mostrar()
On Error GoTo Fin


    

'strSQL = "CN_VENTAS_MUESTRA_DEUDA_POR_CLIENTE  '" & Txt_Porcentaje & "','" & Tip_cliente & "'"

    

Set gridex1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

If gridex1.RowCount = 0 Then Exit Sub

gridex1.Columns("Descripcion").Width = 2500
    gridex1.Columns("Descripcion").Caption = "Descripcion"
    gridex1.Columns("Descripcion").HeaderAlignment = jgexAlignCenter
   
Exit Sub
Fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub

Private Sub Parametros()
    Dim matrizValores() As Variant
    Dim I As Integer
    Dim sTipo As String
    
    
        Set mRs = GetRecordset1(cCONNECT, "CN_VENTAS_MUESTRA_DEUDA_POR_CLIENTE  '" & Txt_Porcentaje & "','" & Tip_cliente & "'")
        
        
        If mRs.RecordCount > 0 Then
            ' Cambiar las dimensiones de la matriz
            ReDim matrizValores(1 To mRs.RecordCount, 0 To 2)
            

            
            ' Ir al inicio de los datos
            mRs.MoveFirst
            I = 1
            Do While Not mRs.EOF
            'Do While Not 5
               ' cargar la matriz
               matrizValores(I, 1) = Left(mRs!Cliente, 10)
               matrizValores(I, 2) = mRs!Saldo_Final
               
               mRs.MoveNext
               I = I + 1
            Loop
            ' asignar la matriz al chart
            MSChart1.ChartData = matrizValores
            
            ' asignar el nombre de la serie
            MSChart1.Plot.SeriesCollection(1).LegendText = COBRANZAS
            ' agregar el titulo
            MSChart1.TitleText = "Deuda Por Clientes"
            ' nombre del eje x
            MSChart1.Plot.Axis(VtChAxisIdX, 0).AxisTitle.Text = "Clientes"
        Else
            ' si no hay datos asignar un arreglo vacio
            MSChart1.ChartData = Array(0, 0)
        End If
    

End Sub




Public Function GetRecordset1(ByVal Connect As String, ByVal SQL As String) As Object 'ADOR.Recordset
  On Error GoTo ehGetRecordset
  Dim objADORs As Object ' CreateObject("ADODB.Recordset") '
  Dim objAdoCn As Object ' New ADODB.Connection '
  
 
  Set objADORs = CreateObject("ADODB.Recordset") 'CreateObject("ADODB.Recordset") '
  Set objAdoCn = CreateObject("ADODB.Connection") ' New ADODB.Connection  '
  objAdoCn.CursorLocation = 3
  objAdoCn.Open Connect
  objAdoCn.CommandTimeout = 900
  objADORs.Open SQL, objAdoCn, 3, 4 ', 4  'adOpenStatic= 3 ,  adLockBatchOptimistic = 4  (orignal)  'cambio desde 24/07/2000 ' 1 adLockReadOnly , ' 4 adCmdStoredProc
  Set GetRecordset1 = objADORs
  Set GetRecordset1.ActiveConnection = objAdoCn
  Set objADORs.ActiveConnection = Nothing
  objAdoCn.Close
  Set objAdoCn = Nothing
 
Exit Function
ehGetRecordset:
  err.Raise err.Number, err.Source, err.Description
  MsgBox err.Description
End Function


Private Sub Reporte()
Dim strSQL As String
On Error GoTo Fin

Dim oo As Object, vRutaLogo As Variant
    
    Screen.MousePointer = 11
    'strSQL = "SELECT Ruta_Logo From SEGURIDAD..SEG_EMPRESAS " & _
    '         "WHERE Cod_Empresa = '" & vemp & "'"
    'vRutaLogo = DevuelveCampo(strSQL, cCONNECT)
    'vRutaLogo = CStr(IIf(IsNull(vRutaLogo), "", vRutaLogo))
          Set oo = CreateObject("excel.application")
          oo.Workbooks.Open vRuta & "\Rpt_Cobranzas_Grafico.XLT"
          oo.displayalerts = False
          oo.Visible = True
    
    oo.Run "REPORTE", gridex1.ADORecordset, cCONNECT
    
    Screen.MousePointer = vbNormal
    'oo.Workbooks.Close
    Set oo = Nothing
Exit Sub
Fin:
    Screen.MousePointer = vbNormal
    MsgBox err.Number
End Sub


Private Sub DrawPieChart()



Draw1.ClearData
  Dim I As Long
    Set mRs = GetRecordset1(cCONNECT, "CN_VENTAS_MUESTRA_DEUDA_POR_CLIENTE  '" & Txt_Porcentaje & "','" & Tip_cliente & "'")
  With Draw1

    .GraphType = dgtPie
    

    .ShowLabel = 1
    .ShowNumbers = 1
    .ShowPercent = 0
    .ShowLegend = 1
    .UseRNDColor = 0
    mRs.MoveFirst

    For I = 1 To mRs.RecordCount
      .AddData Left(mRs!Cliente, 15), CInt(mRs!Porcentaje), Colours(I - 1)
      
       
       mRs.MoveNext
    Next I
    

    .DrawGraph
  End With
End Sub

Private Sub DrawBarChart()
    On Error GoTo SALTO_ERROR

  Draw1.ClearData
  Dim I As Long
  With Draw1
      Set mRs = GetRecordset1(cCONNECT, strSQL)
      
      If mRs.RecordCount = 2 Then
        Colours = Array(vbRed, vbBlue)
      End If
      
      If mRs.RecordCount = 3 Then
      Colours = Array(vbRed, vbBlue, vbGreen)
      End If
      
      If mRs.RecordCount = 4 Then
      Colours = Array(vbRed, vbBlue, vbGreen, vbMagenta)
      End If
      
      If mRs.RecordCount = 5 Then
      Colours = Array(vbRed, vbBlue, vbGreen, vbMagenta, vbYellow)
      End If
      If mRs.RecordCount = 6 Then
      Colours = Array(vbRed, vbBlue, vbGreen, vbMagenta, vbYellow, vbCyan)
      End If
      
      If mRs.RecordCount = 7 Then
      Colours = Array(vbRed, vbBlue, vbGreen, vbMagenta, vbYellow, vbCyan, vbWhite)
      End If
      
      If mRs.RecordCount = 8 Then
      Colours = Array(vbRed, vbBlue, vbGreen, vbMagenta, vbYellow, vbCyan, vbWhite, vbDesktop)
      End If
      
      If mRs.RecordCount = 9 Then
      Colours = Array(vbRed, vbBlue, vbGreen, vbMagenta, vbYellow, vbCyan, vbWhite, vbDesktop, vbTitleBarText)
      End If
      If mRs.RecordCount = 10 Then
      Colours = Array(vbRed, vbBlue, vbGreen, vbMagenta, vbYellow, vbCyan, vbWhite, vbDesktop, vbTitleBarText, vbActiveBorder)
      End If
      
      If mRs.RecordCount = 11 Then
      Colours = Array(vbRed, vbBlue, vbGreen, vbMagenta, vbYellow, vbCyan, &H808080, vbDesktop, &H4000&, &HC0C0FF, &HFF80FF)
      End If
      
      If mRs.RecordCount = 12 Then
      Colours = Array(vbRed, vbBlue, vbGreen, vbMagenta, vbYellow, vbCyan, &H808080, vbDesktop, &H4000&, &HC0C0FF, &HFF80FF, &HE0E0E0)
      End If
      If mRs.RecordCount = 13 Then
      Colours = Array(vbRed, vbBlue, vbGreen, vbMagenta, vbYellow, vbCyan, &H808080, vbDesktop, &H4000&, &HC0C0FF, &HFF80FF, &HE0E0E0, &HFFFFC0)
      End If
      
      
  

    .YGrad = 5
    .YTop = 0
    .OriginY = 250
    .YAxisNegative = 0
    

    .GraphType = dgtPie
    

    .ShowGrid = 1
    .ShowLegend = 1
    
      .PlotAreaColor = &HEEEEEE
    

    
    .ShowBarTotal = 1
    .UseRNDColor = 0
    .AxisTextFont.Size = 1

    mRs.MoveFirst

    For I = 1 To mRs.RecordCount - 1
    
      .AddData Left(mRs!Descripcion, 15), mRs!Porcentaje, Colours(I - 1)
       
       mRs.MoveNext
    Next I
    

    .DrawGraph
    
  End With
  
  Exit Sub
SALTO_ERROR:
    MsgBox err.Description, vbCritical
End Sub


Private Sub Opt_Exportacion_Click()
Tip_cliente = "E"
Txt_Porcentaje.SetFocus
End Sub

Private Sub opt_local_Click()
Tip_cliente = "N"
Txt_Porcentaje.SetFocus
End Sub

Private Sub Option1_Click()

strSQL = "CN_VENTAS_MUESTRA_DEUDA_POR_CLIENTE  '" & Txt_Porcentaje & "','" & Tip_cliente & "'"
mostrar
DrawBarChart
Option1.Value = False

End Sub

Private Sub Option2_Click()
strSQL = "CN_VENTAS_MUESTRA_DEUDA_POR_TIPO_DOCUMENTO  "
mostrar
DrawBarChart
End Sub

Private Sub Option3_Click()
strSQL = "CN_VENTAS_MUESTRA_DEUDA_POR_FEC_VENCIMIENTO  '" & Tip_cliente & "'"
mostrar
DrawBarChart
End Sub

Private Sub Txt_Porcentaje_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
Else
        Call SoloNumeros(Txt_Porcentaje, KeyAscii, False)
End If


End Sub
