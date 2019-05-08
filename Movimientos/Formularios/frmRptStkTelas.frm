VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmRptStkTelas 
   Caption         =   "Stocks Telas Valorizados Mensuales"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3750
      TabIndex        =   6
      Top             =   5895
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "0~0~IMPRIMIR~True~True~&Imprimir~0~0~1~~0~False~False~&Imprimir~~1~0~SALIR~True~True~&Salir~0~0~2~~0~False~False~&Salir~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX gexStkTelas 
      Height          =   4395
      Left            =   60
      TabIndex        =   5
      Top             =   1230
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   7752
      Version         =   "2.0"
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ContScroll      =   -1  'True
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmRptStkTelas.frx":0000
      Column(2)       =   "frmRptStkTelas.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmRptStkTelas.frx":016C
      FormatStyle(2)  =   "frmRptStkTelas.frx":02A4
      FormatStyle(3)  =   "frmRptStkTelas.frx":0354
      FormatStyle(4)  =   "frmRptStkTelas.frx":0408
      FormatStyle(5)  =   "frmRptStkTelas.frx":04E0
      FormatStyle(6)  =   "frmRptStkTelas.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmRptStkTelas.frx":0678
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   9600
      Begin MSComCtl2.DTPicker dtpAnoMes 
         Height          =   315
         Left            =   6225
         TabIndex        =   3
         Top             =   420
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy MMM"
         Format          =   60162051
         CurrentDate     =   37817
      End
      Begin VB.ComboBox cboAlmacen 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   420
         Width           =   4050
      End
      Begin FunctionsButtons.FunctButt fnbBuscar 
         Height          =   495
         Left            =   8235
         TabIndex        =   4
         Top             =   300
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~True~True~&Buscar~0~0~1~~0~False~False~&Buscar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label2 
         Caption         =   "Año / Mes :"
         Height          =   255
         Left            =   5295
         TabIndex        =   2
         Top             =   465
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen"
         Height          =   225
         Left            =   120
         TabIndex        =   0
         Top             =   465
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmRptStkTelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String

Private Sub fnbBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    If cboAlmacen.ListIndex = -1 Then
        MsgBox "Se debe elgir un Alamcén", vbExclamation + vbOKOnly, "Ver"
        Exit Sub
    End If
    
    'Select Case Trim(Right(cboAlmacen, 3)) Case "T"
    If Trim(Right(cboAlmacen, 3)) = "T" Then
        strSQL = "SM_MUESTRA_STOCKS_MENSUALES_TELAS_ACABADAS_VALORIZADAS "
    Else
        strSQL = "SM_MUESTRA_STOCKS_MENSUALES_TELAS_CRUDAS_VALORIZADAS "
    End If
    
    strSQL = strSQL & "'" & Left(cboAlmacen, 2) & "', '" & Format(dtpAnoMes, "yyyy") & _
             "', '" & Format(dtpAnoMes, "mm") & "'"
             
    Set gexStkTelas.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    gexStkTelas.Columns("lote").Width = 450
    gexStkTelas.Columns("Proveedor").Width = 2160
    gexStkTelas.Columns("COD_TELA").Width = 975
    gexStkTelas.Columns("DES_TELA").Width = 1500
    gexStkTelas.Columns("Cod_Comb").Width = 1035
    gexStkTelas.Columns("Des_Comb").Width = 1500
    'gexStkTelas.Columns("Cod_color").Width = 1125
    'gexStkTelas.Columns("NOMBRE_COLOR").Width = 2175
    gexStkTelas.Columns("Cod_Talla").Width = 1065
    gexStkTelas.Columns("Descripcion").Width = 1500
    gexStkTelas.Columns("CALIDAD").Width = 825
    gexStkTelas.Columns("STOCK_FINAL_KGS").Width = 1635
    gexStkTelas.Columns("STOCK_FINAL_UNI").Width = 1590
    gexStkTelas.Columns("PRECIO_UNITARIO").Width = 1605
    gexStkTelas.Columns("IMPORTE_SOLES").Width = 1500
    
    gexStkTelas.SetFocus
    
    gexStkTelas.Tag = Trim(Right(cboAlmacen, 3))
End Sub

Private Sub Form_Load()
    LoadAlmacenes
    dtpAnoMes = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "IMPRIMIR"
        Reporte
    Case "SALIR"
        Unload Me
    End Select
End Sub

Private Sub Reporte()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String, varLogo As Variant, rstAux As ADODB.Recordset

    Screen.MousePointer = 11
    Ruta = vRuta & "\StkTelasValor.xlt"
    'Ruta = App.Path & "\kardex.xlt"
'    Usu = "Usuario : " & vusu
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    
    strSQL = "SELECT Ruta_Logo FROM SEG_EMPRESAS WHERE cod_EMPRESA ='" & Trim(vemp1) & "'"
    varLogo = DevuelveCampo(strSQL, cSEGURIDAD)
    strSQL = IIf(IsNull(varLogo), "", varLogo)
    
    Set rstAux = gexStkTelas.ADORecordset
    
    If Not rstAux Is Nothing Then
        oo.Run "Reporte", rstAux, Mid(cboAlmacen, 6, 30), Format(dtpAnoMes, "yyyy"), _
                Format(dtpAnoMes, "mm"), strSQL, cConnect
    Else
        MsgBox "Se debe presionar Buscar Anter de Imprimir", vbInformation + vbOKOnly, "Imprimir"
        fnbBuscar.SetFocus
    End If
    Set oo = Nothing
    Screen.MousePointer = 0
Exit Sub
hand:
    Screen.MousePointer = 0
    ErrorHandler Err, "GeneraReportes"
    Set oo = Nothing
End Sub

Private Sub LoadAlmacenes()
Dim rstAux As ADODB.Recordset
    strSQL = "SELECT Cod_Almacen, Nom_Almacen, Tip_Presentacion FROM LG_ALMACEN WHERE Tip_Item = 'T'"
    Set rstAux = CargarRecordSetDesconectado(strSQL, cConnect)
    cboAlmacen.Clear
    With rstAux
    If .RecordCount > 0 Then .MoveFirst
    Do Until .EOF
        cboAlmacen.AddItem !Cod_Almacen & " - " & !Nom_Almacen & Space(100) & !Tip_presentacion
        .MoveNext
    Loop
    .Close
    End With
    Set rstAux = Nothing
End Sub
