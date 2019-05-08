VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmShowFactGuias 
   Caption         =   "Emision de Facturas - Almacen de Productos Terminados Hilados"
   ClientHeight    =   7170
   ClientLeft      =   195
   ClientTop       =   795
   ClientWidth     =   11610
   Icon            =   "frmShowFactGuias.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   11610
   WindowState     =   2  'Maximized
   Begin VB.Frame FraBuscar 
      Caption         =   "Opciones de Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11520
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   675
         Left            =   7200
         TabIndex        =   2
         Top             =   150
         Width           =   4170
         _ExtentX        =   7355
         _ExtentY        =   1191
         Custom          =   $"frmShowFactGuias.frx":030A
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1000
         ControlHeigth   =   650
         ControlSeparator=   40
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4980
      Left            =   0
      TabIndex        =   0
      Top             =   945
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   8784
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmShowFactGuias.frx":0444
      Column(2)       =   "frmShowFactGuias.frx":050C
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmShowFactGuias.frx":05B0
      FormatStyle(2)  =   "frmShowFactGuias.frx":06E8
      FormatStyle(3)  =   "frmShowFactGuias.frx":0798
      FormatStyle(4)  =   "frmShowFactGuias.frx":084C
      FormatStyle(5)  =   "frmShowFactGuias.frx":0924
      FormatStyle(6)  =   "frmShowFactGuias.frx":09DC
      FormatStyle(7)  =   "frmShowFactGuias.frx":0ABC
      FormatStyle(8)  =   "frmShowFactGuias.frx":0B68
      ImageCount      =   0
      PrinterProperties=   "frmShowFactGuias.frx":0C18
   End
   Begin GridEX20.GridEX GridEX2 
      Height          =   1020
      Left            =   0
      TabIndex        =   3
      Top             =   6000
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   1799
      Version         =   "2.0"
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
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmShowFactGuias.frx":0DF0
      Column(2)       =   "frmShowFactGuias.frx":0EB8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmShowFactGuias.frx":0F5C
      FormatStyle(2)  =   "frmShowFactGuias.frx":1094
      FormatStyle(3)  =   "frmShowFactGuias.frx":1144
      FormatStyle(4)  =   "frmShowFactGuias.frx":11F8
      FormatStyle(5)  =   "frmShowFactGuias.frx":12D0
      FormatStyle(6)  =   "frmShowFactGuias.frx":1388
      FormatStyle(7)  =   "frmShowFactGuias.frx":1468
      FormatStyle(8)  =   "frmShowFactGuias.frx":1514
      ImageCount      =   0
      PrinterProperties=   "frmShowFactGuias.frx":15C4
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   10680
      Top             =   5760
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmShowFactGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iRowAnterior As Long
Dim iColAnterior As Long
Dim bClickColSelec As Boolean
Dim bCargaGRid As Boolean
Dim bPuedeAutorizar  As Boolean
Dim sTipoDocAutorizar As String


Private Sub Form_Load()

On Error GoTo dprDepurar
    
iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))

If InStr(FunctButt1.FunctionsUser, "AUTORIZARPAGO") <> 0 Then
    bPuedeAutorizar = True
End If
  
Exit Sub
dprDepurar:
errores err.Number
    
End Sub

Private Sub Buscar()

On Error GoTo dprDepurar

Dim sSQL As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle

sSQL = "Ventas_Muestra_Documentos_Facturar"

GridEX1.ClearFields

GridEX1.DefaultGroupMode = jgexDGMExpanded
bCargaGRid = False
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)

'Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Cliente").Index, jgexSortAscending)

MuestraSubTotales
GridEX1.BackColorRowGroup = &H80000005

GridEX1.ColumnHeaderHeight = 500

GridEX1.Columns("Fecha").Width = 1170
GridEX1.Columns("Fecha").Caption = "Fecha"
GridEX1.Columns("Cliente").Width = 2610
GridEX1.Columns("Cliente").Caption = "Cliente"
GridEX1.Columns("Ser_Docum").Width = 465
GridEX1.Columns("Num_Docum_Ventas").Width = 965
GridEX1.Columns("Pedidos").Width = 2190
GridEX1.Columns("Pedidos").Caption = "Nr Pedidos  / Orde Compra"
GridEX1.Columns("Guias").Width = 3000
GridEX1.Columns("Guias").Caption = "Guias"
GridEX1.Columns("Retencion").Caption = "Detraccion"
GridEX1.Columns("Retencion").Width = 825
GridEX1.Columns("Moneda").Width = 720
GridEX1.Columns("Glosa").Width = 3915
GridEX1.Columns("Moneda").Caption = "Moneda"
GridEX1.Columns("Total Cantidad").Width = 900
GridEX1.Columns("Total Cantidad").Caption = "Total Cantidad"

GridEX1.Columns("Imp_Neto").Width = 780
GridEX1.Columns("Imp_Neto").Caption = "Importe Neto"
GridEX1.Columns("Gastos Financieros").Width = 750
GridEX1.Columns("Gastos Financieros").Caption = "Gastos Financieros"
GridEX1.Columns("IGV").Width = 765
GridEX1.Columns("IGV").Caption = "IGV"

GridEX1.Columns("Direccion").Width = 4980
GridEX1.Columns("Direccion").Caption = "Direccion"


GridEX1.Columns("Total Valor Venta").Width = 960
GridEX1.Columns("Total Valor Venta").Caption = "Total Valor Venta"

GridEX1.Columns("Ser_Docum").Caption = "Serie"
GridEX1.Columns("Num_Docum_Ventas").Caption = "Nro Factura"

GridEX1.Columns("SEL").Width = 495

GridEX1.Columns("Fecha").Format = "dd/mm/yyyy"

GridEX1.Columns("Total Cantidad").Format = "#######0.00"
GridEX1.Columns("Total Valor Venta").Format = "#######0.00"

GridEX1.Columns("SEL").ColumnType = jgexCheckBox
GridEX1.Columns("SEL").Visible = True
GridEX1.Columns("SEL").EditType = jgexEditCheckBox
GridEX1.Columns("SEL").Width = 500

GridEX1.Columns("Retencion").ColumnType = jgexCheckBox
GridEX1.Columns("Retencion").Visible = True
GridEX1.Columns("Retencion").EditType = jgexEditCheckBox
GridEX1.Columns("Retencion").Width = 500


SetColores
GridEX1.DefaultGroupMode = jgexDGMCollapsed


GridEX1.ContinuousScroll = True

Exit Sub
dprDepurar:
errores err.Number
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

On Error GoTo dprDepurar

    Dim Msg As Variant
    Select Case ActionName
    Case "BUSCAR"
      Buscar
    Case "IMPRIMIR"
        If GridEX1.RowCount = 0 Then Exit Sub
        Msg = MsgBox("¿Esta seguro de Imprimir estos Documentos?", vbYesNo)
        If Msg = vbNo Then Exit Sub
        Imprime_Docum_Ventas True
        Buscar
    Case "VISTAPRELIMINAR"
        If GridEX1.RowCount = 0 Then Exit Sub
        Preliminar_Docum_Ventas False
        'Call IMPRIMIR(GridEX1.Value(GridEX1.Columns("Num_Corre").Index), GridEX1.Value(GridEX1.Columns("Total Valor Venta").Index), False, GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index))
    Case "SALIR"
       Unload Me
    End Select

Exit Sub
dprDepurar:
errores err.Number
End Sub

Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)

Select Case ColIndex
  Case Is = GridEX1.Columns("Fecha").Index
    Cancel = False
  Case Is = GridEX1.Columns("SEL").Index
    Cancel = False
  Case Is = GridEX1.Columns("Retencion").Index
    Cancel = False
  Case Is = GridEX1.Columns("Glosa").Index
    Cancel = False

  Case Else
    Cancel = True
  End Select
End Sub

Private Sub GridEX1_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    Dim ocol As JSColumn
    Dim oRow As JSRowData
    Dim vCurrentRow As Variant
    Dim oRowGroup As JSRowData
    Dim sProveedor As String, StrSql As String
    
    iColAnterior = LastCol
    iRowAnterior = LastRow
    
    If GridEX1.Row <> 0 Then
        Set oRow = GridEX1.GetRowData(GridEX1.Row)
    End If
  If GridEX1.Row <> 0 Then
    StrSql = "Ventas_Muestra_Detalle_Factura '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "'"
    Set GridEX2.ADORecordset = CargarRecordSetDesconectado(StrSql, cCONNECT)
    GridEX2.Columns("Articulo").Width = 3555
    GridEX2.Columns("Num_Corre").Visible = False
    GridEX2.Columns("Secuencia").Visible = False
    GridEX2.Columns("Origen").Visible = False
  End If
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
Dim strGroupCaption As String

If RowBuffer.RowType = jgexRowTypeGroupHeader Then
    strGroupCaption = RTrim(RowBuffer.GroupCaption) & " (" & RowBuffer.RecordCount & " Documentos " & "" & ") "
    RowBuffer.GroupCaption = strGroupCaption
End If
    
End Sub

Private Sub MuestraSubTotales()
Dim colTemp As JSColumn

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Guias")
colTemp.AggregateFunction = jgexAggregateNone
colTemp.TotalRowPrefix = "SUB TOTAL "

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Total Cantidad")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Total Valor Venta")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

End Sub

Private Sub SetColores()

Dim fmtCon As JSFmtCondition
Dim fmtCond2 As JSFmtCondition
Dim fmtCond3 As JSFmtCondition

Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("SEL").Index, jgexEqual, -1)
    
    With GridEX1.FmtConditions
            .ApplyGroupCondition = True
            .ShowGroupConditionCount = True
            .GroupConditionCountTitle = "Documento(s) Autorizado(s)"
            Set fmtCon = .GroupCondition
    End With
    fmtCon.SetCondition GridEX1.Columns("SEL").Index, jgexEqual, -1
    fmtCon.FormatStyle.FontBold = True
    fmtCon.FormatStyle.BackColor = &HFFFFC0   '&HC0FFC0    ' &HC0E0FF    ' '&HC0FFFF
    
End Sub

Private Sub Imprime_Docum_Ventas(Tipo As Boolean)

On Error GoTo errorx

Dim sSQL As String, Num_Corre As String, rs As New ADODB.Recordset
Dim aMess(4), I As Integer
  
    GridEX1.MoveFirst
    For I = 1 To GridEX1.RowCount
      If GridEX1.Value(GridEX1.Columns("SEL").Index) Then

          sSQL = "Ventas_Actualiza_Datos_Impresion '$' , '$' , '$' , '$', '$' "
          sSQL = VBsprintf(sSQL, _
                 GridEX1.Value(GridEX1.Columns("Num_Corre").Index), _
                 Format(GridEX1.Value(GridEX1.Columns("Fecha").Index), "dd/mm/yyyy"), _
                 IIf(GridEX1.Value(GridEX1.Columns("Retencion").Index), "S", "N"), _
                 GridEX1.Value(GridEX1.Columns("Glosa").Index), "S")

          ExecuteCommandSQL cCONNECT, sSQL
        
        If Imprimir(GridEX1.Value(GridEX1.Columns("Num_Corre").Index), GridEX1.Value(GridEX1.Columns("Total Valor Venta").Index), Tipo, GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index)) = False Then
          MsgBox "Problemas de Impresion con el Documento Nr " & GridEX1.Columns("Num_Docum_Ventas"), vbInformation, "ERROR"
          Buscar
          Exit Sub
        End If
        
      End If
      GridEX1.MoveNext
    Next I
    
    Buscar
    
Exit Sub
Resume
errorx:
    ErrorHandler err, "Autoriza Documentos"
End Sub

Private Sub Preliminar_Docum_Ventas(Tipo As Boolean)

On Error GoTo errorx

Dim sSQL As String, Num_Corre As String, rs As New ADODB.Recordset
Dim aMess(4), I As Integer

  sSQL = "Ventas_Actualiza_Datos_Impresion '$' , '$' , '$' , '$', '$' "
  sSQL = VBsprintf(sSQL, _
         GridEX1.Value(GridEX1.Columns("Num_Corre").Index), _
         Format(GridEX1.Value(GridEX1.Columns("Fecha").Index), "dd/mm/yyyy"), _
         IIf(GridEX1.Value(GridEX1.Columns("Retencion").Index), "S", "N"), _
         GridEX1.Value(GridEX1.Columns("Glosa").Index), "N")
  ExecuteCommandSQL cCONNECT, sSQL
  
  If Imprimir(GridEX1.Value(GridEX1.Columns("Num_Corre").Index), GridEX1.Value(GridEX1.Columns("Total Valor Venta").Index), Tipo, GridEX1.Value(GridEX1.Columns("Cod_TipDoc").Index)) = False Then
    MsgBox "Problemas de Impresion con el Documento Nr " & GridEX1.Columns("Num_Docum_Ventas"), vbInformation, "ERROR"
    Buscar
    Exit Sub
  End If
        
    
Exit Sub
Resume
errorx:
    ErrorHandler err, "Autoriza Documentos"
End Sub


