VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmShowCanjeAutorizar 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Revision de Partes  de Cancelaciones"
   ClientHeight    =   9615
   ClientLeft      =   270
   ClientTop       =   675
   ClientWidth     =   14700
   ForeColor       =   &H00C0FFFF&
   Icon            =   "frmShowCanjeAutorizar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   14700
   Begin VB.Frame FraBuscar 
      Caption         =   "Argumentos de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   14520
      Begin MSComCtl2.DTPicker txtFecha 
         Height          =   300
         Left            =   1080
         TabIndex        =   6
         Top             =   367
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   21757953
         CurrentDate     =   38590
      End
      Begin VB.OptionButton optNotasAbono 
         Caption         =   "&Notas de Abono"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton optCanje_FActuras 
         Caption         =   "Canje de &Facturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton optLetras 
         Caption         =   "&Facturas Vs Letras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   675
         Left            =   12120
         TabIndex        =   2
         Top             =   195
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   1191
         Custom          =   $"frmShowCanjeAutorizar.frx":030A
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1000
         ControlHeigth   =   650
         ControlSeparator=   40
      End
      Begin VB.Label lbSeleccionar 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Marca Todo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   8640
         MouseIcon       =   "frmShowCanjeAutorizar.frx":045C
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "No Marcar Todo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   10320
         MouseIcon       =   "frmShowCanjeAutorizar.frx":0766
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
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
         TabIndex        =   7
         Top             =   420
         Width           =   660
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   8460
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1065
      Width           =   14640
      _ExtentX        =   25823
      _ExtentY        =   14923
      Version         =   "2.0"
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
      Column(1)       =   "frmShowCanjeAutorizar.frx":0A70
      Column(2)       =   "frmShowCanjeAutorizar.frx":0B38
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmShowCanjeAutorizar.frx":0BDC
      FormatStyle(2)  =   "frmShowCanjeAutorizar.frx":0D14
      FormatStyle(3)  =   "frmShowCanjeAutorizar.frx":0DC4
      FormatStyle(4)  =   "frmShowCanjeAutorizar.frx":0E78
      FormatStyle(5)  =   "frmShowCanjeAutorizar.frx":0F50
      FormatStyle(6)  =   "frmShowCanjeAutorizar.frx":1008
      FormatStyle(7)  =   "frmShowCanjeAutorizar.frx":10E8
      FormatStyle(8)  =   "frmShowCanjeAutorizar.frx":1194
      ImageCount      =   0
      PrinterProperties=   "frmShowCanjeAutorizar.frx":1244
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   6435
      Top             =   6825
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmShowCanjeAutorizar"
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
Dim Tip_Doc As String, NroReg As Double
Public codigo As String, Descripcion As String

Private Sub Form_Load()

'  FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name) & "/SALIR"

  iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))

  If InStr(FunctButt1.FunctionsUser, "AUTORIZARPAGO") <> 0 Then
      bPuedeAutorizar = True
  End If
  
  txtFecha = Date
  
  Tip_Doc = "L"

  
End Sub

Private Sub Busca()

Dim sSQL As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle
Dim fmtCon As JSFmtCondition


sSQL = "Cn_Ventas_Emision_Parte_Cancelaciones '" & txtFecha & "','" & Tip_Doc & "','N','S'"

GridEX1.ClearFields

GridEX1.DefaultGroupMode = jgexDGMExpanded

bCargaGRid = False

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)
GridEX1.MoveLast
NroReg = GridEX1.RowCount
GridEX1.MoveFirst

Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Cli").Index, jgexSortAscending)

MuestraSubTotales

GridEX1.BackColorRowGroup = &H80000005

GridEX1.ColumnHeaderHeight = 500

GridEX1.Columns("Documento").Width = 1700
GridEX1.Columns("Fec_EmiDoc").Width = 1140
GridEX1.Columns("Fec_EmiDoc").Caption = "Fecha Emision"
GridEX1.Columns("Fec_VenDoc").Width = 1500
GridEX1.Columns("Fec_VenDoc").Caption = "Fecha Vencimiento"
GridEX1.Columns("sel").Width = 810

'GridEX1.Columns("Letra_Final").Caption = "Nro Letra Final"
'GridEX1.Columns("Letra_Inicial").Caption = "Nro Letra Inicial"
GridEX1.Columns("Importe").Format = "###,###.00"
GridEX1.Columns("Importe").Caption = "Importe Dolares"
GridEX1.Columns("Importe").Width = 1290
GridEX1.Columns("Importe_Soles").Format = "###,###.00"
GridEX1.Columns("Importe_Soles").Caption = "Importe Soles"
GridEX1.Columns("Importe_Soles").Width = 1085

GridEX1.Columns("Moneda").Width = 810
GridEX1.Columns("Num_Ruc").Visible = False
GridEX1.Columns("Cliente").Visible = False
GridEX1.Columns("Num_Corre").Visible = False
GridEX1.Columns("Cli").Visible = False
GridEX1.Columns("Simbolo").Visible = False

GridEX1.Columns("Tipo_Cambio").Format = "###,###.00"
GridEX1.Columns("Saldo").Format = "#.00"

GridEX1.Columns("Tipo_Cambio").Width = 1500
GridEX1.Columns("Saldo").Width = 1500


GridEX1.Columns("SEL").ColumnType = jgexCheckBox
GridEX1.Columns("SEL").Visible = True
GridEX1.Columns("SEL").EditType = jgexEditCheckBox
GridEX1.Columns("SEL").Width = 500

If GridEX1.RowCount > 0 Then
  Me.Caption = "Revision del Parte de Cancelaciones de Fecha : " & txtFecha
Else
  Me.Caption = "Revision del Parte de Cancelaciones "
End If

SetColores

GridEX1.DefaultGroupMode = jgexDGMCollapsed

GridEX1.DefaultGroupMode = jgexDGMExpanded

GridEX1.ContinuousScroll = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

'On Error GoTo drDepurar

Dim Msg As Variant, lvSql As String

Select Case ActionName

Case "BUSCAR"
  Busca
Case "SALIR"
   Unload Me
End Select

Exit Sub
Resume
drDepurar:
  errores err.Number

End Sub

Sub Actualiza_Autorizacion()

On Error GoTo drDepurar

Dim sSQL As String

 If optLetras Then
 
  sSQL = "CN_VENTAS_REVISION_CANJE_LETRAS '$' , '$' , '$' , '$' , '$' , '$' "
  
  sSQL = VBsprintf(sSQL, GridEX1.Value(GridEX1.Columns("Num_Corre").Index), _
                   IIf(GridEX1.Value(GridEX1.Columns("Sel").Index), "S", "N"), _
                   vusu, ComputerName, _
                   GridEX1.Value(GridEX1.Columns("Letra_Inicial").Index), _
                   IIf(Trim(GridEX1.Value(GridEX1.Columns("Letra_Final").Index)) = "", GridEX1.Value(GridEX1.Columns("Letra_Inicial").Index), GridEX1.Value(GridEX1.Columns("Letra_Final").Index)))
Else

  sSQL = "CN_VENTAS_REVISION_CANJE_DOCUMENTOS '$' , '$' , '$' , '$' , '$' "
  
  sSQL = VBsprintf(sSQL, GridEX1.Value(GridEX1.Columns("Num_Corre").Index), _
                      IIf(GridEX1.Value(GridEX1.Columns("Sel").Index), "S", "N"), _
                      vusu, ComputerName, txtFecha)

End If

                         
ExecuteCommandSQL cCONNECT, sSQL
  

Exit Sub

drDepurar:
  errores err.Number

End Sub

Private Sub GridEX1_AfterColEdit(ByVal ColIndex As Integer)

  If ColIndex = GridEX1.Columns("SEL").Index Then Actualiza_Autorizacion

End Sub

Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)

Select Case ColIndex
  Case Is = GridEX1.Columns("SEL").Index
    Cancel = False
  Case Else
    Cancel = True
  End Select
End Sub

Private Sub GridEX1_Click()

'On Error Resume Next
    Dim ColIndex As Long
    Dim oRowData As JSRowData
    Dim SGRUPO As String
    Dim iRow As Long
    Dim I As Long
    Dim sCaptionGroup As String

    bCargaGRid = True

        If GridEX1.RowCount > 0 Then
        ColIndex = GridEX1.Col

        If Not GridEX1.IsGroupItem(GridEX1.Row) Then
            If ColIndex = 0 Then Exit Sub
            If UCase(GridEX1.Columns(ColIndex).Key) = "SEL" Then
                bClickColSelec = True
                SendKeys "{ENTER}"
            End If
        Else
            If GridEX1.IsGroupItem(GridEX1.Row) Then
            End If
        End If
    End If
End Sub

Private Sub GridEX1_DblClick()
  If GridEX1.RowCount = 0 Then Exit Sub
  Load frmMuestraGeneral
  With frmMuestraGeneral
    .Caption = "Detalle Cliente " & GridEX1.Value(GridEX1.Columns("Cliente").Index) & " Documento : " & GridEX1.Value(GridEX1.Columns("Documento").Index)
    .strSQL = "Ventas_Muestra_Cobranzas_del_Documento '" & GridEX1.Value(GridEX1.Columns("NUM_CORRE").Index) & "'"
    .Buscar
    .GridEX1.Columns("Fec_Cobranza").Width = 1170
    .GridEX1.Columns("Tipo_Cobranza").Width = 2055
    .GridEX1.Columns("Status").Width = 1080
    .GridEX1.Columns("Fec_Creacion").Visible = False
    .GridEX1.Columns("Secuencia_Transaccion").Visible = False
    .GridEX1.Columns("Parte_Cobranza").Width = 1290
    .GridEX1.Columns("Glosa").Width = 1770
    .GridEX1.Columns("Importe").Width = 1065
    .GridEX1.Columns("Importe").Format = "###,###.00"
    .GridEX1.Columns("Doc_Pago").Width = 1170
    .GridEX1.Columns("Num_Letra_Canje").Width = 1410
     .Show vbModal
  End With
End Sub

Private Sub GridEX1_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
  Dim ocol As JSColumn
  Dim oRow As JSRowData
  Dim vCurrentRow As Variant
  Dim oRowGroup As JSRowData
  Dim sProveedor As String

  iColAnterior = LastCol
  iRowAnterior = LastRow

  If GridEX1.Row <> 0 Then
      Set oRow = GridEX1.GetRowData(GridEX1.Row)
  End If

'  If GridEX1.RowCount > 0 Then
'    On Error Resume Next
'    TxtDes_Banco = GridEX1.Value(GridEX1.Columns("Banco").Index)
'    txtSer_DocCobra = GridEX1.Value(GridEX1.Columns("Serie").Index) & GridEX1.Value(GridEX1.Columns("Nro_Doc").Index)
'    txtNum_DocCobra = GridEX1.Value(GridEX1.Columns("Comentario").Index)
'  End If
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
Set colTemp = GridEX1.Columns("Fec_VenDoc")
colTemp.AggregateFunction = jgexAggregateNone
colTemp.TotalRowPrefix = "SUB TOTAL "

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Importe")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Importe_Soles")
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
'    fmtCon.FormatStyle.BackColor = &HFFFFC0   '&HC0FFC0    ' &HC0E0FF    ' '&HC0FFFF
    
'    Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("Tip_Color").Index, jgexEqual, 1)
'    fmtCon.FormatStyle.ForeColor = &HFF&

'    fmtCon.FormatStyle.BackColor = &HC000&

'    Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("Tip_Color").Index, jgexEqual, 1)
'    fmtCon.FormatStyle.BackColor = &H8080FF
    
'    Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("Tip_Color").Index, jgexEqual, 2)
'    fmtCon.FormatStyle.BackColor = &H8080FF

    Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("Saldo").Index, jgexNotEqual, 0)
    fmtCon.FormatStyle.ForeColor = &HFF&

    
End Sub

Private Sub Label2_Click()
  If GridEX1.RowCount = 0 Then Exit Sub

  If MsgBox("Esta seguro de Desmarcar todo este Parte de Cancelaciones", vbYesNo, "IMPORTANTE") = vbYes Then Seleecionar_Todo 0

End Sub

Private Sub lbSeleccionar_Click()
  
If GridEX1.RowCount = 0 Then Exit Sub

If MsgBox("Esta seguro de Marcar todo este Parte de Cancelaciones", vbYesNo, "IMPORTANTE") = vbYes Then Seleecionar_Todo 1
  
End Sub

Sub Seleecionar_Todo(bSeleccion As Integer)

On Error GoTo errorx
Dim sSQL As String
Dim aMess(4), I As Integer
  
GridEX1.MoveFirst

For I = 1 To NroReg

  GridEX1.Value(GridEX1.Columns("Sel").Index) = bSeleccion

  Call Actualiza_Autorizacion

  GridEX1.MoveNext

Next I

Mensaje kMESSAGE_INF_PROCESS_SATISFACTO

Exit Sub
Resume
errorx:
    errores err.Number
End Sub

Private Sub optCanje_FActuras_Click()
Tip_Doc = "F"
End Sub

Private Sub optLetras_Click()
Tip_Doc = "L"
End Sub

Private Sub optNotasAbono_Click()
Tip_Doc = "A"
End Sub
