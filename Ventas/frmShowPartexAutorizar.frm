VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmShowPartexAutorizar 
   Caption         =   "Revision de Partes  de Cobranzas"
   ClientHeight    =   9945
   ClientLeft      =   75
   ClientTop       =   930
   ClientWidth     =   14835
   Icon            =   "frmShowPartexAutorizar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9945
   ScaleWidth      =   14835
   Begin VB.TextBox txtNum_DocCobra 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8100
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9600
      Width           =   6240
   End
   Begin VB.TextBox txtSer_DocCobra 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9600
      Width           =   975
   End
   Begin VB.TextBox TxtDes_Banco 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   765
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9600
      Width           =   4215
   End
   Begin VB.Frame FraBuscar 
      BackColor       =   &H00C0FFFF&
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
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   14640
      Begin VB.TextBox lblEstado 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4695
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   675
         Width           =   1410
      End
      Begin VB.TextBox txtNum_Parte 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4695
         TabIndex        =   14
         Top             =   360
         Width           =   1410
      End
      Begin VB.TextBox txtDes 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   3360
         TabIndex        =   13
         Top             =   765
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.TextBox txtCod_Origen 
         Height          =   285
         Left            =   750
         MaxLength       =   1
         TabIndex        =   0
         Text            =   "N"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtDes_Origen 
         Height          =   285
         Left            =   1260
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   675
         Left            =   9345
         TabIndex        =   2
         Top             =   195
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   1191
         Custom          =   $"frmShowPartexAutorizar.frx":030A
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1000
         ControlHeigth   =   650
         ControlSeparator=   40
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
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
         Left            =   7845
         MouseIcon       =   "frmShowPartexAutorizar.frx":04CB
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   405
         Width           =   1395
      End
      Begin VB.Label lbSeleccionar 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
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
         Left            =   6390
         MouseIcon       =   "frmShowPartexAutorizar.frx":07D5
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   405
         Width           =   1035
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Origen :"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   405
         Width           =   555
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Num. Parte Cobranza :"
         Height          =   195
         Left            =   3000
         TabIndex        =   5
         Top             =   390
         Width           =   1605
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   8220
      Left            =   60
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   14640
      _ExtentX        =   25823
      _ExtentY        =   14499
      Version         =   "2.0"
      RecordNavigator =   -1  'True
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
      Column(1)       =   "frmShowPartexAutorizar.frx":0ADF
      Column(2)       =   "frmShowPartexAutorizar.frx":0BA7
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmShowPartexAutorizar.frx":0C4B
      FormatStyle(2)  =   "frmShowPartexAutorizar.frx":0D83
      FormatStyle(3)  =   "frmShowPartexAutorizar.frx":0E33
      FormatStyle(4)  =   "frmShowPartexAutorizar.frx":0EE7
      FormatStyle(5)  =   "frmShowPartexAutorizar.frx":0FBF
      FormatStyle(6)  =   "frmShowPartexAutorizar.frx":1077
      FormatStyle(7)  =   "frmShowPartexAutorizar.frx":1157
      FormatStyle(8)  =   "frmShowPartexAutorizar.frx":1203
      ImageCount      =   0
      PrinterProperties=   "frmShowPartexAutorizar.frx":12B3
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Nro Doc :"
      Height          =   195
      Left            =   5040
      TabIndex        =   11
      Top             =   9645
      Width           =   690
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Comentario :"
      Height          =   195
      Left            =   7080
      TabIndex        =   10
      Top             =   9645
      Width           =   885
   End
   Begin VB.Label Label3 
      Caption         =   "Banco:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   9630
      Width           =   645
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   6435
      Top             =   6825
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmShowPartexAutorizar"
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
Dim Doc As String, NroReg As Integer
Public codigo As String, Descripcion As String, estado As String

Private Sub Form_Load()

'  FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name) & "/SALIR"

  iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))

  If InStr(FunctButt1.FunctionsUser, "AUTORIZARPAGO") <> 0 Then
      bPuedeAutorizar = True
  End If
  
  Call txtCod_Origen_KeyPress(13)

  
End Sub

Sub Encuentra_Parte()
  txtNum_Parte = DevuelveCampo("Ventas_Obtiene_Ultimo_Parte '" & txtCod_Origen & "'", cCONNECT)
End Sub

Private Sub Buscar()

Dim sSQL As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle
Dim fmtCon As JSFmtCondition

If txtNum_Parte = "" Then Exit Sub

sSQL = "Cn_Ventas_Emision_Parte_Cobranzas '" & txtCod_Origen & "','" & txtNum_Parte & "','X'"

gridex1.ClearFields

gridex1.DefaultGroupMode = jgexDGMExpanded

bCargaGRid = False

Set gridex1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)

Set oGroup = gridex1.Groups.Add(gridex1.Columns("Cliente").Index, jgexSortAscending)



MuestraSubTotales

NroReg = gridex1.ADORecordset.RecordCount '  GridEX1.RowCount

gridex1.BackColorRowGroup = &H80000005

gridex1.ColumnHeaderHeight = 500

gridex1.Columns("Documento").SortType = jgexSortTypeString
gridex1.SortKeys.Add gridex1.Columns("Documento").Index, jgexSortAscending
gridex1.SortKeys.Add gridex1.Columns("Secuencia_Transaccion").Index, jgexSortAscending
  

gridex1.Columns("Documento").Width = 1500
gridex1.Columns("Emision").Width = 945
gridex1.Columns("Vencimiento").Width = 1005
gridex1.Columns("Ref_Pago").Width = 1935
gridex1.Columns("Ref_Pago").Caption = "Referencia Pago"
gridex1.Columns("Moneda").Width = 660
gridex1.Columns("Tipo_Cambio").Width = 600
gridex1.Columns("Tipo_Cambio").Caption = "Tipo Cambio"
gridex1.Columns("Importe_Soles").Width = 870
gridex1.Columns("Importe_Soles").Caption = "Importe Soles"
gridex1.Columns("Importe_Soles").Format = "###,###.00"
gridex1.Columns("Importe_Dolares").Width = 870
gridex1.Columns("Importe_Dolares").Caption = "Importe Dolares"
gridex1.Columns("Importe_Dolares").Format = "###,###.00"
gridex1.Columns("Sel").Width = 375
gridex1.Columns("Banco").Width = 1500
gridex1.Columns("Serie").Width = 555
gridex1.Columns("Nro_Doc").Width = 1110
gridex1.Columns("Comentario").Width = 2430

gridex1.Columns("Importe_Total").Width = 870
gridex1.Columns("Importe_Total").Caption = "Importe Total"
gridex1.Columns("Importe_Total").Format = "###,###.00"
gridex1.Columns("Importe_Cancelado").Width = 870
gridex1.Columns("Importe_Cancelado").Caption = "Importe Cancelado"
gridex1.Columns("Importe_Cancelado").Format = "###,###.00"
gridex1.Columns("Importe_Pendiente").Width = 870
gridex1.Columns("Importe_Pendiente").Caption = "Importe Pendiente"
gridex1.Columns("Importe_Pendiente").Format = "###,###.00"

gridex1.Columns("Cliente").Visible = False
gridex1.Columns("Secuencia_Transaccion").Visible = False
gridex1.Columns("Fec_Transaccion").Visible = False
gridex1.Columns("Tip_Color").Visible = False

gridex1.Columns("SEL").ColumnType = jgexCheckBox
gridex1.Columns("SEL").Visible = True
gridex1.Columns("SEL").EditType = jgexEditCheckBox
gridex1.Columns("SEL").Width = 500

If gridex1.RowCount > 0 Then
  Me.Caption = "Revision del Parte de Cobranza Nro " & txtNum_Parte & " de Fecha : " & gridex1.Value(gridex1.Columns("Fec_Transaccion").Index)
Else
  Me.Caption = "Revision del Partes de Cobranzas "
End If

SetColores

gridex1.DefaultGroupMode = jgexDGMCollapsed

gridex1.DefaultGroupMode = jgexDGMExpanded

gridex1.ContinuousScroll = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

On Error GoTo drDepurar

Dim Msg As Variant, lvSql As String

Select Case ActionName

Case "BUSCAR"
  Buscar
Case "CERRARPARTE"

  If gridex1.RowCount = 0 Then Exit Sub

  If MsgBox("Esta seguro de Cerrar el Parte Nro " & txtNum_Parte, vbYesNo, "IMPORTANTE") = vbYes Then
    lvSql = "CN_VENTAS_PARTES_COBRANZA_REVERSION '" & txtCod_Origen & "','" & txtNum_Parte & "'"
    Call ExecuteCommandSQL(cCONNECT, lvSql)
    MsgBox "El Parte se Cerro Satisfactoriamente", vbInformation, "AVISO"
    
  End If
    
Case "ABRIRPARTE"

  If gridex1.RowCount = 0 Then Exit Sub

  If MsgBox("Esta seguro de Abrir el Parte Nro " & txtNum_Parte, vbYesNo, "IMPORTANTE") = vbYes Then
    lvSql = "CN_VENTAS_ANULAR_REVISION_PARTE_COBRANZA '" & txtCod_Origen & "','" & txtNum_Parte & "','" & vusu & "','" & ComputerName & "'"
    Call ExecuteCommandSQL(cCONNECT, lvSql)
    MsgBox "El Parte se Abrio Satisfactoriamente", vbInformation, "AVISO"
  End If

Case "PARTESPENDIENTES"
    Reporte

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

  sSQL = "CN_VENTAS_REVISION_COBRANZA '$' , $ , '$' , '$' , '$' "
  sSQL = VBsprintf(sSQL, gridex1.Value(gridex1.Columns("Fec_Transaccion").Index), _
                         gridex1.Value(gridex1.Columns("Secuencia_Transaccion").Index), _
                         IIf(gridex1.Value(gridex1.Columns("Sel").Index), "S", "N"), _
                         vusu, ComputerName)
  ExecuteCommandSQL cCONNECT, sSQL

Exit Sub

drDepurar:
  errores err.Number

End Sub

Private Sub GridEX1_AfterColEdit(ByVal ColIndex As Integer)

  If ColIndex = gridex1.Columns("SEL").Index Then Actualiza_Autorizacion

End Sub

Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)

Select Case ColIndex
  Case Is = gridex1.Columns("SEL").Index
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

        If gridex1.RowCount > 0 Then
        ColIndex = gridex1.Col

        If Not gridex1.IsGroupItem(gridex1.Row) Then
            If ColIndex = 0 Then Exit Sub
            If UCase(gridex1.Columns(ColIndex).Key) = "SEL" Then
                bClickColSelec = True
                SendKeys "{ENTER}"
            End If
        Else
            If gridex1.IsGroupItem(gridex1.Row) Then
            End If
        End If
    End If
End Sub

Private Sub GridEX1_DblClick()
  If gridex1.RowCount = 0 Then Exit Sub
  Load frmMuestraGeneral
  With frmMuestraGeneral
    .Caption = "Detalle Cliente " & gridex1.Value(gridex1.Columns("Cliente").Index) & " Documento : " & gridex1.Value(gridex1.Columns("Documento").Index)
    .strSQL = "Ventas_Muestra_Cobranzas_del_Documento '" & gridex1.Value(gridex1.Columns("NUM_CORRE").Index) & "'"
    .Buscar
    .gridex1.Columns("Fec_Cobranza").Width = 1170
    .gridex1.Columns("Tipo_Cobranza").Width = 2055
    .gridex1.Columns("Status").Width = 1080
    .gridex1.Columns("Fec_Creacion").Visible = False
    .gridex1.Columns("Secuencia_Transaccion").Visible = False
    .gridex1.Columns("Parte_Cobranza").Width = 1290
    .gridex1.Columns("Glosa").Width = 1770
    .gridex1.Columns("Importe").Width = 1065
    .gridex1.Columns("Importe").Format = "###,###.00"
    .gridex1.Columns("Doc_Pago").Width = 1170
    .gridex1.Columns("Num_Letra_Canje").Width = 1410
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

  If gridex1.Row <> 0 Then
      Set oRow = gridex1.GetRowData(gridex1.Row)
  End If

  If gridex1.RowCount > 0 Then
    On Error Resume Next
    TxtDes_Banco = gridex1.Value(gridex1.Columns("Banco").Index)
    txtSer_DocCobra = gridex1.Value(gridex1.Columns("Serie").Index) & gridex1.Value(gridex1.Columns("Nro_Doc").Index)
    txtNum_DocCobra = gridex1.Value(gridex1.Columns("Comentario").Index)
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

gridex1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = gridex1.Columns("Moneda")
colTemp.AggregateFunction = jgexAggregateNone
colTemp.TotalRowPrefix = "SUB TOTAL "

gridex1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = gridex1.Columns("Importe_Soles")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

gridex1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = gridex1.Columns("Importe_Dolares")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

End Sub

Private Sub SetColores()

Dim fmtCon As JSFmtCondition
Dim fmtCond2 As JSFmtCondition
Dim fmtCond3 As JSFmtCondition

Set fmtCon = gridex1.FmtConditions.Add(gridex1.Columns("SEL").Index, jgexEqual, -1)

    With gridex1.FmtConditions
            .ApplyGroupCondition = True
            .ShowGroupConditionCount = True
            .GroupConditionCountTitle = "Documento(s) Autorizado(s)"
            Set fmtCon = .GroupCondition
    End With
    
    fmtCon.SetCondition gridex1.Columns("SEL").Index, jgexEqual, -1
    fmtCon.FormatStyle.FontBold = True
'    fmtCon.FormatStyle.BackColor = &HFFFFC0   '&HC0FFC0    ' &HC0E0FF    ' '&HC0FFFF
    
    Set fmtCon = gridex1.FmtConditions.Add(gridex1.Columns("Tip_Color").Index, jgexEqual, 1)
    fmtCon.FormatStyle.ForeColor = &HFF&

'    fmtCon.FormatStyle.BackColor = &HC000&

'    Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("Tip_Color").Index, jgexEqual, 1)
'    fmtCon.FormatStyle.BackColor = &H8080FF
    
'    Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("Tip_Color").Index, jgexEqual, 2)
'    fmtCon.FormatStyle.BackColor = &H8080FF
    
End Sub

Private Sub Label1_Click()

If gridex1.RowCount = 0 Then Exit Sub

If MsgBox("Esta seguro de Desmarcar Todo el Parte" & txtNum_Parte, vbYesNo, "IMPORTANTE") = vbYes Then Seleecionar_Todo 0

End Sub

Private Sub lbSeleccionar_Click()

If gridex1.RowCount = 0 Then Exit Sub

If MsgBox("Esta seguro de Seleccionar Todo el Parte" & txtNum_Parte, vbYesNo, "IMPORTANTE") = vbYes Then Seleecionar_Todo 1
  
End Sub

Sub Seleecionar_Todo(bSeleccion As Integer)

On Error GoTo errorx
Dim sSQL As String
Dim aMess(4), I As Integer
  
gridex1.MoveFirst

For I = 0 To NroReg + 1

  gridex1.Value(gridex1.Columns("Sel").Index) = bSeleccion

  sSQL = "CN_VENTAS_REVISION_COBRANZA '$' , $ , '$' , '$' , '$' "
  sSQL = VBsprintf(sSQL, gridex1.Value(gridex1.Columns("Fec_Transaccion").Index), _
                         gridex1.Value(gridex1.Columns("Secuencia_Transaccion").Index), _
                         IIf(gridex1.Value(gridex1.Columns("Sel").Index), "S", "N"), _
                         vusu, ComputerName)
  ExecuteCommandSQL cCONNECT, sSQL

  gridex1.MoveNext

Next I
gridex1.Update
gridex1.Refresh

Mensaje kMESSAGE_INF_PROCESS_SATISFACTO

Exit Sub
Resume
errorx:
    errores err.Number
End Sub


Private Sub txtCod_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Origen", "Des_Origen", " Cn_Origen where ", txtCod_Origen, txtDes_Origen, 1, Me)
  End If
End Sub

Private Sub txtDes_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Origen", "Des_Origen", " Cn_Origen where ", txtCod_Origen, txtDes_Origen, 1, Me)
  End If
End Sub

Private Sub txtNum_Parte_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub txtNum_Parte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Trim(txtNum_Parte) = "" Then
    Call Busca_Opcion_lis("Num_Parte_Cobranza", "convert(varchar,Fec_Transaccion,103) + '  ' + b.Descripcion", "des_status_parte_cobranza", " Cn_Ventas_Partes_Cobranza a , Cn_Ventas_Status_Partes b, Cn_Status_Parte_Cobranza c Where Origen = Flg_Partes and flg_status= flg_status_parte_cobranza and Fec_Transaccion  > '01/09/05' and Origen like '%" & txtCod_Origen & "' and ", txtNum_Parte, txtDes, lblEstado, 1, Me)
  Else
    SendKeys "{TAB}"
  End If
End If
End Sub

Private Sub txtNum_Parte_LostFocus()
 If Len(txtNum_Parte) < 5 Then txtNum_Parte.Text = Format(txtNum_Parte.Text, "00000")
End Sub

Public Sub Reporte()
  
On Error GoTo ErrorImpresion

VB.Screen.MousePointer = vbHourglass

Dim oo As Object, strSQL As String, RS As Object
Set RS = CreateObject("ADODB.Recordset")
Dim RS1 As Object
Set RS1 = CreateObject("ADODB.Recordset")
Set oo = CreateObject("excel.application")

strSQL = "CN_Muestra_Partes_Cobranzas_por_Status '" & txtCod_Origen & "','C'"

Set RS = CargarRecordSetDesconectado(strSQL, cCONNECT)

If RS.RecordCount = 0 Then
  Screen.MousePointer = vbNormal
  MsgBox "No hay Registros que imprimir", vbInformation, "AVISO"
  Exit Sub
End If

oo.Workbooks.Open vRuta & "\rptListadoPartesCobranza.xlt"
oo.Run "REPORTE", RS, txtCod_Origen, txtDes_Origen, cCONNECT

oo.Visible = True
Screen.MousePointer = vbNormal
oo.Visible = True
Set oo = Nothing

Exit Sub
Resume
ErrorImpresion:
    Screen.MousePointer = vbNormal
    Set oo = Nothing
    Error err.Number
End Sub



