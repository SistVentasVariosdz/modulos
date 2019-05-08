VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmShowDocVentas 
   Caption         =   "Revision de Partes  de Cobranzas"
   ClientHeight    =   7665
   ClientLeft      =   225
   ClientTop       =   915
   ClientWidth     =   11550
   Icon            =   "frmShowDocVentas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7665
   ScaleWidth      =   11550
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtNum_DocCobra 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6540
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   7320
      Width           =   4920
   End
   Begin VB.TextBox txtSer_DocCobra 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7320
      Width           =   975
   End
   Begin VB.TextBox TxtDes_Banco 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   765
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7320
      Width           =   2535
   End
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
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   11520
      Begin VB.TextBox txtCod_Origen 
         Height          =   285
         Left            =   990
         MaxLength       =   1
         TabIndex        =   0
         Text            =   "N"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtDes_Origen 
         Height          =   285
         Left            =   1500
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtNum_Parte 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5430
         TabIndex        =   2
         Top             =   360
         Width           =   1410
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   675
         Left            =   9240
         TabIndex        =   3
         Top             =   195
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   1191
         Custom          =   $"frmShowDocVentas.frx":030A
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1000
         ControlHeigth   =   650
         ControlSeparator=   40
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Origen :"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   405
         Width           =   555
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Num. Parte Cobranza :"
         Height          =   195
         Left            =   3600
         TabIndex        =   6
         Top             =   390
         Width           =   1605
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   6180
      Left            =   60
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1065
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   10901
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
      Column(1)       =   "frmShowDocVentas.frx":0397
      Column(2)       =   "frmShowDocVentas.frx":045F
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmShowDocVentas.frx":0503
      FormatStyle(2)  =   "frmShowDocVentas.frx":063B
      FormatStyle(3)  =   "frmShowDocVentas.frx":06EB
      FormatStyle(4)  =   "frmShowDocVentas.frx":079F
      FormatStyle(5)  =   "frmShowDocVentas.frx":0877
      FormatStyle(6)  =   "frmShowDocVentas.frx":092F
      FormatStyle(7)  =   "frmShowDocVentas.frx":0A0F
      FormatStyle(8)  =   "frmShowDocVentas.frx":0ABB
      ImageCount      =   0
      PrinterProperties=   "frmShowDocVentas.frx":0B6B
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Nro Doc :"
      Height          =   195
      Left            =   3600
      TabIndex        =   12
      Top             =   7365
      Width           =   690
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Comentario :"
      Height          =   195
      Left            =   5520
      TabIndex        =   11
      Top             =   7365
      Width           =   885
   End
   Begin VB.Label Label3 
      Caption         =   "Banco:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   7350
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
Attribute VB_Name = "frmShowDocVentas"
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
Dim Doc As String
Public codigo As String, Descripcion As String

Private Sub Form_Load()

'  FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name) & "/SALIR"

  iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))

  If InStr(FunctButt1.FunctionsUser, "AUTORIZARPAGO") <> 0 Then
      bPuedeAutorizar = True
  End If
  
  Call txtCod_Origen_KeyPress(13)
  
  Encuentra_Parte
  
  BUSCAR
  
End Sub
Sub Encuentra_Parte()
  txtNum_Parte = DevuelveCampo("Ventas_Obtiene_Ultimo_Parte '" & txtCod_Origen & "'", cCONNECT)
End Sub
Private Sub BUSCAR()

On Error GoTo drDepurar

Dim sSQL As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle
Dim fmtCon As JSFmtCondition

sSQL = "Ventas_Muestra_Parte_Cobranzas '" & txtCod_Origen & "','" & txtNum_Parte & "'"

GridEX1.ClearFields

GridEX1.DefaultGroupMode = jgexDGMExpanded

bCargaGRid = False

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)

Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Cliente").Index, jgexSortAscending)

MuestraSubTotales

GridEX1.BackColorRowGroup = &H80000005

GridEX1.ColumnHeaderHeight = 500

GridEX1.Columns("Documento").Width = 1500
GridEX1.Columns("Emision").Width = 945
GridEX1.Columns("Vencimiento").Width = 1005
GridEX1.Columns("Ref_Pago").Width = 1380
GridEX1.Columns("Ref_Pago").Caption = "Referencia Pago"
GridEX1.Columns("Moneda").Width = 660
GridEX1.Columns("Tipo_Cambio").Width = 600
GridEX1.Columns("Tipo_Cambio").Caption = "Tipo Cambio"
GridEX1.Columns("Importe_Soles").Width = 870
GridEX1.Columns("Importe_Soles").Caption = "Importe Soles"
GridEX1.Columns("Importe_Soles").Format = "###,###.00"
GridEX1.Columns("Importe_Dolares").Width = 870
GridEX1.Columns("Importe_Dolares").Caption = "Importe Dolares"
GridEX1.Columns("Importe_Dolares").Format = "###,###.00"
GridEX1.Columns("Sel").Width = 375
GridEX1.Columns("Banco").Width = 1500
GridEX1.Columns("Serie").Width = 555
GridEX1.Columns("Nro_Doc").Width = 1110
GridEX1.Columns("Comentario").Width = 2430

GridEX1.Columns("Importe_Total").Width = 870
GridEX1.Columns("Importe_Total").Caption = "Importe Total"
GridEX1.Columns("Importe_Total").Format = "###,###.00"
GridEX1.Columns("Importe_Cancelado").Width = 870
GridEX1.Columns("Importe_Cancelado").Caption = "Importe Cancelado"
GridEX1.Columns("Importe_Cancelado").Format = "###,###.00"
GridEX1.Columns("Importe_Pendiente").Width = 870
GridEX1.Columns("Importe_Pendiente").Caption = "Importe Pendiente"
GridEX1.Columns("Importe_Pendiente").Format = "###,###.00"

GridEX1.Columns("Cliente").Visible = False
GridEX1.Columns("Secuencia_Transaccion").Visible = False
GridEX1.Columns("Fec_Transaccion").Visible = False
GridEX1.Columns("Tip_Color").Visible = False

GridEX1.Columns("SEL").ColumnType = jgexCheckBox
GridEX1.Columns("SEL").Visible = True
GridEX1.Columns("SEL").EditType = jgexEditCheckBox
GridEX1.Columns("SEL").Width = 500


SetColores

GridEX1.DefaultGroupMode = jgexDGMCollapsed

GridEX1.DefaultGroupMode = jgexDGMExpanded

GridEX1.ContinuousScroll = True

Exit Sub
Resume
drDepurar:
  errores Err.Number
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Msg As Variant
    Select Case ActionName
    Case "BUSCAR"
      BUSCAR
    Case "SALIR"
       Unload Me
    End Select
End Sub

Sub Actualiza_Autorizacion()

Dim sSQL As String

  sSQL = "CN_VENTAS_REVISION_COBRANZA '$' , $ , '$' , '$' , '$' "
  sSQL = VBsprintf(sSQL, GridEX1.Value(GridEX1.Columns("Fec_Transaccion").Index), _
                         GridEX1.Value(GridEX1.Columns("Secuencia_Transaccion").Index), _
                         IIf(GridEX1.Value(GridEX1.Columns("Sel").Index), "S", "N"), _
                         vusu, ComputerName)
  ExecuteCommandSQL cCONNECT, sSQL

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
    Dim sGrupo As String
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
    .strSql = "Ventas_Muestra_Cobranzas_del_Documento '" & GridEX1.Value(GridEX1.Columns("NUM_CORRE").Index) & "'"
    .BUSCAR
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

  If GridEX1.RowCount > 0 Then
    On Error Resume Next
    TxtDes_Banco = GridEX1.Value(GridEX1.Columns("Banco").Index)
    txtSer_DocCobra = GridEX1.Value(GridEX1.Columns("Serie").Index) & GridEX1.Value(GridEX1.Columns("Nro_Doc").Index)
    txtNum_DocCobra = GridEX1.Value(GridEX1.Columns("Comentario").Index)
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
Set colTemp = GridEX1.Columns("Moneda")
colTemp.AggregateFunction = jgexAggregateNone
colTemp.TotalRowPrefix = "SUB TOTAL "

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Importe_Soles")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Importe_Dolares")
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
    
    Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("Tip_Color").Index, jgexEqual, 1)
    fmtCon.FormatStyle.ForeColor = &HFF&

'    fmtCon.FormatStyle.BackColor = &HC000&

'    Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("Tip_Color").Index, jgexEqual, 1)
'    fmtCon.FormatStyle.BackColor = &H8080FF
    
'    Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("Tip_Color").Index, jgexEqual, 2)
'    fmtCon.FormatStyle.BackColor = &H8080FF
    
End Sub

Private Sub txtCod_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Origen", "Des_Origen", " Cn_Origen where ", txtCod_Origen, txtDes_Origen, 1, Me)
    Encuentra_Parte
  End If
End Sub

Private Sub txtDes_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Origen", "Des_Origen", " Cn_Origen where ", txtCod_Origen, txtDes_Origen, 1, Me)
    Encuentra_Parte
  End If
End Sub

Private Sub txtNum_Parte_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNum_Parte_LostFocus()
  txtNum_Parte.Text = Format(txtNum_Parte.Text, "00000")
End Sub
