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
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtNum_DocCobra 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8100
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9600
      Width           =   3480
   End
   Begin VB.TextBox txtSer_DocCobra 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4200
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
      TabIndex        =   4
      Top             =   0
      Width           =   14640
      Begin VB.TextBox txtNum_Parte 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4695
         TabIndex        =   14
         Top             =   360
         Width           =   1410
      End
      Begin VB.TextBox txtDes 
         BackColor       =   &H80000000&
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
         Left            =   10200
         TabIndex        =   2
         Top             =   195
         Width           =   4170
         _ExtentX        =   7355
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
      Begin VB.Label lblEstado 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   4725
         TabIndex        =   17
         Top             =   660
         Width           =   1875
      End
      Begin VB.Label Label1 
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
         Left            =   8160
         MouseIcon       =   "frmShowPartexAutorizar.frx":0464
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   405
         Width           =   1395
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
         Left            =   6720
         MouseIcon       =   "frmShowPartexAutorizar.frx":076E
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   405
         Width           =   1035
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Origen :"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   405
         Width           =   555
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
      Column(1)       =   "frmShowPartexAutorizar.frx":0A78
      Column(2)       =   "frmShowPartexAutorizar.frx":0B40
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmShowPartexAutorizar.frx":0BE4
      FormatStyle(2)  =   "frmShowPartexAutorizar.frx":0D1C
      FormatStyle(3)  =   "frmShowPartexAutorizar.frx":0DCC
      FormatStyle(4)  =   "frmShowPartexAutorizar.frx":0E80
      FormatStyle(5)  =   "frmShowPartexAutorizar.frx":0F58
      FormatStyle(6)  =   "frmShowPartexAutorizar.frx":1010
      FormatStyle(7)  =   "frmShowPartexAutorizar.frx":10F0
      FormatStyle(8)  =   "frmShowPartexAutorizar.frx":119C
      ImageCount      =   0
      PrinterProperties=   "frmShowPartexAutorizar.frx":124C
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Nro Doc :"
      Height          =   195
      Left            =   3480
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

Dim ssql As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle
Dim fmtCon As JSFmtCondition

If txtNum_Parte = "" Then Exit Sub

ssql = "Cn_Ventas_Emision_Parte_Cobranzas '" & txtCod_Origen & "','" & txtNum_Parte & "','X'"

GridEX1.ClearFields

GridEX1.DefaultGroupMode = jgexDGMExpanded

bCargaGRid = False

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(ssql, cCONNECT)

NroReg = GridEX1.ADORecordset.RecordCount '  GridEX1.RowCount

Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Cliente").Index, jgexSortAscending)

MuestraSubTotales

GridEX1.BackColorRowGroup = &H80000005

GridEX1.ColumnHeaderHeight = 500

GridEX1.Columns("Documento").Width = 1500
GridEX1.Columns("Emision").Width = 945
GridEX1.Columns("Vencimiento").Width = 1005
GridEX1.Columns("Ref_Pago").Width = 1935
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

If GridEX1.RowCount > 0 Then
  Me.Caption = "Revision del Parte de Cobranza Nro " & txtNum_Parte & " de Fecha : " & GridEX1.Value(GridEX1.Columns("Fec_Transaccion").Index)
Else
  Me.Caption = "Revision del Partes de Cobranzas "
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

On Error GoTo drDepurar

Dim Msg As Variant, lvSql As String

Select Case ActionName

Case "BUSCAR"
  Buscar
Case "CERRARPARTE"

  If GridEX1.RowCount = 0 Then Exit Sub

  If MsgBox("Esta seguro de Cerrar el Parte Nro " & txtNum_Parte, vbYesNo, "IMPORTANTE") = vbYes Then
    lvSql = "CN_VENTAS_PARTES_COBRANZA_REVERSION '" & txtCod_Origen & "','" & txtNum_Parte & "'"
    Call ExecuteCommandSQL(cCONNECT, lvSql)
    MsgBox "El Parte se Cerro Satisfactoriamente", vbInformation, "AVISO"
    
  End If
    
Case "ABRIRPARTE"

  If GridEX1.RowCount = 0 Then Exit Sub

  If MsgBox("Esta seguro de Abrir el Parte Nro " & txtNum_Parte, vbYesNo, "IMPORTANTE") = vbYes Then
    lvSql = "CN_VENTAS_ANULAR_REVISION_PARTE_COBRANZA '" & txtCod_Origen & "','" & txtNum_Parte & "','" & vusu & "','" & ComputerName & "'"
    Call ExecuteCommandSQL(cCONNECT, lvSql)
    MsgBox "El Parte se Abrio Satisfactoriamente", vbInformation, "AVISO"
  End If

Case "SALIR"
   Unload Me
End Select

Exit Sub
Resume
drDepurar:
  errores Err.Number

End Sub

Sub Actualiza_Autorizacion()

On Error GoTo drDepurar

Dim ssql As String

  ssql = "CN_VENTAS_REVISION_COBRANZA '$' , $ , '$' , '$' , '$' "
  ssql = VBsprintf(ssql, GridEX1.Value(GridEX1.Columns("Fec_Transaccion").Index), _
                         GridEX1.Value(GridEX1.Columns("Secuencia_Transaccion").Index), _
                         IIf(GridEX1.Value(GridEX1.Columns("Sel").Index), "S", "N"), _
                         vusu, ComputerName)
  ExecuteCommandSQL cCONNECT, ssql

Exit Sub

drDepurar:
  errores Err.Number

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

Private Sub Label1_Click()

If GridEX1.RowCount = 0 Then Exit Sub

If MsgBox("Esta seguro de Desmarcar Todo el Parte" & txtNum_Parte, vbYesNo, "IMPORTANTE") = vbYes Then Seleecionar_Todo 0

End Sub

Private Sub lbSeleccionar_Click()

If GridEX1.RowCount = 0 Then Exit Sub

If MsgBox("Esta seguro de Seleccionar Todo el Parte" & txtNum_Parte, vbYesNo, "IMPORTANTE") = vbYes Then Seleecionar_Todo 1
  
End Sub

Sub Seleecionar_Todo(bSeleccion As Integer)

On Error GoTo errorx
Dim ssql As String
Dim aMess(4), I As Integer
  
GridEX1.MoveFirst

For I = 1 To NroReg

  GridEX1.Value(GridEX1.Columns("Sel").Index) = bSeleccion

  ssql = "CN_VENTAS_REVISION_COBRANZA '$' , $ , '$' , '$' , '$' "
  ssql = VBsprintf(ssql, GridEX1.Value(GridEX1.Columns("Fec_Transaccion").Index), _
                         GridEX1.Value(GridEX1.Columns("Secuencia_Transaccion").Index), _
                         IIf(GridEX1.Value(GridEX1.Columns("Sel").Index), "S", "N"), _
                         vusu, ComputerName)
  ExecuteCommandSQL cCONNECT, ssql

  GridEX1.MoveNext

Next I

Mensaje kMESSAGE_INF_PROCESS_SATISFACTO

Exit Sub
Resume
errorx:
    errores Err.Number
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
