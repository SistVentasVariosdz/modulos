VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmDetalleLetras 
   Caption         =   "Letras por Planilla"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOperaciones 
      Caption         =   "Operaciones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   8055
      Begin GridEX20.GridEX GridEXSelec 
         Height          =   4200
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   7408
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
         GroupByBoxVisible=   0   'False
         ImageCount      =   2
         ImagePicture1   =   "frmDetalleLetras.frx":0000
         ImagePicture2   =   "frmDetalleLetras.frx":031A
         DataMode        =   1
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "frmDetalleLetras.frx":0634
         Column(2)       =   "frmDetalleLetras.frx":06FC
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmDetalleLetras.frx":07A0
         FormatStyle(2)  =   "frmDetalleLetras.frx":08D8
         FormatStyle(3)  =   "frmDetalleLetras.frx":0988
         FormatStyle(4)  =   "frmDetalleLetras.frx":0A3C
         FormatStyle(5)  =   "frmDetalleLetras.frx":0B14
         FormatStyle(6)  =   "frmDetalleLetras.frx":0BCC
         FormatStyle(7)  =   "frmDetalleLetras.frx":0CAC
         FormatStyle(8)  =   "frmDetalleLetras.frx":0D58
         ImageCount      =   2
         ImagePicture(1) =   "frmDetalleLetras.frx":0E08
         ImagePicture(2) =   "frmDetalleLetras.frx":1122
         PrinterProperties=   "frmDetalleLetras.frx":143C
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3765
      TabIndex        =   2
      Top             =   4755
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   900
      Custom          =   $"frmDetalleLetras.frx":1614
      Orientacion     =   0
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmDetalleLetras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public snumero As Integer
Public codigo As String, Descripcion As String
Public sCOD_BANCO As String
Public sFlg_Status_Letras As String
Public sCod_Tipanex As String
Public sCod_Anxo As String
Public sCuenta As String
Public rsGridOperaciones As ADODB.Recordset
Public scadena As String
Public num As String
Dim rsGridDisponibles As ADODB.Recordset
Dim rsGridSeleccionados As ADODB.Recordset
Dim strNum_Corre_Let_Renov As String, TipoAdd As String
Public tipo As String, xCod_Grupo As Integer, strNum_Corre_Let As String
Dim strSQL  As String
Public sNum_Planilla_Letra As Integer
Dim iRowAnterior As Long
Dim iColAnterior As Long
Dim bClickColSelec As Boolean
Dim bCargaGRid As Boolean
Dim bPuedeAutorizar  As Boolean
Dim sTipoDocAutorizar As String
Dim strOpcion As String
Public strCod_Anxo As String
Dim lvSW As Boolean

'
'Public Sub CargarOperaciones()
'
'Dim ssql As String
'Dim oGroup As GridEX20.JSGroup
'Dim oFormat As JSFormatStyle
'
'Dim fmtCon As JSFmtCondition
'Dim c As JSColumn
'Dim vl As JSValueList
'Dim rsGridCopia As ADODB.Recordset
'Dim oData As Object
'
'ssql = "Ventas_Generar_Detalle_Letras  '$','$','$','$','$'"
'ssql = VBsprintf(ssql, sCOD_BANCO, sFlg_Status_Letras, txtCod_TipAne.Text, txtCod_Anexo.Text, txtRuc.Text)
'
'GridEXDispon.ClearFields
'
'Set GridEXDispon.ADORecordset = CargarRecordSetDesconectado(ssql, cCONNECT)
'
'GridEXDispon.Columns("num_corre").Visible = False
'GridEXDispon.Columns("cod_tipdoc").Caption = "  Tipo Doc."
'GridEXDispon.Columns("ser_docum").Caption = "Serie"
'GridEXDispon.Columns("num_docum_ventas").Caption = "Num. Doc."
'GridEXDispon.Columns("cod_moneda").Caption = "Moneda"
'GridEXDispon.Columns("imp_total").Caption = "Imp. Total"
'
'
'GridEXDispon.SortKeys.Clear
'
'GridEXDispon.BackColorRowGroup = &H80000005
'
'GridEXDispon.DefaultGroupMode = jgexDGMCollapsed
'GridEXDispon.CollapseAll
'
'
'GridEXDispon.ContinuousScroll = True
'
'
'ssql = "Ventas_Generar_Detalle_Letras  '$','$','$','$','$'"
'ssql = VBsprintf(ssql, sCOD_BANCO, sFlg_Status_Letras, txtCod_TipAne.Text, txtCod_Anexo.Text, "")
'
'
'GridEXSelec.ClearFields
'
'Set GridEXSelec.ADORecordset = CargarRecordSetDesconectado(ssql, cCONNECT)
'Set rsGridOperaciones = GridEXSelec.ADORecordset
'
'RefrescaGridOperaciones
'
'GridEXSelec.SortKeys.Clear
'
'End Sub

Private Sub RefrescaGrid(ByRef GridEx As GridEx, ByRef rsGrid As ADODB.Recordset)
    Dim oGroup As JSGroup

    Set GridEx.ADORecordset = rsGrid
    GridEx.Refresh

    GridEx.ColumnHeaderHeight = 500
    GridEx.Columns("num_corre").Visible = False
    GridEx.Columns("cod_tipdoc").Caption = "Tipo Doc."
    GridEx.Columns("ser_docum").Caption = "Serie"
    GridEx.Columns("num_docum_ventas").Caption = "Num. Doc."
    GridEx.Columns("cod_moneda").Caption = "Moneda"
    GridEx.Columns("imp_total").Caption = "Imp. Total"
    GridEx.Columns("fec_emidoc").Caption = "Fec. Emi. Doc"
    GridEx.Columns("fec_vendoc").Caption = "Fec. Ven. Doc."

    GridEx.Columns("cod_tipdoc").Width = 800
    GridEx.Columns("ser_docum").Width = 800
    GridEx.Columns("num_docum_ventas").Width = 800
    GridEx.Columns("cod_moneda").Width = 800
    GridEx.Columns("imp_total").Width = 800
    GridEx.Columns("fec_emidoc").Width = 800
    GridEx.Columns("fec_vendoc").Width = 800

    GridEx.ContinuousScroll = True

    GridEx.Refresh
End Sub

Private Sub Form_Load()
    strSQL = "Ventas_Generar_Detalle_Letras  '$','$','$','$','$','$','$','$'"
    strSQL = VBsprintf(strSQL, 4, sNum_Planilla_Letra, sCOD_BANCO, sFlg_Status_Letras, "", "", "", sCuenta)


    Set GridEXSelec.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

    Set rsGridSeleccionados = GridEXSelec.ADORecordset

    GridEXSelec.Refresh

    GridEXSelec.ColumnHeaderHeight = 500
    GridEXSelec.Columns("num_corre").Visible = False
    GridEXSelec.Columns("cod_tipdoc").Caption = "Tipo Doc."
    GridEXSelec.Columns("ser_docum").Caption = "Serie"
    GridEXSelec.Columns("num_docum_ventas").Caption = "Num. Doc."
    GridEXSelec.Columns("cod_moneda").Caption = "Moneda"
    GridEXSelec.Columns("imp_total").Caption = "Imp. Total"
    GridEXSelec.Columns("fec_emidoc").Caption = "Fec. Emi. Doc"
    GridEXSelec.Columns("fec_vendoc").Caption = "Fec. Ven. Doc."

    GridEXSelec.Columns("cod_tipdoc").Width = 1000
    GridEXSelec.Columns("ser_docum").Width = 1000
    GridEXSelec.Columns("num_docum_ventas").Width = 1000
    GridEXSelec.Columns("cod_moneda").Width = 1000
    GridEXSelec.Columns("imp_total").Width = 1000
    GridEXSelec.Columns("fec_emidoc").Width = 1200
    GridEXSelec.Columns("fec_vendoc").Width = 1200

    GridEXSelec.ContinuousScroll = True

    GridEXSelec.Refresh

End Sub

Sub buscar()
    strSQL = "Ventas_Generar_Detalle_Letras  '$','$','$','$','$','$','$','$'"
    strSQL = VBsprintf(strSQL, 4, sNum_Planilla_Letra, sCOD_BANCO, sFlg_Status_Letras, "", "", "", sCuenta)

    Set GridEXSelec.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
    Set rsGridSeleccionados = GridEXSelec.ADORecordset

    GridEXSelec.Refresh

    GridEXSelec.ColumnHeaderHeight = 500
    GridEXSelec.Columns("num_corre").Visible = False
    GridEXSelec.Columns("cod_tipdoc").Caption = "Tipo Doc."
    GridEXSelec.Columns("ser_docum").Caption = "Serie"
    GridEXSelec.Columns("num_docum_ventas").Caption = "Num. Doc."
    GridEXSelec.Columns("cod_moneda").Caption = "Moneda"
    GridEXSelec.Columns("imp_total").Caption = "Imp. Total"
    GridEXSelec.Columns("fec_emidoc").Caption = "Fec. Emi. Doc"
    GridEXSelec.Columns("fec_vendoc").Caption = "Fec. Ven. Doc."

    GridEXSelec.Columns("cod_tipdoc").Width = 1000
    GridEXSelec.Columns("ser_docum").Width = 1000
    GridEXSelec.Columns("num_docum_ventas").Width = 1000
    GridEXSelec.Columns("cod_moneda").Width = 1000
    GridEXSelec.Columns("imp_total").Width = 1000
    GridEXSelec.Columns("fec_emidoc").Width = 1200
    GridEXSelec.Columns("fec_vendoc").Width = 1200

    GridEXSelec.ContinuousScroll = True

    GridEXSelec.Refresh
End Sub


Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim ssql1 As String
Dim ssql As String
Dim iReg As Integer

    Select Case ActionName
    Case "AGREGAR"
    If sCOD_BANCO = 3 Then
        snumero = 10
    ElseIf sCOD_BANCO = 2 Then
        snumero = 15
    Else
        snumero = 10
    End If
    
        If DevuelveCampo("select count(*) from cn_ventas where num_planilla_letra= '" & sNum_Planilla_Letra & "'", cCONNECT) > snumero Then
                MsgBox "Solo puede ingresar hasta " & snumero & " letras por planilla"
                Exit Sub
        Else
                Load frmMovDetalleLetras
                frmMovDetalleLetras.sCOD_BANCO = sCOD_BANCO
                frmMovDetalleLetras.sCuenta = sCuenta
                frmMovDetalleLetras.sFlg_Status_Letras = sFlg_Status_Letras
                frmMovDetalleLetras.sNum_Planilla_Letra = sNum_Planilla_Letra
                frmMovDetalleLetras.Show vbModal
                Set frmMovDetalleLetras = Nothing
                Call buscar
        End If

    Case "ELIMINAR"
            If MsgBox("Esta seguro de Elimar dicha Letra", vbYesNo, "IMPORTANTE") = vbYes Then
                ssql = "VN_DETALLE_LETRAS '$','$','$','$','$'"
                ssql = VBsprintf(ssql, 2, GridEXSelec.Value(GridEXSelec.Columns("NUM_CORRE").Index), sNum_Planilla_Letra, sCuenta, sCOD_BANCO)
                ExecuteCommandSQL cCONNECT, ssql
                Call buscar
            Else
                Unload Me
            End If

    Case "SALIR"
    
    If sCOD_BANCO = 3 Then
        snumero = 10
    ElseIf sCOD_BANCO = 2 Then
        snumero = 15
    Else
        snumero = 10
    End If
    
        If DevuelveCampo("select count(*) from cn_ventas where num_planilla_letra= '" & sNum_Planilla_Letra & "'", cCONNECT) > snumero Then
                MsgBox "Solo puede ingresar hasta " & snumero & " letras por planilla"
                Exit Sub
        Else
        Unload Me
        End If
    End Select
End Sub

Private Sub RefrescaGridOperaciones()
    Dim oGroup As JSGroup

    If GridEXSelec.Groups.Count > 0 Then
        GridEXSelec.Groups.Remove GridEXSelec.Groups.Count
        GridEXSelec.Refresh
    End If
    Set GridEXSelec.ADORecordset = rsGridOperaciones
    GridEXSelec.DefaultGroupMode = jgexDGMExpanded
    GridEXSelec.ExpandAll

    GridEXSelec.Columns("num_corre").Visible = False
    GridEXSelec.Columns("cod_tipdoc").Caption = "Tipo Doc."
    GridEXSelec.Columns("ser_docum").Caption = "Serie"
    GridEXSelec.Columns("num_docum_ventas").Caption = "Num. Doc."
    GridEXSelec.Columns("cod_moneda").Caption = "Moneda"
    GridEXSelec.Columns("imp_total").Caption = "Imp. Total"

    GridEXSelec.ContinuousScroll = True
    GridEXSelec.RefreshGroups
    GridEXSelec.Refresh
End Sub
