VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form frmShowCN_DocumVoucher 
   Caption         =   "Voucher Contable"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   12450
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDebe 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   8445
      TabIndex        =   5
      Text            =   "0"
      Top             =   5010
      Width           =   1230
   End
   Begin VB.TextBox txtHaber 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9720
      TabIndex        =   4
      Text            =   "0"
      Top             =   5010
      Width           =   1230
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      Height          =   315
      Left            =   8460
      TabIndex        =   3
      Text            =   "Total Debe"
      Top             =   4590
      Width           =   1230
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      Height          =   315
      Left            =   9720
      TabIndex        =   2
      Text            =   "Total Haber"
      Top             =   4590
      Width           =   1230
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4410
      Left            =   90
      TabIndex        =   1
      Top             =   75
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   7779
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmShowCN_DocumVoucher.frx":0000
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   4
      Column(1)       =   "frmShowCN_DocumVoucher.frx":0352
      Column(2)       =   "frmShowCN_DocumVoucher.frx":041A
      Column(3)       =   "frmShowCN_DocumVoucher.frx":04BE
      Column(4)       =   "frmShowCN_DocumVoucher.frx":0562
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmShowCN_DocumVoucher.frx":063E
      FormatStyle(2)  =   "frmShowCN_DocumVoucher.frx":0776
      FormatStyle(3)  =   "frmShowCN_DocumVoucher.frx":0826
      FormatStyle(4)  =   "frmShowCN_DocumVoucher.frx":08DA
      FormatStyle(5)  =   "frmShowCN_DocumVoucher.frx":09B2
      FormatStyle(6)  =   "frmShowCN_DocumVoucher.frx":0A6A
      FormatStyle(7)  =   "frmShowCN_DocumVoucher.frx":0B4A
      FormatStyle(8)  =   "frmShowCN_DocumVoucher.frx":1002
      ImageCount      =   1
      ImagePicture(1) =   "frmShowCN_DocumVoucher.frx":144E
      PrinterProperties=   "frmShowCN_DocumVoucher.frx":17A0
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   3570
      Left            =   11070
      TabIndex        =   0
      Top             =   45
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   6297
      Custom          =   $"frmShowCN_DocumVoucher.frx":1978
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1200
      ControlHeigth   =   550
      ControlSeparator=   50
   End
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   570
      Left            =   105
      TabIndex        =   6
      Top             =   4695
      Visible         =   0   'False
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   1005
      Custom          =   $"frmShowCN_DocumVoucher.frx":1B51
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1200
      ControlHeigth   =   550
      ControlSeparator=   50
   End
End
Attribute VB_Name = "frmShowCN_DocumVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Public sNum_Corre As String
'Public sFlg_TipMondoc As String
'Public oParent As Object
'Public dTipoCambio As Double
'Public sTipAnexo As String
'Public sCod_Anexo As String
'Public sSubdiario As String
'Public sAno_Registro As String
'Public sMes_Registro As String
'Public sNum_Movimiento As String
'Public sTipOpcion As String
'
'Public Function Buscar() As Boolean
'On Error GoTo errores
'Dim sSql As String
'Dim vBookmark As Variant
'sSql = "SM_VOUCHER_CONTABLE '$' ,'$' , '$','$','$','$'"
'sSql = VBsprintf(sSql, sNum_Corre, sTipOpcion, sSubdiario, sAno_Registro, sMes_Registro, sNum_Movimiento)
'
'vBookmark = GridEX1.Row
'GridEX1.ClearFields
'
'Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSql, cCONNECT)
'
'GridEX1.Row = vBookmark
'
'If GridEX1.RowCount > 0 Then
'    txtDebe = Format(GridEX1.Value(GridEX1.Columns("TOTAL_DEBE").Index), "###,##0.00")
'    txtHaber = Format(GridEX1.Value(GridEX1.Columns("TOTAL_HABER").Index), "###,##0.00")
'End If
'
'GridEX1.Columns("IMPORTE").Format = "###,##0.00"
'
'GridEX1.ContinuousScroll = True
'
'GridEX1.FrozenColumns = 2
'Exit Function
'
'errores:
'    errores Err.Number
'End Function
'
'Private Sub Form_Load()
'    sTipOpcion = "1"
'End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If txtDebe <> txtHaber Then
'        Cancel = 1
'    End If
'End Sub
'
'Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
'    Select Case ActionName
'        Case "ADICIONAR"
'            Load frmCN_DocumVoucher
'            frmCN_DocumVoucher.sNum_Corre = sNum_Corre
'            frmCN_DocumVoucher.sFlg_TipMondoc = sFlg_TipMondoc
'            frmCN_DocumVoucher.txtTipodeCambio = dTipoCambio
'            If GridEX1.RowCount > 0 Then
'              frmCN_DocumVoucher.TxtTipoDocSunat = GridEX1.Value(GridEX1.Columns("TIPO").Index)
'              frmCN_DocumVoucher.txtSerie = GridEX1.Value(GridEX1.Columns("SERIE").Index)
'              frmCN_DocumVoucher.txtNum_Docum = GridEX1.Value(GridEX1.Columns("NUMERO").Index)
'            End If
'            Set frmCN_DocumVoucher.oParent = Me
'            frmCN_DocumVoucher.saccion = "I"
'            frmCN_DocumVoucher.Show vbModal
'            Set frmCN_DocumVoucher = Nothing
'        Case "MODIFICAR"
'            If GridEX1.RowCount = 0 Then Exit Sub
'            Load frmCN_DocumVoucher
'            frmCN_DocumVoucher.sNum_Corre = sNum_Corre
'            Set frmCN_DocumVoucher.oParent = Me
'            frmCN_DocumVoucher.saccion = "U"
'            frmCN_DocumVoucher.txtCuenta.Text = GridEX1.Value(GridEX1.Columns("CUENTA").Index)
'            frmCN_DocumVoucher.txtDescripcion = GridEX1.Value(GridEX1.Columns("DESCRIPCION").Index)
'            frmCN_DocumVoucher.txtRuc.Text = GridEX1.Value(GridEX1.Columns("ANALISIS").Index)
'            frmCN_DocumVoucher.sItem = GridEX1.Value(GridEX1.Columns("Item").Index)
'            frmCN_DocumVoucher.TxtTipoDocSunat = GridEX1.Value(GridEX1.Columns("TIPO").Index)
'            frmCN_DocumVoucher.txtSerie = GridEX1.Value(GridEX1.Columns("SERIE").Index)
'            frmCN_DocumVoucher.txtNum_Docum = GridEX1.Value(GridEX1.Columns("NUMERO").Index)
'            'frmCN_DocumVoucher.txtCuenta.Enabled = False
'            'frmCN_DocumVoucher.txtDescripcion.Enabled = False
'            frmCN_DocumVoucher.sFlg_TipMondoc = sFlg_TipMondoc
'            frmCN_DocumVoucher.txtTipodeCambio = GridEX1.Value(GridEX1.Columns("TIPCAM").Index)
'            If GridEX1.Value(GridEX1.Columns("FLG_DEBE_HABER").Index) = "D" Then
'                frmCN_DocumVoucher.txtDebe = GridEX1.Value(GridEX1.Columns("IMPORTE").Index)
'                frmCN_DocumVoucher.txtDebeDol = GridEX1.Value(GridEX1.Columns("DOLARES").Index)
'            Else
'                frmCN_DocumVoucher.txtHaber = GridEX1.Value(GridEX1.Columns("IMPORTE").Index)
'                frmCN_DocumVoucher.txtHaberDol = GridEX1.Value(GridEX1.Columns("DOLARES").Index)
'            End If
'
'            If frmCN_DocumVoucher.TxtTipoDocSunat.Text >= "50" And frmCN_DocumVoucher.TxtTipoDocSunat.Text <= "54" Then
'                frmCN_DocumVoucher.txtTipodeCambio.Enabled = True
'            Else
'                frmCN_DocumVoucher.txtTipodeCambio.Enabled = False
'            End If
'
'            frmCN_DocumVoucher.Show vbModal
'            Set frmCN_DocumVoucher = Nothing
'        Case "ELIMINAR"
'            If GridEX1.RowCount = 0 Then Exit Sub
'            Load frmCN_DocumVoucher
'            frmCN_DocumVoucher.sNum_Corre = sNum_Corre
'            Set frmCN_DocumVoucher.oParent = Me
'            frmCN_DocumVoucher.saccion = "D"
'            frmCN_DocumVoucher.txtCuenta.Text = GridEX1.Value(GridEX1.Columns("CUENTA").Index)
'            frmCN_DocumVoucher.txtDescripcion = GridEX1.Value(GridEX1.Columns("DESCRIPCION").Index)
'            frmCN_DocumVoucher.sFlg_TipMondoc = sFlg_TipMondoc
'            frmCN_DocumVoucher.txtTipodeCambio = dTipoCambio
'            frmCN_DocumVoucher.sItem = GridEX1.Value(GridEX1.Columns("Item").Index)
'            frmCN_DocumVoucher.TxtTipoDocSunat = GridEX1.Value(GridEX1.Columns("TIPO").Index)
'            frmCN_DocumVoucher.txtSerie = GridEX1.Value(GridEX1.Columns("SERIE").Index)
'            frmCN_DocumVoucher.txtNum_Docum = GridEX1.Value(GridEX1.Columns("NUMERO").Index)
'            If GridEX1.Value(GridEX1.Columns("FLG_DEBE_HABER").Index) = "D" Then
'                frmCN_DocumVoucher.txtDebe = GridEX1.Value(GridEX1.Columns("IMPORTE").Index)
'            Else
'                frmCN_DocumVoucher.txtHaber = GridEX1.Value(GridEX1.Columns("IMPORTE").Index)
'            End If
'            frmCN_DocumVoucher.fraDatos.Enabled = False
'            frmCN_DocumVoucher.Show vbModal
'            Set frmCN_DocumVoucher = Nothing
'        Case "IMPRIMIR"
'            Call Imprimir
'        Case "SALIR"
'            Unload Me
'    End Select
'End Sub
'
'Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
'    Select Case ActionName
'        Case "MODIFICARCUADRAR"
'            If GridEX1.RowCount = 0 Then Exit Sub
'            Load frmCN_DocumVoucher
'            frmCN_DocumVoucher.sNum_Corre = sNum_Corre
'            Set frmCN_DocumVoucher.oParent = Me
'            frmCN_DocumVoucher.saccion = "U"
'            frmCN_DocumVoucher.txtCuenta.Text = GridEX1.Value(GridEX1.Columns("CUENTA").Index)
'            frmCN_DocumVoucher.txtDescripcion = GridEX1.Value(GridEX1.Columns("DESCRIPCION").Index)
'            frmCN_DocumVoucher.txtRuc.Text = GridEX1.Value(GridEX1.Columns("ANALISIS").Index)
'            frmCN_DocumVoucher.sItem = GridEX1.Value(GridEX1.Columns("Item").Index)
'            frmCN_DocumVoucher.TxtTipoDocSunat = GridEX1.Value(GridEX1.Columns("TIPO").Index)
'            frmCN_DocumVoucher.txtSerie = GridEX1.Value(GridEX1.Columns("SERIE").Index)
'            frmCN_DocumVoucher.txtNum_Docum = GridEX1.Value(GridEX1.Columns("NUMERO").Index)
'            frmCN_DocumVoucher.sFlg_TipMondoc = sFlg_TipMondoc
'            frmCN_DocumVoucher.txtTipodeCambio = GridEX1.Value(GridEX1.Columns("TIPCAM").Index)
'            If GridEX1.Value(GridEX1.Columns("FLG_DEBE_HABER").Index) = "D" Then
'                frmCN_DocumVoucher.txtDebe = GridEX1.Value(GridEX1.Columns("IMPORTE").Index)
'                frmCN_DocumVoucher.txtDebeDol = GridEX1.Value(GridEX1.Columns("DOLARES").Index)
'            Else
'                frmCN_DocumVoucher.txtHaber = GridEX1.Value(GridEX1.Columns("IMPORTE").Index)
'                frmCN_DocumVoucher.txtHaberDol = GridEX1.Value(GridEX1.Columns("DOLARES").Index)
'            End If
'
'            If frmCN_DocumVoucher.TxtTipoDocSunat.Text >= "50" And frmCN_DocumVoucher.TxtTipoDocSunat.Text <= "54" Then
'                frmCN_DocumVoucher.txtTipodeCambio.Enabled = True
'            Else
'                frmCN_DocumVoucher.txtTipodeCambio.Enabled = False
'            End If
'
'            frmCN_DocumVoucher.txtCuenta.Enabled = False
'            frmCN_DocumVoucher.txtDescripcion.Enabled = False
'            frmCN_DocumVoucher.txtRuc.Enabled = False
'
'            frmCN_DocumVoucher.TxtTipoDocSunat.Enabled = False
'            frmCN_DocumVoucher.txtSerie.Enabled = False
'            frmCN_DocumVoucher.txtNum_Docum.Enabled = False
'
'            'frmCN_DocumVoucher.txtTipodeCambio.Enabled = True
'            'frmCN_DocumVoucher.txtDebe.Enabled = False
'            'frmCN_DocumVoucher.txtDebeDol.Enabled = False
'            'frmCN_DocumVoucher.txtHaber.Enabled = False
'            'frmCN_DocumVoucher.txtHaberDol.Enabled = False
'            frmCN_DocumVoucher.TxtTipoDocSunat.Enabled = False
'            frmCN_DocumVoucher.txtCod_Anexo.Enabled = False
'            frmCN_DocumVoucher.txtCod_TipAne.Enabled = False
'            frmCN_DocumVoucher.txtDes_Anexo.Enabled = False
'            frmCN_DocumVoucher.txtDescripcion.Enabled = False
'            frmCN_DocumVoucher.txtNum_Docum.Enabled = False
'            frmCN_DocumVoucher.txtNum_Ruc.Enabled = False
'            frmCN_DocumVoucher.txtRuc.Enabled = False
'            frmCN_DocumVoucher.txtSerie.Enabled = False
'
'
'
'            frmCN_DocumVoucher.Show vbModal
'            Set frmCN_DocumVoucher = Nothing
'        Case "SALIR"
'            Unload Me
'    End Select
'End Sub
'
'Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
'    Cancel = True
'End Sub
'
'Sub Imprimir()
'On Error GoTo hand
'Dim oo As Object
'Dim Ruta As String
'Dim Cadena1, Cadena2 As String
'
'Cadena1 = "EXEC CN_Muestra_Cabecera_Vocuher_Contabilidad '" & sNum_Corre & "'"
'Cadena2 = "EXEC CN_Muestra_Detalle_Vocuher_Contabilidad '" & sNum_Corre & "'"
'
'    If DevuelveCampo("select cod_tipdoc from cn_docum where num_corre = '" & sNum_Corre & "'", cCONNECT) <> "AT" Then
'        Ruta = vRuta & "\VoucherContabilidad.XLT"
'    Else
'        Ruta = vRuta & "\VoucherAnticipos.XLT"
'    End If
'    Set oo = CreateObject("excel.application")
'    oo.Workbooks.Open Ruta
'    oo.Visible = True
'    oo.DisplayAlerts = False
'    oo.Run "reporte", Cadena1, Cadena2, cCONNECT
'    Set oo = Nothing
'Exit Sub
'hand:
'    ErrorHandler Err, "GeneraReportes"
'    Set oo = Nothing
'End Sub
