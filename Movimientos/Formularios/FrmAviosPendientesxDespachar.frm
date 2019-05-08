VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmAviosPendientesxDespachar 
   Caption         =   "Avios Pendientes por Despachar Acabados"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   13710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Cancelar"
      Height          =   525
      Left            =   6960
      TabIndex        =   2
      Top             =   3360
      Width           =   1245
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   525
      Left            =   5520
      TabIndex        =   1
      Top             =   3360
      Width           =   1245
   End
   Begin GridEX20.GridEX gexLista 
      Height          =   3315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   5847
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      AllowAddNew     =   -1  'True
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   17
      Column(1)       =   "FrmAviosPendientesxDespachar.frx":0000
      Column(2)       =   "FrmAviosPendientesxDespachar.frx":0110
      Column(3)       =   "FrmAviosPendientesxDespachar.frx":0204
      Column(4)       =   "FrmAviosPendientesxDespachar.frx":0300
      Column(5)       =   "FrmAviosPendientesxDespachar.frx":03F4
      Column(6)       =   "FrmAviosPendientesxDespachar.frx":04F0
      Column(7)       =   "FrmAviosPendientesxDespachar.frx":05E4
      Column(8)       =   "FrmAviosPendientesxDespachar.frx":06F0
      Column(9)       =   "FrmAviosPendientesxDespachar.frx":07E4
      Column(10)      =   "FrmAviosPendientesxDespachar.frx":08F0
      Column(11)      =   "FrmAviosPendientesxDespachar.frx":09FC
      Column(12)      =   "FrmAviosPendientesxDespachar.frx":0AF8
      Column(13)      =   "FrmAviosPendientesxDespachar.frx":0BE4
      Column(14)      =   "FrmAviosPendientesxDespachar.frx":0CE4
      Column(15)      =   "FrmAviosPendientesxDespachar.frx":0DF0
      Column(16)      =   "FrmAviosPendientesxDespachar.frx":0F08
      Column(17)      =   "FrmAviosPendientesxDespachar.frx":1004
      FormatStylesCount=   6
      FormatStyle(1)  =   "FrmAviosPendientesxDespachar.frx":1110
      FormatStyle(2)  =   "FrmAviosPendientesxDespachar.frx":1248
      FormatStyle(3)  =   "FrmAviosPendientesxDespachar.frx":12F8
      FormatStyle(4)  =   "FrmAviosPendientesxDespachar.frx":13AC
      FormatStyle(5)  =   "FrmAviosPendientesxDespachar.frx":1484
      FormatStyle(6)  =   "FrmAviosPendientesxDespachar.frx":153C
      ImageCount      =   0
      PrinterProperties=   "FrmAviosPendientesxDespachar.frx":161C
   End
End
Attribute VB_Name = "FrmAviosPendientesxDespachar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrSql As String
Dim Rs_Lista As ADODB.Recordset
Public varCod_Fabrica As String
Public varCod_OrdPro As String
Public varCod_TipMov As String

'Esto es para hacer la grilla unbound
Dim mvaraProducts As Variant
Dim mRecordCount As Long
Dim mRecordsetCols As Long

'Datos para aceptar
Public varCOD_ALMACEN As String
Public varNUM_MOVSTK As String
Public varSec_OrdComp As String

Public Sub CARGA_GRID()
   
   gexLista.AllowAddNew = True
   
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.CursorType = adOpenStatic
   
    Rs_Lista.ActiveConnection = cConnect
   
    'Esta cadena es para devolver el Codigo de Cliente
    StrSql = "EXEC lg_muestra_avios_acabado_por_despachar_NP '" & Me.varCod_Fabrica & "','" & Me.varCod_OrdPro & "'"
    Rs_Lista.Open StrSql
   
    If Rs_Lista.RecordCount > 0 Then
   
        Rs_Lista.MoveLast
        'Get the number of records in recordset
        mRecordCount = Rs_Lista.RecordCount
        Rs_Lista.MoveFirst
        'Retrieve records in a module level variant array used in
        'unbound events
        mvaraProducts = Rs_Lista.GetRows(mRecordCount)
        mRecordsetCols = UBound(mvaraProducts, 1)
        'Set the ItemCount property of the control to the number of records.
        gexLista.ItemCount = mRecordCount
        'BUG: In Unbound mode, the first time the control appears, shows the record selected empty.
        'The Rebind method is used as a turn around for this bug.
        gexLista.ReBind
        
    '    Set gexLista.ADORecordset = CargarRecordSetDesconectado(Strsql, cConnect)

    '    gexLista.Columns.Add "Saldo", jgexText, jgexEditTextBox, "SALDO"
       
    
    End If
    
    SetGeneralGridEX gexLista, 0, 1
    Call Configurar_Grid
        
    If gexLista.RowCount > 0 Then
        gexLista.Enabled = True
        'HabilitaMant Me.FunctButt1, "GENERAR/REVERTIR/IMPRIMIR/SALIR"
    Else
        gexLista.Enabled = False
        'HabilitaMant Me.FunctButt1, "GENERAR/REVERTIR/IMPRIMIR/SALIR"
    End If

gexLista.AllowAddNew = False

End Sub

Public Sub Configurar_Grid()

    Me.gexLista.Columns("Cod_Item").Visible = False
    Me.gexLista.Columns("Cod_Comb").Visible = False
    Me.gexLista.Columns("Cod_Color").Visible = False
    Me.gexLista.Columns("nombre_Medida").Visible = False
        
    Me.gexLista.Columns("CHECK").ColumnType = jgexCheckBox
    Me.gexLista.Columns("CHECK").Caption = "Flag"
    Me.gexLista.Columns("CHECK").Width = 450
    Me.gexLista.Columns("Des_Item").Caption = "Avios"
    Me.gexLista.Columns("Des_Item").Width = 3850
    Me.gexLista.Columns("Nombre_Color").Caption = "Color"
    Me.gexLista.Columns("Nombre_Color").Width = 1900
    Me.gexLista.Columns("MEDIDA").Width = 700
    Me.gexLista.Columns("Cod_Estcli").Caption = "Est. Cliente"
    Me.gexLista.Columns("Cod_Estcli").Width = 1100
    Me.gexLista.Columns("Cod_Destino").Caption = "Destino"
    Me.gexLista.Columns("Cod_Destino").Width = 700
    
    Me.gexLista.Columns("Saldo").Width = 1000
    Me.gexLista.Columns("Requerimiento").Caption = "Cant.Req."
    Me.gexLista.Columns("Requerimiento").Width = 1000
    Me.gexLista.Columns("Enviado_produccion").Caption = "Cant.Entr."
    Me.gexLista.Columns("Enviado_produccion").Width = 1000
   ' Me.gexLista.Columns("Stock_actual").Visible = True

End Sub

Private Sub cmdAceptar_Click()
On Error GoTo hand

If Rs_Lista.RecordCount = 0 Then
    Unload Me
    Exit Sub
End If

Set CadConn = Nothing
CadConn.Open cConnect
Dim j As Integer
gexLista.MoveFirst
For j = 1 To Me.gexLista.RowCount
    If gexLista.Value(gexLista.Columns("CHECK").Index) <> 0 Then
        'If gexLista.Value(gexLista.Columns("FLG_COD_PROV").Index) = "S" And RTrim(gexLista.Value(gexLista.Columns("COD_PROV").Index)) = "" Then
        '  ErrorHandler Err, "COD.PROV ES OBLIGATORIO"
        '   Exit Sub
        'End If
        
        Dim cadena As String
        cadena = " UP_ACTUALIZA_STOCKS_ITEM '" & Me.varCOD_ALMACEN & "','" & Me.varNUM_MOVSTK & "','" & _
                gexLista.Value(gexLista.Columns("Cod_Item").Index) & "','" & _
                gexLista.Value(gexLista.Columns("Cod_Comb").Index) & "','" & _
                gexLista.Value(gexLista.Columns("Cod_Color").Index) & "','" & _
                gexLista.Value(gexLista.Columns("Medida").Index) & "','" & _
                gexLista.Value(gexLista.Columns("Cod_Destino").Index) & "','" & _
                gexLista.Value(gexLista.Columns("Cod_Estcli").Index) & "','',0," & _
                CDbl(gexLista.Value(gexLista.Columns("Saldo").Index)) & ",'I','','','" & vusu & _
                "','','','','','','','','S','',''," & IIf(Trim(gexLista.Value(gexLista.Columns("peso_kgs").Index)) = "", 0, CDbl(gexLista.Value(gexLista.Columns("peso_kgs").Index))) & " ,'','" & varCod_OrdPro & "' "
                
        
        Dim aa As String
        aa = cadena
        
       
        
        
        CadConn.Execute " UP_ACTUALIZA_STOCKS_ITEM '" & Me.varCOD_ALMACEN & "','" & Me.varNUM_MOVSTK & "','" & _
                gexLista.Value(gexLista.Columns("Cod_Item").Index) & "','" & _
                gexLista.Value(gexLista.Columns("Cod_Comb").Index) & "','" & _
                gexLista.Value(gexLista.Columns("Cod_Color").Index) & "','" & _
                gexLista.Value(gexLista.Columns("Medida").Index) & "','" & _
                gexLista.Value(gexLista.Columns("Cod_Destino").Index) & "','" & _
                gexLista.Value(gexLista.Columns("Cod_Estcli").Index) & "','',0," & _
                CDbl(gexLista.Value(gexLista.Columns("Saldo").Index)) & ",'I','','','" & vusu & _
                "','','','','','','','','S','',''," & IIf(Trim(gexLista.Value(gexLista.Columns("peso_kgs").Index)) = "", 0, CDbl(gexLista.Value(gexLista.Columns("peso_kgs").Index))) & " ,'','" & varCod_OrdPro & "' "
                

'" & varCod_OrdPro & "'"

'        CadConn.Execute "UP_ACTUALIZA_STOCKS_ITEM '" & _
'        Me.varCOD_ALMACEN & "','" & _
'        Me.varNUM_MOVSTK & "','" & _
'        gexLista.Value(gexLista.Columns("Cod_Item").Index) & "','" & _
'        gexLista.Value(gexLista.Columns("Cod_Comb").Index) & "','" & _
'        gexLista.Value(gexLista.Columns("Cod_Color").Index) & "','" & _
'        gexLista.Value(gexLista.Columns("Cod_Medida").Index) & "','" & _
'        gexLista.Value(gexLista.Columns("Cod_Destino").Index) & "','" & _
'        gexLista.Value(gexLista.Columns("Cod_Estcli").Index) & "','',0," & _
'        gexLista.Value(gexLista.Columns("Saldo").Index) & ",'" & _
'        "I" & "','" & _
'        "" & "','','" & vusu & "',''"
        
    End If
    gexLista.MoveNext
Next
Set CadConn = Nothing
'FrmDetalleStock.Datos "V", False
FrmDetalleStock.txtPO.Text = ""
FrmDetalleStock.FraPO.Visible = False
FrmDetalleStock.Datos "V", False
Unload Me


Exit Sub
hand:
    ErrorHandler err, "Actualizando"
    Set CadConn = Nothing
'    FrmDetalleStock.Datos "V", False
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

'Private Sub Command1_Click()
'   If UCase(Mid(gexLista.Value(gexLista.Columns("Cod_Item").Index), 1, 2)) = "HI" Then
'        Load frmShowES_OrdProReqItems_Prov
'        frmShowES_OrdProReqItems_Prov.sCod_Almacen = varCOD_ALMACEN
'        Set frmShowES_OrdProReqItems_Prov.oParent = Me
'        frmShowES_OrdProReqItems_Prov.sCod_Item = gexLista.Value(gexLista.Columns("Cod_Item").Index)
'        frmShowES_OrdProReqItems_Prov.sCod_Comb = gexLista.Value(gexLista.Columns("Cod_Comb").Index)
'        frmShowES_OrdProReqItems_Prov.sCod_Color = gexLista.Value(gexLista.Columns("Cod_color").Index)
'        frmShowES_OrdProReqItems_Prov.sCod_Destino = gexLista.Value(gexLista.Columns("Cod_destino").Index)
'        frmShowES_OrdProReqItems_Prov.sCod_Talla = ""
'        frmShowES_OrdProReqItems_Prov.sCod_EstCli = gexLista.Value(gexLista.Columns("Cod_estcli").Index)
'        frmShowES_OrdProReqItems_Prov.BUSCAR
'        frmShowES_OrdProReqItems_Prov.Show vbModal
'        Set frmShowES_OrdProReqItems_Prov = Nothing
'    End If
'End Sub

Private Sub gexLista_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
    If ColIndex = Me.gexLista.Columns("Saldo").ColPosition Or ColIndex = Me.gexLista.Columns("peso_kgs").ColPosition Or ColIndex = Me.gexLista.Columns("CHECK").ColPosition Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub gexLista_Change()
    If Me.gexLista.Col = Me.gexLista.Columns("Saldo").ColPosition Then
        'strSQL = "SELECT Cod_ClaMov FROM LG_TIPOSMOV WHERE Cod_TipMov = '" & Me.varCod_TipMov & "'"
        'If DevuelveCampo(strSQL, cConnect) = "E" Then
        If Val(gexLista.Value(gexLista.Columns("Saldo").Index)) > gexLista.Value(gexLista.Columns("Requerimiento").Index) - gexLista.Value(gexLista.Columns("Enviado_produccion").Index) Then
            gexLista.Value(gexLista.Columns("Saldo").Index) = gexLista.Value(gexLista.Columns("Requerimiento").Index) - gexLista.Value(gexLista.Columns("Enviado_produccion").Index)
            MsgBox "El Saldo excede lo disponible. Sirvase verificar", vbInformation, "Mensaje"
            Exit Sub
        End If
        'End If
    End If
End Sub

Private Sub gexLista_KeyPress(KeyAscii As Integer)
    If Me.gexLista.Col = Me.gexLista.Columns("Saldo").ColPosition Or Me.gexLista.Col = Me.gexLista.Columns("peso_kgs").ColPosition Then
        If KeyAscii = 8 Then
            Exit Sub
        End If
        If KeyAscii = 46 Then
            If InStr(1, gexLista.Value(gexLista.Columns("Saldo").Index), ".") > 0 Then
                KeyAscii = 0
            End If
        Else
            If Not (KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
            If KeyAscii = 39 Or KeyAscii = 13 Then
                KeyAscii = 0
            End If
            
            Dim iPos As Integer
            iPos = InStr(1, gexLista.Value(gexLista.Columns("Saldo").Index), ".")
            
            If iPos = 0 Then
                If Len(gexLista.Value(gexLista.Columns("Saldo").Index)) > 8 Then
                    KeyAscii = 0
                End If
            Else
                If Len(Mid(gexLista.Value(gexLista.Columns("Saldo").Index), iPos + 1, Len(gexLista.Value(gexLista.Columns("Saldo").Index)))) > 3 Then
                    KeyAscii = 0
                End If
            End If
        End If
      
    End If
End Sub

Private Sub gexLista_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Dim i As Integer
    If Rs_Lista.RecordCount > 0 Then
        For i = 1 To Values.ColCount
            Values(i) = mvaraProducts(i - 1, RowIndex - 1)
        Next
    End If
End Sub

Private Sub gexLista_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
Dim i As Integer

On Error GoTo EH_UnboundUpdate

    If Rs_Lista.RecordCount > 0 Then
        'Decrease 1 to row index because the array is zero based
        RowIndex = RowIndex - 1
        For i = 1 To Values.ColCount
            mvaraProducts(i - 1, RowIndex) = Values(i)
        Next
    End If
    
    Exit Sub
    
EH_UnboundUpdate:
    MsgBox err.Description, vbExclamation, "Unbound Sample"
End Sub


