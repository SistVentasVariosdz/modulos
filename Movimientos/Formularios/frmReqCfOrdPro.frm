VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmReqCfOrdPro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requerimientos por Comprar"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   11040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Muestra Detalle"
      Height          =   525
      Left            =   7710
      TabIndex        =   3
      Top             =   3300
      Width           =   1245
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   525
      Left            =   3435
      TabIndex        =   2
      Top             =   3285
      Width           =   1245
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Cancelar"
      Height          =   525
      Left            =   6270
      TabIndex        =   1
      Top             =   3300
      Width           =   1245
   End
   Begin GridEX20.GridEX gexLista 
      Height          =   3075
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   5424
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
      ColumnsCount    =   21
      Column(1)       =   "frmReqCfOrdPro.frx":0000
      Column(2)       =   "frmReqCfOrdPro.frx":0110
      Column(3)       =   "frmReqCfOrdPro.frx":0230
      Column(4)       =   "frmReqCfOrdPro.frx":0350
      Column(5)       =   "frmReqCfOrdPro.frx":0470
      Column(6)       =   "frmReqCfOrdPro.frx":0588
      Column(7)       =   "frmReqCfOrdPro.frx":0674
      Column(8)       =   "frmReqCfOrdPro.frx":076C
      Column(9)       =   "frmReqCfOrdPro.frx":0884
      Column(10)      =   "frmReqCfOrdPro.frx":0988
      Column(11)      =   "frmReqCfOrdPro.frx":0AA0
      Column(12)      =   "frmReqCfOrdPro.frx":0B8C
      Column(13)      =   "frmReqCfOrdPro.frx":0CAC
      Column(14)      =   "frmReqCfOrdPro.frx":0DA0
      Column(15)      =   "frmReqCfOrdPro.frx":0EA4
      Column(16)      =   "frmReqCfOrdPro.frx":0FA0
      Column(17)      =   "frmReqCfOrdPro.frx":10C4
      Column(18)      =   "frmReqCfOrdPro.frx":11B0
      Column(19)      =   "frmReqCfOrdPro.frx":12C0
      Column(20)      =   "frmReqCfOrdPro.frx":13D0
      Column(21)      =   "frmReqCfOrdPro.frx":14FC
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmReqCfOrdPro.frx":1638
      FormatStyle(2)  =   "frmReqCfOrdPro.frx":1770
      FormatStyle(3)  =   "frmReqCfOrdPro.frx":1820
      FormatStyle(4)  =   "frmReqCfOrdPro.frx":18D4
      FormatStyle(5)  =   "frmReqCfOrdPro.frx":19AC
      FormatStyle(6)  =   "frmReqCfOrdPro.frx":1A64
      ImageCount      =   0
      PrinterProperties=   "frmReqCfOrdPro.frx":1B44
   End
End
Attribute VB_Name = "frmReqCfOrdPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim Rs_Lista As ADODB.Recordset
Public varCod_Fabrica As String
Public varCod_OrdPro As String
Public varNum_SecOrd  As String
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
    strSQL = "EXEC UP_SEL_CFORDPRO_AVIOSREQ '" & Me.varCod_Fabrica & "','" & Me.varCod_OrdPro & "','" & Me.varNum_SecOrd & "','" & Me.varCod_TipMov & "'"
    Rs_Lista.Open strSQL
   
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

    Me.gexLista.Columns("Cod_Fabrica").Visible = False
    Me.gexLista.Columns("Cod_OrdPro").Visible = False
    Me.gexLista.Columns("Num_SecOrd").Visible = False
    Me.gexLista.Columns("Cod_Item").Visible = False
    Me.gexLista.Columns("Cod_Comb").Visible = False
    Me.gexLista.Columns("Cod_Color").Visible = False
    Me.gexLista.Columns("Cod_Medida").Visible = False
    
    
    Me.gexLista.Columns("CHECK").ColumnType = jgexCheckBox
    Me.gexLista.Columns("CHECK").Caption = "Flag"
    Me.gexLista.Columns("CHECK").Width = 500
    Me.gexLista.Columns("ITEM").Caption = "Avios"
    Me.gexLista.Columns("ITEM").Width = 2500
    Me.gexLista.Columns("Cod_UniMed").Caption = "U.M."
    Me.gexLista.Columns("Cod_UniMed").Width = 400
    Me.gexLista.Columns("COMBINACION").Caption = "Combinación"
    Me.gexLista.Columns("COMBINACION").Width = 1500
    Me.gexLista.Columns("COLOR").Caption = "Color"
    Me.gexLista.Columns("COLOR").Width = 1400
    Me.gexLista.Columns("MEDIDA").Caption = "Medida"
    Me.gexLista.Columns("MEDIDA").Width = 800
    Me.gexLista.Columns("Cod_Estcli").Caption = "Est. Cliente"
    Me.gexLista.Columns("Cod_Estcli").Width = 1000
    Me.gexLista.Columns("Cod_Destino").Caption = "Destino"
    Me.gexLista.Columns("Cod_Destino").Width = 800
    Me.gexLista.Columns("Cod_Prov").Caption = "Cod.Prov."
    Me.gexLista.Columns("Cod_Prov").Width = 800
    Me.gexLista.Columns("Can_Requerida").Caption = "Cant. Req"
    Me.gexLista.Columns("Can_Requerida").Width = 1100
    Me.gexLista.Columns("Can_Entregada").Caption = "Cant. Ent"
    Me.gexLista.Columns("Can_Entregada").Width = 1100

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
        
        CadConn.Execute "UP_ACTUALIZA_STOCKS_ITEM '" & _
        Me.varCOD_ALMACEN & "','" & _
        Me.varNUM_MOVSTK & "','" & _
        gexLista.Value(gexLista.Columns("Cod_Item").Index) & "','" & _
        gexLista.Value(gexLista.Columns("Cod_Comb").Index) & "','" & _
        gexLista.Value(gexLista.Columns("Cod_Color").Index) & "','" & _
        gexLista.Value(gexLista.Columns("Cod_Medida").Index) & "','" & _
        gexLista.Value(gexLista.Columns("Cod_Destino").Index) & "','" & _
        gexLista.Value(gexLista.Columns("Cod_Estcli").Index) & "','',0," & _
        gexLista.Value(gexLista.Columns("Saldo").Index) & ",'" & _
        "I" & "','" & _
        "" & "','" & _
        gexLista.Value(gexLista.Columns("Num_SecOrd").Index) & "','" & vusu & "','" & gexLista.Value(gexLista.Columns("Cod_Prov").Index) & "',''" & _
        ",'','','','','','S','',''," & IIf(Trim(gexLista.Value(gexLista.Columns("peso_kgs").Index)) = "", 0, CDbl(gexLista.Value(gexLista.Columns("peso_kgs").Index)))
        
    End If
    gexLista.MoveNext
Next
Set CadConn = Nothing
'FrmDetalleStock.Datos "V", False
Unload Me
Exit Sub
hand:
    ErrorHandler err, "Actualizando"
    Set CadConn = Nothing
'    FrmDetalleStock.Datos "V", False
    Unload Me
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
   If UCase(Mid(gexLista.Value(gexLista.Columns("Cod_Item").Index), 1, 2)) = "HI" Then
        Load frmShowES_OrdProReqItems_Prov
        frmShowES_OrdProReqItems_Prov.sCod_Almacen = varCOD_ALMACEN
        Set frmShowES_OrdProReqItems_Prov.oParent = Me
        frmShowES_OrdProReqItems_Prov.sCod_Item = gexLista.Value(gexLista.Columns("Cod_Item").Index)
        frmShowES_OrdProReqItems_Prov.sCod_Comb = gexLista.Value(gexLista.Columns("Cod_Comb").Index)
        frmShowES_OrdProReqItems_Prov.sCod_Color = gexLista.Value(gexLista.Columns("Cod_color").Index)
        frmShowES_OrdProReqItems_Prov.sCod_destino = gexLista.Value(gexLista.Columns("Cod_destino").Index)
        frmShowES_OrdProReqItems_Prov.sCod_Talla = ""
        frmShowES_OrdProReqItems_Prov.sCod_EstCli = gexLista.Value(gexLista.Columns("Cod_estcli").Index)
        frmShowES_OrdProReqItems_Prov.BUSCAR
        frmShowES_OrdProReqItems_Prov.Show vbModal
        Set frmShowES_OrdProReqItems_Prov = Nothing
    End If
End Sub

Private Sub gexLista_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
    If ColIndex = Me.gexLista.Columns("Saldo").ColPosition Or ColIndex = Me.gexLista.Columns("Peso_kgs").ColPosition Or ColIndex = Me.gexLista.Columns("CHECK").ColPosition Or ColIndex = Me.gexLista.Columns("cod_prov").ColPosition Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub gexLista_Change()
    If Me.gexLista.Col = Me.gexLista.Columns("Saldo").ColPosition Then
        strSQL = "SELECT Cod_ClaMov FROM LG_TIPOSMOV WHERE Cod_TipMov = '" & Me.varCod_TipMov & "'"
        If DevuelveCampo(strSQL, cConnect) = "E" Then
            If Val(gexLista.Value(gexLista.Columns("Saldo").Index)) > gexLista.Value(gexLista.Columns("Can_Entregada").Index) Then
                gexLista.Value(gexLista.Columns("Saldo").Index) = gexLista.Value(gexLista.Columns("Can_Entregada").Index)
                MsgBox "El Saldo no puede exceder la cantidad atendida. Sirvase verificar", vbInformation, "Mensaje"
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub gexLista_KeyPress(KeyAscii As Integer)
    If Me.gexLista.Col = Me.gexLista.Columns("Saldo").ColPosition Or Me.gexLista.Col = Me.gexLista.Columns("Peso_Kgs").ColPosition Then
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
