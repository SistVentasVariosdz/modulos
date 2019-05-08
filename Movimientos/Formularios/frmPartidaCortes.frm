VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form frmPartidaCortes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cortes Partidas"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPanos 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6195
      TabIndex        =   5
      Text            =   "0"
      Top             =   4500
      Width           =   1185
   End
   Begin VB.TextBox txtBultos 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6195
      TabIndex        =   4
      Text            =   "0"
      Top             =   4155
      Width           =   1185
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2490
      TabIndex        =   6
      Top             =   4365
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmPartidaCortes.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX gexOrdenes 
      Height          =   3135
      Left            =   60
      TabIndex        =   3
      Top             =   645
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5530
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ContScroll      =   -1  'True
      AllowColumnDrag =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorBkg    =   -2147483624
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmPartidaCortes.frx":0096
      Column(2)       =   "frmPartidaCortes.frx":015E
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmPartidaCortes.frx":0202
      FormatStyle(2)  =   "frmPartidaCortes.frx":033A
      FormatStyle(3)  =   "frmPartidaCortes.frx":03EA
      FormatStyle(4)  =   "frmPartidaCortes.frx":049E
      FormatStyle(5)  =   "frmPartidaCortes.frx":0576
      FormatStyle(6)  =   "frmPartidaCortes.frx":062E
      ImageCount      =   0
      PrinterProperties=   "frmPartidaCortes.frx":070E
   End
   Begin FunctionsButtons.FunctButt fnbBuscar 
      Height          =   495
      Left            =   6135
      TabIndex        =   2
      Top             =   15
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.TextBox txtCod_OrdProv 
      Height          =   345
      Left            =   1335
      TabIndex        =   1
      Top             =   135
      Width           =   1875
   End
   Begin VB.Label Label4 
      Caption         =   "Total Paños :"
      Height          =   225
      Left            =   5145
      TabIndex        =   10
      Top             =   4545
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "Total Bultos :"
      Height          =   225
      Left            =   5145
      TabIndex        =   9
      Top             =   4200
      Width           =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "Total Kilos :"
      Height          =   225
      Left            =   5145
      TabIndex        =   8
      Top             =   3870
      Width           =   885
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   285
      Left            =   6195
      TabIndex        =   7
      Top             =   3840
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Partida"
      Height          =   225
      Left            =   390
      TabIndex        =   0
      Top             =   195
      Width           =   885
   End
End
Attribute VB_Name = "frmPartidaCortes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_TipMov As String, sCod_Almacen As String, sNum_MovStk As String, _
       Codigo As String, Descripcion As String, Paso As Boolean
Dim rstAux As ADODB.Recordset, strSQL As String, sTit As String, sErr As String

Private Sub fnbBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo ErrBusq
Dim iCol As Integer
    
    sTit = "Buscar Ordenes de Corte x Partida"
    
    Screen.MousePointer = 11
    
    strSQL = "EXEC TX_SM_CO_ORDPRO_ORDTRA '" & txtCod_OrdProv & "', '" & sCod_TipMov & "'"
    Set gexOrdenes.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    
    gexOrdenes.Columns("Sel").Width = 360
    gexOrdenes.Columns("Co_CodOrdPro").Width = 570
    gexOrdenes.Columns("Cod_Tela").Width = 825
    gexOrdenes.Columns("Des_Tela").Width = 1500
    gexOrdenes.Columns("Cod_Comb").Width = 105
    gexOrdenes.Columns("Des_Comb").Width = 1140
    gexOrdenes.Columns("Cod_Color").Width = 150
    gexOrdenes.Columns("Des_Color").Width = 1110
    gexOrdenes.Columns("Cod_Medida").Width = 345
    gexOrdenes.Columns("Cod_Calidad").Width = 120
    gexOrdenes.Columns("Stock").Width = 690
    gexOrdenes.Columns("Kilos").Width = 690
    
    gexOrdenes.Columns("Sel").Caption = "Sel"
    gexOrdenes.Columns("Co_CodOrdPro").Caption = "O/Corte"
    gexOrdenes.Columns("Cod_Tela").Caption = "C.Tela"
    gexOrdenes.Columns("Des_Tela").Caption = "Desc.Tela"
    gexOrdenes.Columns("Cod_Comb").Visible = False
    gexOrdenes.Columns("Des_Comb").Caption = "Combinacion"
    gexOrdenes.Columns("Cod_Color").Visible = False
    gexOrdenes.Columns("Des_Color").Caption = "Color"
    gexOrdenes.Columns("Cod_Medida").Caption = "Med."
    gexOrdenes.Columns("Cod_Calidad").Visible = False
    gexOrdenes.Columns("Stock").Caption = "Stock"
    gexOrdenes.Columns("Kilos").Caption = "Movim."
    gexOrdenes.Columns("Cod_TipOrdTra").Visible = False
    gexOrdenes.Columns("Cod_OrdTra").Visible = False
    
    gexOrdenes.Columns("Stock").Format = "0.00"
    gexOrdenes.Columns("Kilos").Format = "0.00"
    
    gexOrdenes.Columns("Sel").ColumnType = jgexCheckBox
    
    TotalSel
    
    For iCol = 1 To gexOrdenes.Columns.Count
        If gexOrdenes.Columns(iCol).Key <> "Sel" And _
        gexOrdenes.Columns(iCol).Key <> "Kilos" Then
            gexOrdenes.Columns(iCol).EditType = jgexEditNone
        End If
    Next iCol
    
    If gexOrdenes.RowCount > 0 Then SendKeys "{TAB}"
    
    Screen.MousePointer = 0
Exit Sub
ErrBusq:
    sErr = Err.Description
    MsgBox sErr, vbCritical + vbOKOnly, sTit
End Sub

Private Sub fnbBuscar_GotFocus()
    fnbBuscar_ActionClick 0, 0, ""
End Sub

Private Sub Form_Load()
    Paso = False
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "ACEPTAR"
        GuardarOrdenes
    Case "CANCELAR"
        Unload Me
    End Select
    
End Sub

Private Sub GuardarOrdenes()
On Error GoTo ErrGuardar
Dim cntAux As New ADODB.Connection, dPanos As Double, dBultos As Double, _
    bPrimero As Boolean
    
    sTit = "Guardar Movimiento"
    
    If Not IsNumeric(txtBultos) Then
        MsgBox "Nro. de Bultos Invalido", vbExclamation + vbOKOnly, sTit
        txtBultos.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtPanos) Then
        MsgBox "Nro. de Paños Invalido", vbExclamation + vbOKOnly, sTit
        txtPanos.SetFocus
        Exit Sub
    End If
    
    If CDbl(txtBultos) < 0 Then
        MsgBox "Nro. de Bultos debe ser mayor o igual que 0", vbExclamation + vbOKOnly, sTit
        txtBultos.SetFocus
        Exit Sub
    End If
    
    If CDbl(txtPanos) < 0 Then
        MsgBox "Nro. de Paños debe ser mayor o igual que 0", vbExclamation + vbOKOnly, sTit
        txtPanos.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    bPrimero = True
    
    cntAux.Open cConnect
    cntAux.BeginTrans
    gexOrdenes.Update
    Set rstAux = gexOrdenes.ADORecordset.Clone
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
            If CBool(!Sel) Then
                
                If bPrimero Then
                    dPanos = txtPanos
                    dBultos = txtBultos
                    bPrimero = False
                Else
                    dPanos = 0
                    dBultos = 0
                End If
                
                strSQL = "EXEC UP_ACT_STOCKSTELCOR 'I', '" & sCod_Almacen & _
                "', '" & sNum_MovStk & "', '', '" & !Co_CodOrdPro & "', '" & _
                !Cod_TipOrdTra & "', '" & !Cod_OrdTra & "', '" & !Cod_Tela & "', '" & _
                !Cod_Comb & "', '" & !Cod_color & "', '" & !Cod_Medida & "', '" & _
                !Cod_Calidad & "', " & !Kilos & ", " & dBultos & ", " & dPanos & _
                ", '', '" & vusu & "'"
                
                cntAux.Execute strSQL, adExecuteNoRecords
            End If
            .MoveNext
        Loop
    End With
    cntAux.CommitTrans: Set cntAux = Nothing
    rstAux.Close: Set rstAux = Nothing
    Screen.MousePointer = 0
    
    Unload Me
Exit Sub
ErrGuardar:
    sErr = Err.Description
    cntAux.RollbackTrans: Set cntAux = Nothing
    rstAux.Close: Set rstAux = Nothing
    Screen.MousePointer = 0
    MsgBox sErr, vbCritical + vbOKOnly, "Guardar Mov. Ordenes"
End Sub

Private Sub gexOrdenes_AfterUpdate()
    TotalSel
End Sub

Private Sub gexOrdenes_BeforeColUpdate(ByVal Row As Long, ByVal ColIndex As Integer, ByVal OldValue As String, ByVal Cancel As GridEX20.JSRetBoolean)
    If gexOrdenes.Row <= 0 Then Exit Sub
    'If Paso Then Exit Sub
    Paso = True
    
    If ColIndex = gexOrdenes.Columns("Sel").Index Then
        'gexOrdenes.Value(gexOrdenes.Columns("Kilos").Index) <= 0 And
        If CDbl(OldValue) <> gexOrdenes.Value(ColIndex) Then
            If CBool(gexOrdenes.Value(ColIndex)) Then
                gexOrdenes.Value(gexOrdenes.Columns("Kilos").Index) = gexOrdenes.Value(gexOrdenes.Columns("Stock").Index)
            Else
                gexOrdenes.Value(gexOrdenes.Columns("Kilos").Index) = 0
            End If
'        Else
'            gexOrdenes.Value(gexOrdenes.Columns("Kilos").Index) = 0
        End If
    ElseIf ColIndex = gexOrdenes.Columns("Kilos").Index Then
        If CDbl(gexOrdenes.Value(ColIndex)) <= gexOrdenes.Value(gexOrdenes.Columns("Stock").Index) Then
            If CDbl(gexOrdenes.Value(ColIndex)) <= 0 Then
                gexOrdenes.Value(gexOrdenes.Columns("Sel").Index) = 0
            Else
                gexOrdenes.Value(gexOrdenes.Columns("Sel").Index) = -1
            End If
        Else
            gexOrdenes.Value(ColIndex) = 0
        End If
    End If
    
    Paso = False
    
End Sub

Private Sub TxtBultos_GotFocus()
    SelectionText txtBultos
End Sub

Private Sub txtBultos_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCod_OrdProv_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaPartidaOrden
        SendKeys "{TAB}"
    End If
End Sub

Private Sub BuscaPartidaOrden()
On Error GoTo ErrPartidaLote
    
    txtCod_OrdProv = Trim(txtCod_OrdProv)
    
    strSQL = "SELECT d.Cod_OrdProv, 0 AS Aux FROM LG_MOVISTK a, LG_ORDCOMPITEM b, " & _
    "CO_ORDPRO_TELAS c, TX_ORDTRA d WHERE a.Cod_Almacen = '" & _
    sCod_Almacen & "' AND   a.Num_MovStk = '" & sNum_MovStk & "' " & _
    "AND   a.Ser_OrdComp = b.Ser_OrdComp AND   a.Cod_OrdComp = b.Cod_OrdComp " & _
    "AND   b.Cod_Item = c.Cod_Tela AND   b.Cod_Comb = c.Cod_Comb " & _
    "AND   b.Cod_Color = c.Cod_Color AND   c.Cod_TipOrdTra = d.Cod_TipOrdTra " & _
    "AND   c.Cod_OrdTra = d.Cod_OrdTra AND   d.Cod_OrdProv LIKE '%" & _
    txtCod_OrdProv & "%' " & "GROUP BY d.Cod_OrdProv ORDER BY d.Cod_OrdProv"
    
    Set frmBusqGeneral.oParent = Me
    frmBusqGeneral.sQuery = strSQL
    frmBusqGeneral.CARGAR_DATOS
    frmBusqGeneral.gexList.Columns("Cod_OrdProv").Width = 3500
    frmBusqGeneral.gexList.Columns("Cod_OrdProv").Caption = "Partida / Lote"
    frmBusqGeneral.gexList.Columns("Aux").Visible = False
    
    Codigo = ""
    Descripcion = ""
    
    If frmBusqGeneral.gexList.RowCount > 1 Then
        frmBusqGeneral.Show 1
    ElseIf frmBusqGeneral.gexList.RowCount = 1 Then
        frmBusqGeneral.cmdAceptar_Click
    Else
        frmBusqGeneral.cmdCancelar_Click
    End If
    txtCod_OrdProv = Codigo
    
Exit Sub
ErrPartidaLote:
    sErr = Err.Description
    MsgBox sErr, vbCritical + vbOKOnly, "Buscar Partida por Orden de Compra"
End Sub

Private Function TotalSel() As Double
Dim dTot As Double
    dTot = 0
    gexOrdenes.Update
    Set rstAux = gexOrdenes.ADORecordset.Clone
    If rstAux.RecordCount > 0 Then rstAux.MoveFirst
    Do Until rstAux.EOF
        dTot = dTot + CDbl(rstAux!Kilos)
        rstAux.MoveNext
    Loop
    rstAux.Close: Set rstAux = Nothing
    
    lblTotal = Format(dTot, "0.00")
    
    TotalSel = dTot
End Function

Private Sub txtPanos_GotFocus()
    SelectionText txtPanos
End Sub

Private Sub txtPanos_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
