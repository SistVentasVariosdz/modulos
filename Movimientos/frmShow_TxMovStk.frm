VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmShow_TxMovStk 
   Caption         =   "Movimientos de Tejeduria"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   5250
      Left            =   45
      TabIndex        =   14
      Top             =   1545
      Width           =   10830
      Begin GridEX20.GridEX gexMov 
         Height          =   4965
         Left            =   90
         TabIndex        =   11
         Top             =   195
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   8758
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmShow_TxMovStk.frx":0000
         Column(2)       =   "frmShow_TxMovStk.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmShow_TxMovStk.frx":016C
         FormatStyle(2)  =   "frmShow_TxMovStk.frx":02A4
         FormatStyle(3)  =   "frmShow_TxMovStk.frx":0354
         FormatStyle(4)  =   "frmShow_TxMovStk.frx":0408
         FormatStyle(5)  =   "frmShow_TxMovStk.frx":04E0
         FormatStyle(6)  =   "frmShow_TxMovStk.frx":0598
         ImageCount      =   0
         PrinterProperties=   "frmShow_TxMovStk.frx":0678
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   3600
         Left            =   9540
         TabIndex        =   15
         Top             =   165
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   6350
         Custom          =   $"frmShow_TxMovStk.frx":0850
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   25
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Argumentos de Busqueda"
      ForeColor       =   &H8000000D&
      Height          =   1500
      Left            =   45
      TabIndex        =   12
      Top             =   45
      Width           =   10845
      Begin VB.TextBox txtSer_Guia 
         Height          =   285
         Left            =   6810
         TabIndex        =   7
         Top             =   810
         Width           =   510
      End
      Begin VB.TextBox txtCod_OrdTra 
         Height          =   285
         Left            =   6810
         TabIndex        =   9
         Top             =   1110
         Width           =   1305
      End
      Begin VB.TextBox txtNum_MovStk 
         Height          =   285
         Left            =   6810
         TabIndex        =   5
         Top             =   210
         Width           =   900
      End
      Begin VB.TextBox txtNumero_Guia 
         Height          =   285
         Left            =   7335
         TabIndex        =   8
         Top             =   810
         Width           =   1320
      End
      Begin VB.OptionButton OptMov 
         Caption         =   "Movimiento"
         Height          =   195
         Left            =   5610
         TabIndex        =   1
         Top             =   240
         Width           =   1125
      End
      Begin VB.OptionButton OptFecha 
         Caption         =   "Fecha"
         Height          =   195
         Left            =   5610
         TabIndex        =   2
         Top             =   540
         Width           =   795
      End
      Begin VB.OptionButton OptOT 
         Caption         =   "O.T."
         Height          =   195
         Left            =   5610
         TabIndex        =   4
         Top             =   1140
         Width           =   810
      End
      Begin VB.OptionButton OptGuia 
         Caption         =   "Guia"
         Height          =   195
         Left            =   5610
         TabIndex        =   3
         Top             =   840
         Width           =   1050
      End
      Begin VB.ComboBox cboAlmacen 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtpFec_MovStk 
         Height          =   285
         Left            =   6825
         TabIndex        =   6
         Top             =   510
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Format          =   130220033
         CurrentDate     =   37270
      End
      Begin FunctionsButtons.FunctButt fnbBuscar 
         Height          =   525
         Left            =   9360
         TabIndex        =   10
         Top             =   600
         Width           =   1230
         _ExtentX        =   2117
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~4~~0~Verdadero~Falso~&Buscar~"
         Orientacion     =   1
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Almacen:"
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   13
         Top             =   345
         Width           =   660
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   645
      Top             =   6435
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmShow_TxMovStk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrSql As String, sCod_Almacen As String, sNum_MovStk As String, sFec_MovStk As String, _
    sSer_Guia As String, sNumero_Guia As String, sCod_OrdTra As String, _
    sopcion As String


Private Sub cboAlmacen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub dtpFec_MovStk_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub fnbBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo Fin
Dim sTit As String
    sCod_Almacen = ""
    sNum_MovStk = "": sFec_MovStk = "": sSer_Guia = ""
    sNumero_Guia = "": sCod_OrdTra = "": sopcion = ""

    sTit = "Busqueda de Movimientos"

    If cboAlmacen.ListIndex = -1 Then
        MsgBox "Se debe elegir un Almacen", vbOKOnly + vbExclamation, sTit
        cboAlmacen.SetFocus
        Exit Sub
    End If

    sCod_Almacen = Left(cboAlmacen, 2)

    Select Case True
    Case OptMov
        sopcion = "1"
        txtNum_MovStk = Trim(txtNum_MovStk)
        If Len(txtNum_MovStk) <> 6 Then
            MsgBox "Numero de Movimiento Inválido", vbExclamation + vbOKOnly
            txtNum_MovStk.SetFocus
            Exit Sub
        End If
        sNum_MovStk = txtNum_MovStk
    Case OptFecha
        sopcion = "2"
        sFec_MovStk = Format(dtpFec_MovStk, "dd/mm/yyyy")
    Case OptGuia
        sopcion = "3"
        txtSer_Guia = Trim(txtSer_Guia)
        txtNumero_Guia = Trim(txtNumero_Guia)
        If Len(txtSer_Guia) <> 3 Or Len(txtNumero_Guia) <> 8 Then
            MsgBox "Nro de Guia Invalido", vbExclamation + vbOKOnly
            txtSer_Guia.SetFocus
            Exit Sub
        End If
        sSer_Guia = txtSer_Guia
        sNumero_Guia = txtNumero_Guia
    Case OptOT
        sopcion = "4"
        txtCod_OrdTra = Trim(txtCod_OrdTra)
        If Len(txtCod_OrdTra) <> 5 Then
            MsgBox "Nro de Orden Invalido", vbExclamation + vbOKOnly
            txtCod_OrdTra.SetFocus
            Exit Sub
        End If
        sCod_OrdTra = txtCod_OrdTra
    End Select

    Screen.MousePointer = 11

    StrSql = "EXEC TX_SM_MUESTRA_RESUMEN_MOV_TELA_CRUDA_SEGUN_OPCION '" & sopcion & _
    "', '" & sCod_Almacen & "', '" & sFec_MovStk & "', '" & sNumero_Guia & "', '" & _
    sCod_OrdTra & "', '" & sNum_MovStk & "'"

    Set gexMov.ADORecordset = CargarRecordSetDesconectado(StrSql, cConnect)

    gexMov.Columns("COD_CLASE_ROLLO").Caption = "Clase"
    gexMov.Columns("NUM_MOVSTK").Caption = "Nro.Mov."
    gexMov.Columns("FEC_MOVSTK").Caption = "Fecha"
    gexMov.Columns("GUIA").Caption = "Guia"
    gexMov.Columns("COD_ORDTRA").Caption = "O.T."
    gexMov.Columns("ABR_CLIENTE").Caption = "Abr."
    gexMov.Columns("KGS").Caption = "Kgs"
    gexMov.Columns("ROLLOS").Caption = "Rollos"
    gexMov.Columns("OC").Caption = "O/C"
    gexMov.Columns("DES_TELA").Caption = "Tela Tejeduria"
    gexMov.Columns("COD_FAMGRUPO").Caption = "Fam."
    gexMov.Columns("DES_TIPMOV").Caption = "Movimiento"
    gexMov.Columns("COD_CALIDAD").Caption = "Cal"

    gexMov.Columns("COD_CLASE_ROLLO").Width = 500
    gexMov.Columns("Num_MovStk").Width = 795
    gexMov.Columns("FEC_MOVSTK").Width = 1000
    gexMov.Columns("GUIA").Width = 1110
    gexMov.Columns("COD_ORDTRA").Width = 525
    gexMov.Columns("ABR_CLIENTE").Width = 735
    gexMov.Columns("KGS").Width = 825
    gexMov.Columns("ROLLOS").Width = 780
    gexMov.Columns("OC").Width = 930
    gexMov.Columns("DES_TELA").Width = 2700
    gexMov.Columns("COD_FAMGRUPO").Width = 480
    gexMov.Columns("DES_TIPMOV").Width = 2250
    gexMov.Columns("COD_CALIDAD").Width = 360

    gexMov.Columns("COD_TELA_TEJEDURIA").Visible = False
    gexMov.Columns("Cod_Almacen").Visible = False
    gexMov.Columns("Cod_TipMov").Visible = False
    gexMov.Columns("Ser_Guia").Visible = False
    gexMov.Columns("Numero_Guia").Visible = False
    gexMov.Columns("Observaciones").Visible = False

    Screen.MousePointer = 0
Exit Sub
Fin:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical + vbOKOnly, sTit
End Sub

Private Sub fnbBuscar_GotFocus()
    fnbBuscar_ActionClick 0, 0, ""
End Sub

Private Sub Form_Load()
    'OptFecha = True
    'FillAlmacen
    'fnbBuscar_ActionClick 0, 0, ""
    'FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp1, Me.Name)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If Not oParent Is Nothing Then oParent.DropWindowList Me.Tag
End Sub

'Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
'
'    Select Case ActionName
'    Case "ADICIONAR"
'        If cboAlmacen.ListIndex = -1 Then
'            MsgBox "Se debe elegir un Almacen", vbExclamation + vbOKOnly, "Adicionar Movimiento"
'            Exit Sub
'        End If
'        frmMovRollo.sAccion = "I"
'        frmMovRollo.sNum_MovStk = ""
'        frmMovRollo.sCod_Almacen = Left(cboAlmacen, 2)
'        frmMovRollo.Show vbModal
'        If Not frmMovRollo.bCancel Then
'            AddRollo "I", frmMovRollo.sCod_Almacen, _
'                    Left(frmMovRollo.cboTipMov, 3), _
'                    frmMovRollo.sNum_MovStk
'        End If
'        RefrescaDatos frmMovRollo.sNum_MovStk
'        Unload frmMovRollo
'    Case "MODIFICAR"
'        UpCabMov
'    Case "ADDROLLOS"
'        If gexMov.RowCount = 0 Then Exit Sub
'        AddRollo "I", gexMov.Value(gexMov.Columns("Cod_Almacen").Index), _
'                gexMov.Value(gexMov.Columns("Cod_TipMov").Index), _
'                gexMov.Value(gexMov.Columns("Num_MovStk").Index)
'        RefrescaDatos gexMov.Value(gexMov.Columns("Num_MovStk").Index)
'    Case "IMPRIMIRVOUCHER"
'           If gexMov.RowCount = 0 Then Exit Sub
'           VoucherExcel
'
'    Case "VERROLLOS"
'        If gexMov.RowCount = 0 Then Exit Sub
'        frmVerDetRollos.sCod_Almacen = gexMov.Value(gexMov.Columns("Cod_Almacen").Index)
'        frmVerDetRollos.sCod_TipMov = gexMov.Value(gexMov.Columns("Cod_TipMov").Index)
'        frmVerDetRollos.sNum_MovStk = gexMov.Value(gexMov.Columns("Num_MovStk").Index)
'        frmVerDetRollos.Caption = "Movimientos de Rollos : " & frmVerDetRollos.sNum_MovStk
'        frmVerDetRollos.BUSCAR
'        frmVerDetRollos.Show vbModal
'        RefrescaDatos gexMov.Value(gexMov.Columns("Num_MovStk").Index)
'    Case "SALIR"
'        Unload Me
'    End Select
'End Sub
'
'Private Sub UpCabMov()
'Dim rstAux As ADODB.Recordset
'
'    If gexMov.RowCount = 0 Then Exit Sub
'    'Llenar Datos
'    With frmMovRollo
'        .sAccion = "U"
'        .sNum_MovStk = gexMov.Value(gexMov.Columns("Num_MovStk").Index)
'        .sCod_Almacen = gexMov.Value(gexMov.Columns("Cod_Almacen").Index)
'
'        StrSql = "SELECT Tip_Accion, Flg_Despacho_Masivo FROM TX_TIPOSMOV " & _
'                 "WHERE Cod_TipMov = '" & gexMov.Value(gexMov.Columns("Cod_TipMov").Index) & "'"
'        Set rstAux = CargarRecordSetDesconectado(StrSql, cConnect)
'
'        If rstAux!Flg_Despacho_Masivo = "S" Then
'            MsgBox "Este Movimiento pertenece a un Despacho Masivo, " & _
'                   "No se puede Modificar", vbInformation + vbOKOnly, "Modificar Movimiento"
'            GoTo Fin
'        End If
'
'        .optExterno = (rstAux!Tip_Accion = "E")
'        .optInterno = (rstAux!Tip_Accion = "I")
'        rstAux.Close
'
'        BuscaCombo gexMov.Value(gexMov.Columns("Cod_TipMov").Index), 1, frmMovRollo.cboTipMov
'        .dtpMov = gexMov.Value(gexMov.Columns("Fec_MovStk").Index)
'        .dtpMov = gexMov.Value(gexMov.Columns("Fec_MovStk").Index)
'
'        StrSql = "SELECT a.Cod_Proveedor, a.Cod_CenCost, a.Cod_Cliente, " & _
'                 "ISNULL(b.Abr_Cliente, '') AS Abr_Cliente, " & _
'                 "ISNULL(a.Cod_Turno, '') AS Cod_Turno " & _
'                 "FROM TX_MOVISTK a, TX_CLIENTE b " & _
'                 "WHERE a.Cod_Almacen = '" & frmMovRollo.sCod_Almacen & "' " & _
'                 "AND   a.Num_MovStk = '" & frmMovRollo.sNum_MovStk & "' " & _
'                 "AND   a.Cod_Cliente *= b.Cod_Cliente_Tex"
'        Set rstAux = CargarRecordSetDesconectado(StrSql, cConnect)
'        If rstAux.RecordCount > 0 Then
'            rstAux.MoveFirst
'            .txtCod_Turno = rstAux!Cod_Turno
'            If Trim(frmMovRollo.txtCod_Turno) <> "" Then .BuscaTurno 1
'            .txtAbr_Cliente = rstAux!abr_cliente
'            If Trim(frmMovRollo.txtAbr_Cliente) <> "" Then .BuscaCliente 1
'            .txtCod_CenCost = rstAux!Cod_CenCost
'            If Trim(frmMovRollo.txtCod_CenCost) <> "" Then .BuscaCenCost 1
'            .txtCod_Proveedor = rstAux!Cod_Proveedor
'            If Trim(frmMovRollo.txtCod_Proveedor) <> "" Then .BuscaProveedor 1
'        End If
'
'        If gexMov.Value(gexMov.Columns("Rollos").Index) > 0 Then
'            .txtCod_Turno.Enabled = False
'            .txtDes_Turno.Enabled = False
'        End If
'
'        .txtSer_Guia = gexMov.Value(gexMov.Columns("Ser_Guia").Index)
'        .txtNumero_Guia = gexMov.Value(gexMov.Columns("Numero_Guia").Index)
'
'        .TxtObs = gexMov.Value(gexMov.Columns("Observaciones").Index)
'
'        .Frame1.Enabled = False
'        .Show vbModal
'    End With
'    RefrescaDatos frmMovRollo.sNum_MovStk
'Fin:
'    Unload frmMovRollo
'
'End Sub
'
'Public Sub AddRollo(Accion As String, Cod_Almacen, cod_tipmov As String, Num_MovStk As String)
'Dim rstMov As ADODB.Recordset
'
'    StrSql = "SELECT Des_TipMov, Cod_Calidad, Cod_ClaMov, Cod_TipAnx, Tip_PtMp, Flg_Devolucion_Rollos_Tejeduria " & _
'             "FROM LG_TIPOSMOV " & _
'             "WHERE Cod_TipMov = '" & cod_tipmov & "'"
'
'    Set rstMov = CargarRecordSetDesconectado(StrSql, cConnect)
'
'    If rstMov.RecordCount > 0 Then rstMov.MoveFirst
'
'    If IIf(IsNull(rstMov!Flg_Devolucion_Rollos_Tejeduria), "", rstMov!Flg_Devolucion_Rollos_Tejeduria) = "S" Then
'        With frmMovRolloDevolucionDet
'            .sAccion = Accion
'            .sCod_Almacen = Cod_Almacen
'            .sCod_TipMov = cod_tipmov
'            .sDes_TipMov = IIf(IsNull(rstMov!DES_TIPMOV), "", rstMov!DES_TIPMOV)
'            .sCod_Calidad = IIf(IsNull(rstMov!Cod_Calidad), "", rstMov!Cod_Calidad)
'            .sCod_ClaMov = IIf(IsNull(rstMov!Cod_ClaMov), "", rstMov!Cod_ClaMov)
'            .sCod_TipAnx = IIf(IsNull(rstMov!Cod_TipAnx), "", rstMov!Cod_TipAnx)
'            .sTip_PtMP = IIf(IsNull(rstMov!Tip_PtMp), "", rstMov!Tip_PtMp)
'            .sFlg_Devolucion_Rollos_Tejeduria = IIf(IsNull(rstMov!Flg_Devolucion_Rollos_Tejeduria), "", rstMov!Flg_Devolucion_Rollos_Tejeduria)
'
'            .sNum_MovStk = Num_MovStk
'            .sNum_Secuencia = ""
'            .Caption = "Det. Movimiento de Devolucion: " & .sNum_MovStk & " - " & .sDes_TipMov
'            rstMov.Close
'            .HabilitaCantidades
'            .LimpiaForm
'            .Show vbModal
'        End With
'    Else
'        With frmMovRolloDet
'            .sAccion = Accion
'            .sCod_Almacen = Cod_Almacen
'            .sCod_TipMov = cod_tipmov
'            .sDes_TipMov = IIf(IsNull(rstMov!DES_TIPMOV), "", rstMov!DES_TIPMOV)
'            .sCod_Calidad = IIf(IsNull(rstMov!Cod_Calidad), "", rstMov!Cod_Calidad)
'            .sCod_ClaMov = IIf(IsNull(rstMov!Cod_ClaMov), "", rstMov!Cod_ClaMov)
'            .sCod_TipAnx = IIf(IsNull(rstMov!Cod_TipAnx), "", rstMov!Cod_TipAnx)
'            '.sTip_PtMP = IIf(IsNull(rstMov!Tip_PtMp), "", rstMov!Tip_PtMp)
'            '.sFlg_Devolucion_Rollos_Tejeduria = IIf(IsNull(rstMov!Flg_Devolucion_Rollos_Tejeduria), "", rstMov!Flg_Devolucion_Rollos_Tejeduria)
'            .sNum_MovStk = Num_MovStk
'            .sNum_Secuencia = ""
'            .Caption = "Detalle de Movimiento : " & .sNum_MovStk & " - " & .sDes_TipMov
'            rstMov.Close
'            .HabilitaCantidades
'            .LimpiaForm
'            .Show vbModal
'        End With
'    End If
'End Sub
'
'Private Sub FillAlmacen()
'Dim rstAux As ADODB.Recordset
'
'    StrSql = "SELECT Cod_Almacen, Nom_Almacen FROM TX_ALMACEN " & _
'             "WHERE  Tip_Item = 'T' " & _
'             "AND    Tip_Presentacion = 'C' " & _
'             "AND    ISNULL(Flg_Pre_Tenido, 'N') = 'N'"
'    Set rstAux = CargarRecordSetDesconectado(StrSql, cConnect)
'    cboAlmacen.Clear
'    With rstAux
'        If .RecordCount > 0 Then .MoveFirst
'        Do Until .EOF
'            cboAlmacen.AddItem !Cod_Almacen & " " & !Nom_Almacen
'            .MoveNext
'        Loop
'        .Close
'    End With
'    If cboAlmacen.ListCount > 0 Then cboAlmacen.ListIndex = 0
'    Set rstAux = Nothing
'End Sub
'
'Private Sub OptFecha_Click()
'    BusquedaVisible
'End Sub
'
'Private Sub OptFecha_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then dtpFec_MovStk.SetFocus
'End Sub
'
'Private Sub OptGuia_Click()
'    BusquedaVisible
'End Sub
'
'Private Sub OptGuia_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then txtSer_Guia.SetFocus
'End Sub
'
'Private Sub optMov_Click()
'    BusquedaVisible
'End Sub
'
'Private Sub OptMov_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then txtNum_MovStk.SetFocus
'End Sub
'
'Private Sub OptOT_Click()
'    BusquedaVisible
'End Sub
'
'Private Sub BusquedaVisible()
'    txtNum_MovStk = ""
'    dtpFec_MovStk = Date
'    txtSer_Guia = ""
'    txtNumero_Guia = ""
'    txtCod_OrdTra = ""
'
'    txtNum_MovStk.Visible = OptMov
'    dtpFec_MovStk.Visible = OptFecha
'    txtSer_Guia.Visible = OptGuia
'    txtNumero_Guia.Visible = OptGuia
'    txtCod_OrdTra.Visible = OptOT
'End Sub
'
'Private Sub optOT_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then txtCod_OrdTra.SetFocus
'End Sub
'
'Private Sub txtCod_OrdTra_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        txtCod_OrdTra = Format(txtCod_OrdTra, "00000")
'        SendKeys "{TAB}"
'    End If
'End Sub
'
'Private Sub txtNum_MovStk_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        txtNum_MovStk = Format(txtNum_MovStk, "000000")
'        SendKeys "{TAB}"
'    End If
'End Sub
'
'Private Sub txtNumero_Guia_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        txtNumero_Guia = Format(txtNumero_Guia, "00000000")
'        SendKeys "{TAB}"
'    End If
'End Sub
'
'Private Sub RefrescaDatos(NumMov As String)
'    If OptMov Then txtNum_MovStk = NumMov
'    fnbBuscar_ActionClick 0, 0, ""
'    gexMov.Find gexMov.Columns("Num_MovStk").Index, jgexEqual, NumMov
'End Sub
'
'Private Sub txtSer_Guia_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        txtSer_Guia = Format(txtSer_Guia, "000")
'        SendKeys "{TAB}"
'    End If
'End Sub
'
'Private Sub VoucherExcel()
'On Error GoTo ErrExcelVou
'Dim oo As Object, vRutaLogo As Variant
'Dim cliente As String
'Dim StrSql1 As String
'
'    StrSql = "SELECT Ruta_Logo From SEGURIDAD..SEG_EMPRESAS " & _
'             "WHERE Cod_Empresa = '" & vemp & "'"
'    vRutaLogo = DevuelveCampo(StrSql, cConnect)
'
'    vRutaLogo = CStr(IIf(IsNull(vRutaLogo), "", vRutaLogo))
'
'    StrSql1 = "select nom_cliente from tx_cliente  " & _
'             " WHERE abr_cliente = '" & Trim(gexMov.Value(gexMov.Columns("abr_cliente").Index)) & "'"
'    cliente = DevuelveCampo(StrSql1, cConnect)
'
'    Screen.MousePointer = 11
'
'    Set oo = CreateObject("excel.application")
'
'    oo.Workbooks.Open vRuta & "\Vouchertejeduria.XLT"
'    oo.DisplayAlerts = False
'    oo.Visible = True
'
'    oo.Run "REPORTE", Left(cboAlmacen, 2), gexMov.Value(gexMov.Columns("Num_MovStk").Index), CStr(vRutaLogo), cboAlmacen, gexMov.Value(gexMov.Columns("DES_TIPMOV").Index), gexMov.Value(gexMov.Columns("FEC_MOVSTK").Index), gexMov.Value(gexMov.Columns("GUIA").Index), gexMov.Value(gexMov.Columns("Des_Proveedor").Index), gexMov.Value(gexMov.Columns("Fec_creacion").Index), cliente, cConnect
'
'    'oo.Workbooks.Close
'    Set oo = Nothing
'    Screen.MousePointer = 0
'Exit Sub
'ErrExcelVou:
'    Screen.MousePointer = 0
'    MsgBox err.Description, vbCritical + vbOKOnly, "Imprimir Voucher Formato Excel"
'End Sub
'
'
'
