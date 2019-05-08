VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAddLetras 
   Caption         =   "Adición Letra Cabecera"
   ClientHeight    =   2505
   ClientLeft      =   2445
   ClientTop       =   1485
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   9900
   Begin VB.Frame frTransacciones 
      BorderStyle     =   0  'None
      Height          =   2730
      Left            =   15
      TabIndex        =   10
      Top             =   -240
      Width           =   9735
      Begin VB.Frame Frame3 
         Height          =   1710
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   9615
         Begin VB.TextBox txtDes_Origen 
            Height          =   285
            Left            =   6585
            TabIndex        =   5
            Top             =   675
            Width           =   1575
         End
         Begin VB.TextBox txtCod_Origen 
            Height          =   285
            Left            =   6165
            MaxLength       =   1
            TabIndex        =   4
            Top             =   675
            Width           =   375
         End
         Begin VB.TextBox TxtCod_Banco 
            Height          =   285
            Left            =   1305
            TabIndex        =   0
            Top             =   255
            Width           =   735
         End
         Begin VB.TextBox TxtDes_Banco 
            Height          =   285
            Left            =   2085
            TabIndex        =   1
            Top             =   255
            Width           =   2415
         End
         Begin VB.TextBox TxtObservacion 
            Height          =   285
            Left            =   1320
            TabIndex        =   6
            Top             =   1230
            Width           =   7965
         End
         Begin VB.TextBox txtCuenta_Des 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   6900
            MaxLength       =   30
            TabIndex        =   3
            Top             =   225
            Width           =   2370
         End
         Begin VB.TextBox txtCuenta_Cod 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   6120
            MaxLength       =   11
            TabIndex        =   2
            Top             =   225
            Width           =   720
         End
         Begin MSComCtl2.DTPicker DTPFecha 
            Height          =   300
            Left            =   1305
            TabIndex        =   9
            Top             =   690
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   529
            _Version        =   393216
            Format          =   94109697
            CurrentDate     =   38590
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Status :"
            Height          =   195
            Left            =   5400
            TabIndex        =   15
            Top             =   750
            Width           =   540
         End
         Begin VB.Label Label3 
            Caption         =   "Banco :"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   285
            Width           =   645
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Funcionario"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   1215
            Width           =   825
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta :"
            Height          =   195
            Left            =   5400
            TabIndex        =   12
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha de Presentación :"
            Height          =   375
            Left            =   135
            TabIndex        =   11
            Top             =   630
            Width           =   1590
         End
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   7215
         TabIndex        =   7
         Top             =   2100
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmAddLetras.frx":0000
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
   End
End
Attribute VB_Name = "frmAddLetras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public codigo As String, Descripcion As String, strOption As String, strCod_Anxo As String, lfSalvar As Boolean
Public sTipoBusq As String, sNum_Planilla_Letra As Integer
Dim RsGrid1 As ADODB.Recordset
Dim RsGrid2 As ADODB.Recordset
Dim intTransaccion As Integer, vrTotalTransaccion As Double
Dim strSQL As String, intCancel As Integer

Private Sub chkDiferido_Click()

If chkDiferido Then
  txtFec_Diferido.Visible = True
  lbDiferido.Visible = True
Else
  txtFec_Diferido.Visible = False
  lbDiferido.Visible = False
  txtFec_Diferido.Text = ""
End If

End Sub

Private Sub CmdBack_Click()
    If gexGrid2.RowCount = 0 Then Exit Sub
    RsGrid1.AddNew
    RsGrid1.Fields("Correlativo").Value = gexGrid2.Value(gexGrid1.Columns("Correlativo").Index)
    RsGrid1.Fields("Numero").Value = gexGrid2.Value(gexGrid1.Columns("Numero").Index)
    RsGrid1.Fields("Fecha").Value = gexGrid2.Value(gexGrid1.Columns("Fecha").Index)
    RsGrid1.Fields("Tipo_Cambio").Value = gexGrid2.Value(gexGrid1.Columns("Tipo_Cambio").Index)
    RsGrid1.Fields("Moneda").Value = gexGrid2.Value(gexGrid1.Columns("Moneda").Index)
    RsGrid1.Fields("Monto_Origen1").Value = gexGrid2.Value(gexGrid1.Columns("Monto_Origen1").Index)
    RsGrid1.Fields("Monto_Origen").Value = gexGrid2.Value(gexGrid1.Columns("Monto_Origen").Index)
    RsGrid1.Fields("Monto_Aceptado").Value = gexGrid2.Value(gexGrid1.Columns("Monto_Aceptado").Index)
    RsGrid1.Fields("Cod_Cobranza").Value = gexGrid2.Value(gexGrid1.Columns("Cod_Cobranza").Index)
    RsGrid1.Fields("Debe_Haber").Value = gexGrid2.Value(gexGrid1.Columns("Debe_Haber").Index)
    RsGrid1.Fields("Observacion").Value = gexGrid2.Value(gexGrid1.Columns("Observacion").Index)
    RsGrid1.Fields("Tran_TipMonDoc").Value = gexGrid2.Value(gexGrid1.Columns("Tran_TipMonDoc").Index)
    RsGrid1.Fields("Doc_TipMonDoc").Value = gexGrid2.Value(gexGrid1.Columns("Doc_TipMonDoc").Index)
    RsGrid1.Fields("Otro_Tip_Cambio").Value = gexGrid2.Value(gexGrid1.Columns("Otro_Tip_Cambio").Index)


    RsGrid1.Update

    RsGrid2.MoveFirst
    Call BuscaCampo(RsGrid2, "Correlativo", gexGrid2.Value(gexGrid2.Columns("Correlativo").Index))

    RsGrid2.Delete

    Set gexGrid1.ADORecordset = RsGrid1
    'ConfigurarGrid gexGrid1
    Set gexGrid2.ADORecordset = RsGrid2
    'ConfigurarGrid gexGrid2

    If RsGrid1.RecordCount = 0 Then
        TxtMonto1.Text = "0.00"
    Else
        TxtMonto1 = CALCULA_MONTO_TOTAL(RsGrid1)
    End If

    If RsGrid2.RecordCount = 0 Then
        TxtMonto2.Text = "0.00"
    Else
        TxtMonto2 = CALCULA_MONTO_TOTAL(RsGrid2)
    End If
End Sub

Private Sub cmdBackAll_Click()
Dim I As Integer

  If gexGrid2.RowCount > 0 Then
      VB.Screen.MousePointer = 11
      gexGrid2.Redraw = False
      gexGrid1.Redraw = False
      gexGrid2.MoveFirst

      For I = 1 To gexGrid2.RowCount
          CmdBack_Click
      Next


      gexGrid2.Redraw = True
      gexGrid1.Redraw = True

      VB.Screen.MousePointer = 0
  End If

End Sub

Private Sub cmdDocumentos_Click()

  If DevuelveCampo("select count(*) from tg_moneda where cod_moneda='" & Trim(txtCod_Moneda.Text) & "'", cCONNECT) = 0 Then
    MsgBox "Seleccione una Moneda Valida", vbInformation, Me.Caption
    txtCod_Moneda.SetFocus
    Exit Sub
  End If

  If strCod_Anxo = "" Then
    MsgBox "Seleccione una Cliente", vbInformation, Me.Caption
    txtNum_Ruc.SetFocus
    Exit Sub
  End If

  If Me.WindowState <> 2 Then Me.Height = 7590

  frTransacciones.Visible = False
  frFacturas.Top = 0
  frFacturas.Visible = True

  fncBuscar.SetFocus

End Sub

Private Sub cmdNext_Click()

Dim Valor As Double, varCorrelativo As String

If gexGrid1.RowCount = 0 Then Exit Sub

If gexGrid1.EditMode = jgexEditModeOn Then
  MsgBox "Salga del Modo de Edicion de la Grilla" & vbCr & "Haga Click en la columna Numero", vbInformation, "IMPORTANTE"
  Exit Sub
End If

Valor = txt_ImpTotal_Doc_Cobra.Text - (gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) + CDbl(TxtMonto2))
If Valor < -10 And txt_ImpTotal_Doc_Cobra.Text <> 0 Then
   If MsgBox("Con este Documento el importe excederia en " & Valor & "  al importe del Documento de Cobranza de " & txt_ImpTotal_Doc_Cobra.Text, vbYesNo, "AVISO") = vbNo Then Exit Sub
End If

RsGrid2.AddNew
RsGrid2.Fields("Correlativo").Value = gexGrid1.Value(gexGrid1.Columns("Correlativo").Index)
RsGrid2.Fields("Numero").Value = gexGrid1.Value(gexGrid1.Columns("Numero").Index)
RsGrid2.Fields("Fecha").Value = gexGrid1.Value(gexGrid1.Columns("Fecha").Index)
RsGrid2.Fields("Moneda").Value = gexGrid1.Value(gexGrid1.Columns("Moneda").Index)
RsGrid2.Fields("Tipo_Cambio").Value = gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index)
RsGrid2.Fields("Monto_Origen1").Value = gexGrid1.Value(gexGrid1.Columns("Monto_Origen1").Index)
RsGrid2.Fields("Monto_Origen").Value = gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index)
RsGrid2.Fields("Monto_Aceptado").Value = gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index)
RsGrid2.Fields("Cod_Cobranza").Value = gexGrid1.Value(gexGrid1.Columns("Cod_Cobranza").Index)
RsGrid2.Fields("Debe_Haber").Value = gexGrid1.Value(gexGrid1.Columns("Debe_Haber").Index)
RsGrid2.Fields("Observacion").Value = gexGrid1.Value(gexGrid1.Columns("Observacion").Index)
RsGrid2.Fields("Tran_TipMonDoc").Value = gexGrid1.Value(gexGrid1.Columns("Tran_TipMonDoc").Index)
RsGrid2.Fields("Doc_TipMonDoc").Value = gexGrid1.Value(gexGrid1.Columns("Doc_TipMonDoc").Index)
RsGrid2.Fields("Otro_Tip_Cambio").Value = gexGrid1.Value(gexGrid1.Columns("Otro_Tip_Cambio").Index)

RsGrid2.Update

RsGrid1.MoveFirst
Call BuscaCampo(RsGrid1, "Correlativo", gexGrid1.Value(gexGrid1.Columns("Correlativo").Index))

RsGrid1.Delete

Set gexGrid1.ADORecordset = RsGrid1
'ConfigurarGrid gexGrid1
Set gexGrid2.ADORecordset = RsGrid2
'ConfigurarGrid gexGrid2

If RsGrid1.RecordCount = 0 Then
    TxtMonto1.Text = "0.00"
Else
    TxtMonto1 = CALCULA_MONTO_TOTAL(RsGrid1)
End If

If RsGrid2.RecordCount = 0 Then
    TxtMonto2.Text = "0.00"
Else
    TxtMonto2 = CALCULA_MONTO_TOTAL(RsGrid2)
End If

End Sub

Private Sub CmdNextAll_Click()
Dim I As Integer
    If gexGrid1.RowCount > 0 Then
        gexGrid2.Redraw = False
        gexGrid1.Redraw = False
        gexGrid1.MoveFirst

        For I = 1 To gexGrid1.RowCount
            cmdNext_Click
        Next

        gexGrid2.Redraw = True
        gexGrid1.Redraw = True

    End If
End Sub

Private Sub cmdObtieneComp_Click()

Dim RS As Object
Set RS = CreateObject("ADODB.Recordset")

Set RS = CargarRecordSetDesconectado("Select Cod_TipDoc,Ser_Docum,Num_Docum From Tg_Bancos_Cuentas Where Cod_Banco = '" & TxtCod_Banco & "' and Sec_Cuenta_Banco = '" & txtCuenta_Cod & "' and Fec_Transaccion_Cobranza = '" & txtFecha.Text & "' and Cod_Usuario = '" & vusu & "'", cCONNECT)

If Not (RS.BOF And RS.EOF) Then

  txtCod_TipDocCobra = RS!Cod_TipDoc
  txtSer_DocCobra = RS!Ser_Docum
  txtNum_DocCobra = RS!Num_Docum

  Set RS = Nothing

End If

End Sub

Private Sub fncBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
  Case "BUSCAR"
    CARGA_FACTURAS_PENDIENTES
  Case "CERRAR"
    frFacturas.Visible = False
    frTransacciones.Visible = True
    Resumen_Facturas
    Cambio_Apariencia
    If txtCod_TipCobra = "013" Then txt_ImpTotal_Doc_Cobra.Text = CDbl(txtTotDoc.Text)
    cmdDocumentos.SetFocus
End Select
End Sub
Sub Calcula_Monto_Letra()

Dim dbTotFac As Double, I As Integer, dbMontoLetra As Double, Nro_Letra As Integer, VrBookMark, dbTotLetras As Double

dbTotFac = CDbl(txtTotFact)
Nro_Letra = gexLetra.RowCount
txtTotLetras = 0

If dbTotFac = 0 Then Exit Sub
If Nro_Letra = 0 Then Exit Sub

gexLetra.MoveFirst
dbMontoLetra = 0

VrBookMark = gexLetra.Row

For I = 1 To Nro_Letra

  If I = Nro_Letra Then
    dbMontoLetra = dbTotFac - dbTotLetras
  Else
    dbMontoLetra = Format(dbTotFac / Nro_Letra, "###,###.00")
  End If

  dbTotLetras = dbTotLetras + dbMontoLetra

  gexLetra.Value(gexLetra.Columns("Imp_Total").Index) = dbMontoLetra
  gexLetra.MoveNext
Next I

gexLetra.Row = VrBookMark
txtTotLetras = Format(dbTotLetras, "###,###.00")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = intCancel
End Sub



Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim sMessage As Long
Dim strSQL As String

On Error GoTo dprError

Select Case ActionName
  Case "GRABAR"
If TxtCod_Banco.Text = "" Or txtCuenta_Cod.Text = "" Or txtCod_Origen.Text = "" Then
    MsgBox "Debe de ingresar todos los campos"
Else
    If MsgBox("Esta seguro de incorporar la Letra a la Planilla Actual", vbYesNo, "IMPORTANTE") = vbYes Then

    If sTipoBusq = 1 Then
        sNum_Planilla_Letra = DevuelveCampo("select ISNULL(max(num_planilla_letra), 0) from Cn_Ventas_Planilla_Letras ", cCONNECT)
        sNum_Planilla_Letra = sNum_Planilla_Letra + 1
     End If


      If lfSalvar_Datos Then
        lfSalvar = True
        With frmTransacciones
          .inpFec_Emi.Text = txtFecha.Text
          .txtCod_Origen = txtCod_Origen
          .txtDes_Origen = txtDes_Origen
        End With
        intCancel = 0
        Unload Me

      End If
      intCancel = 0
      frmControl_Letras.Buscar
      Unload Me
    End If
End If
  Case "CANCELAR"
    If MsgBox("Esta seguro de eliminar la Letra a la Planilla Actual", vbYesNo, "IMPORTANTE") = vbYes Then
      lfSalvar = False
      intCancel = 0
      Unload Me
    End If
End Select

Exit Sub

dprError:

errores err.Number

End Sub

Private Function lfSalvar_Datos() As Boolean

On Error GoTo hand

Dim SQL As String



  SQL = "Ventas_Generar_Control_Letras '" & sTipoBusq & "','" & sNum_Planilla_Letra & "','" & TxtCod_Banco.Text & "','" _
          & DTPFecha & "','" & txtCuenta_Cod.Text & "', '81', '" & txtCod_Origen.Text & "','" & TxtObservacion.Text & "'"

  Call ExecuteCommandSQL(cCONNECT, SQL)


Exit Function

hand:

errores err.Number

lfSalvar_Datos = False

End Function

Private Function lfGenera_Cuadre() As Boolean

On Error GoTo hand

Dim RS As ADODB.Recordset
Dim SQL As String
Dim I As Integer
Dim strCod_Cobranza As String, strFlg_Debe_Haber As String

gridex1.Redraw = False



SQL = "TM_VENTAS_TRANSACCIONES_COBRANZAS_INSERT '" & strOption & "'," & intTransaccion & ",'" & txtFecha.Text & "'," & 0 & ",'" _
      & txtCod_TipCobra & "','" & txtCod_TipAne & "','" & strCod_Anxo & "','" & TxtCod_Banco & "','" & txtCuenta_Cod & "','" _
      & txtCod_TipDocCobra & "','" & txtSer_DocCobra & "','" & txtNum_DocCobra & "','" & txtCod_Moneda & "','" _
      & Des_Apos(TxtObservacion) & "','" & vusu & "','" & ComputerName & "','" & txtCod_Origen & "','" & IIf(chkDiferido, "S", "N") & "'," _
      & IIf(txtFec_Diferido.Text <> "", "'" & txtFec_Diferido.Text & "'", "Null")

Set RS = CargarRecordSetDesconectado(SQL, cCONNECT)

If Not (RS.BOF Or RS.EOF) Then
  intTransaccion = RS!Num_Transaccion
  strCod_Cobranza = RS!Cod_Concepto_Cobranza
  strFlg_Debe_Haber = RS!Flg_Debe_Haber
End If

strOption = "I"

If txt_ImpTotal_Doc_Cobra.Text <> 0 Then
  SQL = "TM_Ventas_Transacciones_Cobranzas_DETALLE_MAN '" & strOption & "'," & intTransaccion & ",'" & txtFecha.Text & "'," _
        & 0 & "," & 0 & ",'" & strCod_Cobranza & "','" & strFlg_Debe_Haber & "',''," & txt_ImpTotal_Doc_Cobra.Text & ",''," & txtTipo_Cambio & ",'S'," & txtOtro_Tipo_Cambio
  Call ExecuteCommandSQL(cCONNECT, SQL)
End If

gexGrid2.MoveFirst
For I = 1 To gexGrid2.RowCount
  SQL = "TM_Ventas_Transacciones_Cobranzas_DETALLE_MAN '" & strOption & "'," & intTransaccion & ",'" & txtFecha.Text & "',0,0,'" _
        & gexGrid2.Value(gexGrid2.Columns("Cod_Cobranza").Index) & "','" & gexGrid2.Value(gexGrid2.Columns("Debe_Haber").Index) & "','" _
        & gexGrid2.Value(gexGrid2.Columns("Correlativo").Index) & "','" & gexGrid2.Value(gexGrid2.Columns("Monto_Origen").Index) & "','" _
        & gexGrid2.Value(gexGrid2.Columns("Observacion").Index) & "'," & gexGrid2.Value(gexGrid2.Columns("Tipo_Cambio").Index) & ",'S'," _
        & gexGrid2.Value(gexGrid2.Columns("Otro_Tip_Cambio").Index)

  Call ExecuteCommandSQL(cCONNECT, SQL)
  gexGrid2.MoveNext
Next

gridex1.MoveFirst
For I = 1 To gridex1.RowCount
  If gridex1.Value(gridex1.Columns("Sel").Index) Then
    SQL = "TM_Ventas_Transacciones_Cobranzas_DETALLE_MAN '" & strOption & "'," & intTransaccion & ",'" & txtFecha.Text & "',0,0,'" _
          & gridex1.Value(gridex1.Columns("Cod").Index) & "','" & gridex1.Value(gridex1.Columns("flg").Index) & "','" & "" & "','" _
          & IIf(gridex1.Value(gridex1.Columns("flg").Index) = "D", gridex1.Value(gridex1.Columns("Imp_Debe").Index), gridex1.Value(gridex1.Columns("Imp_Haber").Index)) & "','" _
          & gridex1.Value(gridex1.Columns("Observacion").Index) & "'," & txtTipo_Cambio & ",'S'," & txtOtro_Tipo_Cambio
    Call ExecuteCommandSQL(cCONNECT, SQL)
  End If
  gridex1.MoveNext
Next

strOption = "U"

gridex1.Redraw = True

With frmTransaccionesAddCuadre
  .strSQL = "TM_VENTAS_MUESTRA_CUADRE_COBRANZAS " & intTransaccion
  .intNum_Transaccion = intTransaccion
  .strCod_Anexo = strCod_Anxo
  .strCod_TipAnexo = txtCod_TipAne
  .strCod_Moneda = txtCod_Moneda
  .dFecha = txtFecha.Text
  .CARGA_GRID
  .Caption = "Detalle de Transaccion del Cliente " & txtDes_TipAne
  .Show 1
  lfGenera_Cuadre = .lfAceptar
End With

Exit Function
Resume
hand:

gridex1.Redraw = True

errores err.Number

Set RS = Nothing

lfGenera_Cuadre = False

End Function


Private Sub gexGrid1_AfterColEdit(ByVal ColIndex As Integer)

Dim dbImporte As Double

  Select Case ColIndex

  Case Is = gexGrid1.Columns("Monto_Aceptado").Index

    If CDbl(gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index)) <= 0 Then
      MsgBox "El Monto del documento debe ser Mayor a Cero", vbInformation, "AVISO"
      gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) = gexGrid1.Value(gexGrid1.Columns("Monto_Origen1").Index)
      gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) = DevuelveCampo("select dbo.Convierte_Importe_Moneda_Destino('" & gexGrid1.Value(gexGrid1.Columns("Moneda").Index) & "'," & gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) & "," & gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index) & ",'" & txtCod_Moneda & "','" & txtFecha.Text & "'," & gexGrid1.Value(gexGrid1.Columns("Otro_Tip_Cambio").Index) & ")", cCONNECT)
      Exit Sub
    End If

    gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) = DevuelveCampo("select dbo.Convierte_Importe_Moneda_Destino('" & txtCod_Moneda & "'," & gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) & "," & gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index) & ",'" & gexGrid1.Value(gexGrid1.Columns("Moneda").Index) & "','" & txtFecha.Text & "'," & gexGrid1.Value(gexGrid1.Columns("Otro_Tip_Cambio").Index) & ")", cCONNECT)

   If CDbl(gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index)) > CDbl(gexGrid1.Value(gexGrid1.Columns("Monto_Origen1").Index)) Then
      MsgBox "No se puede ingresar un monto Mayor al pendiente del Documento", vbInformation, "AVISO"
      gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) = gexGrid1.Value(gexGrid1.Columns("Monto_Origen1").Index)
      gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) = DevuelveCampo("select dbo.Convierte_Importe_Moneda_Destino('" & gexGrid1.Value(gexGrid1.Columns("Moneda").Index) & "'," & gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) & "," & gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index) & ",'" & txtCod_Moneda & "','" & txtFecha.Text & "'," & gexGrid1.Value(gexGrid1.Columns("Otro_Tip_Cambio").Index) & ")", cCONNECT)
   End If

    If Abs(gexGrid1.Value(gexGrid1.Columns("Monto_Origen1").Index) - gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index)) = 0.01 Then
        gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) = gexGrid1.Value(gexGrid1.Columns("Monto_Origen1").Index)
        gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) = DevuelveCampo("select dbo.Convierte_Importe_Moneda_Destino('" & gexGrid1.Value(gexGrid1.Columns("Moneda").Index) & "'," & gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) & "," & gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index) & ",'" & txtCod_Moneda & "','" & txtFecha.Text & "'," & gexGrid1.Value(gexGrid1.Columns("Otro_Tip_Cambio").Index) & ")", cCONNECT)
    End If

'    SendKeys "{ENTER}"

  Case Is = gexGrid1.Columns("Tipo_Cambio").Index

    If Trim(gexGrid1.Value(gexGrid1.Columns("Doc_TipMonDoc").Index)) <> "" Then

      If txtCod_Moneda <> gexGrid1.Value(gexGrid1.Columns("Moneda").Index) Then

        gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) = IIf(txtCod_Moneda = "SOL", gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) * gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index), Format(gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) / gexGrid1.Value(gexGrid1.Columns("Tipo_Cambio").Index), "###,###.00"))
      Else
        gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) = gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index)
      End If

    End If

 '   SendKeys "{ENTER}"

  Case Is = gexGrid1.Columns("Otro_Tip_Cambio").Index
    gexGrid1.Value(gexGrid1.Columns("Monto_Aceptado").Index) = Format(gexGrid1.Value(gexGrid1.Columns("Monto_Origen").Index) * gexGrid1.Value(gexGrid1.Columns("Otro_Tip_Cambio").Index), "###,###.00")
    TxtMonto1 = CALCULA_MONTO_TOTAL(RsGrid1)

    SendKeys "{ENTER}"

  End Select

End Sub

Private Sub gexGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
  Select Case ColIndex
  Case Is = gexGrid1.Columns("Monto_Aceptado").Index
    Cancel = False
  Case Is = gexGrid1.Columns("Tipo_Cambio").Index
    Cancel = False
  Case Is = gexGrid1.Columns("Observacion").Index
    Cancel = False
  Case Is = gexGrid1.Columns("Otro_Tip_Cambio").Index
    If Trim(gexGrid1.Value(gexGrid1.Columns("Doc_TipMonDoc").Index)) <> "" Then Cancel = True Else Cancel = False
  Case Else
    Cancel = True
  End Select
End Sub

Private Sub gexGrid1_BeforeColUpdate(ByVal Row As Long, ByVal ColIndex As Integer, ByVal OldValue As String, ByVal Cancel As GridEX20.JSRetBoolean)
    'If gexGrid1.Columns("Monto_ORIGEN").Index = ColIndex Then
    '    If OldValue = gexGrid1.Value(gexGrid1.Columns("Monto_ORIGEN").Index) Then
    '        gexGrid1.Value(gexGrid1.Columns("Monto_ORIGEN").Index) = OldValue
    '    End If
    'End If
End Sub

Private Sub GridEX1_AfterColEdit(ByVal ColIndex As Integer)
  Select Case ColIndex
  Case Is = gridex1.Columns("Sel").Index
    If Not gridex1.Value(gridex1.Columns("Sel").Index) Then
      gridex1.Value(gridex1.Columns("Imp_Debe").Index) = 0
      gridex1.Value(gridex1.Columns("Imp_Haber").Index) = 0
      gridex1.Value(gridex1.Columns("Observacion").Index) = ""
    End If
  End Select
End Sub

Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
  Select Case ColIndex
  Case Is = gridex1.Columns("Sel").Index
    Cancel = False
  Case Is = gridex1.Columns("Imp_Debe").Index
    Cancel = IIf(gridex1.Value(gridex1.Columns("Flg").Index) = "D" And gridex1.Value(gridex1.Columns("sel").Index), False, True)
  Case Is = gridex1.Columns("Imp_Haber").Index
    Cancel = IIf(gridex1.Value(gridex1.Columns("Flg").Index) = "H" And gridex1.Value(gridex1.Columns("sel").Index), False, True)
  Case Is = gridex1.Columns("Observacion").Index
    Cancel = IIf(gridex1.Value(gridex1.Columns("sel").Index), False, True)
  Case Else
    Cancel = True
  End Select
End Sub

Private Sub txt_ImpTotal_Doc_Cobra_Change()
 If txt_ImpTotal_Doc_Cobra = "" Then txt_ImpTotal_Doc_Cobra = 0
End Sub

Private Sub txt_ImpTotal_Doc_Cobra_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub txt_ImpTotal_Doc_Cobra_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
  SoloNumeros txt_ImpTotal_Doc_Cobra, KeyAscii, True, 2, 9
End Sub

Private Sub TxtCod_Banco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_Banco", "Nom_Banco", "Tg_Banco where ", TxtCod_Banco, TxtDes_Banco, 1, Me)
    txtCuenta_Cod = ""
    txtCuenta_Des = ""

  End If
End Sub

Private Sub txtCod_Moneda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    'Call Busca_Opcion("Cod_Moneda", "Nom_Moneda", "TG_Moneda where ", txtCod_Moneda, txtDes_Moneda, 1, Me)
    Limpia_Doc
  End If
End Sub

Private Sub txtCod_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        If KeyAscii = 13 Then Call Busca_Opcion("flg_status_letra", "descripcion", "cn_status_letras where flg_planilla_letra='s' and ", txtCod_Origen, txtDes_Origen, 1, Me)

  End If
End Sub

Private Sub txtCod_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    'Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAne, txtDes_TipAnex, 1, Me)
    Limpia_Doc
  End If
End Sub

Private Sub txtCod_TipCobra_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    'Call Busca_Opcion_Store("Cn_Ventas_Muestra_Tipos_Cobranza_Permitidos '" & vusu & "'", txtCod_TipCobra, txtDes_TipCobra, 1, Me)
    Cambio_Apariencia
  End If
End Sub

Private Sub txtCod_TipDocCobra_KeyPress(KeyAscii As Integer)
  'If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipDoc", "Des_TipDoc", " CN_TiposDocum where Flg_Doc_Cobranza = '*' and ", txtCod_TipDocCobra, txtDes_DocCobra, 1, Me)
End Sub

Private Sub txtCuenta_Cod_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    txtCuenta_Cod = Format(txtCuenta_Cod, "000")
    Call Busca_Opcion("Sec_Cuenta_Banco", "cod_cuenta", "Tg_Bancos_Cuentas where Cod_Banco ='" & TxtCod_Banco & "' and ", txtCuenta_Cod, txtCuenta_Des, 1, Me)

  End If
End Sub

Private Sub txtCuenta_Des_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Sec_Cuenta_Banco", "cod_cuenta", "Tg_Bancos_Cuentas where Cod_Banco ='" & TxtCod_Banco & "' and ", txtCuenta_Cod, txtCuenta_Des, 2, Me)
    'Check_Moneda_Cuenta
  End If
End Sub

Private Sub TxtDes_Banco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_Banco", "Nom_Banco", "Tg_Banco where ", TxtCod_Banco, TxtDes_Banco, 2, Me)
    txtCuenta_Cod = ""
    txtCuenta_Des = ""
    'Check_Moneda_Cuenta
  End If
End Sub

Private Sub txtDes_Moneda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    'Call Busca_Opcion("Cod_Moneda", "Nom_Moneda", "TG_Moneda where ", txtCod_Moneda, txtDes_Moneda, 2, Me)
    Limpia_Doc
  End If
End Sub


Private Sub txtDes_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Origen", "Des_Origen", " Cn_Origen where ", txtCod_Origen, txtDes_Origen, 2, Me)
End Sub

Private Sub txtDes_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    'Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 2, Me)
    SendKeys "{TAB}"
    Limpia_Doc
  End If
End Sub

Private Sub txtDes_TipCobra_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    'Call Busca_Opcion_Store("Cn_Ventas_Muestra_Tipos_Cobranza_Permitidos '" & vusu & "'", txtCod_TipCobra, txtDes_TipCobra, 2, Me)
    Cambio_Apariencia
  End If
End Sub

Private Sub txtFec_Diferido_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    txtTipo_Cambio = DevuelveCampo("Select dbo.SM_OBTIENE_TIPO_CAMBIO('" & txtFecha.Text & "')", cCONNECT)
    txtOtro_Tipo_Cambio = DevuelveCampo("Select dbo.SM_OBTIENE_TIPO_CAMBIO_EUROS('" & txtFecha.Text & "')", cCONNECT)
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtNum_DocCobra_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    'Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 1, Me)
    Limpia_Doc
  End If
End Sub

Private Sub txtNumeroPendiente_Change()
  Call gexGrid1.Find(gexGrid1.Columns("Numero").Index, jgexContains, txtNumeroPendiente)
End Sub

Private Sub txtNumeroPendiente_GotFocus()
  SendKeys "{Home}+{End}"
End Sub

Private Sub txtNumeroPendiente_KeyPress(KeyAscii As Integer)
If gexGrid1.RowCount > 0 And KeyAscii = 13 Then
  If Not gexGrid1.Find(gexGrid1.Columns("Numero").Index, jgexContains, txtNumeroPendiente) Or txtNumeroPendiente = "" Then
    MsgBox "Factura no se encuentra en la Lista", vbInformation, "AVISO"
    Exit Sub
  End If
  Call cmdNext_Click
  txtNumeroPendiente = ""
End If
End Sub

Private Sub txtNumeroxCancelar_Change()
  Call gexGrid2.Find(gexGrid2.Columns("Numero").Index, jgexContains, txtNumeroxCancelar)
End Sub

Private Sub txtNumeroxCancelar_KeyPress(KeyAscii As Integer)
If gexGrid2.RowCount > 0 And KeyAscii = 13 Then
  If Not gexGrid2.Find(gexGrid2.Columns("Numero").Index, jgexContains, txtNumeroxCancelar) Or txtNumeroxCancelar = "" Then
    MsgBox "Factura no se encuentra en la Lista", vbInformation, "AVISO"
    Exit Sub
  End If
  Call CmdBack_Click
  txtNumeroxCancelar = ""
End If
End Sub

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtOtro_Tipo_Cambio_LostFocus()
  If txtOtro_Tipo_Cambio = "" Then txtOtro_Tipo_Cambio = 0
End Sub

Private Sub txtSer_DocCobra_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Public Sub CARGA_FACTURAS_PENDIENTES()
On Error GoTo hand
Dim SQL As String

Set RsGrid1 = CreateObject("ADODB.Recordset")
RsGrid1.CursorLocation = adUseClient


SQL = "Ventas_Muestra_Docum_Pedientes_Cobranzas '" & txtCod_TipAne & "','" & strCod_Anxo & "','" & txtCod_Moneda & "','" & txtFecha.Text & "'"
Set RsGrid1 = CargarRecordSetDesconectado(SQL, cCONNECT)

Set gexGrid1.ADORecordset = RsGrid1
'ConfigurarGrid gexGrid1

If RsGrid1.RecordCount Then

    TxtMonto1 = CALCULA_MONTO_TOTAL(RsGrid1)

    Set RsGrid2 = CreateObject("ADODB.Recordset")
    RsGrid2.CursorLocation = adUseClient
    Set RsGrid2.ActiveConnection = Nothing

    RsGrid2.Fields.Append RsGrid1.Fields("Correlativo").Name, RsGrid1.Fields(0).Type, RsGrid1.Fields(0).DefinedSize
    RsGrid2.Fields.Append RsGrid1.Fields("Numero").Name, RsGrid1.Fields(1).Type, RsGrid1.Fields(1).DefinedSize
    RsGrid2.Fields.Append RsGrid1.Fields("Fecha").Name, adDate
    RsGrid2.Fields.Append RsGrid1.Fields("Moneda").Name, RsGrid1.Fields(3).Type, RsGrid1.Fields(3).DefinedSize
    RsGrid2.Fields.Append RsGrid1.Fields("Tipo_Cambio").Name, adDouble
    RsGrid2.Fields.Append RsGrid1.Fields("Monto_Origen1").Name, adDouble
    RsGrid2.Fields.Append RsGrid1.Fields("Monto_Origen").Name, adDouble
    RsGrid2.Fields.Append RsGrid1.Fields("Monto_Aceptado").Name, adDouble
    RsGrid2.Fields.Append RsGrid1.Fields("Cod_Cobranza").Name, RsGrid1.Fields(8).Type, RsGrid1.Fields(8).DefinedSize
    RsGrid2.Fields.Append RsGrid1.Fields("Debe_Haber").Name, RsGrid1.Fields(9).Type, RsGrid1.Fields(9).DefinedSize
    RsGrid2.Fields.Append RsGrid1.Fields("Otro_Tip_Cambio").Name, adDouble
    RsGrid2.Fields.Append RsGrid1.Fields("Observacion").Name, RsGrid1.Fields(11).Type, RsGrid1.Fields(11).DefinedSize
    RsGrid2.Fields.Append RsGrid1.Fields("Tran_TipMonDoc").Name, RsGrid1.Fields(12).Type, RsGrid1.Fields(12).DefinedSize
    RsGrid2.Fields.Append RsGrid1.Fields("Doc_TipMonDoc").Name, RsGrid1.Fields(13).Type, RsGrid1.Fields(13).DefinedSize


    RsGrid2.Open

    Set gexGrid2.ADORecordset = RsGrid2
    'ConfigurarGrid gexGrid2

    TxtMonto2 = 0
End If

txtNumeroPendiente.SetFocus

Exit Sub
Resume
hand:
ErrorHandler err, "CARGA_FACTURAS_PENDIENTES"
End Sub

Sub ConfigurarGrid(mGridEx As GridEx)
    mGridEx.Columns(2).Width = 1305
    mGridEx.Columns(3).Width = 945
    mGridEx.Columns(4).Width = 720
    gexGrid1.Columns(5).Width = 1065
    gexGrid1.Columns(6).Width = 1500
    gexGrid1.Columns(7).Width = 1365
    mGridEx.Columns("Correlativo").Visible = False
    mGridEx.Columns("Monto_Origen1").Visible = False
    mGridEx.Columns("Cod_Cobranza").Visible = False
    mGridEx.Columns("Debe_Haber").Visible = False
    mGridEx.Columns("Tran_TipMonDoc").Visible = False
    mGridEx.Columns("Doc_TipMonDoc").Visible = False
    mGridEx.Columns("Monto_Origen").Format = "###,###.00"
    mGridEx.Columns("Monto_Aceptado").Format = "###,###.00"
End Sub

Private Function CALCULA_MONTO_TOTAL(ByVal mRs As ADODB.Recordset) As String
Dim Monto As Double
Dim I As Integer
    Monto = 0
    mRs.MoveFirst
    For I = 1 To mRs.RecordCount
        Monto = Monto + mRs.Fields("Monto_Aceptado").Value
        mRs.MoveNext
    Next
CALCULA_MONTO_TOTAL = Format(Monto, "###,###.00")
End Function

Sub Resumen_Facturas()

Dim Monto As Double

Dim I As Integer

    Monto = 0
    txtDocumentos = ""

    If gexGrid2.RowCount > 0 Then

      gexGrid2.MoveFirst
      For I = 1 To gexGrid2.RowCount
          Monto = Monto + gexGrid2.Value(gexGrid2.Columns("Monto_Aceptado").Index)
          txtDocumentos = txtDocumentos + " " + gexGrid2.Value(gexGrid2.Columns("Numero").Index)
          gexGrid2.MoveNext
      Next

    End If

txtTotDoc = Format(Monto, "###,###.00")

End Sub

Sub Limpia_Doc()
  Set gexGrid1.ADORecordset = Nothing
  Set gexGrid2.ADORecordset = Nothing
  txtDocumentos.Text = ""
  txtTotDoc.Text = 0
End Sub


Sub Cambio_Apariencia()
On Error Resume Next
  If DevuelveCampo("select Flg_Cobranza_Simple from Cn_Ventas_Tipos_Cobranza where Cod_Tipcobranza ='" & txtCod_TipCobra & "'", cCONNECT) = "N" Then
   If Me.WindowState <> 2 Then Me.Height = 7710
    FunctButt1.Top = 6940
    gridex1.Visible = True
  Else
    If Me.WindowState <> 2 Then Me.Height = 5325
    FunctButt1.Top = 4560
    gridex1.Visible = False
  End If
End Sub

Private Sub txtOtro_Tipo_Cambio_GotFocus()
  SendKeys "{HOME}+{END}"
End Sub

Private Sub txtOtro_Tipo_Cambio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
  SoloNumeros txtOtro_Tipo_Cambio, KeyAscii, True, 4, 3
End Sub

Private Sub txtTipo_Cambio_GotFocus()
  SendKeys "{HOME}+{END}"
End Sub

Private Sub TxtTipo_Cambio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
  SoloNumeros txtTipo_Cambio, KeyAscii, True, 4, 3
End Sub

Private Sub txtTipo_Cambio_LostFocus()
  If txtTipo_Cambio = "" Then txtTipo_Cambio = 0
End Sub
