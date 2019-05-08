VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmTransacciones 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Registro de Transacciones de Cobranzas"
   ClientHeight    =   7665
   ClientLeft      =   285
   ClientTop       =   720
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   13230
   Begin VB.TextBox txtNum_Ruc 
      Height          =   285
      Left            =   1950
      TabIndex        =   16
      Top             =   1200
      Width           =   1380
   End
   Begin VB.TextBox txtDes_Anexo 
      Height          =   285
      Left            =   3420
      TabIndex        =   17
      Top             =   1200
      Width           =   4245
   End
   Begin VB.TextBox txtCod_Anxo 
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   8895
      MaxLength       =   4
      TabIndex        =   19
      Top             =   1200
      Width           =   750
   End
   Begin VB.TextBox txtCod_TipAnex 
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   8520
      MaxLength       =   1
      TabIndex        =   18
      Top             =   1200
      Width           =   315
   End
   Begin VB.OptionButton OptRangoFecha 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Rango  Fecha"
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
      Height          =   225
      Left            =   120
      TabIndex        =   13
      Top             =   885
      Width           =   1620
   End
   Begin VB.OptionButton optParteCobranza 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Num. Parte"
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
      Height          =   225
      Left            =   90
      TabIndex        =   12
      Top             =   480
      Width           =   1350
   End
   Begin VB.OptionButton optFecha 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fecha"
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
      Height          =   225
      Left            =   90
      TabIndex        =   11
      Top             =   120
      Value           =   -1  'True
      Width           =   1155
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   555
      Left            =   12000
      TabIndex        =   4
      Top             =   120
      Width           =   1065
   End
   Begin VB.TextBox txtNum_Parte 
      Height          =   285
      Left            =   1950
      TabIndex        =   3
      Top             =   465
      Width           =   1410
   End
   Begin VB.TextBox TxtDes_Banco 
      Height          =   285
      Left            =   6420
      TabIndex        =   8
      Top             =   465
      Width           =   4575
   End
   Begin VB.TextBox TxtCod_Banco 
      Height          =   285
      Left            =   5670
      TabIndex        =   7
      Top             =   465
      Width           =   615
   End
   Begin VB.TextBox txtDes_Origen 
      Height          =   285
      Left            =   6420
      TabIndex        =   2
      Top             =   105
      Width           =   1575
   End
   Begin VB.TextBox txtCod_Origen 
      Height          =   285
      Left            =   5670
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "N"
      Top             =   105
      Width           =   375
   End
   Begin NumBoxProject.NumBox inpFec_Emi 
      Height          =   285
      Left            =   1950
      TabIndex        =   0
      Top             =   105
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   503
      TypeVal         =   3
      Mask            =   "99/99/9999"
      Formato         =   "dd/MM/yyyy"
      AllowedMask     =   -1
      MaskLen         =   10
      Aling           =   2
      Text            =   ""
      CanEmpty        =   -1
      ShowError       =   0
      Locked          =   0   'False
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DecimalNumber   =   0
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5340
      Left            =   0
      TabIndex        =   6
      Top             =   1635
      Width           =   13200
      _ExtentX        =   23283
      _ExtentY        =   9419
      Version         =   "2.0"
      RecordNavigator =   -1  'True
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmTransacciones.frx":0000
      Column(2)       =   "frmTransacciones.frx":00C8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmTransacciones.frx":016C
      FormatStyle(2)  =   "frmTransacciones.frx":02A4
      FormatStyle(3)  =   "frmTransacciones.frx":0354
      FormatStyle(4)  =   "frmTransacciones.frx":0408
      FormatStyle(5)  =   "frmTransacciones.frx":04E0
      FormatStyle(6)  =   "frmTransacciones.frx":0598
      FormatStyle(7)  =   "frmTransacciones.frx":0678
      FormatStyle(8)  =   "frmTransacciones.frx":0724
      ImageCount      =   0
      PrinterProperties=   "frmTransacciones.frx":07D4
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   675
      Left            =   30
      TabIndex        =   5
      Top             =   6975
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   1191
      Custom          =   $"frmTransacciones.frx":09AC
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1150
      ControlHeigth   =   650
      ControlSeparator=   75
   End
   Begin NumBoxProject.NumBox inpFec_EmiIni 
      Height          =   285
      Left            =   1950
      TabIndex        =   14
      Top             =   840
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   503
      TypeVal         =   3
      Mask            =   "99/99/9999"
      Formato         =   "dd/MM/yyyy"
      AllowedMask     =   -1
      MaskLen         =   10
      Aling           =   2
      Text            =   ""
      CanEmpty        =   -1
      ShowError       =   0
      Locked          =   0   'False
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DecimalNumber   =   0
   End
   Begin NumBoxProject.NumBox inpFec_EmiFin 
      Height          =   285
      Left            =   3435
      TabIndex        =   15
      Top             =   840
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   503
      TypeVal         =   3
      Mask            =   "99/99/9999"
      Formato         =   "dd/MM/yyyy"
      AllowedMask     =   -1
      MaskLen         =   10
      Aling           =   2
      Text            =   ""
      CanEmpty        =   -1
      ShowError       =   0
      Locked          =   0   'False
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DecimalNumber   =   0
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cliente"
      Height          =   195
      Left            =   1320
      TabIndex        =   21
      Top             =   1245
      Width           =   480
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Anexo:"
      Height          =   210
      Left            =   7920
      TabIndex        =   20
      Tag             =   "Document Type"
      Top             =   1237
      Width           =   555
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   3840
      Top             =   240
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Banco:"
      Height          =   255
      Left            =   4920
      TabIndex        =   10
      Top             =   480
      Width           =   525
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Origen :"
      Height          =   195
      Left            =   4920
      TabIndex        =   9
      Top             =   150
      Width           =   555
   End
End
Attribute VB_Name = "frmTransacciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public codigo As String, Descripcion As String
Public sTipoBusq As String

Private Sub cmdBuscar_Click()
If OptRangoFecha.Value = True Then
    If Not IsDate(inpFec_EmiIni.Text) Or Not IsDate(inpFec_EmiFin.Text) Then
      MsgBox "Rango de fecha invalido no se ingreso una de las fechas", vbInformation, "AVISO"
      Exit Sub
    End If
    If CDate(inpFec_EmiIni.Text) >= CDate(inpFec_EmiFin.Text) Then
      MsgBox "Rango de fecha invalido fecha inicial mayor a la fecha final", vbInformation, "AVISO"
      Exit Sub
    End If
'    If DateDiff("y", CDate(inpFec_EmiIni.Text), CDate(inpFec_EmiFin.Text)) > 3 Then
'      MsgBox "Rango de fecha invalido maximo 3 meses", vbInformation, "AVISO"
'      Exit Sub
'    End If

End If
  Buscar
End Sub
Sub Buscar()

Dim strSQL
On Error GoTo errores

strSQL = "Ventas_Muestra_Transacciones_Cobranzas '" & txtCod_Origen & "','" & inpFec_Emi.Text & "','" & TxtCod_Banco & "','" & txtNum_Parte & "','" & vusu & "','" & sTipoBusq & "','" & inpFec_EmiIni.Text & "','" & inpFec_EmiFin.Text & "','" & txtCod_TipAnex.Text & "','" & txtCod_Anxo.Text & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

Dim colTemp As JSColumn

GridEX1.ColumnHeaderHeight = 500

GridEX1.Columns("Fecha").Width = 1020
GridEX1.Columns("Secuencia").Width = 390
GridEX1.Columns("Num_Parte_Cobranza").Width = 750
GridEX1.Columns("Num_Parte_Cobranza").Caption = "Parte Cobranza"
GridEX1.Columns("Secuencia").Caption = "Sec"
GridEX1.Columns("Origen").Visible = False
GridEX1.Columns("Des_Origen").Visible = False
GridEX1.Columns("Des_Cobranza").Caption = "Tipo Cobranza"
GridEX1.Columns("Cod_TipCobranza").Visible = False
GridEX1.Columns("Nro_Cuenta").Width = 1920
GridEX1.Columns("Nro_Cuenta").Caption = "Cuenta"
GridEX1.Columns("Des_Cobranza").Width = 1500
GridEX1.Columns("Cod_TipAnex").Visible = False
GridEX1.Columns("Cod_AnxCon").Visible = False
GridEX1.Columns("Num_Transaccion").Caption = "Num Transaccion"
GridEX1.Columns("Num_Transaccion").Width = 960
GridEX1.Columns("Des_Anexo").Width = 3240
GridEX1.Columns("Des_Anexo").Caption = "Cliente"
GridEX1.Columns("Des_Anexo").Width = 3030
GridEX1.Columns("Cod_Banco").Visible = False
GridEX1.Columns("Banco").Width = 1635
GridEX1.Columns("Cod_TipDoc_Cobranza").Visible = False
GridEX1.Columns("Tipo_Doc_Cobranza").Width = 2235
GridEX1.Columns("Serie").Width = 450
GridEX1.Columns("Numero").Width = 1065
GridEX1.Columns("Moneda").Width = 765
GridEX1.Columns("Observacion").Width = 1065
GridEX1.Columns("Total_Debe").Width = 975
GridEX1.Columns("Total_Haber").Width = 1020
GridEX1.Columns("Num_Ruc").Visible = False
GridEX1.Columns("Nom_Moneda").Visible = False
GridEX1.Columns("Fec_Pago_Prog_Diferido").Visible = False
GridEX1.Columns("Sec_Cuenta_Banco").Visible = False

Exit Sub
Resume
errores:
    errores err.Number
End Sub

Private Sub Form_Activate()
inpFec_Emi.SetFocus
End Sub

Private Sub Form_Load()

Dim origen As String

If DevuelveCampo("Select max(Fec_transaccion) from Cn_Ventas_Partes_Cobranza ", cCONNECT) <> "NULL" Then
    inpFec_Emi.Text = DevuelveCampo("Select max(Fec_transaccion) from Cn_Ventas_Partes_Cobranza ", cCONNECT)

End If

'inpFec_Emi.Text = DevuelveCampo("Select max(Fec_transaccion) from Cn_Ventas_Partes_Cobranza ", cCONNECT)
origen = DevuelveCampo("select isnull(Origen,'') from cn_ventas_control_usuario where cod_usuario = '" & vusu & "'", cCONNECT)

FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name) & "/SALIR"

Encuentra_Parte

sTipoBusq = "1"

If origen <> "*" Then
  txtCod_Origen.Text = origen
  Call txtCod_Origen_KeyPress(13)
  txtCod_Origen.Enabled = False
  txtDes_Origen.Enabled = False
End If

End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

Dim varSecuencia As Integer

On Error GoTo hand

Select Case ActionName
  Case Is = "AGREGAR"
    With frmTransaccionesAdd
      .strOption = "I"
      .txtFecha.Text = inpFec_Emi.Text
      If Not txtCod_Origen.Enabled Then
        .txtCod_Origen.Enabled = False
        .txtDes_Origen.Enabled = False
      End If
      
      .txtCod_Origen = txtCod_Origen
      .txtDes_Origen = txtDes_Origen
      .Show 1
      If .lfSalvar Then
        Buscar
      Else
        FunctButt1.SetFocus
      End If
    End With
    
  Case Is = "MODIFICAR"
    If GridEX1.RowCount = 0 Then Exit Sub
    With frmTransaccionesUpd
      .txtCod_TipCobra.Enabled = False
      .txtDes_TipCobra.Enabled = False
      .txtFecha.Text = GridEX1.Value(GridEX1.Columns("Fecha").Index)
      .txtFecha.Enabled = False
      .intSecuencia = GridEX1.Value(GridEX1.Columns("Secuencia").Index)
      .txtCod_TipCobra.Text = GridEX1.Value(GridEX1.Columns("Cod_TipCobranza").Index)
      .txtDes_TipCobra.Text = GridEX1.Value(GridEX1.Columns("Des_Cobranza").Index)
      .txtCod_Origen.Text = GridEX1.Value(GridEX1.Columns("Origen").Index)
      .txtDes_Origen.Text = GridEX1.Value(GridEX1.Columns("Des_Origen").Index)
      .txtNum_Ruc.Text = GridEX1.Value(GridEX1.Columns("Num_Ruc").Index)
      .txtCod_TipAne.Text = GridEX1.Value(GridEX1.Columns("Cod_TipAnex").Index)
      .txtDes_TipAne.Text = GridEX1.Value(GridEX1.Columns("Des_Anexo").Index)
      .txtCod_Moneda.Text = GridEX1.Value(GridEX1.Columns("Moneda").Index)
      .TxtCod_Banco.Text = GridEX1.Value(GridEX1.Columns("Cod_Banco").Index)
      .TxtDes_Banco.Text = GridEX1.Value(GridEX1.Columns("Banco").Index)
      .txtCuenta_Cod.Text = GridEX1.Value(GridEX1.Columns("Sec_Cuenta_Banco").Index)
      .txtCuenta_Des.Text = GridEX1.Value(GridEX1.Columns("Nro_Cuenta").Index)
      .txtCod_TipDocCobra.Text = GridEX1.Value(GridEX1.Columns("Cod_TipDoc_Cobranza").Index)
      .txtDes_DocCobra.Text = GridEX1.Value(GridEX1.Columns("Tipo_Doc_Cobranza").Index)
      .txtSer_DocCobra.Text = GridEX1.Value(GridEX1.Columns("Serie").Index)
      .txtNum_DocCobra.Text = GridEX1.Value(GridEX1.Columns("Numero").Index)
      .TxtObservacion.Text = GridEX1.Value(GridEX1.Columns("Observacion").Index)
      .chkDiferido = IIf(GridEX1.Value(GridEX1.Columns("Flg_Diferido").Index) = "S", 1, 0)
      If .chkDiferido Then
        .txtFec_Diferido.Visible = True
        .lbDiferido.Visible = True
        .txtFec_Diferido.Text = GridEX1.Value(GridEX1.Columns("Fec_Pago_Prog_Diferido").Index)
      End If
      .strOption = "U"
      varSecuencia = GridEX1.Value(GridEX1.Columns("Secuencia").Index)
      If .lfAceptar = False Then .Show 1
      Buscar
      Call GridEX1.Find(GridEX1.Columns("Secuencia").Index, jgexEqual, varSecuencia)
    End With
  Case Is = "VERDETALLE"
    If GridEX1.RowCount = 0 Then Exit Sub
    With frmTransaccionesUpdCuadre
      .Caption = "Conceptos de " & GridEX1.Value(GridEX1.Columns("Des_Cobranza").Index) & " del Cliente " & GridEX1.Value(GridEX1.Columns("Des_Anexo").Index)
      .strSQL = "Ventas_Muestra_Detalle_Cobranzas '" & GridEX1.Value(GridEX1.Columns("Fecha").Index) & "','" & GridEX1.Value(GridEX1.Columns("Secuencia").Index) & "'"
      .strCod_TipAnexo = GridEX1.Value(GridEX1.Columns("Cod_TipAnex").Index)
      .strCod_Anexo = GridEX1.Value(GridEX1.Columns("Cod_AnxCon").Index)
      .dFecha = GridEX1.Value(GridEX1.Columns("Fecha").Index)
      .intSecuencia = GridEX1.Value(GridEX1.Columns("Secuencia").Index)
      .strCod_Moneda = GridEX1.Value(GridEX1.Columns("Moneda").Index)
      .CARGA_GRID
      varSecuencia = GridEX1.Value(GridEX1.Columns("Secuencia").Index)
      .Show 1
      Buscar
      Call GridEX1.Find(GridEX1.Columns("Secuencia").Index, jgexEqual, varSecuencia)
    End With
  Case Is = "IMPRESION"
    Reporte
  Case Is = "ABRIRCERRAR"
    With frmTransaccionesStatus
      .txtCod_Origen = txtCod_Origen
      .txtDes_Origen = txtDes_Origen
      .txtFecha_Cierre.Text = DevuelveCampo("select isnull(max(Fec_Transaccion),getdate()) from CN_VENTAS_PARTES_COBRANZA where Origen = '" & txtCod_Origen & "'", cCONNECT)
      .txtFecha_Nuevo.Text = Date
      .Show 1
      If .lfAceptar Then
        Call inpFec_Emi_KeyPress(13)
        Buscar
      End If
    End With
  Case Is = "ABRRIRDIAR"
      With frmTransaccionesStatusReversion
      .txtCod_Origen = txtCod_Origen
      .txtDes_Origen = txtDes_Origen
      .Show 1
      Buscar
    End With
  Case Is = "ELIMINAR"
    If GridEX1.RowCount = 0 Then Exit Sub
    
    If DevuelveCampo("select count(*) from Cn_Ventas_Transacciones_Cobranzas_Detalle where Fec_Transaccion = '" & GridEX1.Value(GridEX1.Columns("Fecha").Index) & "' and secuencia = " & GridEX1.Value(GridEX1.Columns("Secuencia").Index), cCONNECT) > 0 Then
      MsgBox "Elimine Primero el Detalle de la Transaccion", vbInformation, "IMPORTATEN"
      Exit Sub
    End If
    
    If MsgBox("Esta seguro de Eliminar esta Transaccion", vbYesNo, "IMPORTANTE") = vbYes Then
      lvSql = "CN_VENTAS_TRANSACCIONES_COBRANZAS_MAN 'D','" & GridEX1.Value(GridEX1.Columns("Fecha").Index) & "'," _
              & GridEX1.Value(GridEX1.Columns("Secuencia").Index) & ",'" & GridEX1.Value(GridEX1.Columns("Cod_TipCobranza").Index) & "','" _
              & GridEX1.Value(GridEX1.Columns("Cod_TipAnex").Index) & "','" & GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index) & "','" _
              & GridEX1.Value(GridEX1.Columns("Cod_Banco").Index) & "','" & GridEX1.Value(GridEX1.Columns("Sec_Cuenta_Banco").Index) & "','" _
              & GridEX1.Value(GridEX1.Columns("Cod_TipDoc_Cobranza").Index) & "','" _
              & GridEX1.Value(GridEX1.Columns("Serie").Index) & "','" & GridEX1.Value(GridEX1.Columns("Numero").Index) & "','" _
              & GridEX1.Value(GridEX1.Columns("Moneda").Index) & "','','" & vusu & "','" & ComputerName & "','" _
              & GridEX1.Value(GridEX1.Columns("Origen").Index) & "','" & GridEX1.Value(GridEX1.Columns("Flg_Diferido").Index) & "','S'," _
              & IIf(IsNull(GridEX1.Value(GridEX1.Columns("Fec_Pago_Prog_Diferido").Index)), "NULL", "'" & GridEX1.Value(GridEX1.Columns("Fec_Pago_Prog_Diferido").Index) & "'")
      Call ExecuteCommandSQL(cCONNECT, lvSql)
      Buscar
    End If
  Case Is = "CHEQUEDIFERIDO"
      If GridEX1.RowCount = 0 Then Exit Sub
      With frmTransaccionesDetalle_Cheque_Diferido
        .dFecha = GridEX1.Value(GridEX1.Columns("Fecha").Index)
        .intSecuencia = GridEX1.Value(GridEX1.Columns("Secuencia").Index)
        .Caption = UCase("Aplicacion del Cheques Diferidos Banco : " & GridEX1.Value(GridEX1.Columns("Banco").Index) & " Cuenta :" & GridEX1.Value(GridEX1.Columns("Nro_Cuenta").Index))
        .Buscar
        .Show 1
      End With
  Case Is = "TRANSFINAN"
      Trans_Finanzas
  Case Is = "VERVOUCHER"
        If GridEX1.RowCount = 0 Then Exit Sub
        MuestraVoucher2
  Case Is = "SALIR"
    Unload Me
End Select

Exit Sub
Resume
hand:

errores err.Number

End Sub

Sub Trans_Finanzas()

On Error GoTo dprDepurar

Dim sSql As String, iSecuencia As Integer

If GridEX1.RowCount = 0 Then Exit Sub

If MsgBox("Esta seguro de Transferir esta transaccion ha Finanzas", vbYesNo + vbCritical, "AVISO") = vbYes Then
  iSecuencia = GridEX1.Value(GridEX1.Columns("Secuencia").Index)
  sSql = "Fi_Genera_Movimiento_Finanza '" & GridEX1.Value(GridEX1.Columns("Fecha").Index) & "'," & GridEX1.Value(GridEX1.Columns("Secuencia").Index) & ",'S'"
  ExecuteCommandSQL cCONNECT, sSql
  Buscar
  Call GridEX1.Find(GridEX1.Columns("Secuencia").Index, jgexEqual, iSecuencia)
  MsgBox "La Transferencia se hizo satisfactoriamente", vbInformation, "AVISO"
End If

Exit Sub

dprDepurar:
  
  errores err.Number

End Sub
Public Sub Reporte()
  
On Error GoTo ErrorImpresion

VB.Screen.MousePointer = vbHourglass
Dim sempresas As String
Dim oo As Object, strSQL As String, rs As Object, RS1 As Object
Set rs = CreateObject("ADODB.Recordset")
Set RS1 = CreateObject("ADODB.Recordset")
Set oo = CreateObject("excel.application")

strSQL = "SELECT Des_Empresa from seg_empresas where Cod_Empresa = '" & vemp & "'"
sempresas = Trim(DevuelveCampo(strSQL, cSEGURIDAD))

strSQL = "Cn_Ventas_Emision_Parte_Cobranzas '" & txtCod_Origen & "','" & txtNum_Parte & "'"

Set rs = CargarRecordSetDesconectado(strSQL, cCONNECT)

strSQL = "CN_VENTAS_EMITE_PARTE_COBRANZAS_RESUMEN1 '" & txtCod_Origen & "','" & txtNum_Parte & "'"

Set RS1 = CargarRecordSetDesconectado(strSQL, cCONNECT)

If rs.RecordCount = 0 Then
  Screen.MousePointer = vbNormal
  MsgBox "No hay Registros que imprimir", vbInformation, "AVISO"
  Exit Sub
End If

oo.Workbooks.Open vRuta & "\rptParteCobranza.xlt"
oo.Run "REPORTE", rs, RS1, inpFec_Emi.Text, txtNum_Parte, txtCod_Origen, cCONNECT, txtDes_Origen, sempresas

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



Private Sub inpFec_Emi_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
    Encuentra_Parte
  End If
End Sub

Private Sub inpFec_Emi_LostFocus()
  Encuentra_Parte
End Sub

Private Sub inpFec_EmiIni_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
    'Encuentra_Parte
  End If
End Sub

Private Sub inpFec_EmiIni_LostFocus()
  'Encuentra_Parte
  If Len(inpFec_EmiIni.Text) > 0 Then OptRangoFecha.Value = True
End Sub

Private Sub inpFec_EmiFin_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
    'Encuentra_Parte
  End If
End Sub

Private Sub inpFec_EmiFin_LostFocus()
  'Encuentra_Parte
  If Len(inpFec_EmiIni.Text) > 0 Then OptRangoFecha.Value = True
End Sub


Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
  txtCod_Anxo.Text = ""
  txtCod_TipAnex.Text = ""
  txtDes_Anexo.Text = ""
  If KeyAscii = 13 And Len(Trim(txtNum_Ruc.Text)) > 0 Then
    BUSCARUC 1
  ElseIf Len(Trim(txtNum_Ruc.Text)) = 0 Then
    SendKeys "{TAB}"
  End If
End Sub
Private Sub txtNum_Ruc_GotFocus()
    SelectionText txtNum_Ruc
End Sub

Private Sub txtDes_Anexo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(txtDes_Anexo.Text) <> "" Then
            If Len(Trim(txtDes_Anexo)) > 2 Then
                Call BUSCA_ANEXO(2, 1)
            Else
                Aviso "Debe ingresar al menos 3 caracteres del Nombre requerido", 1
                Exit Sub
            End If
        Else
            cmdBuscar.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub BUSCARUC(Opcion As Integer)

On Error GoTo Fin
Dim strSQL As String
Dim oTipo As New frmBusqGeneral

    strSQL = "SELECT num_ruc as Ruc,Des_Anexo Descripcion FROM CN_AnexosContables "
    txtNum_Ruc = Trim(txtNum_Ruc)
    
    strSQL = strSQL & " where num_ruc like '%" & txtNum_Ruc & "%' and Cod_TipAnex <>'P'"
    
    txtNum_Ruc = ""
        
    Set oTipo.oParent = Me
    
    oTipo.sQuery = strSQL
    oTipo.Cargar_Datos
    oTipo.DGridLista.Columns(1).Width = 4350.047
    oTipo.Show 1
    If codigo <> "" Then
      txtNum_Ruc = Trim(codigo)
      txtDes_Anexo = Trim(Descripcion)
      
      strSQL = "SELECT Cod_TipAnEx FROM CN_AnexosContables WHERE num_ruc = '" & txtNum_Ruc.Text & "' and Cod_TipAnex <>'P'"
      txtCod_TipAnex.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
      strSQL = "SELECT Cod_Anxo FROM CN_AnexosContables WHERE num_ruc = '" & txtNum_Ruc.Text & "' and Cod_TipAnex <>'P'"
      txtCod_Anxo.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
      
      SendKeys "{TAB}"
    End If
    Set oTipo = Nothing
    
Exit Sub
Resume
Fin:
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda (" & Opcion & ")"
End Sub
Sub BUSCA_ANEXO(Tipo As Integer, Ubic As Integer)
Dim strSQL As String
Dim iLen As Integer
    Select Case Tipo
        Case 2:
        
                Dim oTipo As New frmBusqGeneral
                Dim rs As Object
                Set rs = CreateObject("ADODB.Recordset")
                Set oTipo.oParent = Me
                If Ubic = 1 Then
                    oTipo.sQuery = "SELECT Cod_Anxo as Código, Des_Anexo as Descripción FROM CN_AnexosContables WHERE Cod_TipAnEX <> 'P' AND Des_Anexo like '%" & Trim(txtDes_Anexo.Text) & "%'"
                Else
                End If
                oTipo.Cargar_Datos
                oTipo.Top = txtDes_Anexo.Top + txtDes_Anexo.Height
                oTipo.Left = txtDes_Anexo.Left
                oTipo.DGridLista.Columns(1).Width = 4800
                oTipo.Show 1
                If codigo <> "" Then
                    If Ubic = 1 Then
                        txtCod_Anxo.Text = Trim(codigo)
                        txtDes_Anexo.Text = Trim(Descripcion)
                        strSQL = "SELECT num_ruc FROM CN_AnexosContables WHERE Cod_TipAnEX <> 'P' AND Cod_Anxo = '" & txtCod_Anxo.Text & "'"
                        txtNum_Ruc = Trim(DevuelveCampo(strSQL, cCONNECT))

                        strSQL = "SELECT Cod_TipAnEX FROM CN_AnexosContables WHERE Cod_TipAnEX <> 'P' AND Cod_Anxo = '" & txtCod_Anxo.Text & "'"
                        txtCod_TipAnex.Text = Trim(DevuelveCampo(strSQL, cCONNECT))

                        cmdBuscar.SetFocus
                    Else
                    End If
                End If
                Set oTipo = Nothing
                Set rs = Nothing
                
    End Select
    
End Sub

Private Sub optFecha_Click()
    sTipoBusq = "1"
End Sub

Private Sub optParteCobranza_Click()
    sTipoBusq = "2"
End Sub

Private Sub OptRangoFecha_Click()
    sTipoBusq = "3"
End Sub
Private Sub TxtCod_Banco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Banco", "Nom_Banco", "Tg_Banco where ", TxtCod_Banco, TxtDes_Banco, 1, Me)
End Sub

Private Sub TxtDes_Banco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_Banco", "Nom_Banco", "Tg_Banco where ", TxtCod_Banco, TxtDes_Banco, 2, Me)
    Encuentra_Parte
  End If
End Sub

Private Sub txtDes_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Origen", "Des_Origen", " Cn_Origen where ", txtCod_Origen, txtDes_Origen, 2, Me)
End Sub

Private Sub txtCod_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Origen", "Des_Origen", " Cn_Origen where ", txtCod_Origen, txtDes_Origen, 1, Me)
    Encuentra_Parte
  End If
End Sub

Private Sub txtNum_Parte_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Sub Encuentra_Parte()
  txtNum_Parte.Text = DevuelveCampo("Select isnull(MAX(Num_Parte_Cobranza),'') from Cn_Ventas_Partes_Cobranza where Fec_Transaccion = '" & inpFec_Emi.Text & "' and Origen = '" & txtCod_Origen & "'  ", cCONNECT)
End Sub

Private Sub MuestraVoucher2()

On Error GoTo errx
Dim sSql As String
Dim rsAsientos As ADODB.Recordset


If GridEX1.RowCount = 0 Then Exit Sub

sSql = "FI_Muestra_Data_Asientos_Cobranzas '$' ,'$'"
sSql = VBsprintf(sSql, GridEX1.Value(GridEX1.Columns("fecha").Index), GridEX1.Value(GridEX1.Columns("secuencia").Index))

Set rsAsientos = GetDataSet(cCONNECT, sSql)

With rsAsientos
  
  If .BOF Or .EOF Then
    MsgBox "No se le ha Generado Voucher", vbInformation, "AVISO"
    Exit Sub
  End If

  Load frmShowVoucher
  frmShowVoucher.sCod_TipoDiario = !Cod_TipoDiario
  frmShowVoucher.sano = !Ano_Contable
  frmShowVoucher.smes = !Mes_Contable
  frmShowVoucher.lNum_Registro = !Num_Registro
  frmShowVoucher.sFec_Transaccion = GridEX1.Value(GridEX1.Columns("fecha").Index)
  frmShowVoucher.sSecuencia = GridEX1.Value(GridEX1.Columns("secuencia").Index)
  'frmShowVoucher.Num_Corre = GridEX1.Value(GridEX1.Columns("Num_Corre").Index)
  'frmShowVoucher.dImporte = GridEX1.Value(GridEX1.Columns("Imp_Total").Index)
  'frmShowVoucher.sFlg_Status = GridEX1.Value(GridEX1.Columns("Estatus_Letra").Index)
  frmShowVoucher.Buscar
  frmShowVoucher.FunctButt1.ChangeProperty "ENABLED", 1, False
  frmShowVoucher.Show vbModal
  Set frmShowVoucher = Nothing
  
End With

Set rsAsientos = Nothing

Exit Sub

errx:
    errores err.Number

End Sub



