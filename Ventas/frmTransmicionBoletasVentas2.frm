VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransmicionBoletasVentas 
   Caption         =   "Transmision de Boletas del Almacen 61 de Saldos"
   ClientHeight    =   8715
   ClientLeft      =   330
   ClientTop       =   735
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   10080
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBoletas 
      Caption         =   "Transmisión Boletas de Venta - Credito"
      Height          =   3840
      Left            =   2220
      TabIndex        =   14
      Top             =   1410
      Visible         =   0   'False
      Width           =   5385
      Begin VB.TextBox TxtCod_Banco 
         Height          =   285
         Left            =   1725
         TabIndex        =   27
         Top             =   2145
         Width           =   735
      End
      Begin VB.TextBox TxtDes_Banco 
         Height          =   285
         Left            =   2490
         TabIndex        =   26
         Top             =   2145
         Width           =   2415
      End
      Begin VB.TextBox txtCuenta_Des 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   25
         Top             =   2565
         Width           =   2370
      End
      Begin VB.TextBox txtCuenta_Cod 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   1725
         MaxLength       =   11
         TabIndex        =   24
         Top             =   2565
         Width           =   720
      End
      Begin VB.ComboBox Cbo 
         Height          =   315
         Left            =   1740
         TabIndex        =   23
         Top             =   1665
         Width           =   3510
      End
      Begin MSComCtl2.DTPicker dtpBoletasCredito 
         Height          =   300
         Left            =   1785
         TabIndex        =   15
         Top             =   315
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   23920641
         CurrentDate     =   38338
      End
      Begin FunctionsButtons.FunctButt FunctButt4 
         Height          =   510
         Left            =   1500
         TabIndex        =   16
         Top             =   3090
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmTransmicionBoletasVentas2.frx":0000
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin MSComCtl2.DTPicker dtpFecFinAutoriz 
         Height          =   300
         Left            =   1755
         TabIndex        =   18
         Top             =   1245
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   23920641
         CurrentDate     =   38338
      End
      Begin MSComCtl2.DTPicker dtpFecIniAutoriz 
         Height          =   300
         Left            =   1770
         TabIndex        =   20
         Top             =   795
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   23920641
         CurrentDate     =   38338
      End
      Begin VB.Label Label8 
         Caption         =   "Banco:"
         Height          =   255
         Left            =   135
         TabIndex        =   29
         Top             =   2115
         Width           =   645
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta :"
         Height          =   195
         Left            =   150
         TabIndex        =   28
         Top             =   2595
         Width           =   600
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo de Movim:"
         Height          =   390
         Left            =   120
         TabIndex        =   22
         Top             =   1710
         Width           =   1500
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Autorización Inicial:"
         Height          =   360
         Left            =   135
         TabIndex        =   21
         Top             =   780
         Width           =   1605
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha Autorización Final:"
         Height          =   390
         Left            =   120
         TabIndex        =   19
         Top             =   1245
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Transacción:"
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Top             =   375
         Width           =   1425
      End
   End
   Begin VB.Frame fraRevertir 
      Caption         =   "Reversión de Asientos Tienda"
      Height          =   1830
      Left            =   3360
      TabIndex        =   10
      Top             =   5475
      Visible         =   0   'False
      Width           =   3225
      Begin FunctionsButtons.FunctButt FunctButt3 
         Height          =   510
         Left            =   360
         TabIndex        =   11
         Top             =   1050
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmTransmicionBoletasVentas2.frx":0097
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   1350
         TabIndex        =   13
         Top             =   555
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   23920641
         CurrentDate     =   38338
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   570
         TabIndex        =   12
         Top             =   630
         Width           =   495
      End
   End
   Begin VB.Frame fraTransmitir 
      Caption         =   "Transmisión de Asientos Tienda"
      Height          =   1830
      Left            =   3420
      TabIndex        =   6
      Top             =   3390
      Visible         =   0   'False
      Width           =   3225
      Begin MSComCtl2.DTPicker dtpTransmitir 
         Height          =   300
         Left            =   1275
         TabIndex        =   7
         Top             =   570
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   23920641
         CurrentDate     =   38338
      End
      Begin FunctionsButtons.FunctButt FunctOKCancel 
         Height          =   510
         Left            =   360
         TabIndex        =   9
         Top             =   1050
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmTransmicionBoletasVentas2.frx":012E
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   570
         TabIndex        =   8
         Top             =   630
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   9855
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   7935
         TabIndex        =   1
         Top             =   300
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   900
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   491
         ControlSeparator=   110
      End
      Begin MSComCtl2.DTPicker dtpFecInicio 
         Height          =   300
         Left            =   2160
         TabIndex        =   0
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   23920641
         CurrentDate     =   38338
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Boletas de Saldos Hasta:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   540
         Width           =   1785
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   6300
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   11113
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
      RowHeaders      =   -1  'True
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmTransmicionBoletasVentas2.frx":01C5
      Column(2)       =   "frmTransmicionBoletasVentas2.frx":028D
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmTransmicionBoletasVentas2.frx":0331
      FormatStyle(2)  =   "frmTransmicionBoletasVentas2.frx":0469
      FormatStyle(3)  =   "frmTransmicionBoletasVentas2.frx":0519
      FormatStyle(4)  =   "frmTransmicionBoletasVentas2.frx":05CD
      FormatStyle(5)  =   "frmTransmicionBoletasVentas2.frx":06A5
      FormatStyle(6)  =   "frmTransmicionBoletasVentas2.frx":075D
      FormatStyle(7)  =   "frmTransmicionBoletasVentas2.frx":083D
      FormatStyle(8)  =   "frmTransmicionBoletasVentas2.frx":08E9
      ImageCount      =   0
      PrinterProperties=   "frmTransmicionBoletasVentas2.frx":0999
   End
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   930
      Left            =   600
      TabIndex        =   5
      Top             =   7650
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   1640
      Custom          =   $"frmTransmicionBoletasVentas2.frx":0B71
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   900
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   -120
      Top             =   1920
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmTransmicionBoletasVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public codigo As String, Descripcion As String

Sub Transmitir()

On Error GoTo ErrTransmitir

    Screen.MousePointer = vbHourglass
    
    Call ExecuteCommandSQL(cCONNECT, "Ventas_Transmite_Boletas_Saldos '" & vusu & "','" & dtpFecInicio & "'")
    
    Screen.MousePointer = vbDefault
    
    MsgBox "La Transmision se ha llevado con exito", vbInformation, "AVISO"
    
    Exit Sub
    
ErrTransmitir:

    Screen.MousePointer = vbDefault

    ErrorHandler err, "Transmision"
End Sub

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Form_Load()
Dim sSeguridad  As String

  sSeguridad = get_botones1(Me, vper, vemp, Me.Name)

  Me.FunctButt2.FunctionsUser = sSeguridad
  

  dtpFecInicio = Date - 2
  dtpTransmitir = Date - 2
  
  dtpBoletasCredito = Date
  dtpFecIniAutoriz = Date
  dtpFecFinAutoriz = Date
  FillMovs
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

Select Case ActionName
Case Is = "BUSCAR"
  Buscar
End Select

End Sub

Private Sub Buscar()

Dim sSQL As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle

sSQL = "Ventas_Muesta_Boletas_Por_Transmitir '" & dtpFecInicio & "'"

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)

GridEX1.Columns("Tip").Width = 360
GridEX1.Columns("Serie").Width = 495
GridEX1.Columns("Nro_Doc").Width = 795
GridEX1.Columns("Anexo").Width = 3150
GridEX1.Columns("Ruc").Width = 1095
GridEX1.Columns("Moneda").Width = 720
GridEX1.Columns("Emision").Width = 945
GridEX1.Columns("Tipo_Cambio").Width = 1065
GridEX1.Columns("Imp_Neto").Width = 865
GridEX1.Columns("IGV").Width = 660
GridEX1.Columns("Imp_Total").Width = 840
  
End Sub



Sub TransmitirAsientosTienda()

On Error GoTo ErrTransmitir

    Screen.MousePointer = vbHourglass
    
    Call ExecuteCommandSQL(cCONNECT, "Ventas_Captura_Cobranzas_Boletas_Venta '" & dtpTransmitir.Value & "','" & vusu & "','" & ComputerName & "'")
    
    Screen.MousePointer = vbDefault
    
    MsgBox "La Transmision de Asientos de Tienda se ha llevado con exito", vbInformation, "AVISO"
    
    Exit Sub
    
ErrTransmitir:

    Screen.MousePointer = vbDefault

    ErrorHandler err, "Transmision Asiento"
End Sub

Sub RevertirAsientosTienda()

On Error GoTo ErrTransmitir

    Screen.MousePointer = vbHourglass
    
    Call ExecuteCommandSQL(cCONNECT, "Ventas_Revierte_Captura_Cobranzas_Boletas_Venta '" & dtpTransmitir.Value & "','" & vusu & "','" & ComputerName & "'")
    
    Screen.MousePointer = vbDefault
    
    MsgBox "La Reversión de Asientos de Tienda se ha llevado con exito", vbInformation, "AVISO"
    
    Exit Sub
    
ErrTransmitir:

    Screen.MousePointer = vbDefault

    ErrorHandler err, "Reversión Asiento"
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case Is = "TRANSMITIR"
  'If GridEX1.RowCount = 0 Then
  '  MsgBox "NO REGISTROS HA TRANSMITIR", vbInformation, "AVISO"
  '  Exit Sub
  'End If
  Transmitir
  Buscar
Case Is = "TRANSMITEASIENTOS"
    Me.fraTransmitir.Visible = True
    Me.fraRevertir.Visible = False
    Me.fraBoletas.Visible = False
Case Is = "REVIERTEASIENTOS"
    Me.fraTransmitir.Visible = False
    Me.fraRevertir.Visible = True
    Me.fraBoletas.Visible = False
Case Is = "AUTORIZABOLETAS"
    Me.fraBoletas.Visible = True
    Me.fraTransmitir.Visible = False
    Me.fraRevertir.Visible = False
Case Is = "GENERARTRANSBV"
    'If GridEX1.RowCount = 0 Then
    '   MsgBox "NO EXISTEN REGISTROS HA TRANSMITIR", vbInformation, "AVISO"
    '   Exit Sub
    'End If
    If MsgBox("Esta seguro de generar la transmisión de datos", vbInformation + vbYesNo, "AVISO") = vbYes Then
        Call GenerarTransmisionBoleta
    End If
Case Is = "RECIBIRTRANSBV"
    frmRecepcionInformacionBoletas.Show 1
Case Is = "SALIR"
  Unload Me
End Select
End Sub


Sub GenerarTransmisionBoleta()

On Error GoTo errores
 Dim sql As String
 Dim RutaTienda As String
 
 RutaTienda = DevuelveCampo(" select ruta_datos_tienda from vt_control   ", cCONNECT)
 
  
  If ExecuteSQL(cCONNECT, "Ventas_Transmite_Boletas_Saldos_Tienda '" & vusu & "','" & dtpFecInicio & "'") = -1 Then
  
      EjecBatch (RutaTienda & "\TRANS_V.BAT") ', vbMaximizedFocus
      EjecBatch (RutaTienda & "\TRANS_VA.BAT") ', vbMaximizedFocus
      EjecBatch (RutaTienda & "\TRANS_VI.BAT") ', vbMaximizedFocus
  
     MsgBox "El texto fue generado con éxito", vbInformation, "Generación"
    
      sql = " exec VT_BORRA_INFORMACION_BOLETA_ENVIO_TIENDA "
  
      If ExecuteSQL(cCONNECT, sql) = -1 Then
         MsgBox "Proceso culminado satisfactoriamente", vbInformation, "Eliminación"
      End If
  Else
   MsgBox "Ocurrio un imprevisto , informe a sistemas del problema", vbCritical, "Aviso"
  End If
  Call Buscar
Exit Sub

errores:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub

Private Sub FunctButt3_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            RevertirAsientosTienda
        Case "CANCELAR"
            Me.fraRevertir.Visible = False
    End Select

End Sub

Private Sub FunctButt4_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            TransmitirBoletasCredito
        Case "CANCELAR"
            Me.fraBoletas.Visible = False
    End Select
    
End Sub

Private Sub FunctOKCancel_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            TransmitirAsientosTienda
        Case "CANCELAR"
            Me.fraTransmitir.Visible = False
    End Select
End Sub


Sub TransmitirBoletasCredito()

On Error GoTo ErrTransmitir

    Screen.MousePointer = vbHourglass
    
    If Cbo.ListIndex = -1 Then Exit Sub
    
    Call ExecuteCommandSQL(cCONNECT, "Ventas_Captura_Cobranzas_Boletas_Venta_CREDITO '" & dtpBoletasCredito.Value & "','" & vusu & "','" & ComputerName & "','" & dtpFecIniAutoriz.Value & "','" & dtpFecFinAutoriz.Value & "','" & Left(Cbo.Text, 3) & "','" & TxtCod_Banco.Text & "','" & txtCuenta_Cod & "'")
    
    Screen.MousePointer = vbDefault
    
    MsgBox "La Transmision de Asientos de Tienda se ha llevado con exito", vbInformation, "AVISO"
    Me.fraBoletas.Visible = False
    Exit Sub
    
ErrTransmitir:

    Screen.MousePointer = vbDefault

    ErrorHandler err, "Transmision Asiento"
End Sub


Private Sub FillMovs()

Dim rstAux As ADODB.Recordset
Dim StrSql As String
    
StrSql = "select cod_tipmov , des_tipmov from lg_tiposmov where cod_tipmov in ('VS1','VS2') order by 1 "
         
Set rstAux = CargarRecordSetDesconectado(StrSql, cCONNECT)
Cbo.Clear
With rstAux
    If .RecordCount > 0 Then .MoveFirst
    Do Until .EOF
        Cbo.AddItem !Cod_TipMov & " " & !Des_Tipmov
        .MoveNext
    Loop
    .Close
End With
If Cbo.ListCount > 0 Then Cbo.ListIndex = 0
Set rstAux = Nothing
    
End Sub



Private Sub TxtCod_Banco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_Banco", "Nom_Banco", "Tg_Banco where ", TxtCod_Banco, TxtDes_Banco, 1, Me)
    txtCuenta_Cod = ""
    txtCuenta_Des = ""
    txtCuenta_Cod.SetFocus
  End If
End Sub


Private Sub txtCuenta_Cod_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    txtCuenta_Cod = Format(txtCuenta_Cod, "000")
    Call Busca_Opcion("Sec_Cuenta_Banco", "cod_cuenta", "Tg_Bancos_Cuentas where Cod_Banco ='" & TxtCod_Banco & "' and ", txtCuenta_Cod, txtCuenta_Des, 1, Me)
    FunctButt4.SetFocus
  End If
End Sub

Private Sub txtCuenta_Des_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Sec_Cuenta_Banco", "cod_cuenta", "Tg_Bancos_Cuentas where Cod_Banco ='" & TxtCod_Banco & "' and ", txtCuenta_Cod, txtCuenta_Des, 2, Me)
    
  End If
End Sub

Private Sub TxtDes_Banco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_Banco", "Nom_Banco", "Tg_Banco where ", TxtCod_Banco, TxtDes_Banco, 2, Me)
    txtCuenta_Cod = ""
    txtCuenta_Des = ""
    
  End If
End Sub

