VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmTransaccionesAddCuadreMan 
   ClientHeight    =   2070
   ClientLeft      =   1050
   ClientTop       =   1470
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2070
   ScaleWidth      =   10830
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3240
      TabIndex        =   7
      Top             =   1440
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmTransaccionesAddCuadreMan.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   10455
      Begin VB.TextBox txtOtro_Tipo_Cambio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9105
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   720
         Width           =   825
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   5
         Text            =   "0"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TxtTipo_Cambio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3480
         MaxLength       =   8
         TabIndex        =   6
         Text            =   "0"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtCod_Moneda 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5490
         MaxLength       =   8
         TabIndex        =   16
         Top             =   720
         Width           =   480
      End
      Begin VB.TextBox txtImp_Convertido 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7320
         MaxLength       =   8
         TabIndex        =   15
         Text            =   "0"
         Top             =   720
         Width           =   1215
      End
      Begin VB.Frame frDocumento 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   8535
         Begin VB.TextBox txtSer_Docum 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   5400
            MaxLength       =   3
            TabIndex        =   2
            Top             =   120
            Width           =   600
         End
         Begin VB.TextBox txtCod_TipDoc 
            Height          =   285
            Left            =   1440
            MaxLength       =   4
            TabIndex        =   0
            Top             =   120
            Width           =   600
         End
         Begin VB.TextBox txtDes_TipDoc 
            Height          =   285
            Left            =   2160
            TabIndex        =   1
            Top             =   120
            Width           =   2505
         End
         Begin VB.TextBox txtNum_Docum 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   7200
            MaxLength       =   8
            TabIndex        =   3
            Top             =   120
            Width           =   1200
         End
         Begin VB.Label Label12 
            Caption         =   "Tipo Documento :"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   135
            Width           =   1335
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Serie :"
            Height          =   195
            Left            =   4800
            TabIndex        =   11
            Top             =   165
            Width           =   450
         End
         Begin VB.Label Label5 
            Caption         =   "Numero :"
            Height          =   255
            Left            =   6270
            TabIndex        =   10
            Top             =   135
            Width           =   735
         End
      End
      Begin VB.Frame frAnticipo 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   8415
         Begin VB.TextBox txtNro_Anticipo 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   4
            Top             =   120
            Width           =   1080
         End
         Begin VB.Label Label1 
            Caption         =   "Nro Anticipo :"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   135
            Width           =   975
         End
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Otro Tipo Cambio"
         Height          =   195
         Left            =   8880
         TabIndex        =   22
         Top             =   285
         Width           =   1230
      End
      Begin VB.Label Label15 
         Caption         =   "Monto Origen :"
         Height          =   255
         Left            =   6120
         TabIndex        =   18
         Top             =   735
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Moneda :"
         Height          =   255
         Left            =   4680
         TabIndex        =   17
         Top             =   735
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "T./C.:"
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   735
         Width           =   495
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Monto Aceptado :"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   765
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmTransaccionesAddCuadreMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dFecha As Date, intSecuencia As Integer, strOption As String, _
       lfAceptar As Boolean, strNum_Corre As String, strCod_TipAnex As String, strCod_Anexo As String, _
       strStore As String, intNum_Transaccion As Long, strCod_Moneda As String, strStore_Carga As String
Public codigo As String, Descripcion As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim sMessage As Long
Dim strSQL As String

On Error GoTo dprError

Select Case ActionName
  Case "GRABAR"
    If MsgBox("Esta seguro de Actualizar un Concepto ", vbYesNo, "IMPORTANTE") = vbYes Then
      If lfSalvar_Datos Then
        Unload Me
        lfAceptar = True
      End If
    End If
  Case "CANCELAR"
      Unload Me
      lfAceptar = False
End Select

Exit Sub

dprError:

errores err.Number
End Sub
Private Function lfSalvar_Datos() As Boolean

On Error GoTo hand

Dim SQL As String

If Not frDocumento.Visible Then
  SQL = strStore & " '" & strOption & "'," & intNum_Transaccion & ",'" & dFecha & "'," _
      & intSecuencia & ",'" & strCod_TipAnex & "','" & strCod_Anexo & "'," & txtNro_Anticipo.Text & "," _
      & txtImporte.Text & "," & txtTipo_Cambio.Text & "," & txtOtro_Tipo_Cambio
Else
  SQL = strStore & " '" & strOption & "'," & intNum_Transaccion & ",'" & dFecha & "'," _
      & intSecuencia & ",'" & strNum_Corre & "'," & txtImporte.Text & "," & txtTipo_Cambio.Text & "," & txtOtro_Tipo_Cambio
End If
      
Call ExecuteCommandSQL(cCONNECT, SQL)

lfSalvar_Datos = True

Exit Function
Resume
hand:

errores err.Number

lfSalvar_Datos = False

End Function

Sub Carga_Busqueda()
If frAnticipo.Visible Then
  With frmTransaccionesAddCuadreManAde
    .Caption = Me.Caption
    .txtCod_TipAne = strCod_TipAnexo
    .strCod_Anxo = strCod_Anexo
    .strStore_Carga = strStore_Carga
    .strCod_Moneda = strCod_Moneda
    .dFecha = dFecha
    .CARGA_GRID
    .Show 1
  End With
Else
  With frmTransaccionesAddCuadreManDoc
    .Caption = Me.Caption
    .txtCod_TipAne = strCod_TipAnexo
    .strCod_Anxo = strCod_Anexo
    .strStore_Carga = strStore_Carga
    .strCod_Moneda = strCod_Moneda
    .dFecha = dFecha
    .CARGA_GRID
    .Show 1
  End With
End If
End Sub

Private Sub txtCod_TipDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If txtCod_TipDoc.Text = "" Then
      Carga_Busqueda
  Else
    Call Busca_Opcion("Cod_TipDoc", "Des_TipDoc", "CN_TiposDocum where Flg_Doc_Ventas = '*' and ", txtCod_TipDoc, txtDes_TipDoc, 1, Me)
'    Ecuentra_Documento
  End If
End If
End Sub

Private Sub txtDes_TipDoc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_TipDoc", "Des_TipDoc", "CN_TiposDocum where Flg_Doc_Ventas = '*' and ", txtCod_TipDoc, txtDes_TipDoc, 2, Me)
'    Ecuentra_Documento
  End If
End Sub

Private Sub txtImporte_Change()
  If txtImporte = "" Then txtImporte = 0
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
    Calcula_Importe_Converido
  End If
  SoloNumeros txtImporte, KeyAscii, True, 2, 9
End Sub

Private Sub txtNro_Anticipo_KeyPress(KeyAscii As Integer)
  If txtNro_Anticipo.Text = "" Then Carga_Busqueda
  If KeyAscii = 13 Then SendKeys "{TAB}"
  If KeyAscii >= 48 And KeyAscii <= 57 _
    Or KeyAscii = 8 Or KeyAscii = 13 Then Else KeyAscii = 0
End Sub

Private Sub txtNum_Docum_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
'    Ecuentra_Documento
  End If
  If KeyAscii >= 48 And KeyAscii <= 57 _
    Or KeyAscii = 8 Or KeyAscii = 13 Then Else KeyAscii = 0
End Sub

Private Sub txtNum_Docum_LostFocus()
'  txtNum_Docum = Format(txtNum_Docum, "00000000")
End Sub

Private Sub txtSer_Docum_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
'    Ecuentra_Documento
  End If
  If KeyAscii >= 48 And KeyAscii <= 57 _
    Or KeyAscii = 8 Or KeyAscii = 13 Then Else KeyAscii = 0
End Sub

Private Sub txtSer_Docum_LostFocus()
  txtSer_Docum = Format(txtSer_Docum, "000")
End Sub

Private Sub TxtTipo_Cambio_Change()
  If txtTipo_Cambio = "" Then txtTipo_Cambio = 0
End Sub

Private Sub TxtTipo_Cambio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
    Calcula_Importe_Converido
  End If
  SoloNumeros txtTipo_Cambio, KeyAscii, True, 4, 2
End Sub

Sub Ecuentra_Documento()

On Error GoTo Fin

Dim rstAux As Object
Set rstAux = CreateObject("ADODB.Recordset")
Dim strSQL As String

strSQL = "Ventas_Muestra_Doc_Pagos_Detalle '" & strCod_TipAnex & "','" & strCod_Anexo & "','" & txtCod_TipDoc & "','" & txtSer_Docum & "','" & txtNum_Docum & "','" & dFecha & "'"

Set rstAux = CargarRecordSetDesconectado(strSQL, cCONNECT)

With rstAux
  If Not (.BOF And .EOF) Then
    strNum_Corre = !Num_Corre
    txtCod_Moneda.Text = !Cod_Moneda
    txtImporte.Text = !Imp_Pendiente
    txtTipo_Cambio.Text = !Tipo_Cambio
  Else
    strNum_Corre = ""
    txtCod_Moneda.Text = ""
    txtImporte.Text = 0
    txtTipo_Cambio.Text = 0
  End If
  .Close
End With

Set rstAux = Nothing
    
Exit Sub
Resume
Fin:
On Error Resume Next
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, "Búsqueda de Descuento (" & opcion & ")"
End Sub

Public Sub Calcula_Importe_Converido()

  txtImp_Convertido = DevuelveCampo("select dbo.Convierte_Importe_Moneda_Destino('" & strCod_Moneda & "'," & txtImporte.Text & "," & txtTipo_Cambio.Text & ",'" & txtCod_Moneda & "','" & dFecha & "'," & txtOtro_Tipo_Cambio & ")", cCONNECT)

End Sub

