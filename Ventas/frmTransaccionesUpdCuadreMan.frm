VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmTransaccionesUpdCuadreMan 
   ClientHeight    =   2670
   ClientLeft      =   1095
   ClientTop       =   1800
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2670
   ScaleWidth      =   10680
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   10455
      Begin VB.TextBox txtCod_Moneda 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5610
         MaxLength       =   8
         TabIndex        =   9
         Top             =   960
         Width           =   480
      End
      Begin VB.TextBox TxtTipo_Cambio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3600
         MaxLength       =   8
         TabIndex        =   8
         Text            =   "0"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtImp_Convertido 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7560
         TabIndex        =   10
         Text            =   "0"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtOtro_Tipo_Cambio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9120
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   915
         Width           =   825
      End
      Begin VB.Frame frImporte 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   860
         Width           =   2535
         Begin VB.TextBox txtImporte 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1440
            TabIndex        =   7
            Text            =   "0"
            Top             =   75
            Width           =   975
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Monto Aceptado :"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   1275
         End
      End
      Begin VB.TextBox txtObservacion 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   1440
         Width           =   6945
      End
      Begin VB.Frame frDocumento 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   8655
         Begin VB.TextBox txtSer_Docum 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   5400
            MaxLength       =   3
            TabIndex        =   2
            Top             =   240
            Width           =   600
         End
         Begin VB.TextBox txtCod_TipDoc 
            Height          =   285
            Left            =   1440
            MaxLength       =   4
            TabIndex        =   0
            Top             =   240
            Width           =   600
         End
         Begin VB.TextBox txtDes_TipDoc 
            Height          =   285
            Left            =   2160
            TabIndex        =   1
            Top             =   240
            Width           =   2505
         End
         Begin VB.TextBox txtNum_Docum 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   7440
            MaxLength       =   8
            TabIndex        =   3
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label12 
            Caption         =   "Tipo Documento :"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   255
            Width           =   1335
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Serie :"
            Height          =   195
            Left            =   4800
            TabIndex        =   18
            Top             =   285
            Width           =   450
         End
         Begin VB.Label Label5 
            Caption         =   "Numero :"
            Height          =   255
            Left            =   6270
            TabIndex        =   17
            Top             =   255
            Width           =   735
         End
      End
      Begin VB.Frame frAnticipo 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   8415
         Begin VB.TextBox txtNro_Anticipo 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   4
            Top             =   0
            Width           =   1080
         End
         Begin VB.Label Label6 
            Caption         =   "Nro Anticipo :"
            Height          =   255
            Left            =   0
            TabIndex        =   25
            Top             =   15
            Width           =   975
         End
      End
      Begin VB.Frame frConcepto 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   8535
         Begin VB.TextBox txtCod_Cobranza 
            Height          =   285
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   5
            Top             =   240
            Width           =   600
         End
         Begin VB.TextBox txtDes_Cobranza 
            Height          =   285
            Left            =   2400
            TabIndex        =   6
            Top             =   240
            Width           =   2505
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto Cobranza:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   255
            Width           =   1455
         End
      End
      Begin VB.Label Label15 
         Caption         =   "Monto Origen :"
         Height          =   255
         Left            =   6240
         TabIndex        =   29
         Top             =   975
         Width           =   1095
      End
      Begin VB.Label Label27 
         Caption         =   "T./C.:"
         Height          =   255
         Left            =   3000
         TabIndex        =   28
         Top             =   975
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Moneda :"
         Height          =   255
         Left            =   4800
         TabIndex        =   27
         Top             =   975
         Width           =   735
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Otro Tipo Cambio"
         Height          =   195
         Left            =   9000
         TabIndex        =   26
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "Observacion :"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1455
         Width           =   1095
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3960
      TabIndex        =   13
      Top             =   2040
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmTransaccionesUpdCuadreMan.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmTransaccionesUpdCuadreMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dFecha As Date, intSecuencia As Integer, intSecuencia_Det As Integer, strOption As String, _
       lfAceptar As Boolean, strCod_Moneda As String, strNum_Corre As String, strStore_Carga As String, strStore_Man As String
Public codigo As String, Descripcion As String, strTipo_Det As String, strCod_Anexo As String, strCod_TipAnexo, strStore As String

Private Sub Imp_Flete_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

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

Dim SQL As String, dbTipoCambio As Double

If frConcepto.Visible Then dbTipoCambio = DevuelveCampo("Select dbo.Sm_Obtiene_Tipo_Cambio_Venta('" & dFecha & "')", cCONNECT) Else dbTipoCambio = TxtTipo_Cambio.Text

Select Case strTipo_Det

Case Is = "1"
  SQL = strStore & " '" & strOption & "','" & dFecha & "'," & intSecuencia & "," _
        & intSecuencia_Det & ",'" & txtCod_Cobranza.Text & "','" & "" & "','" & strNum_Corre & "'," _
        & IIf(frDocumento.Visible, txtImp_Convertido.Text, txtImporte.Text) & ",'" _
        & TxtObservacion & "'," & dbTipoCambio & "," & txtOtro_Tipo_Cambio
Case Is = "2"
  SQL = strStore & " '" & strOption & "','" & dFecha & "'," & intSecuencia & ",'" & strCod_TipAnexo & "','" _
      & strCod_Anexo & "'," & txtNro_Anticipo.Text & "," & txtImporte.Text & "," & TxtTipo_Cambio.Text
Case Is = "3"
  If InStr(strStore, "Notas_Abono") <> 0 Then
    SQL = strStore & " '" & strOption & "','" & dFecha & "'," & intSecuencia & ",'" & strNum_Corre & "'," _
          & txtImporte.Text & "," & TxtTipo_Cambio.Text & ",'" & vusu & "','" & txtOtro_Tipo_Cambio.Text & "'"
    
  Else
    SQL = strStore & " '" & strOption & "','" & dFecha & "'," & intSecuencia & ",'" & strNum_Corre & "'," _
          & txtImporte.Text & "," & TxtTipo_Cambio.Text & ",'" & vusu & "'"
  End If
Case Is = "4"
  SQL = strStore & " '" & strOption & "','" & dFecha & "'," & intSecuencia & "," _
        & intSecuencia_Det & ",'" & txtCod_Cobranza.Text & "','" & "" & "','" & "" & "'," _
        & IIf(frDocumento.Visible, txtImp_Convertido.Text, txtImporte.Text) & ",'" _
        & TxtObservacion & "'," & dbTipoCambio
        
End Select
      
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
  With frmTransaccionesUpdCuadreManAde
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
  With frmTransaccionesUpdCuadreManDoc
    .Caption = Me.Caption
    .txtCod_TipAne = strCod_TipAnexo
    .strCod_Anxo = strCod_Anexo
    .strStore_Carga = strStore_Carga
    .strCod_Moneda = strCod_Moneda
    .dFecha = dFecha
    .iSecuencia = intSecuencia
    .CARGA_GRID
    .Show 1
  End With
End If

End Sub

Private Sub txtCod_Cobranza_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Concepto_Cobranza", "Descripcion", "Cn_Ventas_Conceptos_Cobranza where Flg_Seleccionable = 'S' and ", txtCod_Cobranza, txtDes_Cobranza, 1, Me)
End Sub

Private Sub txtCod_TipDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If txtCod_TipDoc.Text = "" Then
      Carga_Busqueda
  Else
    Call Busca_Opcion("Cod_TipDoc", "Des_TipDoc", "CN_TiposDocum where Flg_Doc_Ventas = '*' and ", txtCod_TipDoc, txtDes_TipDoc, 1, Me)
  End If
End If
End Sub

Private Sub txtDes_Cobranza_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Concepto_Cobranza", "Descripcion", "Cn_Ventas_Conceptos_Cobranza where Flg_Seleccionable = 'S' and ", txtCod_Cobranza, txtDes_Cobranza, 2, Me)
End Sub

Private Sub txtDes_TipDoc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_TipDoc", "Des_TipDoc", "CN_TiposDocum where Flg_Doc_Ventas = '*' and ", txtCod_TipDoc, txtDes_TipDoc, 2, Me)
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

Private Sub txtImporte_LostFocus()
  Calcula_Importe_Converido
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
  End If
  If KeyAscii >= 48 And KeyAscii <= 57 _
    Or KeyAscii = 8 Or KeyAscii = 13 Then Else KeyAscii = 0
End Sub

Private Sub txtNum_Docum_LostFocus()
  txtNum_Docum = Format(txtNum_Docum, "00000000")
End Sub

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtSer_Docum_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
  If KeyAscii >= 48 And KeyAscii <= 57 _
    Or KeyAscii = 8 Or KeyAscii = 13 Then Else KeyAscii = 0
End Sub

Private Sub txtSer_Docum_LostFocus()
  txtSer_Docum = Format(txtSer_Docum, "000")
End Sub

Private Sub TxtTipo_Cambio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
    Calcula_Importe_Converido
  End If
  SoloNumeros TxtTipo_Cambio, KeyAscii, True, 4, 2
End Sub

Sub Ecuentra_Documento()

On Error GoTo Fin

Dim rstAux As Object

Set rstAux = CreateObject("ADODB.Recordset")
Dim strSQL As String

strSQL = "Ventas_Muestra_Doc_Cobranza_Detalle '" & dFecha & "'," & intSecuencia & ",'" & txtCod_TipDoc & "','" & txtSer_Docum & "','" & txtNum_Docum & "'"

Set rstAux = CargarRecordSetDesconectado(strSQL, cCONNECT)

With rstAux
  If Not (.BOF And .EOF) Then
    strNum_Corre = !Num_Corre
    txtCod_Moneda.Text = !Cod_Moneda
    TxtTipo_Cambio.Text = !Tipo_Cambio
    
    If strCod_Moneda = txtCod_Moneda Then
      txtImporte.Text = !Imp_Pendiente
    Else
      If !Cod_Moneda = "SOL" Then
        txtImporte.Text = !Imp_Pendiente / !Tipo_Cambio
      Else
        txtImporte.Text = !Imp_Pendiente * !Tipo_Cambio
      End If
    End If
    
    Calcula_Importe_Converido
  Else
    strNum_Corre = ""
    txtCod_Moneda.Text = ""
    txtImporte.Text = 0
    TxtTipo_Cambio.Text = 0
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

If Trim(strNum_Corre) <> "" Then
  txtImp_Convertido = DevuelveCampo("select dbo.Convierte_Importe_Moneda_Destino('" & strCod_Moneda & "'," & txtImporte.Text & "," & TxtTipo_Cambio.Text & ",'" & txtCod_Moneda & "','" & dFecha & "'," & txtOtro_Tipo_Cambio & ")", cCONNECT)
Else
  txtImp_Convertido = txtImporte
End If

End Sub

Private Sub txtTipo_Cambio_LostFocus()
  Calcula_Importe_Converido
End Sub
