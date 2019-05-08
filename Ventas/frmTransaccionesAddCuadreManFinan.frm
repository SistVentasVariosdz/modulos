VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmTransaccionesAddCuadreManFinan 
   ClientHeight    =   3480
   ClientLeft      =   1545
   ClientTop       =   1200
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3480
   ScaleWidth      =   9135
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   8895
      Begin VB.TextBox txtTipo_Cambio_Otros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7080
         MaxLength       =   8
         TabIndex        =   8
         Text            =   "0"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtNro_Doc_Otros 
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   3
         Top             =   840
         Width           =   1905
      End
      Begin VB.TextBox txtNum_Cuota 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         TabIndex        =   4
         Text            =   "0"
         Top             =   840
         Width           =   825
      End
      Begin VB.TextBox txtCod_Moneda 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1320
         Width           =   480
      End
      Begin VB.TextBox TxtTipo_Cambio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5040
         MaxLength       =   8
         TabIndex        =   7
         Text            =   "0"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtImp_Convertido 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   8
         TabIndex        =   9
         Text            =   "0"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   5
         Text            =   "0"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtDes_Cobranza 
         Height          =   285
         Left            =   2520
         TabIndex        =   2
         Top             =   360
         Width           =   6105
      End
      Begin VB.TextBox txtCod_Cobranza 
         Height          =   285
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   1
         Top             =   360
         Width           =   600
      End
      Begin VB.TextBox txtObservacion 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Top             =   2280
         Width           =   6945
      End
      Begin VB.Label Label2 
         Caption         =   "Observacion :"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2295
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto Cobranza:"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   375
         Width           =   1455
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Monto Aceptado :"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   1365
         Width           =   1275
      End
      Begin VB.Label Label15 
         Caption         =   "Monto Origen :"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1815
         Width           =   1095
      End
      Begin VB.Label lbFinan 
         AutoSize        =   -1  'True
         Caption         =   "Nro Financiamiento  :"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   885
         Width           =   1500
      End
      Begin VB.Label Label4 
         Caption         =   "T./C. Otro:"
         Height          =   255
         Left            =   6240
         TabIndex        =   16
         Top             =   1335
         Width           =   855
      End
      Begin VB.Label lbCuota 
         AutoSize        =   -1  'True
         Caption         =   "Nro Cuota  :"
         Height          =   195
         Left            =   3840
         TabIndex        =   15
         Top             =   885
         Width           =   855
      End
      Begin VB.Label Label27 
         Caption         =   "T./C.:"
         Height          =   255
         Left            =   4560
         TabIndex        =   14
         Top             =   1335
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Moneda :"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   1335
         Width           =   735
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3120
      TabIndex        =   11
      Top             =   2880
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmTransaccionesAddCuadreManFinan.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmTransaccionesAddCuadreManFinan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dFecha As Date, intSecuencia As Integer, intSecuencia_Det As Integer, StrOption As String, _
       lfAceptar As Boolean, strCod_Moneda As String, strNum_Corre As String, strStore_Carga As String, strStore_Man As String, _
      strNum_Corre_Otros As String, intNum_Transaccion As String
Public codigo As String, Descripcion As String, strTipo_Det As String, strCod_Anexo As String, strCod_TipAnexo, strStore As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim sMessage As Long
Dim StrSql As String

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

errores Err.Number
End Sub

Private Function lfSalvar_Datos() As Boolean

On Error GoTo hand

Dim SQL As String, dbTipoCambio As Double

SQL = "Tm_Ventas_Transacciones_Cobranzas_Detalle_Man_Finan '" & StrOption & "'," & intNum_Transaccion & ",'" & dFecha & "'," _
       & intSecuencia & "," & intSecuencia_Det & ",'" & txtCod_Cobranza.Text & "','" & "" & "','" & "" & "'," _
       & Format(txtImporte.Text, "######.00") & ",'" _
       & TxtObservacion & "'," & TxtTipo_Cambio & ",'S'," & txtTipo_Cambio_Otros & ",'" & strNum_Corre_Otros & " '," & txtNum_Cuota
         
Call ExecuteCommandSQL(cCONNECT, SQL)

lfSalvar_Datos = True

Exit Function
Resume
hand:

errores Err.Number

lfSalvar_Datos = False

End Function

Private Sub txtCod_Cobranza_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Concepto_Cobranza", "Descripcion", "Cn_Ventas_Conceptos_Cobranza Where Flg_Financiamientos = 'S' and ", txtCod_Cobranza, txtDes_Cobranza, 1, Me)
End Sub

Private Sub txtDes_Cobranza_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Concepto_Cobranza", "Descripcion", "Cn_Ventas_Conceptos_Cobranza Where Flg_Financiamientos = 'S' and ", txtCod_Cobranza, txtDes_Cobranza, 2, Me)
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

Private Sub txtNro_Doc_Otros_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Busca_Opcion_Cuota "TM_Muestra_Cuotas_Financiamientos_Cobranzas " & intNum_Transaccion & ",'" & dFecha & "','" & txtCod_Cobranza & "'", txtNro_Doc_Otros, txtNum_Cuota, 1, Me
End Sub

Private Sub txtNum_Cuota_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Busca_Opcion_Cuota "TM_Muestra_Cuotas_Financiamientos_Cobranzas " & intNum_Transaccion & ",'" & dFecha & "','" & txtCod_Cobranza & "'", txtNro_Doc_Otros, txtNum_Cuota, 2, Me
End Sub

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtTipo_Cambio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
    Calcula_Importe_Converido
  End If
  SoloNumeros TxtTipo_Cambio, KeyAscii, True, 4, 2
End Sub

Public Sub Calcula_Importe_Converido()
If strCod_Moneda = txtCod_Moneda Then
  txtImp_Convertido.Text = txtImporte.Text
ElseIf TxtTipo_Cambio.Text <> 0 Then
  If txtCod_Moneda = "SOL" Then
    txtImp_Convertido.Text = txtImporte.Text * TxtTipo_Cambio.Text
  Else
    txtImp_Convertido.Text = txtImporte.Text / TxtTipo_Cambio.Text
  End If
End If
End Sub

Private Sub TxtTipo_Cambio_LostFocus()
  Calcula_Importe_Converido
End Sub


Sub Busca_Opcion_Cuota(strStore As String, txtCod As TextBox, txtDes As TextBox, Opcion As Integer, frmME As Form)

On Error GoTo Fin

Dim rstAux As ADODB.Recordset

    txtCod = ""
    txtDes = ""
    With frmBusqGeneral
        Set .oParent = frmME
        .SQuery = strStore
        .Cargar_Datos
        
        codigo = ""
        .DGridLista.Columns("Num_Corre").Visible = False
        .DGridLista.Columns("Tipo_Cambio").Visible = False
        .DGridLista.Columns("Num_Cuota").Width = 930
        .DGridLista.Columns("Nro_Financiamiento").Width = 1500
        .DGridLista.Columns("Moneda").Width = 765
        .DGridLista.Columns("Importe_Pendiente").Format = "###,###.00"
        Set rstAux = .DGridLista.ADORecordset
        
        If rstAux.RecordCount > 1 Then
          .Show vbModal
        Else
          frmME.codigo = ".."
        End If
        If frmME.codigo <> "" And rstAux.RecordCount > 0 Then
            strNum_Corre_Otros = Trim(rstAux!Num_Corre)
            txtImporte = Format(rstAux!Importe_Pendiente, "###,###.00")
            txtCod_Moneda = Trim(rstAux!Moneda)
            TxtTipo_Cambio = Format(rstAux!Tipo_Cambio, "###,###.0000")
            txtDes = Trim(rstAux!Num_Cuota)
            txtCod = Trim(rstAux!Nro_Financiamiento)
            Calcula_Importe_Converido
            Select Case Opcion
            Case 1: SendKeys "{TAB}": SendKeys "{TAB}"
            Case 2: SendKeys "{TAB}"
            End Select
        Else
            SendKeys "{TAB}"
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Resume
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & Opcion & ")"
End Sub

Private Sub txtTipo_Cambio_Otros_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
    Calcula_Importe_Converido
  End If
  SoloNumeros txtTipo_Cambio_Otros, KeyAscii, True, 4, 2
End Sub
