VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmAdelantosAdd 
   ClientHeight    =   4170
   ClientLeft      =   675
   ClientTop       =   1635
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   7005
   Begin VB.Frame Frame1 
      Height          =   3405
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   6735
      Begin VB.Frame fraAnticipoOtraMoneda 
         Height          =   960
         Left            =   90
         TabIndex        =   21
         Top             =   2355
         Visible         =   0   'False
         Width           =   6525
         Begin NumBoxProject.NumBox txtAnticipo_Moneda_Deposito 
            Height          =   285
            Left            =   3405
            TabIndex        =   22
            Top             =   315
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            TypeVal         =   2
            Mask            =   "9,999,999,999.99"
            Formato         =   "#,###,###,###.##"
            AllowedMask     =   -1
            MaskLen         =   10
            Aling           =   3
            Text            =   "0.00"
            CanEmpty        =   -1
            ShowError       =   0
            Locked          =   0   'False
            Enabled         =   -1  'True
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DecimalNumber   =   2
         End
         Begin NumBoxProject.NumBox txtTipoCambioNegociado 
            Height          =   285
            Left            =   1110
            TabIndex        =   25
            Top             =   315
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            TypeVal         =   2
            Mask            =   "9,999,999,999.99"
            Formato         =   "#,###,###,###.##"
            AllowedMask     =   -1
            MaskLen         =   10
            Aling           =   3
            Text            =   "0.00"
            CanEmpty        =   -1
            ShowError       =   0
            Locked          =   0   'False
            Enabled         =   -1  'True
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DecimalNumber   =   2
         End
         Begin VB.Label lblMonedaTransaccion 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   4710
            TabIndex        =   26
            Top             =   345
            Width           =   45
         End
         Begin VB.Label Label6 
            Caption         =   "T/Cambio Dolares Negociado"
            Height          =   570
            Left            =   150
            TabIndex        =   24
            Top             =   240
            Width           =   990
         End
         Begin VB.Label Label4 
            Caption         =   "Importe en Moneda Depósito:"
            Height          =   570
            Left            =   2445
            TabIndex        =   23
            Top             =   225
            Width           =   1050
         End
      End
      Begin VB.TextBox txtSec_Parte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5910
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "0"
         Top             =   1320
         Width           =   600
      End
      Begin VB.TextBox txtDes_TipCobra 
         Height          =   285
         Left            =   4425
         TabIndex        =   7
         Top             =   1695
         Width           =   2085
      End
      Begin VB.TextBox txtCod_TipCobra 
         Height          =   285
         Left            =   3705
         MaxLength       =   4
         TabIndex        =   6
         Top             =   1695
         Width           =   600
      End
      Begin VB.TextBox txtNum_Ruc 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7755
         MaxLength       =   11
         TabIndex        =   17
         Top             =   240
         Width           =   1545
      End
      Begin VB.TextBox txtDes_TipAne 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         Top             =   240
         Width           =   5265
      End
      Begin VB.TextBox TxtObservacion 
         Height          =   285
         Left            =   1215
         TabIndex        =   8
         Top             =   2055
         Width           =   5325
      End
      Begin VB.TextBox txtCod_Moneda 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   2
         Top             =   1335
         Width           =   600
      End
      Begin VB.TextBox txtDes_Moneda 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Top             =   1335
         Width           =   2385
      End
      Begin VB.TextBox txtNro_Anticipo 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   0
         Top             =   600
         Width           =   1200
      End
      Begin NumBoxProject.NumBox txtFecha 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   960
         Width           =   1200
         _ExtentX        =   2117
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
      Begin NumBoxProject.NumBox txt_Importe 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   1680
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         TypeVal         =   2
         Mask            =   "9,999,999,999.99"
         Formato         =   "#,###,###,###.##"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   3
         Text            =   "0.00"
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sec Parte :"
         Height          =   195
         Left            =   4920
         TabIndex        =   20
         Top             =   1350
         Width           =   795
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cobranza:"
         Height          =   195
         Left            =   2505
         TabIndex        =   19
         Top             =   1740
         Width           =   1080
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   285
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Importe :"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1725
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Observacion :"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   2100
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1005
         Width           =   660
      End
      Begin VB.Label Label11 
         Caption         =   "Moneda :"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1365
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Numero :"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   615
         Width           =   735
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2250
      TabIndex        =   9
      Top             =   3555
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "0~0~GRABAR~True~True~&Grabar~0~0~1~~0~False~False~&Grabar~~1~0~CANCELAR~True~True~&Cancelar~0~0~2~~0~False~False~&Cancelar~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmAdelantosAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strCod_TipAnex As String, strCod_Anexo As String, strOption As String, _
       lfAceptar As Boolean, intNum_Anticipo As String
Public codigo As String, Descripcion As String

Private Sub Form_Load()
    txtTipoCambioNegociado.DecimalNumber = 4
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

Dim sMessage As Long
Dim strSQL As String

On Error GoTo dprError

Select Case ActionName
  Case "GRABAR"
    If MsgBox("Esta seguro de actualizar los anticipos ", vbYesNo, "IMPORTANTE") = vbYes Then
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
Dim sCod_Moneda_Transaccion As String

sCod_Moneda_Transaccion = DevuelveCampo("SELECT DBO.CN_Ventas_Obtiene_Moneda_Transacc_Cobranza ('" & txtFecha.Text & "'," & txtSec_Parte.Text & ")", cCONNECT)
lblMonedaTransaccion.Caption = sCod_Moneda_Transaccion

If RTrim(sCod_Moneda_Transaccion) <> "" And RTrim(sCod_Moneda_Transaccion) <> txtCod_Moneda.Text Then
    If txtAnticipo_Moneda_Deposito.Text = 0 Then
         fraAnticipoOtraMoneda.Visible = True
         Exit Function
    End If
  
Else
    txtAnticipo_Moneda_Deposito.Text = 0
    txtTipoCambioNegociado.Text = 0
    fraAnticipoOtraMoneda.Visible = False
End If

SQL = "Ventas_Man_Adelantos '" & strOption & "','" & strCod_TipAnex & "','" & strCod_Anexo & "'," _
      & txtNro_Anticipo.Text & ",'','" & txtFecha.Text & "','" & txtCod_Moneda.Text & "'," & txt_Importe.Text & ",'" _
      & TxtObservacion & "','" & txtCod_TipCobra.Text & "'," & txtSec_Parte & "," & txtAnticipo_Moneda_Deposito.Text & "," & txtTipoCambioNegociado.Text
      
Call ExecuteCommandSQL(cCONNECT, SQL)

lfSalvar_Datos = True

Exit Function

hand:

errores Err.Number

lfSalvar_Datos = False

End Function

Private Sub txt_Importe_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtAnticipo_Moneda_Deposito_GotFocus()
    txtAnticipo_Moneda_Deposito.Text = Round(txtTipoCambioNegociado.Text * txt_Importe.Text, 2)
End Sub

Private Sub txtAnticipo_Moneda_Deposito_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FunctButt1.SetFocus
    End If
End Sub

Private Sub txtCod_Moneda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Moneda", "Nom_Moneda", "TG_Moneda where ", txtCod_Moneda, txtDes_Moneda, 1, Me)
End Sub

Private Sub txtDes_Moneda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Moneda", "Nom_Moneda", "TG_Moneda where ", txtCod_Moneda, txtDes_Moneda, 2, Me)
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNro_Anticipo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
  If KeyAscii >= 48 And KeyAscii <= 57 _
    Or KeyAscii = 8 Or KeyAscii = 13 Then Else KeyAscii = 0
End Sub
  
Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCod_TipCobra_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion_Store("Cn_Ventas_Muestra_Tipos_Cobranza_Permitidos_Adelantos  ", txtCod_TipCobra, txtDes_TipCobra, 1, Me)
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtDes_TipCobra_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion_Store("Cn_Ventas_Muestra_Tipos_Cobranza_Permitidos_Adelantos ", txtCod_TipCobra, txtDes_TipCobra, 2, Me)
  End If
End Sub

Private Sub txtSec_Parte_Change()
  If txtSec_Parte = "" Then txtSec_Parte = 0
End Sub

Private Sub txtSec_Parte_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
  If KeyAscii >= 48 And KeyAscii <= 57 _
    Or KeyAscii = 8 Or KeyAscii = 13 Then Else KeyAscii = 0
End Sub

Private Sub txtTipoCambioNegociado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        
        txtAnticipo_Moneda_Deposito.SetFocus
    End If
    
End Sub
