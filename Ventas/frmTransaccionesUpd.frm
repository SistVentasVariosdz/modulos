VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmTransaccionesUpd 
   Caption         =   "Modifica Documento de Cobranza"
   ClientHeight    =   3585
   ClientLeft      =   465
   ClientTop       =   2355
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3585
   ScaleWidth      =   9705
   Begin VB.Frame Frame3 
      Height          =   2895
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   9615
      Begin VB.TextBox txtDes_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   645
         Width           =   5265
      End
      Begin VB.TextBox txtDes_TipAnex 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4680
         MaxLength       =   4
         TabIndex        =   24
         Text            =   "C"
         Top             =   645
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.TextBox txtDes_TipCobra 
         Height          =   285
         Left            =   2085
         TabIndex        =   1
         Top             =   255
         Width           =   1905
      End
      Begin VB.TextBox txtCod_TipCobra 
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   0
         Top             =   255
         Width           =   735
      End
      Begin VB.TextBox txtNum_Ruc 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7740
         MaxLength       =   11
         TabIndex        =   6
         Top             =   645
         Width           =   1545
      End
      Begin VB.TextBox txtCod_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6600
         MaxLength       =   4
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "C"
         Top             =   645
         Width           =   360
      End
      Begin VB.TextBox txtCod_Moneda 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6120
         MaxLength       =   4
         TabIndex        =   13
         Top             =   1350
         Width           =   720
      End
      Begin VB.TextBox txtDes_DocCobra 
         Height          =   285
         Left            =   2085
         TabIndex        =   12
         Top             =   1365
         Width           =   2415
      End
      Begin VB.TextBox txtCod_TipDocCobra 
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   11
         Top             =   1365
         Width           =   735
      End
      Begin VB.TextBox TxtCod_Banco 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   990
         Width           =   735
      End
      Begin VB.TextBox TxtDes_Banco 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2085
         TabIndex        =   8
         Top             =   990
         Width           =   2415
      End
      Begin VB.TextBox txtSer_DocCobra 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   15
         Top             =   1725
         Width           =   735
      End
      Begin VB.TextBox txtNum_DocCobra 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6120
         MaxLength       =   8
         TabIndex        =   16
         Top             =   1725
         Width           =   3150
      End
      Begin VB.TextBox TxtObservacion 
         Height          =   285
         Left            =   1320
         TabIndex        =   20
         Top             =   2460
         Width           =   7965
      End
      Begin VB.CheckBox chkDiferido 
         Alignment       =   1  'Right Justify
         Caption         =   "&Diferido"
         Height          =   255
         Left            =   5400
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2145
         Width           =   1455
      End
      Begin VB.TextBox txtDes_Origen 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5400
         TabIndex        =   3
         Top             =   255
         Width           =   1665
      End
      Begin VB.TextBox txtCod_Origen 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4680
         MaxLength       =   1
         TabIndex        =   2
         Top             =   255
         Width           =   600
      End
      Begin VB.TextBox txt_ImpTotal_Doc_Cobra 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2085
         MaxLength       =   15
         TabIndex        =   17
         Text            =   "0"
         Top             =   2130
         Width           =   1200
      End
      Begin VB.TextBox txtCuenta_Cod 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6120
         MaxLength       =   11
         TabIndex        =   9
         Top             =   990
         Width           =   720
      End
      Begin VB.TextBox txtCuenta_Des 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6915
         MaxLength       =   30
         TabIndex        =   10
         Top             =   990
         Width           =   2370
      End
      Begin VB.TextBox txtDes_Moneda 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6900
         TabIndex        =   14
         Top             =   1350
         Width           =   2385
      End
      Begin NumBoxProject.NumBox txtFecha 
         Height          =   285
         Left            =   7905
         TabIndex        =   4
         Top             =   255
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
      Begin NumBoxProject.NumBox txtFec_Diferido 
         Height          =   285
         Left            =   7905
         TabIndex        =   19
         Top             =   2130
         Visible         =   0   'False
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
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   7200
         TabIndex        =   38
         Top             =   300
         Width           =   660
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cobranza:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   690
         Width           =   570
      End
      Begin VB.Label Label28 
         Caption         =   "R.U.C."
         Height          =   255
         Left            =   7080
         TabIndex        =   35
         Top             =   660
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Moneda :"
         Height          =   255
         Left            =   5400
         TabIndex        =   34
         Top             =   1380
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tip Doc Cobra:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   1410
         Width           =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "Banco:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1020
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Serie :"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   1770
         Width           =   450
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Observacion :"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   2445
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Imp Total Doc Cobranza:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   2175
         Width           =   1770
      End
      Begin VB.Label Label4 
         Caption         =   "Origen :"
         Height          =   255
         Left            =   4080
         TabIndex        =   28
         Top             =   270
         Width           =   615
      End
      Begin VB.Label lbDiferido 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   7200
         TabIndex        =   27
         Top             =   2175
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta :"
         Height          =   195
         Left            =   5400
         TabIndex        =   26
         Top             =   1035
         Width           =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Numero :"
         Height          =   195
         Left            =   5400
         TabIndex        =   25
         Top             =   1770
         Width           =   645
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   7080
      TabIndex        =   21
      Top             =   3000
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmTransaccionesUpd.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmTransaccionesUpd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public codigo As String, Descripcion As String, strOption As String, strCod_Anxo As String, intSecuencia As Integer, lfAceptar As Boolean

Dim RsGrid1 As ADODB.Recordset
Dim RsGrid2 As ADODB.Recordset
Dim intTransaccion As Integer
Dim strSQL As String

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

Private Sub Form_Load()
  lfAceptar = False
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim sMessage As Long
Dim strSQL As String

On Error GoTo dprError

Select Case ActionName
  Case "GRABAR"
  
    If MsgBox("Esta seguro de Actualizar una Transaccion ", vbYesNo, "IMPORTANTE") = vbYes Then
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

SQL = "CN_VENTAS_TRANSACCIONES_COBRANZAS_MAN '" & strOption & "','" & txtFecha.Text & "'," & intSecuencia & ",'" _
      & txtCod_TipCobra & "','" & txtCod_TipAne & "','" & strCod_Anxo & "','" & TxtCod_Banco & "','" & txtCuenta_Cod & "','" _
      & txtCod_TipDocCobra & "','" & txtSer_DocCobra & "','" & txtNum_DocCobra & "','" & txtCod_Moneda & "','" _
      & Des_Apos(TxtObservacion) & "','" & vusu & "','" & ComputerName & "','" & txtCod_Origen & "','" & IIf(chkDiferido, "S", "N") & "','S'," _
      & IIf(txtFec_Diferido.Text <> "", "'" & txtFec_Diferido.Text & "'", "Null")
      
Call ExecuteCommandSQL(cCONNECT, SQL)

lfSalvar_Datos = True

Exit Function

hand:

errores err.Number

lfSalvar_Datos = False

End Function

Private Sub txt_ImpTotal_Doc_Cobra_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtCod_Banco_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Call Busca_Opcion("Cod_Banco", "Nom_Banco", "Tg_Banco where ", TxtCod_Banco, TxtDes_Banco, 1, Me)
End Sub

Private Sub txtCod_Moneda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Moneda", "Nom_Moneda", "TG_Moneda where ", txtCod_Moneda, txtDes_Moneda, 1, Me)
End Sub

Private Sub txtCod_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Origen", "Des_Origen", " Cn_Origen where ", txtCod_Origen, txtDes_Origen, 1, Me)
End Sub

Private Sub txtCod_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAne, txtDes_TipAnex, 1, Me)
End Sub

Private Sub txtCod_TipCobra_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Tipcobranza", "Descripcion", " Cn_Ventas_Tipos_Cobranza where ", txtCod_TipCobra, txtDes_TipCobra, 1, Me)
End Sub

Private Sub txtCod_TipDocCobra_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipDoc", "Des_TipDoc", " CN_TiposDocum where Flg_Doc_Cobranza = '*' and ", txtCod_TipDocCobra, txtDes_DocCobra, 1, Me)
End Sub

Private Sub TxtDes_Banco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Banco", "Nom_Banco", "Tg_Banco where ", TxtCod_Banco, TxtDes_Banco, 1, Me)
End Sub

Private Sub txtDes_Moneda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Moneda", "Nom_Moneda", "TG_Moneda where ", txtCod_Moneda, txtDes_Moneda, 2, Me)
End Sub


Private Sub txtDes_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Origen", "Des_Origen", " Cn_Origen where ", txtCod_Origen, txtDes_Origen, 1, Me)
End Sub

Private Sub txtDes_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion_Anexo("Num_Ruc", "Des_Anexo", " Cn_AnexosContables where cod_tipanex = '" & txtCod_TipAne & "' and ", txtNum_Ruc, txtDes_TipAne, 2, Me)
End Sub

Private Sub txtDes_TipCobra_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Tipcobranza", "Descripcion", " Cn_Ventas_Tipos_Cobranza where ", txtCod_TipCobra, txtDes_TipCobra, 1, Me)
End Sub

Private Sub txtFec_Diferido_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNum_DocCobra_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion_Anexo("Num_Ruc", "Des_Anexo", " Cn_AnexosContables where cod_tipanex = '" & txtCod_TipAne & "' and ", txtNum_Ruc, txtDes_TipAne, 1, Me)
End Sub

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtSer_DocCobra_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


