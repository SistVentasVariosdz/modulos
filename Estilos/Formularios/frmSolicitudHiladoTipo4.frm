VERSION 5.00
Begin VB.Form frmSolicitudHiladoTipo4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitudes de Hilado"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   660
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   7125
      Begin VB.Label lblTipo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Tipo 4 POLYCOTTON"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2565
         TabIndex        =   40
         Top             =   300
         Width           =   1860
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6780
      Left            =   15
      TabIndex        =   2
      Top             =   690
      Width           =   7110
      Begin VB.CheckBox chkColor_Fibra_Poly_Otro 
         Alignment       =   1  'Right Justify
         Height          =   200
         Left            =   5370
         TabIndex        =   19
         Top             =   1710
         Width           =   200
      End
      Begin VB.CheckBox chkColor_Fibra_Poly_Negro 
         Alignment       =   1  'Right Justify
         Height          =   200
         Left            =   5370
         TabIndex        =   18
         Top             =   1335
         Width           =   200
      End
      Begin VB.CheckBox chkColor_Fibra_Poly_Blanco 
         Alignment       =   1  'Right Justify
         Height          =   200
         Left            =   5370
         TabIndex        =   17
         Top             =   975
         Width           =   200
      End
      Begin VB.TextBox txtNE 
         Height          =   285
         Left            =   855
         MaxLength       =   20
         TabIndex        =   16
         Top             =   180
         Width           =   1425
      End
      Begin VB.CheckBox chkProc_3_Otros 
         Alignment       =   1  'Right Justify
         Height          =   200
         Left            =   5370
         TabIndex        =   15
         Top             =   2940
         Width           =   200
      End
      Begin VB.CheckBox chkProc_1_Tanguis 
         Alignment       =   1  'Right Justify
         Height          =   200
         Left            =   5370
         TabIndex        =   14
         Top             =   2175
         Width           =   200
      End
      Begin VB.CheckBox chkProc_2_Americano 
         Alignment       =   1  'Right Justify
         Height          =   200
         Left            =   5370
         TabIndex        =   13
         Top             =   2550
         Width           =   200
      End
      Begin VB.CheckBox chkSentido_Torsion_Z 
         Alignment       =   1  'Right Justify
         Caption         =   "TORSION Z"
         Height          =   210
         Left            =   3800
         TabIndex        =   12
         Top             =   3945
         Width           =   1770
      End
      Begin VB.CheckBox chkMet_Hilatura_1 
         Alignment       =   1  'Right Justify
         Caption         =   "PEINADO"
         Height          =   240
         Left            =   3800
         TabIndex        =   11
         Top             =   5490
         Width           =   1770
      End
      Begin VB.CheckBox chkSentido_Torsion_S 
         Alignment       =   1  'Right Justify
         Caption         =   "TORSION S"
         Height          =   195
         Left            =   3800
         TabIndex        =   10
         Top             =   3675
         Width           =   1770
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   285
         Left            =   1095
         MaxLength       =   50
         TabIndex        =   9
         Top             =   5055
         Width           =   5865
      End
      Begin VB.TextBox txtPorc_Fibra_Poly_Blanco 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5760
         TabIndex        =   8
         Text            =   "0"
         Top             =   945
         Width           =   840
      End
      Begin VB.TextBox txtPorc_Fibra_Poly_Negro 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5760
         TabIndex        =   7
         Text            =   "0"
         Top             =   1305
         Width           =   825
      End
      Begin VB.TextBox txtPorc_Fibra_Poly_Otro 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5760
         TabIndex        =   6
         Text            =   "0"
         Top             =   1680
         Width           =   825
      End
      Begin VB.CheckBox chkMet_Hilatura_2 
         Alignment       =   1  'Right Justify
         Caption         =   "CARDADO"
         Height          =   195
         Left            =   3800
         TabIndex        =   5
         Top             =   5790
         Width           =   1770
      End
      Begin VB.CheckBox chkMet_Hilatura_3 
         Alignment       =   1  'Right Justify
         Caption         =   "OPEN END"
         Height          =   285
         Left            =   3800
         TabIndex        =   4
         Top             =   6030
         Width           =   1770
      End
      Begin VB.TextBox txtAlpha 
         Height          =   285
         Left            =   1095
         MaxLength       =   10
         TabIndex        =   3
         Top             =   6375
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NE"
         Height          =   195
         Left            =   180
         TabIndex        =   38
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "COMPOSICION"
         Height          =   195
         Left            =   255
         TabIndex        =   37
         Top             =   660
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "FIBRA ALGODON CRUDA"
         Height          =   195
         Left            =   255
         TabIndex        =   36
         Top             =   2175
         Width           =   1920
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "OBS :"
         Height          =   195
         Left            =   255
         TabIndex        =   35
         Top             =   5145
         Width           =   420
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   3015
         Left            =   120
         Top             =   540
         Width           =   6795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "COLOR DE FIBRA DE POLIESTER"
         Height          =   195
         Left            =   255
         TabIndex        =   34
         Top             =   1005
         Width           =   2550
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "COLOR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3800
         TabIndex        =   33
         Top             =   660
         Width           =   645
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   1125
         Left            =   5280
         Top             =   900
         Width           =   1395
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         Index           =   0
         X1              =   5265
         X2              =   6660
         Y1              =   1620
         Y2              =   1635
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         Index           =   1
         X1              =   6660
         X2              =   5280
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   0
         X1              =   5700
         X2              =   5700
         Y1              =   885
         Y2              =   1995
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         Height          =   1125
         Left            =   5280
         Top             =   2115
         Width           =   435
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         Index           =   2
         X1              =   5685
         X2              =   5265
         Y1              =   2445
         Y2              =   2460
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         Index           =   3
         X1              =   5685
         X2              =   5280
         Y1              =   2820
         Y2              =   2835
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "CRUDO"
         Height          =   195
         Left            =   3800
         TabIndex        =   32
         Top             =   1005
         Width           =   585
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "NEGRO"
         Height          =   195
         Left            =   3800
         TabIndex        =   31
         Top             =   1365
         Width           =   585
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "OTRO"
         Height          =   195
         Left            =   3800
         TabIndex        =   30
         Top             =   1725
         Width           =   465
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "TANGUIS"
         Height          =   195
         Left            =   3800
         TabIndex        =   29
         Top             =   2175
         Width           =   720
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "AMERICANO"
         Height          =   195
         Left            =   3800
         TabIndex        =   28
         Top             =   2565
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "OTROS"
         Height          =   195
         Left            =   3800
         TabIndex        =   27
         Top             =   2955
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL"
         Height          =   195
         Left            =   3800
         TabIndex        =   26
         Top             =   3285
         Width           =   525
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   5940
         TabIndex        =   25
         Top             =   3300
         Width           =   120
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "SENTIDO DE TORSION"
         Height          =   195
         Left            =   255
         TabIndex        =   24
         Top             =   3720
         Width           =   1770
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   60
         X2              =   6960
         Y1              =   4230
         Y2              =   4230
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "MUESTRA DEL TONO DESEADO"
         Height          =   195
         Left            =   255
         TabIndex        =   23
         Top             =   4440
         Width           =   2490
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         Height          =   690
         Left            =   3015
         Top             =   4335
         Width           =   3945
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "ALPHA"
         Height          =   195
         Left            =   255
         TabIndex        =   22
         Top             =   6450
         Width           =   645
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "METODO DE HILANDURIA"
         Height          =   195
         Left            =   255
         TabIndex        =   21
         Top             =   5520
         Width           =   2010
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   6105
         TabIndex        =   20
         Top             =   630
         Width           =   120
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   500
      Left            =   1080
      TabIndex        =   1
      Top             =   7620
      Width           =   1350
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   500
      Left            =   4485
      TabIndex        =   0
      Top             =   7620
      Width           =   1350
   End
End
Attribute VB_Name = "frmSolicitudHiladoTipo4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrSQL As String
Dim sTipo As String
Public varCod_HilTel As String
Dim Rs_Lista As ADODB.Recordset

Public Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
End Function

Sub Cargar_Datos()
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.ActiveConnection = cCONNECT
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    'Esta cadena es para devolver el Codigo de Cliente
    StrSQL = "EXEC UP_SEL_IT_HILADO_DATOS_DESARROLLO '" & Me.varCod_HilTel & "'"
    Rs_Lista.Open StrSQL

    'Aqui cargamos los datos
    If Rs_Lista.RecordCount > 0 Then
    
        sTipo = "U"
    
        Me.txtAlpha.Text = Trim(Rs_Lista("Alpha").Value)
        Me.txtNE.Text = Trim(Rs_Lista("NE").Value)
        Me.txtObservaciones.Text = Trim(Rs_Lista("Observaciones").Value)
        Me.txtPorc_Fibra_Poly_Blanco.Text = Rs_Lista("Porc_Fibra_Poly_Blanco").Value
        Me.txtPorc_Fibra_Poly_Negro.Text = Rs_Lista("Porc_Fibra_Poly_Negro").Value
        Me.txtPorc_Fibra_Poly_Otro.Text = Rs_Lista("Porc_Fibra_Poly_Otro").Value
        
        If Rs_Lista("Color_Fibra_Poly_Blanco").Value = "*" Then
            Me.chkColor_Fibra_Poly_Blanco.Value = 1
        End If
        If Rs_Lista("Color_Fibra_Poly_Negro").Value = "*" Then
            Me.chkColor_Fibra_Poly_Negro.Value = 1
        End If
        If Rs_Lista("Color_Fibra_Poly_Otro").Value = "*" Then
            Me.chkColor_Fibra_Poly_Otro.Value = 1
        End If
        If Rs_Lista("Met_Hilatura_1").Value = "*" Then
            Me.chkMet_Hilatura_1.Value = 1
        End If
        If Rs_Lista("Met_Hilatura_2").Value = "*" Then
            Me.chkMet_Hilatura_2.Value = 1
        End If
        If Rs_Lista("Met_Hilatura_3").Value = "*" Then
            Me.chkMet_Hilatura_3.Value = 1
        End If
        If Rs_Lista("Proc_1_Tanguis").Value = "*" Then
            Me.chkProc_1_Tanguis.Value = 1
        End If
        If Rs_Lista("Proc_2_Americano").Value = "*" Then
            Me.chkProc_2_Americano.Value = 1
        End If
        If Rs_Lista("Proc_3_Otros").Value = "*" Then
            Me.chkProc_3_Otros.Value = 1
        End If
        If Rs_Lista("Sentido_Torsion_S").Value = "*" Then
            Me.chkSentido_Torsion_S.Value = 1
        End If
        If Rs_Lista("Sentido_Torsion_Z").Value = "*" Then
            Me.chkSentido_Torsion_Z.Value = 1
        End If
    Else
        sTipo = "I"
    End If
End Sub

Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    Dim Rs As New ADODB.Recordset
    On Error GoTo Salvar_DatosErr
    Dim StrSQL As String
    
    Con.ConnectionString = cCONNECT
    Con.Open
    
        Con.BeginTrans

        StrSQL = "EXEC UP_MAN_IT_HILADO_DATOS_DESARROLLO '" & _
        sTipo & "','" & _
        Me.varCod_HilTel & "','" & _
        Trim(Me.txtNE.Text) & "','" & _
        IIf(Me.chkProc_1_Tanguis.Value <> 0, "*", "") & "','" & _
        IIf(Me.chkProc_2_Americano.Value <> 0, "*", "") & "','" & _
        IIf(Me.chkProc_3_Otros.Value <> 0, "*", "") & "','" & _
        IIf(Me.chkSentido_Torsion_S.Value <> 0, "*", "") & "','" & _
        IIf(Me.chkSentido_Torsion_Z.Value <> 0, "*", "") & "','" & _
        IIf(Me.chkMet_Hilatura_1.Value <> 0, "*", "") & "','" & _
        IIf(Me.chkMet_Hilatura_2.Value <> 0, "*", "") & "','" & _
        IIf(Me.chkMet_Hilatura_3.Value <> 0, "*", "") & "','" & _
        Trim(Me.txtAlpha.Text) & "','" & _
        Trim(Me.txtObservaciones.Text) & "','"
        
        StrSQL = StrSQL & _
        "" & "'," & _
        "0" & ",'" & _
        IIf(Me.chkColor_Fibra_Poly_Negro.Value <> 0, "*", "") & "'," & _
        Me.txtPorc_Fibra_Poly_Negro & ",'" & _
        IIf(Me.chkColor_Fibra_Poly_Otro.Value <> 0, "*", "") & "'," & _
        Me.txtPorc_Fibra_Poly_Otro & ",'" & _
        IIf(Me.chkColor_Fibra_Poly_Blanco.Value <> 0, "*", "") & "'," & _
        Me.txtPorc_Fibra_Poly_Blanco & ",'" & _
        "" & "','" & _
        "" & "'"
        
        Con.Execute StrSQL
       
        Con.CommitTrans
        Dim amensaje As New clsMessages
        amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
        Informa "", amensaje
        
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub

Private Sub cmdGrabar_Click()
    If VALIDA_DATOS Then
        Call Me.SALVAR_DATOS
        Unload Me
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub txtPorc_Fibra_Poly_Blanco_KeyPress(KeyAscii As Integer)
    Call SoloNumeros(Me.txtPorc_Fibra_Poly_Blanco, KeyAscii, True, 2, 5)
End Sub

Private Sub txtPorc_Fibra_Poly_Blanco_LostFocus()
    If txtPorc_Fibra_Poly_Blanco.Text = "" Then
        txtPorc_Fibra_Poly_Blanco.Text = "0"
    End If
End Sub

Private Sub txtPorc_Fibra_Poly_Negro_KeyPress(KeyAscii As Integer)
    Call SoloNumeros(Me.txtPorc_Fibra_Poly_Negro, KeyAscii, True, 2, 5)
End Sub

Private Sub txtPorc_Fibra_Poly_Negro_LostFocus()
    If Trim(txtPorc_Fibra_Poly_Negro.Text) = "" Then
        txtPorc_Fibra_Poly_Negro.Text = "0"
    End If
End Sub

Private Sub txtPorc_Fibra_Poly_Otro_KeyPress(KeyAscii As Integer)
    Call SoloNumeros(Me.txtPorc_Fibra_Poly_Otro, KeyAscii, True, 2, 5)
End Sub

Private Sub txtPorc_Fibra_Poly_Otro_LostFocus()
    If Trim(txtPorc_Fibra_Poly_Otro.Text) = "" Then
        txtPorc_Fibra_Poly_Otro.Text = "0"
    End If
End Sub
