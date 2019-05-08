VERSION 5.00
Begin VB.Form frmSolicitudHiladoTipo1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de Hilado"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   660
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   7125
      Begin VB.Label lblTipo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Tipo 1 ALGODON"
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
         Left            =   2775
         TabIndex        =   23
         Top             =   285
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4560
      Left            =   15
      TabIndex        =   2
      Top             =   690
      Width           =   7110
      Begin VB.TextBox txtNE 
         Height          =   285
         Left            =   855
         TabIndex        =   12
         Top             =   180
         Width           =   1425
      End
      Begin VB.CheckBox chkProc_3_Otros 
         Alignment       =   1  'Right Justify
         Height          =   200
         Left            =   5370
         TabIndex        =   11
         Top             =   1725
         Width           =   200
      End
      Begin VB.CheckBox chkProc_1_Tanguis 
         Alignment       =   1  'Right Justify
         Height          =   200
         Left            =   5370
         TabIndex        =   10
         Top             =   960
         Width           =   200
      End
      Begin VB.CheckBox chkProc_2_Americano 
         Alignment       =   1  'Right Justify
         Height          =   200
         Left            =   5370
         TabIndex        =   9
         Top             =   1335
         Width           =   200
      End
      Begin VB.CheckBox chkSentido_Torsion_Z 
         Alignment       =   1  'Right Justify
         Caption         =   "TORSION Z"
         Height          =   210
         Left            =   3800
         TabIndex        =   8
         Top             =   2625
         Width           =   1770
      End
      Begin VB.CheckBox chkMet_Hilatura_1 
         Alignment       =   1  'Right Justify
         Caption         =   "PEINADO"
         Height          =   240
         Left            =   3800
         TabIndex        =   7
         Top             =   3180
         Width           =   1770
      End
      Begin VB.CheckBox chkSentido_Torsion_S 
         Alignment       =   1  'Right Justify
         Caption         =   "TORSION S"
         Height          =   195
         Left            =   3800
         TabIndex        =   6
         Top             =   2355
         Width           =   1770
      End
      Begin VB.CheckBox chkMet_Hilatura_2 
         Alignment       =   1  'Right Justify
         Caption         =   "CARDADO"
         Height          =   195
         Left            =   3800
         TabIndex        =   5
         Top             =   3480
         Width           =   1770
      End
      Begin VB.CheckBox chkMet_Hilatura_3 
         Alignment       =   1  'Right Justify
         Caption         =   "OPEN END"
         Height          =   285
         Left            =   3800
         TabIndex        =   4
         Top             =   3720
         Width           =   1770
      End
      Begin VB.TextBox txtAlpha 
         Height          =   285
         Left            =   1095
         MaxLength       =   10
         TabIndex        =   3
         Top             =   4065
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NE"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "COMPOSICION"
         Height          =   195
         Left            =   255
         TabIndex        =   20
         Top             =   660
         Width           =   1125
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   1665
         Left            =   120
         Top             =   540
         Width           =   6795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "PROCEDENCIA DE FIBRA"
         Height          =   195
         Left            =   255
         TabIndex        =   19
         Top             =   1005
         Width           =   1935
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         Height          =   1125
         Left            =   5280
         Top             =   900
         Width           =   435
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         Index           =   2
         X1              =   5685
         X2              =   5265
         Y1              =   1230
         Y2              =   1245
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         Index           =   3
         X1              =   5685
         X2              =   5280
         Y1              =   1605
         Y2              =   1620
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "TANGUIS"
         Height          =   195
         Left            =   3795
         TabIndex        =   18
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "AMERICANO"
         Height          =   195
         Left            =   3795
         TabIndex        =   17
         Top             =   1350
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "OTROS"
         Height          =   195
         Left            =   3795
         TabIndex        =   16
         Top             =   1740
         Width           =   570
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "SENTIDO DE TORSION"
         Height          =   195
         Left            =   255
         TabIndex        =   15
         Top             =   2400
         Width           =   1770
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   60
         X2              =   6960
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "ALPHA"
         Height          =   195
         Left            =   255
         TabIndex        =   14
         Top             =   4140
         Width           =   645
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "METODO DE HILANDURIA"
         Height          =   195
         Left            =   255
         TabIndex        =   13
         Top             =   3210
         Width           =   2010
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   500
      Left            =   1080
      TabIndex        =   1
      Top             =   5385
      Width           =   1350
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   500
      Left            =   4485
      TabIndex        =   0
      Top             =   5385
      Width           =   1350
   End
End
Attribute VB_Name = "frmSolicitudHiladoTipo1"
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
        'Me.txtObservaciones.Text = Rs_Lista("Observaciones").Value
        'Me.txtPorc_Fibra_Poly_Crudo.Text = Rs_Lista("Porc_Fibra_Poly_Crudo").Value
        'Me.txtPorc_Fibra_Poly_Negro.Text = Rs_Lista("Porc_Fibra_Poly_Negro").Value
        'Me.txtPorc_Fibra_Poly_Otro.Text = Rs_Lista("Porc_Fibra_Poly_Otro").Value
        
'        If Rs_Lista("Color_Fibra_Poly_Crudo").Value = "*" Then
'            Me.chkColor_Fibra_Poly_Crudo.Value = 1
'        End If
'        If Rs_Lista("Color_Fibra_Poly_Negro").Value = "*" Then
'            Me.chkColor_Fibra_Poly_Negro.Value = 1
'        End If
'        If Rs_Lista("Color_Fibra_Poly_Otro").Value = "*" Then
'            Me.chkColor_Fibra_Poly_Otro.Value = 1
'        End If
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
        "" & "','"
        
        StrSQL = StrSQL & _
        "" & "'," & _
        "0" & ",'" & _
        "" & "'," & _
        "0" & ",'" & _
        "" & "'," & _
        "0" & ",'" & _
        "" & "'," & _
        "0" & ",'" & _
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

