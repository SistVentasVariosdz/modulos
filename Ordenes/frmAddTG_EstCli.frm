VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmAddTG_EstCli 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creación de Estilo de Cliente"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Add Client Style "
   Begin VB.Frame fraAddEstPro 
      Caption         =   "Ingreso de Estilo Propio"
      Height          =   2355
      Left            =   120
      TabIndex        =   14
      Top             =   2535
      Visible         =   0   'False
      Width           =   6915
      Begin VB.TextBox txtDes_GruTal 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2025
         MaxLength       =   50
         TabIndex        =   19
         Top             =   1170
         Width           =   4335
      End
      Begin VB.TextBox txtCod_GruTal 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1275
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1170
         Width           =   690
      End
      Begin VB.TextBox txtDes_TipPre 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2025
         MaxLength       =   50
         TabIndex        =   17
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox txtCod_TipPre 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1290
         MaxLength       =   20
         TabIndex        =   6
         Top             =   720
         Width           =   690
      End
      Begin VB.TextBox txtDes_EstPro2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   5
         Top             =   300
         Width           =   4335
      End
      Begin FunctionsButtons.FunctButt funAceptar 
         Height          =   510
         Left            =   2235
         TabIndex        =   8
         Top             =   1710
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   "7~0~ACEPTAR~True~True~&Aceptar~0~0~4~~0~True~False~&Ok~~1~0~CANCELAR~True~True~&Cancelar~0~0~3~~0~False~True~&Cancel~"
         Orientacion     =   0
         Style           =   1
         Language        =   1
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Grupo de Talla"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Tag             =   "Id Style"
         Top             =   1185
         Width           =   1050
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Prenda:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Tag             =   "Id Style"
         Top             =   735
         Width           =   1125
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Descripción :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Tag             =   "Description"
         Top             =   330
         Width           =   945
      End
   End
   Begin VB.TextBox txtDes_EstPro 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   13
      Top             =   1230
      Width           =   4335
   End
   Begin VB.TextBox txtCod_EstPro 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1245
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1230
      Width           =   1095
   End
   Begin VB.TextBox txttelaestilo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1245
      MaxLength       =   300
      TabIndex        =   2
      Top             =   810
      Width           =   5505
   End
   Begin VB.TextBox txtIdestilo 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1245
      MaxLength       =   20
      TabIndex        =   0
      Top             =   90
      Width           =   3480
   End
   Begin VB.TextBox txtNomestilo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1260
      MaxLength       =   50
      TabIndex        =   1
      Top             =   450
      Width           =   4335
   End
   Begin FunctionsButtons.FunctButt acbForm 
      Height          =   510
      Left            =   2370
      TabIndex        =   4
      Top             =   1770
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "7~0~ACEPTAR~True~True~&Aceptar~0~0~4~~0~True~False~&Ok~~1~0~CANCELAR~True~True~&Cancelar~0~0~3~~0~False~True~&Cancel~"
      Orientacion     =   0
      Style           =   1
      Language        =   1
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Label Etiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Id Estilo Propio :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   75
      TabIndex        =   12
      Tag             =   "Id Style"
      Top             =   1245
      Width           =   1125
   End
   Begin VB.Label Etiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Tela :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   60
      TabIndex        =   11
      Tag             =   "Fabric"
      Top             =   840
      Width           =   390
   End
   Begin VB.Label Etiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Id Estilo :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   75
      TabIndex        =   10
      Tag             =   "Id Style"
      Top             =   105
      Width           =   1170
   End
   Begin VB.Label Etiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Descripción :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   75
      TabIndex        =   9
      Tag             =   "Description"
      Top             =   495
      Width           =   945
   End
End
Attribute VB_Name = "frmAddTG_EstCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Cliente As String
Public sCod_TemCli As String
Public sCod_EStCli  As String
Public oParent As Object
Public bOk As Boolean
Public sFlag As String
Public sModoAddEstCli As String

Private Sub acbForm_ActionClick(ByVal index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo errores
    Dim vbuff
    Dim obj As clsTG_Cliente

    Select Case ActionName
        Case "ACEPTAR"
            If Me.txtIdestilo.Text = "" Then
                Mensaje kMESSAGE_ERR_NOTEMPTY
                Me.txtIdestilo.SetFocus
                Exit Sub
            End If
        
            If sModoAddEstCli <> "SOLO ASIGNACION" Then
                If Me.txtNomestilo.Text = "" Then
                    Mensaje kMESSAGE_ERR_NOTEMPTY
                    Me.txtNomestilo.SetFocus
                    Exit Sub
                End If
                
                If Me.txttelaestilo.Text = "" Then
                    Mensaje kMESSAGE_ERR_NOTEMPTY
                    Me.txttelaestilo.SetFocus
                    Exit Sub
                End If
            End If
            sCod_EStCli = Me.txtIdestilo.Text
                        
            sFlag = "COD_ESTPRO"
            If Not Filtrar(sFlag, Me, txtCod_EstPro, txtDes_estpro, False) Then
                Me.fraAddEstPro.Visible = True
                Me.fraAddEstPro.Top = Me.acbForm.Top
                Exit Sub
            End If
                                                
            Set obj = New clsTG_Cliente
            obj.ConexionString = cCONNECT
            If sModoAddEstCli = "SOLO ASIGNACION" Then
                obj.AddEStCli sCod_Cliente, sCod_TemCli, Me.txtIdestilo.Text, Me.txtNomestilo.Text, Me.txttelaestilo.Text, Me.txtCod_EstPro.Text, "ASIG"
            Else
                obj.AddEStCli sCod_Cliente, sCod_TemCli, Me.txtIdestilo.Text, Me.txtNomestilo.Text, Me.txttelaestilo.Text, Me.txtCod_EstPro.Text, "ADIC"
            End If
            Set obj = Nothing
            bOk = True
            sCod_EStCli = Me.txtIdestilo.Text
            
            Unload Me
        Case "CANCELAR"
            Unload Me
    End Select
    
    Exit Sub
errores:
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    ErrorHandler Err, Err.Description

    
End Sub

Private Sub funAceptar_ActionClick(ByVal index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "ACEPTAR"
        If RTrim(Me.txtDes_EstPro2.Text) = "" Then
            txtDes_EstPro2.SetFocus
            Exit Sub
        End If
        If RTrim(Me.txtCod_TipPre.Text) = "" Then
            txtCod_TipPre.SetFocus
            Exit Sub
        End If
        If RTrim(Me.txtCod_GruTal.Text) = "" Then
            txtCod_GruTal.SetFocus
            Exit Sub
        End If
        AdicionarEstPro
        Me.fraAddEstPro.Visible = False
    Case "CANCELAR"
        Me.fraAddEstPro.Visible = False
        
    End Select
End Sub

Private Sub txtCod_TipPre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sFlag = "COD_TIPPRE"
        If Filtrar(sFlag, Me, txtCod_TipPre, txtDes_TipPre) Then
            Me.txtCod_GruTal.SetFocus
        End If
    End If

End Sub

Private Sub txtCod_Grutal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sFlag = "COD_GRUTAL"
        If Filtrar(sFlag, Me, txtCod_GruTal, txtDes_GruTal) Then
            Me.funAceptar.SetFocus
        End If
    End If

End Sub

Private Sub txtCod_EstPro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sFlag = "COD_ESTPRO"
        If Not Filtrar(sFlag, Me, txtCod_EstPro, txtDes_estpro, False) Then
            Me.fraAddEstPro.Visible = True
            Me.fraAddEstPro.Top = Me.acbForm.Top
            Me.txtDes_EstPro2.Text = Me.txtNomestilo.Text
            Me.txtDes_EstPro2.SetFocus
        Else
            acbForm.SetFocus
        End If
    End If

End Sub

Private Sub AdicionarEstPro()
On Error GoTo errores
    Dim vbuff
    Dim obj As clsTG_PurOrd
    Dim sCod_Estpro As String
    
    Set obj = New clsTG_PurOrd
    obj.ConexionString = cCONNECT
    sCod_Estpro = obj.AddEstPro(Me.txtDes_EstPro2.Text, Me.txtCod_TipPre.Text, Me.txtCod_GruTal.Text)
    Set obj = Nothing
    
    Me.txtCod_EstPro.Text = sCod_Estpro
    sFlag = "COD_ESTPRO"
    Filtrar sFlag, Me, txtCod_EstPro, txtDes_estpro, False
    'Mensaje kMESSAGE_INF_NEW_CODIGO
    Exit Sub
errores:
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    ErrorHandler Err, Err.Description
End Sub

