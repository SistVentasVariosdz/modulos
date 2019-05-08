VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmDivPre 
   Caption         =   "División de Prenda"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Garment Divition"
   Begin VB.TextBox txtCod_DivPre 
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
      TabIndex        =   1
      Top             =   60
      Width           =   2310
   End
   Begin VB.TextBox txtDes_DivPre 
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
      MaxLength       =   30
      TabIndex        =   0
      Top             =   465
      Width           =   4335
   End
   Begin FunctionsButtons.FunctButt acbForm 
      Height          =   510
      Left            =   1590
      TabIndex        =   2
      Top             =   1020
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
      Left            =   90
      TabIndex        =   4
      Tag             =   "Description"
      Top             =   495
      Width           =   945
   End
   Begin VB.Label Etiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Id Div Prenda :"
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
      Left            =   90
      TabIndex        =   3
      Tag             =   "Id Divition"
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "frmDivPre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Cliente As String
Public bOk As Boolean
Public sCod_DivPRe As String
Public oParent As Object

Private Sub acbForm_ActionClick(ByVal index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo errores
    Dim vbuff
    Dim obj As clsTG_LotColTal
    Dim sTipColor As String
    
    Select Case ActionName
    Case "ACEPTAR"
        If txtCod_DivPre.Text = "" Then
            If txtCod_DivPre.Enabled Then
                txtCod_DivPre.SetFocus
            End If
            Exit Sub
        End If
        
        If txtDes_DivPre.Text = "" Then
            If txtDes_DivPre.Enabled Then
                txtDes_DivPre.SetFocus
            End If
            Exit Sub
        End If
                
        Set obj = New clsTG_LotColTal
        obj.ConexionString = cCONNECT
        obj.AddDivPre txtCod_DivPre.Text, Me.txtDes_DivPre.Text
        Set obj = Nothing
        sCod_DivPRe = Me.txtCod_DivPre.Text
        bOk = True
        
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

