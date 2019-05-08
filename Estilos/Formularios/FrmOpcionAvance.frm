VERSION 5.00
Begin VB.Form FrmOpcionAvance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciones de visualización"
   ClientHeight    =   2580
   ClientLeft      =   2550
   ClientTop       =   1485
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4245
   Begin VB.OptionButton optopcion 
      Caption         =   "Hilo Crudo"
      Height          =   195
      Index           =   4
      Left            =   630
      TabIndex        =   7
      Top             =   1575
      Width           =   1215
   End
   Begin VB.OptionButton optopcion 
      Caption         =   "Hilo Color"
      Height          =   195
      Index           =   3
      Left            =   630
      TabIndex        =   6
      Top             =   1245
      Width           =   1215
   End
   Begin VB.OptionButton optopcion 
      Caption         =   "Tela Cruda"
      Height          =   195
      Index           =   2
      Left            =   630
      TabIndex        =   5
      Top             =   915
      Width           =   1215
   End
   Begin VB.OptionButton optopcion 
      Caption         =   "Tela Acabada"
      Height          =   195
      Index           =   1
      Left            =   630
      TabIndex        =   4
      Top             =   585
      Width           =   1395
   End
   Begin VB.OptionButton optopcion 
      Caption         =   "Avios"
      Height          =   195
      Index           =   0
      Left            =   630
      TabIndex        =   3
      Top             =   270
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   2535
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame fraOpciones 
      Caption         =   "Opciones de visualización de avance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   90
      TabIndex        =   0
      Top             =   15
      Width           =   4065
   End
End
Attribute VB_Name = "FrmOpcionAvance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public varCod_GrupoPro As String
Public oParent As Object

Private Sub cmdAceptar_Click()
    oParent.varOpcionAvances = strSQL
    Unload Me
    Set oParent = Nothing
End Sub

Private Sub cmdSalir_Click()
    oParent.varOpcionAvances = ""
    Unload Me
    Set oParent = Nothing
End Sub

Private Sub Form_Load()
    optopcion(0).Value = True
End Sub

Private Sub optopcion_Click(Index As Integer)
    Select Case Index
        Case 0:
                strSQL = "EXEC SM_AVANCES_AVIOS '" & varCod_GrupoPro & "','I'"
        Case 1:
                strSQL = "EXEC SM_AVANCES_TELA_TENIDA '" & varCod_GrupoPro & "','T','T'"
        Case 2:
                strSQL = "EXEC SM_AVANCES_TELA_CRUDA '" & varCod_GrupoPro & "','T','C'"
        Case 3:
                strSQL = "EXEC SM_AVANCES_HILO_TENIDO '" & varCod_GrupoPro & "','H','T'"
        Case 4:
                strSQL = "EXEC SM_AVANCES_HILO_CRUDO '" & varCod_GrupoPro & "','H','C'"
    End Select
    
End Sub


