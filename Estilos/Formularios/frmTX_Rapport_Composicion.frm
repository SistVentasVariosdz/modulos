VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmTX_Rapport_Composicion 
   Caption         =   "Composición del Rapport"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2100
      TabIndex        =   4
      Top             =   1260
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "0~0~ACEPTAR~True~True~&Aceptar~0~0~1~~0~False~False~&Aceptar~~1~0~CANCELAR~True~True~&Cancelar~1~0~3~~0~False~False~&Cancelar~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   1170
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.TextBox TxtSecuencia 
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
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   2
         Top             =   630
         Width           =   1410
      End
      Begin VB.TextBox txtPorcentaje 
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
         Left            =   5070
         TabIndex        =   3
         Top             =   630
         Width           =   1410
      End
      Begin VB.TextBox txtCodRapport 
         Enabled         =   0   'False
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
         Left            =   1470
         MaxLength       =   8
         TabIndex        =   1
         Top             =   210
         Width           =   1395
      End
      Begin VB.TextBox txtDesRapport 
         Enabled         =   0   'False
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
         Left            =   2985
         TabIndex        =   5
         Top             =   210
         Width           =   3555
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje :"
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
         Left            =   3675
         TabIndex        =   8
         Tag             =   "Family :"
         Top             =   675
         Width           =   855
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Secuencia :"
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
         Left            =   105
         TabIndex        =   7
         Tag             =   "Family :"
         Top             =   675
         Width           =   855
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Rapport:"
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
         Left            =   120
         TabIndex        =   6
         Tag             =   "Code"
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmTX_Rapport_Composicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Opcion As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    Call GRABAR_RAPPORT_COMPOSICION
Case "CANCELAR"
Unload Me
End Select
End Sub

Sub GRABAR_RAPPORT_COMPOSICION()
Dim con As New ADODB.Connection
On Error GoTo Salvar_DatosErr
Dim StrSql As String
Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    con.ConnectionString = cCONNECT
    con.Open
    
    con.BeginTrans

    StrSql = "EXEC UP_MAN_TX_RAPPORT_COMPOSICION '" & Opcion & "'," & Me.txtCodRapport.Text & ",'" & TxtSecuencia.Text & "'," & txtPorcentaje.Text
                
    con.Execute StrSql
    con.CommitTrans
    
    Screen.MousePointer = vbDefault
    Unload Me
    
    Exit Sub
Salvar_DatosErr:
    con.RollbackTrans
    Set con = Nothing
    Screen.MousePointer = vbDefault
    ErrorHandler Err, "Mantenimiento Rapport Composicion"
End Sub



Private Sub txtCodRapport_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


Private Sub txtPorcentaje_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtSecuencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

