VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmTX_Rapport_Comb 
   Caption         =   "Mantenimiento Combinaciones Rapport"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2100
      TabIndex        =   3
      Top             =   1260
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "0~0~ACEPTAR~True~True~&Aceptar~0~0~1~~0~False~False~&Aceptar~~1~0~CANCELAR~True~True~&Cancelar~0~0~2~~0~False~False~&Cancelar~"
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
      Begin VB.TextBox TxtDes_Comb 
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
         Left            =   1485
         TabIndex        =   2
         Top             =   660
         Width           =   3870
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
         Left            =   2670
         TabIndex        =   4
         Top             =   210
         Width           =   3870
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
         Left            =   1500
         TabIndex        =   1
         Top             =   210
         Width           =   1080
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Des. Comb.:"
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
         Left            =   120
         TabIndex        =   6
         Tag             =   "Family :"
         Top             =   660
         Width           =   870
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
         TabIndex        =   5
         Tag             =   "Code"
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmTX_Rapport_Comb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public opcion As String
Public ultcomb As String
Public bOK As Boolean

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    If VALIDA = False Then
        Exit Sub
    End If
    Call GRABAR_RAPPORT_COMB
    
Case "CANCELAR"
    Unload Me
End Select
End Sub


Sub GRABAR_RAPPORT_COMB()
Dim con As New ADODB.Connection
On Error GoTo Salvar_DatosErr
Dim StrSql As String
Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    If opcion = "I" Then
        ultcomb = DevuelveCampo("SELECT ISNULL(MAX(rapport_comb),0) FROM TX_RAPPORT_COMB WHERE RAPPORT_NUMBER=" & CInt(Me.txtCodRapport.Text), cCONNECT)
        If ultcomb = "" Then
            ultcomb = "001"
        End If
        ultcomb = Right("000" & Trim(CInt(ultcomb) + 1), 3)
    End If
    
    con.ConnectionString = cCONNECT
    con.Open
    
    con.BeginTrans

    StrSql = "EXEC UP_MAN_TX_RAPPORT_COMB '" & opcion & "'," & Me.txtCodRapport.Text & ",'" & ultcomb & "','" & Me.TxtDes_Comb.Text & "','" & vusu & "','" & Format(Now, "DD/MM/YYYY") & "','" & ComputerName & "'"
                
    con.Execute StrSql
    con.CommitTrans
    
    Screen.MousePointer = vbDefault
    bOK = True
    
    Unload Me
    
    Exit Sub
Salvar_DatosErr:
    con.RollbackTrans
    Set con = Nothing
    Screen.MousePointer = vbDefault
    ErrorHandler Err, "GRABAR_RAPPORT_COMB"
End Sub
Function VALIDA() As Boolean
    VALIDA = False
    
    If Trim(TxtDes_Comb.Text) = "" Then
        MsgBox "Ingrese Descripcion Comb."
        VALIDA = False
        Exit Function
    End If
    VALIDA = True
    
End Function

Private Sub txtCodRapport_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


Private Sub TxtDes_Comb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
