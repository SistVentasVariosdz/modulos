VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmMerma 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mermas Especificas"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmMerma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1485
      TabIndex        =   1
      Top             =   1440
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
      Height          =   1185
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   5205
      Begin VB.TextBox TxtMer_Tintoreria 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4320
         TabIndex        =   7
         Text            =   "0"
         Top             =   690
         Width           =   720
      End
      Begin VB.TextBox TxtMer_Tejeduria 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1545
         TabIndex        =   6
         Text            =   "0"
         Top             =   660
         Width           =   735
      End
      Begin VB.TextBox TxtTela 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   645
         TabIndex        =   3
         Top             =   210
         Width           =   4380
      End
      Begin VB.Label Label3 
         Caption         =   "Merma Tintoreria Telas :"
         Height          =   225
         Left            =   2475
         TabIndex        =   5
         Top             =   705
         Width           =   1785
      End
      Begin VB.Label Label2 
         Caption         =   "Merma Tejeduria :"
         Height          =   210
         Left            =   150
         TabIndex        =   4
         Top             =   705
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Tela :"
         Height          =   225
         Left            =   105
         TabIndex        =   2
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmMerma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public varCod_Tela As String
Dim StrSql As String

Sub SALVAR_DATOS()
On Error GoTo Err

StrSql = "EXEC UP_MAN_TELA_MERMA '" & varCod_Tela & "'," & TxtMer_Tintoreria & "," & TxtMer_Tejeduria
Call ExecuteSQL(cCONNECT, StrSql)
Exit Sub

Err:
    ErrorHandler Err, "SALVAR_DATOS"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            SALVAR_DATOS
            Unload Me
        Case "CANCELAR"
            Unload Me
    End Select
End Sub

Private Sub TxtMer_Tejeduria_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtMer_Tintoreria.SetFocus
    Else
        Call SoloNumeros(TxtMer_Tejeduria, KeyAscii, True, 2, 3)
    End If
End Sub

Private Sub TxtMer_Tejeduria_LostFocus()
    If Trim(TxtMer_Tejeduria.Text) = "" Then
        TxtMer_Tejeduria.Text = 0
    Else
        If TxtMer_Tejeduria.Text > 100 Then
            MsgBox "El Porcentaje no puede ser mayor a 100", vbCritical, Me.Caption
            TxtMer_Tejeduria.SetFocus
        End If
    End If
End Sub

Private Sub TxtMer_Tintoreria_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FunctButt1.SetFocus
    Else
        Call SoloNumeros(TxtMer_Tintoreria, KeyAscii, True, 2, 3)
    End If
End Sub

Private Sub TxtMer_Tintoreria_LostFocus()
    If Trim(TxtMer_Tintoreria.Text) = "" Then
        TxtMer_Tintoreria.Text = 0
    Else
        If TxtMer_Tintoreria.Text > 100 Then
            MsgBox "El Porcentaje no puede ser mayor a 100", vbCritical, Me.Caption
            TxtMer_Tintoreria.SetFocus
        End If
    End If
End Sub
