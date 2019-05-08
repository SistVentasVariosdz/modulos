VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmOtLote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignacion de Lote"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   825
      TabIndex        =   2
      Top             =   930
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
   Begin VB.TextBox txtLote 
      Height          =   300
      Left            =   1665
      TabIndex        =   0
      Top             =   375
      Width           =   2085
   End
   Begin VB.Label Label1 
      Caption         =   "Lote"
      Height          =   285
      Left            =   375
      TabIndex        =   1
      Top             =   405
      Width           =   1215
   End
End
Attribute VB_Name = "frmOtLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bCancel As Boolean

Private Sub Form_Load()
    bCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.Visible Then
        Me.Hide
        Cancel = 200
    End If
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    If ActionName = "ACEPTAR" Then
        AsignarLote
    Else
        Unload Me
    End If
End Sub

Private Sub AsignarLote()
    txtLote = Trim(txtLote)
    If txtLote = "" Then
        MsgBox "Lote Invalido", vbExclamation + vbOKOnly, "Asignar Lote"
        Exit Sub
    End If
    
    bCancel = False
    
    Unload Me
End Sub

Private Sub txtLote_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
