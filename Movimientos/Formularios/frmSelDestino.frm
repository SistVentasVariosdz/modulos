VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmSelDestino 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar Almacen Destino"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboAlmacen 
      Height          =   315
      Left            =   1485
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   225
      Width           =   2505
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1170
      TabIndex        =   2
      Top             =   750
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmSelDestino.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Label Label1 
      Caption         =   "Almacen"
      Height          =   255
      Left            =   300
      TabIndex        =   1
      Top             =   270
      Width           =   1155
   End
End
Attribute VB_Name = "frmSelDestino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bCancel As Boolean, sCod_Almacen As String, sCod_TipMov As String
Dim strSQL As String, sErr As String, sTit As String

Private Sub cboAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    bCancel = True
End Sub

Public Sub MostrarAlm()
    strSQL = "EXEC LG_SM_MUESTRA_ALMACEN_TRANS '" & sCod_Almacen & "', '" & _
             sCod_TipMov & "'"
    LlenaCombo cboAlmacen, strSQL, cConnect
    If cboAlmacen.ListCount > 0 Then cboAlmacen.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.Visible Then
        Cancel = 200
        Me.Hide
    End If
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    If ActionName = "ACEPTAR" Then bCancel = False
    Unload Me
End Sub
