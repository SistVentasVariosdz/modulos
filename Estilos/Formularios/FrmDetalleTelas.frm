VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form FrmDetalleTelas 
   Caption         =   "Alternativas Peso / Ancho"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2400
      TabIndex        =   1
      Top             =   1560
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmDetalleTelas.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame FraDatos 
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7335
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   285
         TabIndex        =   0
         Top             =   705
         Width           =   6840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   135
         TabIndex        =   3
         Top             =   375
         Width           =   840
      End
   End
End
Attribute VB_Name = "FrmDetalleTelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vCod_Tela As String
Public vAccion As String
Public vRuta As String
Public oParent As Object
Dim strSQL As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    Call Grabar
Case "CANCELAR"
    Unload Me
End Select
End Sub


Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub


Sub Grabar()
On Error GoTo errGrabar

If Trim(TxtDescripcion.Text) = "" Then
    MsgBox "Debe ingresar una descripción"
End If


strSQL = "tx_up_man_Tela_DatTecnicos_cabecera '" & vAccion & "','" & vCod_Tela & "','" & _
            vRuta & "','" & Trim(TxtDescripcion.Text) & "'"
            
ExecuteCommandSQL cCONNECT, strSQL

oParent.CARGA_GRID
Unload Me

Exit Sub
errGrabar:
    MsgBox Err.Description, vbCritical, "Grabar"
End Sub
