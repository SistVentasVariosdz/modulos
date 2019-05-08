VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form FrmCopiarEstiloNew 
   Caption         =   "Copiar Estilo"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmCopiarEstiloNew.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4575
      Begin VB.TextBox TxtCod_EstiloNew 
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Estilo"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   330
         Width           =   900
      End
   End
End
Attribute VB_Name = "FrmCopiarEstiloNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrSQL As String
Public varCod_Cliente As String, varCod_TemCli_origen As String, varCod_EstCli As String


Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    Call Grabar
Case "SALIR"
    Unload Me
End Select
End Sub

Sub Grabar()
On Error GoTo errGrabar

StrSQL = "es_copia_estilo_cliente_a_otro '" & varCod_Cliente & "','" & varCod_TemCli_origen & "','" & varCod_EstCli & "','" & Trim(TxtCod_EstiloNew.Text) & "'"

ExecuteSQL cCONNECT, StrSQL

Unload Me
Exit Sub
errGrabar:
    ErrorHandler Err, "Grabar"
End Sub

Private Sub TxtCod_EstiloNew_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
