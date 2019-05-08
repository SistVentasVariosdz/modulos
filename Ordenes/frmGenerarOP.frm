VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmGenerarOP 
   Caption         =   "Generación de O/P"
   ClientHeight    =   2040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4305
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraGenerar 
      Caption         =   "Modo de Generar O/P"
      Height          =   1935
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   4245
      Begin VB.OptionButton optModoGenerar 
         Caption         =   "Agrupado"
         Height          =   315
         Index           =   1
         Left            =   540
         TabIndex        =   2
         Top             =   750
         Width           =   1575
      End
      Begin VB.OptionButton optModoGenerar 
         Caption         =   "Simple"
         Height          =   315
         Index           =   0
         Left            =   510
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   1575
      End
      Begin FunctionsButtons.FunctButt acbForm 
         Height          =   510
         Left            =   870
         TabIndex        =   3
         Top             =   1230
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmGenerarOP.frx":0000
         Orientacion     =   0
         Style           =   1
         Language        =   1
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
   End
End
Attribute VB_Name = "frmGenerarOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Cliente As String
Public sCod_PurOrd As String
Public oParent As Object
Public bOk As Boolean
Private Sub acbForm_ActionClick(ByVal index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "ACEPTAR"
        GenerarOP
        bOk = True
        Unload Me
    Case "CANCELAR"
        Unload Me
    End Select
End Sub

Private Sub GenerarOP()
On Error GoTo errores
    Dim obj As clsTG_PurOrd
    Dim sFlg_Modo As String
    
    If optModoGenerar(0).value = True Then
        sFlg_Modo = "1"
    Else
        sFlg_Modo = "2"
    End If
    
    Set obj = New clsTG_PurOrd
    obj.ConexionString = cCONNECT
    obj.GenerarOP sCod_Cliente, sCod_PurOrd, sFlg_Modo, vusu
    Set obj = Nothing
    
    'Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
    MsgBox "Cambios realizados correctamente", vbInformation
    oParent.BUSCAR
    oParent.BuscarEStilos
    Exit Sub
errores:
    If Err.Number <> 91 Then
        ErrorHandler Err, Err.Description
    Else
        Resume Next
    End If
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If

End Sub
