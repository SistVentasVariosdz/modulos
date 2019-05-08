VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmLetraControlStatusProtesto 
   Caption         =   "Protesto de  Letras"
   ClientHeight    =   2070
   ClientLeft      =   1620
   ClientTop       =   1785
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2070
   ScaleWidth      =   3840
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3495
      Begin NumBoxProject.NumBox txtFecha 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         TypeVal         =   3
         Mask            =   "99/99/9999"
         Formato         =   "dd/MM/yyyy"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   2
         Text            =   ""
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   0
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   405
         Width           =   660
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmLetraControlStatusProtesto.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmLetraControlStatusProtesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lfSalvar As Boolean, strNum_Corre As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

On Error GoTo hand
  
Select Case ActionName
  Case "ACEPTAR"
    If MsgBox("Esta seguro de Efectuar es Cambio de Estatus ", vbYesNo, "IMPORTANTE") = vbYes Then
      lfSalvar_Datos
      Unload Me
      lfAceptar = True
    End If
  Case "CANCELAR"
      Unload Me
      lfAceptar = False
End Select

Exit Sub

hand:

errores Err.Number

End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub lfSalvar_Datos()

Dim SQL As String

SQL = "Cn_Ventas_Protesto_Letras '" & strNum_Corre & "','" & txtFecha.Text & "'"
Call ExecuteCommandSQL(cCONNECT, SQL)

End Sub

