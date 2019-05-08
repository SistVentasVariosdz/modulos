VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmTransaccionesStatusReversion 
   Caption         =   "Control Partes de Cobranzas"
   ClientHeight    =   2070
   ClientLeft      =   1620
   ClientTop       =   1785
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   3840
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   3495
      Begin VB.TextBox txtNum_Parte 
         Height          =   285
         Left            =   1830
         TabIndex        =   5
         Top             =   720
         Width           =   1410
      End
      Begin VB.TextBox txtCod_Origen 
         Height          =   285
         Left            =   1110
         MaxLength       =   1
         TabIndex        =   0
         Text            =   "N"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtDes_Origen 
         Height          =   285
         Left            =   1665
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Num. Parte Cobranza :"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   750
         Width           =   1605
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Origen :"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   285
         Width           =   555
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmTransaccionesStatusReversion.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmTransaccionesStatusReversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public codigo As String, Descripcion As String, lfSalvar As Boolean

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

Dim sMessage As Long
Dim strSQL As String

On Error GoTo dprError

Select Case ActionName
  Case "GRABAR"
  
    If MsgBox("Esta seguro Abrir este Parte de Cobranza ", vbYesNo, "IMPORTANTE") = vbYes Then
      If lfSalvar_Datos Then
        Unload Me
        lfAceptar = True
      End If
    End If
    
  Case "CANCELAR"
      Unload Me
      lfAceptar = False
End Select

Exit Sub

dprError:

errores Err.Number
End Sub

Private Function lfSalvar_Datos() As Boolean

On Error GoTo hand

Dim SQL As String

SQL = "CN_VENTAS_PARTES_COBRANZA_ABRIR '" & txtCod_Origen.Text & "','" & txtNum_Parte & "'"
Call ExecuteCommandSQL(cCONNECT, SQL)

lfSalvar_Datos = True

Exit Function

hand:

errores Err.Number

lfSalvar_Datos = False

End Function


Private Sub txtCod_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Origen", "Des_Origen", " Cn_Origen where ", txtCod_Origen, txtDes_Origen, 1, Me)
End Sub

Private Sub txtDes_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Origen", "Des_Origen", " Cn_Origen where ", txtCod_Origen, txtDes_Origen, 2, Me)
End Sub

Private Sub txtNum_Parte_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
