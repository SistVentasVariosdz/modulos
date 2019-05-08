VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmTransaccionesUpdCuadreManTipoCambio 
   ClientHeight    =   1950
   ClientLeft      =   3810
   ClientTop       =   2115
   ClientWidth     =   2835
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1950
   ScaleWidth      =   2835
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2535
      Begin VB.TextBox TxtTipo_Cambio_Otro 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Text            =   "0"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TxtTipo_Cambio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   2
         Text            =   "0"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "T./C. Otros :"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   735
         Width           =   885
      End
      Begin VB.Label Label27 
         Caption         =   "T./C.:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   255
         Width           =   495
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmTransaccionesUpdCuadreManTipoCambio.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmTransaccionesUpdCuadreManTipoCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dFecha As Date, intSecuencia As Integer, intSecuencia_Det As Integer, StrOption As String, _
       lfAceptar As Boolean, strCod_Moneda As String, strNum_Corre As String, strStore_Carga As String, strStore_Man As String
Public codigo As String, Descripcion As String, strTipo_Det As String, strCod_Anexo As String, strCod_TipAnexo, strStore As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim sMessage As Long
Dim strSQL As String

On Error GoTo dprError

Select Case ActionName
  Case "GRABAR"
    If MsgBox("Esta seguro de el tipo de Cambio ", vbYesNo, "IMPORTANTE") = vbYes Then
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

errores err.Number
End Sub

Private Function lfSalvar_Datos() As Boolean

On Error GoTo hand

SQL = strStore & " '" & StrOption & "','" & dFecha & "'," & intSecuencia & "," & intSecuencia_Det & "," & txtTipo_Cambio & "," & TxtTipo_Cambio_Otro
      
Call ExecuteCommandSQL(cCONNECT, SQL)

lfSalvar_Datos = True

Exit Function
Resume
hand:

errores err.Number

lfSalvar_Datos = False

End Function

Private Sub TxtTipo_Cambio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
  SoloNumeros txtTipo_Cambio, KeyAscii, True, 4, 2
End Sub

Private Sub TxtTipo_Cambio_Otro_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
  SoloNumeros TxtTipo_Cambio_Otro, KeyAscii, True, 8, 2
End Sub


