VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmLiquidacionDiaria 
   Caption         =   "Liquidacion Diaria de Boletas Ventas"
   ClientHeight    =   2550
   ClientLeft      =   2685
   ClientTop       =   2280
   ClientWidth     =   3315
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2550
   ScaleWidth      =   3315
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   3015
      Begin NumBoxProject.NumBox txtImp_Dol 
         Height          =   285
         Left            =   1620
         TabIndex        =   2
         Top             =   1200
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         TypeVal         =   2
         Mask            =   "9,999,999,999.99"
         Formato         =   "#,###,###,###.##"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   3
         Text            =   "0.00"
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   2
      End
      Begin NumBoxProject.NumBox txtImp_Soles 
         Height          =   285
         Left            =   1620
         TabIndex        =   1
         Top             =   720
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         TypeVal         =   2
         Mask            =   "9,999,999,999.99"
         Formato         =   "#,###,###,###.##"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   3
         Text            =   "0.00"
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   2
      End
      Begin NumBoxProject.NumBox txtFecha 
         Height          =   285
         Left            =   1620
         TabIndex        =   0
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Importe Soles :"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   765
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Importe Dolares :"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1245
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   285
         Width           =   660
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   480
      TabIndex        =   3
      Top             =   1920
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmLiquidacionDiaria.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmLiquidacionDiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  txtFecha.Text = Date
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim sMessage As Long
Dim strSQL As String

On Error GoTo dprError

Select Case ActionName
  Case "ACEPTAR"
  
    If MsgBox("Esta seguro de efectuar la liquidacion boletas ", vbYesNo, "IMPORTANTE") = vbYes Then
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

SQL = "VT_LIQUIDACION_DIARIA '" & txtFecha.Text & "'," & txtImp_Soles.Text & "," & txtImp_Dol.Text
Call ExecuteCommandSQL(cCONNECT, SQL)

lfSalvar_Datos = True

Exit Function

hand:

errores Err.Number

lfSalvar_Datos = False

End Function

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtImp_Dol_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtImp_Soles_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
