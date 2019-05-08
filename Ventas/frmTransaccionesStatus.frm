VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmTransaccionesStatus 
   Caption         =   "Control Partes de Cobranzas"
   ClientHeight    =   2580
   ClientLeft      =   1620
   ClientTop       =   1785
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2580
   ScaleWidth      =   3510
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   3255
      Begin VB.TextBox txtCod_Origen 
         Height          =   285
         Left            =   990
         MaxLength       =   1
         TabIndex        =   0
         Text            =   "N"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtDes_Origen 
         Height          =   285
         Left            =   1500
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin NumBoxProject.NumBox txtFecha_Cierre 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   720
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
      Begin NumBoxProject.NumBox txtFecha_Nuevo 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   1200
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
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Origen :"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   285
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Nuevo Parte :"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1245
         Width           =   1485
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Cierre:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   765
         Width           =   945
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmTransaccionesStatus.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmTransaccionesStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public codigo As String, Descripcion As String, lfAceptar As Boolean

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

Dim sMessage As Long
Dim strSQL As String

On Error GoTo dprError

Select Case ActionName
  Case "GRABAR"
  
    If MsgBox("Esta seguro efectuar el cambio de Cirre ", vbYesNo, "IMPORTANTE") = vbYes Then
      If lfSalvar_Datos Then
        frmTransacciones.inpFec_Emi.Text = txtFecha_Nuevo.Text
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

SQL = "CN_VENTAS_PARTES_COBRANZA_GENERACION '" & txtCod_Origen.Text & "'," & IIf(txtFecha_Cierre.Text <> "", "'" & txtFecha_Cierre.Text & "'", "NULL") & ",'" & txtFecha_Nuevo.Text & "'"
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

Private Sub txtFecha_Cierre_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFecha_Nuevo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
