VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmLetraControlStatus 
   Caption         =   "Control Status Letra"
   ClientHeight    =   2595
   ClientLeft      =   825
   ClientTop       =   1740
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2595
   ScaleWidth      =   5775
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   5535
      Begin VB.TextBox TxtNom_Banco 
         Height          =   315
         Left            =   1620
         TabIndex        =   3
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox TxtCod_Banco 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton optCartera 
         Caption         =   "&Cartera"
         Height          =   375
         Left            =   3360
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optPorAbonar 
         Caption         =   "&Aceptada"
         Height          =   375
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin NumBoxProject.NumBox txtFecha 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   1080
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
         Caption         =   "Banco:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   735
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1125
         Width           =   540
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1680
      TabIndex        =   5
      Top             =   1920
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmLetraControlStatus.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmLetraControlStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lfSalvar As Boolean, strNum_Corre As String
Public codigo As String, Descripcion As String

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

errores err.Number

End Sub

Private Sub TxtCod_Banco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("cod_banco", "nom_banco", "tg_banco where flg_operativo ='*' and ", TxtCod_Banco, TxtNom_Banco, 1, Me)
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub lfSalvar_Datos()

Dim SQL As String, strStatus As String

If optPorAbonar Then
  strStatus = "A"
Else
  strStatus = "C"
End If

SQL = "Cn_Ventas_Actualiza_Status_Letras '" & strNum_Corre & "','" & strStatus & "','" & txtFecha.Text & "','" & TxtCod_Banco & "'"

Call ExecuteCommandSQL(cCONNECT, SQL)

End Sub

Private Sub TxtNom_Banco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("cod_banco", "nom_banco", "tg_banco where flg_operativo ='*' and ", TxtCod_Banco, TxtNom_Banco, 1, Me)
End Sub
