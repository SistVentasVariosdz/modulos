VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmLetraControlStatusAbono 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Estatus Letra Abonada"
   ClientHeight    =   2670
   ClientLeft      =   660
   ClientTop       =   1155
   ClientWidth     =   6570
   Icon            =   "frmLetraControlStatusAbono.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6570
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1905
      TabIndex        =   7
      Top             =   2040
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmLetraControlStatusAbono.frx":030A
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   1920
      Left            =   0
      TabIndex        =   8
      Top             =   -45
      Width           =   6480
      Begin VB.OptionButton optCartera 
         Caption         =   "&Cartera"
         Height          =   375
         Left            =   4920
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optCobranzaGarantia 
         Caption         =   "En Cobranza &Garantia"
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optDescuento 
         Caption         =   "En &Descuento"
         Height          =   375
         Left            =   360
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optCobranza 
         Caption         =   "En C&obranza"
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtCod_Banco 
         Height          =   315
         Left            =   1590
         TabIndex        =   3
         Top             =   975
         Width           =   735
      End
      Begin VB.TextBox TxtNom_Banco 
         Height          =   315
         Left            =   2340
         TabIndex        =   4
         Top             =   975
         Width           =   3855
      End
      Begin VB.TextBox txtNumLetraBanco 
         Height          =   315
         Left            =   1590
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1320
         Width           =   1410
      End
      Begin NumBoxProject.NumBox txtFecha 
         Height          =   315
         Left            =   4830
         TabIndex        =   6
         Top             =   1350
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
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
         Left            =   120
         TabIndex        =   11
         Top             =   990
         Width           =   885
      End
      Begin VB.Label Label10 
         Caption         =   "Num. Letra Banco :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1350
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha :"
         Height          =   225
         Left            =   4200
         TabIndex        =   9
         Top             =   1395
         Width           =   570
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   195
         X2              =   6255
         Y1              =   840
         Y2              =   840
      End
   End
End
Attribute VB_Name = "frmLetraControlStatusAbono"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Codigo As String, Descripcion As String, lfSalvar As Boolean, strNum_Corre As String, strStatus As String

Private Sub Form_Load()
  strStatus = "D"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

On Error GoTo hand

    Select Case ActionName
        Case "ACEPTAR"
            SALVAR_DATOS
            Unload Me
            lfSalvar = True
        Case "CANCELAR"
            Unload Me
            lfSalvar = False
    End Select
    
    Exit Sub
hand:
errores Err.Number

End Sub

Sub SALVAR_DATOS()

Dim SQL As String

SQL = "Exec Cn_Ventas_Actualiza_Status_Letras '" & strNum_Corre & "','" & strStatus & "','" & txtFecha.Text & "','" & TxtCod_Banco & "','" & txtNumLetraBanco & "'"
ExecuteCommandSQL cCONNECT, SQL
  
End Sub

Private Sub optCartera_Click()
  strStatus = "C"
End Sub

Private Sub optCobranza_Click()
  strStatus = "B"
End Sub

Private Sub optCobranzaGarantia_Click()
  strStatus = "G"
End Sub

Private Sub optDescuento_Click()
  strStatus = "D"
End Sub

Private Sub TxtCod_Banco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("cod_banco", "nom_banco", "tg_banco where flg_operativo ='*' and ", TxtCod_Banco, TxtNom_Banco, 1, Me)
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtNom_Banco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("cod_banco", "nom_banco", "tg_banco where flg_operativo ='*' and ", TxtCod_Banco, TxtNom_Banco, 2, Me)
End Sub

Private Sub txtNumLetraBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

