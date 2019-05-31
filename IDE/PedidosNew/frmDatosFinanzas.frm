VERSION 5.00
Begin VB.Form frmDatosFinanzas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos Finanzas"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDatosFinanzas.frx":0000
   ScaleHeight     =   2850
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Chage of PO"
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   510
      Left            =   2520
      TabIndex        =   9
      Tag             =   "&Cancel"
      Top             =   2160
      Width           =   1485
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Tag             =   "&Accept"
      Top             =   2160
      Width           =   1470
   End
   Begin VB.Frame fraActualiza 
      Height          =   585
      Left            =   30
      TabIndex        =   7
      Top             =   1470
      Width           =   5070
      Begin VB.CheckBox chkcommit 
         Caption         =   "Commit"
         Height          =   255
         Left            =   3120
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkop 
         Caption         =   "PO"
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraDatos 
      Height          =   1500
      Left            =   30
      TabIndex        =   0
      Top             =   -15
      Width           =   5085
      Begin VB.TextBox txtEstilo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1365
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1005
         Width           =   2955
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1365
         MaxLength       =   20
         TabIndex        =   5
         Top             =   645
         Width           =   2940
      End
      Begin VB.TextBox txtPO 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1365
         MaxLength       =   20
         TabIndex        =   4
         Top             =   285
         Width           =   2940
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Estilo :"
         Height          =   195
         Left            =   225
         TabIndex        =   3
         Tag             =   "Style :"
         Top             =   1035
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   210
         TabIndex        =   2
         Tag             =   "Client :"
         Top             =   705
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PO :"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Tag             =   "PO :"
         Top             =   345
         Width           =   315
      End
   End
End
Attribute VB_Name = "frmDatosFinanzas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSql As String

Public varCod_Cliente, varCod_EstCli, varCod_LotPurOrd, varCod_TemCli As String

Public bNivelPO As Boolean

Dim Rs_Lista    As New ADODB.Recordset

Public oParent  As Object

Sub CAMBIO()

    Dim oCone As ADODB.Connection

    Dim spo   As String, scommit As String

    On Error GoTo AceptaErr
    
    Set oCone = New ADODB.Connection
    
    oCone.CursorLocation = adUseClient
    oCone.Open cCONNECT
    oCone.CommandTimeout = 4000
           
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.ActiveConnection = oCone
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    If chkop.value = 0 Then
        spo = "N"
    Else
        spo = "S"
    End If
    
    If chkcommit.value = 0 Then
        scommit = "N"
    Else
        scommit = "S"
    End If

    strSql = "EXEC TG_ACTUALIZA_IF_PO '" & varCod_Cliente & "','" & txtPO.Text & "','" & spo & "','" & scommit & "'"

    Rs_Lista.Open strSql
    MsgBox "Los cambios se guardaron exitosamente", vbInformation, "PO Change"
    Call oParent.BUSCAR
    Unload Me

    Exit Sub

AceptaErr:
    
    If Err.Number <> 91 Then
        ErrorHandler Err, "Error en Cambio"
    Else

        Resume Next

    End If
   
End Sub

Private Sub cmdAceptar_Click()

    Dim opcion As Integer

    Call CAMBIO
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

