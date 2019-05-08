VERSION 5.00
Begin VB.Form frmSelecFamilias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de Familias"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   3525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Seleccionar Todo"
      Height          =   240
      Left            =   315
      TabIndex        =   4
      Top             =   2955
      Width           =   2460
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   525
      Left            =   1860
      TabIndex        =   3
      Top             =   3255
      Width           =   1245
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   555
      Left            =   270
      TabIndex        =   2
      Top             =   3270
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Caption         =   "Familias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   3345
      Begin VB.ListBox LstFamilias 
         Height          =   2310
         Left            =   225
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   285
         Width           =   2940
      End
   End
End
Attribute VB_Name = "frmSelecFamilias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Dim Strsql As String
Dim Rs_Lista As ADODB.Recordset

Private Sub Check1_Click()
Dim i As Integer
    If Check1.Value = 1 Then
            For i = 0 To LstFamilias.ListCount - 1
                LstFamilias.Selected(i) = True
            Next
    Else
            For i = 0 To LstFamilias.ListCount - 1
                LstFamilias.Selected(i) = False
            Next
    End If
    LstFamilias.ListIndex = 0
End Sub

Private Sub cmdAceptar_Click()
    oParent.varCadena_Familias = CADENA_FAMILIAS(Me.LstFamilias)
    oParent.varCancelImpresion = 0
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    oParent.varCancelImpresion = 1
    Unload Me
End Sub

Public Sub CARGA_FAMILIAS()
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.ActiveConnection = cCONNECT
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    'Esta cadena es para devolver el Codigo de Cliente
    Strsql = "SELECT Cod_FamTela + ' - ' + RTRIM(Des_FamTela) FROM TX_FAMTELA where cod_tipfamtela = 'N' AND flg_status ='S' ORDER BY Des_FamTela"
    
    Rs_Lista.Open Strsql
    
    Me.LstFamilias.Clear
    If Rs_Lista.RecordCount > 0 Then
        Rs_Lista.MoveFirst
        While Not Rs_Lista.EOF And Not Rs_Lista.BOF
            Me.LstFamilias.AddItem Rs_Lista(0).Value
            Rs_Lista.MoveNext
        Wend
    Else
        Me.LstFamilias.Enabled = False
        Exit Sub
    End If
    
    Rs_Lista.Close
    Set Rs_Lista = Nothing
    
End Sub

Public Function CADENA_FAMILIAS(ByRef pListBox As Object) As String
    Dim Contador As Integer
    Dim Cadena As String    'Este es el prov del resultado
    
    Cadena = ""
    If pListBox.ListCount < 1 Then
        CADENA_FAMILIAS = ""
        Exit Function
    End If
    For Contador = 0 To pListBox.ListCount - 1
        If pListBox.Selected(Contador) Then
            Cadena = Cadena & "." & Mid(pListBox.List(Contador), 1, 2) & ".,"
            'MsgBox List1.List(nro)
        End If
    Next
    Cadena = Trim(Cadena)
    If Cadena <> "" Then
        Cadena = Mid(Cadena, 1, Len(Cadena) - 1)
    End If
    
    CADENA_FAMILIAS = Cadena
    
End Function
