VERSION 5.00
Begin VB.Form frmAddTX_Ordtra_Ordenes 
   Caption         =   "Selección de O/P "
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5730
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Filtro de O/P's x Orden de Compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   45
      TabIndex        =   9
      Top             =   15
      Width           =   6060
      Begin VB.TextBox txtCOD_ORDCOMP 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3705
         TabIndex        =   13
         Top             =   270
         Width           =   1950
      End
      Begin VB.TextBox txtSER_ORDCOMP 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   11
         Top             =   270
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   2895
         TabIndex        =   12
         Top             =   315
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Serie :"
         Height          =   195
         Left            =   210
         TabIndex        =   10
         Top             =   315
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   510
      Left            =   4035
      TabIndex        =   2
      Top             =   5145
      Width           =   1395
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   510
      Left            =   765
      TabIndex        =   1
      Top             =   5145
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      Left            =   45
      TabIndex        =   0
      Top             =   735
      Width           =   6060
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3375
         TabIndex        =   15
         Text            =   "O/P 's Seleccionadas"
         Top             =   240
         Width           =   2565
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   135
         TabIndex        =   14
         Text            =   "O/P's Disponibles"
         Top             =   225
         Width           =   2535
      End
      Begin VB.CommandButton cmdIzqAll 
         Caption         =   "<<"
         Height          =   480
         Left            =   2805
         TabIndex        =   5
         Top             =   2535
         Width           =   480
      End
      Begin VB.CommandButton cmdIzq 
         Caption         =   "<"
         Height          =   480
         Left            =   2805
         TabIndex        =   6
         Top             =   2070
         Width           =   480
      End
      Begin VB.CommandButton cmdDer 
         Caption         =   ">"
         Height          =   480
         Left            =   2805
         TabIndex        =   7
         Top             =   1605
         Width           =   480
      End
      Begin VB.ListBox lstOrdProSelec 
         BackColor       =   &H80000018&
         Height          =   3570
         ItemData        =   "frmAddTX_Ordtra_Ordenes.frx":0000
         Left            =   3375
         List            =   "frmAddTX_Ordtra_Ordenes.frx":0002
         TabIndex        =   4
         Top             =   540
         Width           =   2580
      End
      Begin VB.ListBox lstOrdPro 
         BackColor       =   &H80000018&
         Height          =   3570
         ItemData        =   "frmAddTX_Ordtra_Ordenes.frx":0004
         Left            =   120
         List            =   "frmAddTX_Ordtra_Ordenes.frx":0006
         TabIndex        =   3
         Top             =   540
         Width           =   2580
      End
      Begin VB.CommandButton cmdDerAll 
         Caption         =   ">>"
         Height          =   480
         Left            =   2805
         TabIndex        =   8
         Top             =   1140
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmAddTX_Ordtra_Ordenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim sTipo As String
Public varSer_OrdComp As String
Public varCod_OrdComp As String, Flg_Requerimiento As String

Public varCod_TipOrdTra As String
Public varCod_OrdTra As String

Public oParent As frmMovAlmacenAnexo

Private Sub cmdCancelar_Click()
    oParent.varCancelar = True
    Unload Me
End Sub

Private Sub cmdGrabar_Click()
    
    If Me.lstOrdProSelec.ListCount > 0 Or Flg_Requerimiento = "N" Then
        
        Call oParent.ANADE_PARTIDA
        
        Me.varCod_TipOrdTra = "TI"
        Me.varCod_OrdTra = oParent.txtCod_Ordtra1er
        
        Call SALVAR_DATOS
        Unload Me
    Else
        MsgBox "No existen O/P seleccionadas. Sirvase verificar", vbInformation, "Mensaje"
    End If
    
End Sub

Private Sub Form_Load()
    sTipo = "I"
End Sub

Public Sub CARGA_DATOS()
    strSQL = "EXEC SM_MUESTRA_ORDENES_PRODUCCION_ORDEN_COMPRA '" & Me.varSer_OrdComp & "','" & Me.varCod_OrdComp & "'"
    
    Dim Rs_Lista As ADODB.Recordset
    Set Rs_Lista = New ADODB.Recordset
    
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.ActiveConnection = cConnect
    
    Rs_Lista.Open strSQL
    Me.lstOrdPro.Clear
    If Rs_Lista.RecordCount > 0 Then
        Rs_Lista.MoveFirst
        Do Until Rs_Lista.EOF
            Me.lstOrdPro.AddItem Rs_Lista("COD_ORDPRO").Value & "-" & Rs_Lista("DES_ESTPRO").Value & Space(100) & Rs_Lista("COD_FABRICA").Value
            Rs_Lista.MoveNext
        Loop
    End If
    
End Sub

Private Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    Dim i As Integer
    On Error GoTo Salvar_DatosErr
    Dim strSQL As String
    
    Con.ConnectionString = cConnect
    Con.Open
    
        Con.BeginTrans

        For i = 1 To Me.lstOrdProSelec.ListCount
        
            strSQL = "EXEC UP_MAN_TX_OrdTra_Ordenes '" & _
            sTipo & "','" & _
            Me.varCod_TipOrdTra & "','" & _
            Me.varCod_OrdTra & "','" & _
            Right(Me.lstOrdProSelec.List(i - 1), 3) & "','" & _
            Left(Me.lstOrdProSelec.List(i - 1), 5) & "',NULL,0"
        
            Con.Execute strSQL
            
        Next

        Con.CommitTrans
       
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub

Private Sub cmdDerAll_Click()
    Call ComboBoxToComboBox(Me.lstOrdPro, Me.lstOrdProSelec, 1)
End Sub

Private Sub cmdDer_Click()
    If lstOrdPro.ListIndex = -1 Then
        Exit Sub
    Else
        Call ComboBoxToComboBox(Me.lstOrdPro, Me.lstOrdProSelec, 0)
        'If varFlg_OrdPro = True Then
            'Me.lstOrdProSelec.ListIndex = Me.lstOrdProSelec.ListCount - 1
            'Call lstOrdProSelec_DblClick
        'End If
    End If
End Sub

Private Sub cmdIzq_Click()
    If lstOrdProSelec.ListIndex = -1 Then
        Exit Sub
    Else
        Call ComboBoxToComboBox(lstOrdProSelec, lstOrdPro, 0)
    End If
End Sub
Private Sub cmdIzqAll_Click()
    Call ComboBoxToComboBox(lstOrdProSelec, lstOrdPro, 1)
End Sub
