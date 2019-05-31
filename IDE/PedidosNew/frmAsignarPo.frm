VERSION 5.00
Begin VB.Form frmAsignarPo 
   Caption         =   "Asignar PO"
   ClientHeight    =   2430
   ClientLeft      =   5145
   ClientTop       =   4545
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   5250
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtEstLoteNumero 
         Height          =   285
         Left            =   2760
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtEstLotehija 
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtpohija 
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Estilo Cliente :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Purchase Order :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmAsignarPo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public oParent     As Object

Public Codigo      As String

Public Descripcion As String

Private Sub cmdAceptar_Click()
    SALVAR_DATOS
    Unload Me
End Sub

Sub SALVAR_DATOS()

    Dim Con As New ADODB.Connection

    Dim rs  As New ADODB.Recordset

    On Error GoTo Salvar_DatosErr

    Dim strSql As String
    
    Con.ConnectionString = cCONNECT
    Con.Open
    
    Con.BeginTrans

    strSql = "EXEC Tg_Asigna_Po_Hija_Madre '" & oParent.COD_CLIENTE & "','" & oParent.cod_purord & "','" & oParent.cod_lotpurord & "','" & oParent.cod_estcli & "','" & txtpohija & "','" & txtEstLotehija & "','" & txtEstLoteNumero & "'"
        
    Con.Execute strSql
       
    Con.CommitTrans

    Dim amensaje As New clsMessages

    amensaje.Codigo = MESSAGECODE.kMESSAGE_INF_DATA_SAVE
    Informa "", amensaje
        
    oParent.cod_purodHija = txtpohija
    oParent.cod_lotpurordHija = txtEstLotehija
    oParent.cod_estcliHija = txtEstLoteNumero

    Exit Sub

Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub txtEstLotehija_keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        BuscaLote (1)
    End If

End Sub

Private Sub txtEstLoteNumero_keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        BuscaLote (2)
    End If

End Sub

Private Sub txtpohija_keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtEstLotehija.SetFocus
    End If

End Sub

Sub BuscaLote(Tipo As Integer)

    Dim oTipo As New frmBusqGeneral3

    Dim rs    As New ADODB.Recordset

    Set oTipo.oParent = Me

    If Tipo = 1 Then
        oTipo.sQuery = "Tg_Muestra_LotEst '" & oParent.COD_CLIENTE & "','" & txtpohija & "','" & txtEstLotehija & "',''"
    ElseIf Tipo = 2 Then
        oTipo.sQuery = "Tg_Muestra_LotEst '" & oParent.COD_CLIENTE & "','" & txtpohija & "','','" & txtEstLoteNumero & "'"
    End If

    oTipo.Cargar_Datos
    oTipo.Show 1

    If Codigo <> "" Then
        txtEstLotehija.Text = Trim(Codigo)
        txtEstLoteNumero.Text = Trim(Descripcion)
        Codigo = "": Descripcion = ""
        cmdAceptar.SetFocus
    Else
        txtEstLotehija.Text = ""
        txtEstLoteNumero.Text = ""
    End If

    Set oTipo = Nothing
    Set rs = Nothing

End Sub
