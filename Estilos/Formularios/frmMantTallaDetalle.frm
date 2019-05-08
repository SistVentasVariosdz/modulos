VERSION 5.00
Begin VB.Form frmMantTallaDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tallas"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   480
      Left            =   3360
      TabIndex        =   9
      Top             =   4350
      Width           =   1230
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Grabar"
      Height          =   480
      Left            =   1440
      TabIndex        =   8
      Top             =   4350
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tallas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Left            =   75
      TabIndex        =   0
      Top             =   405
      Width           =   5940
      Begin VB.CommandButton cmdIzqAll 
         Caption         =   "<<"
         Height          =   480
         Left            =   2730
         TabIndex        =   6
         Top             =   2430
         Width           =   480
      End
      Begin VB.CommandButton cmdIzq 
         Caption         =   "<"
         Height          =   480
         Left            =   2730
         TabIndex        =   5
         Top             =   1965
         Width           =   480
      End
      Begin VB.CommandButton cmdDer 
         Caption         =   ">"
         Height          =   480
         Left            =   2730
         TabIndex        =   4
         Top             =   1500
         Width           =   480
      End
      Begin VB.CommandButton cmdDerAll 
         Caption         =   ">>"
         Height          =   480
         Left            =   2730
         TabIndex        =   3
         Top             =   1035
         Width           =   480
      End
      Begin VB.ListBox lstTallasSelec 
         Height          =   3375
         ItemData        =   "frmMantTallaDetalle.frx":0000
         Left            =   3375
         List            =   "frmMantTallaDetalle.frx":0002
         TabIndex        =   2
         Top             =   300
         Width           =   2350
      End
      Begin VB.ListBox LstTallas 
         Height          =   3375
         ItemData        =   "frmMantTallaDetalle.frx":0004
         Left            =   195
         List            =   "frmMantTallaDetalle.frx":0006
         TabIndex        =   1
         Top             =   300
         Width           =   2350
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Selección deTallas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2190
      TabIndex        =   7
      Top             =   120
      Width           =   1635
   End
End
Attribute VB_Name = "frmMantTallaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Strsql As String
Dim Rs_Lista As ADODB.Recordset
Dim CadConn  As New ADODB.Connection
Public varCod_grutal As String

Public Sub CARGA_TALLAS()

    'Aqui llenamos con los colores pendientes
    Strsql = "SELECT Cod_Talla + SPACE(100) + '0' FROM TG_TALLA WHERE cod_talla NOT IN (SELECT cod_talla FROM ES_TALLASDET WHERE cod_grutal = '" & varCod_grutal & "')"
    Call LlenaCombo(Me.LstTallas, Strsql, cCONNECT)
    
    'Aqui llenamos con los colores usados
    'Strsql = "SELECT cod_talla + SPACE(100) + '1' FROM ES_TALLASDET WHERE cod_grutal = '" & varCod_grutal & "'"
     Strsql = " exec ES_MUESTRA_TALLAS_ASIGNADAS_GRUPO '" & varCod_grutal & "'"
    Call LlenaCombo(Me.lstTallasSelec, Strsql, cCONNECT)
    
    
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo ErrorAceptar:

    Set CadConn = Nothing
    CadConn.Open cCONNECT
    Dim j As Integer
    
    Dim varItem As String
    Dim LonItem As Integer
    
    For j = 0 To LstTallas.ListCount - 1
    
        LonItem = Len(LstTallas.List(j))
        varItem = Trim(Mid(LstTallas.List(j), 1, LonItem - 1))
        
        If Right(LstTallas.List(j), 1) = 2 Then
            'Aqui ponemos la validacion de la eliminacion
            Strsql = "SELECT COUNT(*) FROM Es_OrdproColTal a , Es_ordpro b , Es_estpro c " & _
                    "WHERE A.Cod_fabrica = B.Cod_fabrica And A.Cod_ordpro = B.Cod_ordpro " & _
                    "AND b.Cod_estpro = c.Cod_estpro " & _
                    "AND c.Cod_grutal = '" & varCod_grutal & "' " & _
                    "AND A.Cod_talla = '" & varItem & "'"
        
            If DevuelveCampo(Strsql, cCONNECT) <> 0 Then
                Call MsgBox("La talla " & varItem & " no puede ser eliminada, por que posee referencias. Sirvase verificar", vbInformation, "Tallas")
                Exit Sub
            End If
        
            'Aqui se pone el codigo para la eliminacion
            Strsql = "EXEC UP_MAN_TALLASDET 'D','" & varCod_grutal & "','" & varItem & "',0"
            CadConn.Execute Strsql
            
        End If
        
    Next
    'Set CadConn = Nothing
    
    'Esto es para adicionar las tallas
    For j = 0 To lstTallasSelec.ListCount - 1
    
        LonItem = Len(lstTallasSelec.List(j))
        varItem = Trim(Mid(lstTallasSelec.List(j), 1, LonItem - 1))
        
        If Right(lstTallasSelec.List(j), 1) = 3 Then
'            'Aqui ponemos la validacion de la eliminacion
'            Strsql = "SELECT COUNT(*) FROM Es_OrdproColTal a , Es_ordpro b , Es_estpro c " & _
'                    "WHERE A.Cod_fabrica = B.Cod_fabrica And A.Cod_ordpro = B.Cod_ordpro " & _
'                    "AND b.Cod_estpro = c.Cod_estpro " & _
'                    "AND c.Cod_grutal = '" & varCod_grutal & "' " & _
'                    "AND A.Cod_talla = '" & varItem & "'"
'
'            If DevuelveCampo(Strsql, cCONNECT) <> 0 Then
'                Call MsgBox("La talla no puede ser eliminada, por que posee referencias. Sirvase verificar", vbInformation, "Tallas")
'                Exit Sub
'            End If
        
            'Aqui se pone el codigo para la eliminacion
            Strsql = "EXEC UP_MAN_TALLASDET 'I','" & varCod_grutal & "','" & varItem & "',0"
            CadConn.Execute Strsql
            
        End If
        
    Next
    Set CadConn = Nothing
    
    Call MsgBox("La actualizacion de tallas fue exitosa", vbInformation, "Tallas")
    Call Me.CARGA_TALLAS
    
    'Unload Me
    Exit Sub
ErrorAceptar:
    Set CadConn = Nothing
    ErrorHandler Err, "Error Aceptar"


End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdDer_Click()
    If LstTallas.ListIndex = -1 Then
        Exit Sub
    Else
        Call ListBoxToListBox(Me.LstTallas, lstTallasSelec, 0)
    End If
End Sub

Public Sub ListBoxToListBox(ByRef lstOrigen As Object, ByRef lstDestino As Object, ByVal iModal As Integer)
'iModal 0 Pasa item actual
'       1 Pasa todos los items

Dim i As Long
Dim j As Long
Dim longLista As Integer
Dim varItem As String
Dim varSelec As String

If iModal = 0 Then
    If lstOrigen.ListIndex <> -1 Then
        lstDestino.AddItem ""
        For i = 0 To 0 ' lstOrigen.ColumnCount - 1
            longLista = Len(lstOrigen.List(lstOrigen.ListIndex))
            varItem = lstOrigen.List(lstOrigen.ListIndex)
            Select Case Right(varItem, 1)
                Case "0":   varSelec = "3"
                Case "1":   varSelec = "2"
                Case "2":   varSelec = "1"
                Case "3":   varSelec = "0"
            End Select
            
            lstDestino.List(lstDestino.ListCount - 1) = Mid(lstOrigen.List(lstOrigen.ListIndex), 1, longLista - 1) & varSelec
            'lstDestino.List(lstDestino.ListCount - 1) = lstOrigen.List(lstOrigen.ListIndex)
            
        Next
        lstOrigen.RemoveItem lstOrigen.ListIndex
    End If
Else
    For j = 0 To lstOrigen.ListCount - 1
        If RTrim(lstOrigen.List(j)) <> "" Then
            lstDestino.AddItem ""
            For i = 0 To 0  ' lstOrigen.ColumnCount - 1
                longLista = Len(lstOrigen.List(j))
                varItem = lstOrigen.List(j)
                Select Case Right(varItem, 1)
                    Case "0":   varSelec = "3"
                    Case "1":   varSelec = "2"
                    Case "2":   varSelec = "1"
                    Case "3":   varSelec = "0"
                End Select
            
                lstDestino.List(lstDestino.ListCount - 1) = Mid(lstOrigen.List(j), 1, longLista - 1) & varSelec
                'lstDestino.List(lstDestino.ListCount - 1) = lstOrigen.List(lstOrigen.ListIndex)
            
                'lstDestino.List(lstDestino.ListCount - 1) = lstOrigen.List(j)
            Next
        End If
    Next
    
    For j = lstOrigen.ListCount - 1 To 0 Step -1
        lstOrigen.RemoveItem j
    Next
    
End If
End Sub

Private Sub cmdDerAll_Click()
    Call ListBoxToListBox(Me.LstTallas, lstTallasSelec, 1)
End Sub

Private Sub cmdIzq_Click()
    If lstTallasSelec.ListIndex = -1 Then
        Exit Sub
    Else
        Call ListBoxToListBox(Me.lstTallasSelec, LstTallas, 0)
    End If

End Sub

Private Sub cmdIzqAll_Click()
    Call ListBoxToListBox(Me.lstTallasSelec, LstTallas, 1)
End Sub

Private Sub Form_Load()

End Sub

Private Sub LstTallas_DblClick()
    cmdDer_Click
End Sub

Private Sub LstTallas_KeyPress(KeyAscii As Integer)
    Dim varindex As Integer
    If KeyAscii = 13 Then
        If LstTallas.ListCount < 1 Or LstTallas.ListIndex = -1 Then
            Exit Sub
        End If
        
        varindex = LstTallas.ListIndex
        cmdDer_Click
        
        If varindex > 0 Then
            varindex = varindex - 1
        Else
            varindex = 0
        End If
        
        If LstTallas.ListCount > 0 Then
            LstTallas.ListIndex = varindex
        End If
    End If
End Sub

Private Sub lstTallasSelec_DblClick()
    cmdIzq_Click
End Sub

Private Sub lstTallasSelec_KeyPress(KeyAscii As Integer)
    Dim varindex  As Integer
    If KeyAscii = 13 Then
        If lstTallasSelec.ListCount < 1 Or lstTallasSelec.ListIndex = -1 Then
            Exit Sub
        End If
        
        varindex = lstTallasSelec.ListIndex
        cmdIzq_Click
        
        If varindex > 0 Then
            varindex = varindex - 1
        Else
            varindex = 0
        End If
        
        If lstTallasSelec.ListCount > 0 Then
            lstTallasSelec.ListIndex = varindex
        End If
    End If
End Sub
