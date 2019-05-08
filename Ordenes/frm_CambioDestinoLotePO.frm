VERSION 5.00
Begin VB.Form frm_CambioDestinoLotePO 
   Caption         =   "Destino Lote PO"
   ClientHeight    =   630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   630
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "&Actualizar"
      Height          =   285
      Left            =   6450
      TabIndex        =   3
      Top             =   180
      Width           =   1395
   End
   Begin VB.TextBox txtDes_DestinoLOT 
      Height          =   285
      Left            =   2220
      MaxLength       =   30
      TabIndex        =   1
      Top             =   180
      Width           =   4050
   End
   Begin VB.TextBox txtCod_DestinoLOT 
      Height          =   285
      Left            =   1575
      MaxLength       =   3
      TabIndex        =   0
      Top             =   180
      Width           =   615
   End
   Begin VB.Label labels 
      Caption         =   "Destino"
      Height          =   255
      Index           =   15
      Left            =   90
      TabIndex        =   2
      Tag             =   "Destination"
      Top             =   195
      Width           =   1200
   End
End
Attribute VB_Name = "frm_CambioDestinoLotePO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sCodCli As String
Public sCodPurOrd As String
Public sCodLote As String
Public sEstCli As String
Public sCodDestino As String
Public sDesDestino As String
Public sCod_DestinoLOT As String

Private Sub cmdActualizar_Click()

    On Error GoTo SALTO_ERROR

    
    If Trim(txtCod_DestinoLOT) = "" Then MsgBox "Ingrese DESTINO DEL LOTE", vbInformation, "Mensaje": txtCod_DestinoLOT.SetFocus: Exit Sub
    
    Strsql = "EXEC SM_TG_CAMBIA_DESTINO_LOTE '" & sCodCli & "'," & _
                                        "'" & sCodPurOrd & "'," & _
                                        "'" & sCodLote & "'," & _
                                        "'" & sEstCli & "'," & _
                                        "'" & Trim(txtCod_DestinoLOT.Text) & "'"
                                         
    Call ExecuteCommandSQL(cCONNECT, Strsql)
    MsgBox "El cambio de destino realizado satisfactoriamente", vbInformation, Me.Caption
    Unload Me
    Exit Sub
    
SALTO_ERROR:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Load()
    txtCod_DestinoLOT = sCodDestino
    txtDes_DestinoLOT = sDesDestino
End Sub

Private Sub txtCod_DestinoLOT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then

        If Filtrar("COD_DESTINOLOT", Me, txtCod_DestinoLOT, txtDes_DestinoLOT) Then
            cmdActualizar.SetFocus
        End If
    End If

End Sub
