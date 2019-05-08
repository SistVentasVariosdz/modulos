VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMantPartidasArancelarias 
   Caption         =   "Partidas Arancelarias Item"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   8475
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraLista 
      Caption         =   "Lista:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   15
      TabIndex        =   6
      Top             =   -15
      Width           =   8415
      Begin GridEX20.GridEX gexLista 
         Height          =   2385
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   4207
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmMantPartidasArancelarias.frx":0000
         Column(2)       =   "frmMantPartidasArancelarias.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmMantPartidasArancelarias.frx":016C
         FormatStyle(2)  =   "frmMantPartidasArancelarias.frx":02A4
         FormatStyle(3)  =   "frmMantPartidasArancelarias.frx":0354
         FormatStyle(4)  =   "frmMantPartidasArancelarias.frx":0408
         FormatStyle(5)  =   "frmMantPartidasArancelarias.frx":04E0
         FormatStyle(6)  =   "frmMantPartidasArancelarias.frx":0598
         ImageCount      =   0
         PrinterProperties=   "frmMantPartidasArancelarias.frx":0678
      End
   End
   Begin VB.Frame fraDetalle 
      Caption         =   "Detalle :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3405
      Left            =   30
      TabIndex        =   0
      Top             =   2775
      Width           =   8415
      Begin VB.TextBox txtComposicion_Ingles 
         Height          =   300
         Left            =   1140
         MaxLength       =   200
         TabIndex        =   11
         Top             =   2895
         Width           =   7110
      End
      Begin VB.TextBox txtComposicion 
         Height          =   300
         Left            =   1140
         MaxLength       =   200
         TabIndex        =   10
         Top             =   2445
         Width           =   7110
      End
      Begin VB.TextBox txtDes_AdicionalPartida 
         Height          =   855
         Left            =   1140
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   1425
         Width           =   7125
      End
      Begin VB.TextBox txtsec_Partida_Arancelaria 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1155
         MaxLength       =   3
         TabIndex        =   1
         Top             =   210
         Width           =   630
      End
      Begin VB.TextBox txtDes_Partida 
         Height          =   675
         Left            =   1140
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   645
         Width           =   7125
      End
      Begin VB.Label Label5 
         Caption         =   "% Composic. Inglés :"
         Height          =   420
         Left            =   180
         TabIndex        =   13
         Top             =   2835
         Width           =   990
      End
      Begin VB.Label Label3 
         Caption         =   "% Composic. :"
         Height          =   285
         Left            =   165
         TabIndex        =   12
         Top             =   2490
         Width           =   990
      End
      Begin VB.Label Label4 
         Caption         =   "Descrip. Adicional :"
         Height          =   495
         Left            =   165
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Secuencia :"
         Height          =   195
         Left            =   165
         TabIndex        =   5
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción :"
         Height          =   195
         Left            =   165
         TabIndex        =   2
         Top             =   660
         Width           =   930
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2475
      TabIndex        =   8
      Top             =   6360
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantPartidasArancelarias.frx":0850
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmMantPartidasArancelarias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim sTipo As String
Public sNum_Partida_Arancelaria As String

Public Sub HABILITA_DATOS()
    Me.txtSec_Partida_Arancelaria.Enabled = False
    
    Me.txtDes_Partida.Enabled = True
    Me.txtDes_AdicionalPartida.Enabled = True
    Me.txtComposicion.Enabled = True
    Me.txtComposicion_Ingles.Enabled = True
    Me.txtDes_Partida.Locked = False
    Me.txtDes_AdicionalPartida.Locked = False
    Me.txtComposicion.Locked = False
    Me.txtComposicion_Ingles.Locked = False
End Sub

Public Sub INHABILITA_DATOS()
    Me.txtSec_Partida_Arancelaria.Enabled = False
    Me.txtDes_Partida.Locked = True
    Me.txtDes_AdicionalPartida.Locked = True
    Me.txtComposicion.Locked = True
    Me.txtComposicion_Ingles.Locked = True
End Sub

Public Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
                
        If Trim(Me.txtDes_Partida.Text) = "" Or (Me.txtDes_AdicionalPartida.Text) = "" Then
            VALIDA_DATOS = False
            Call MsgBox("La Descripción no puede estar vacia. Sirvase verficar", vbInformation, "Mensaje")
            Exit Function
        End If
        
End Function

Public Sub LIMPIAR_DATOS()
    Me.txtSec_Partida_Arancelaria.Text = ""
    Me.txtDes_Partida.Text = ""
    Me.txtDes_AdicionalPartida.Text = ""
    Me.txtComposicion.Text = ""
    Me.txtComposicion_Ingles.Text = ""
End Sub

Public Sub CARGA_DATOS()
    If gexLista.RowCount > 0 Then
        Me.txtSec_Partida_Arancelaria.Text = gexLista.Value(gexLista.Columns("SEC_PARTIDA_ARANCELARIA").Index)
        Me.txtDes_Partida.Text = gexLista.Value(gexLista.Columns("DES_PARTIDA").Index)
        Me.txtDes_AdicionalPartida = gexLista.Value(gexLista.Columns("DES_ADICIONAL_PARTIDA").Index)
        Me.txtComposicion.Text = gexLista.Value(gexLista.Columns("COMPOSICION").Index)
        Me.txtComposicion_Ingles = gexLista.Value(gexLista.Columns("COMPOSICION_INGLES").Index)
    End If
End Sub

Public Sub CARGA_GRID()
    
    strSQL = "EXEC UP_SEL_TG_Partida_Arancelaria_Detalle  '" & sNum_Partida_Arancelaria & "'"

    Set gexLista.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
    

    Call CONFIGURAR_GRID
    
    If gexLista.RowCount > 0 Then
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
    Else
        HabilitaMant Me.MantFunc1, "ADICIONAR"
    End If

End Sub

Private Sub SALVAR_DATOS()
   Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Con.ConnectionString = cCONNECT
    Con.Open
    
    Con.BeginTrans
        
        
        
        strSQL = "EXEC UP_MAN_Tg_Partida_Arancelaria_Detalle '" & sTipo & "', '" & _
        sNum_Partida_Arancelaria & "','" & _
        Trim(Me.txtSec_Partida_Arancelaria.Text) & "','" & _
        Trim(Me.txtDes_Partida.Text) & "','" & _
        Trim(Me.txtDes_AdicionalPartida.Text) & "','" & _
        Trim(Me.txtComposicion.Text) & "','" & Trim(Me.txtComposicion_Ingles) & "'"
        
    Con.Execute strSQL

    Con.CommitTrans
    Dim amensaje As New clsMessages
    
    LIMPIAR_DATOS

    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub

Private Sub ELIMINAR_DATOS()
    Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
    Con.ConnectionString = cCONNECT
    Con.Open
    
    Con.BeginTrans
        
        strSQL = "EXEC UP_MAN_Tg_Partida_Arancelaria_Detalle '" & "D" & "', '" & _
        sNum_Partida_Arancelaria & "','" & _
        Trim(Me.txtSec_Partida_Arancelaria.Text) & "','" & _
        Trim(Me.txtDes_Partida.Text) & "','" & _
        Trim(Me.txtDes_AdicionalPartida.Text) & "','" & _
        Trim(Me.txtComposicion.Text) & "','" & Trim(Me.txtComposicion_Ingles) & "'"

    
    Con.Execute strSQL
   
    Con.CommitTrans
    Dim amensaje As New clsMessages

    Exit Sub
    
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Eliminar_Datos"
End Sub

Private Sub Form_Load()
    Call INHABILITA_DATOS
End Sub

Private Sub gexLista_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    Call CARGA_DATOS
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim Eliminar As Integer
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            LIMPIAR_DATOS
            HABILITA_DATOS
            Me.txtDes_Partida.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Case "MODIFICAR"
            sTipo = "U"
            HABILITA_DATOS
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Case "ELIMINAR"
            sTipo = "D"
            If VALIDA_DATOS Then
                Eliminar = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?.", vbInformation + vbYesNo, "Mensaje")
                If Eliminar = vbYes Then
                    Call ELIMINAR_DATOS
                    Call LIMPIAR_DATOS
                    Call Me.CARGA_GRID
                End If
            End If
        Case "GRABAR"
            If VALIDA_DATOS Then
                SALVAR_DATOS
                sTipo = ""
                Call Me.CARGA_GRID
                Call INHABILITA_DATOS
            End If
        Case "DESHACER"
            INHABILITA_DATOS
            sTipo = ""
            LIMPIAR_DATOS
            Call Me.CARGA_GRID
        Case "SALIR"
            sTipo = ""
            Unload Me
    End Select
End Sub

Public Sub CONFIGURAR_GRID()
End Sub

Private Sub txtComposicion_Ingles_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        MantFunc1.SetFocus
    End If
End Sub

Private Sub txtComposicion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtComposicion_Ingles.SetFocus
    End If
End Sub

Private Sub txtDes_Partida_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtDes_AdicionalPartida.SetFocus
    End If
End Sub

Private Sub txtDes_AdicionalPartida_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtComposicion.SetFocus
    End If
End Sub

