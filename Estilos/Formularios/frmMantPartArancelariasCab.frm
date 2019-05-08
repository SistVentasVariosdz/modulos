VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMantPartArancelariasCab 
   Caption         =   "Partidas Arancelarias"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
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
      Height          =   1575
      Left            =   45
      TabIndex        =   2
      Top             =   2790
      Width           =   8415
      Begin VB.TextBox txtDes_Partida 
         Height          =   675
         Left            =   1140
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   645
         Width           =   7125
      End
      Begin VB.TextBox txtNum_Partida_Arancelaria 
         Height          =   285
         Left            =   1155
         MaxLength       =   15
         TabIndex        =   3
         Top             =   210
         Width           =   1755
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción :"
         Height          =   195
         Left            =   165
         TabIndex        =   6
         Top             =   660
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número :"
         Height          =   195
         Left            =   165
         TabIndex        =   5
         Top             =   285
         Width           =   645
      End
   End
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
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin GridEX20.GridEX gexLista 
         Height          =   2385
         Left            =   120
         TabIndex        =   1
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
         Column(1)       =   "frmMantPartArancelariasCab.frx":0000
         Column(2)       =   "frmMantPartArancelariasCab.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmMantPartArancelariasCab.frx":016C
         FormatStyle(2)  =   "frmMantPartArancelariasCab.frx":02A4
         FormatStyle(3)  =   "frmMantPartArancelariasCab.frx":0354
         FormatStyle(4)  =   "frmMantPartArancelariasCab.frx":0408
         FormatStyle(5)  =   "frmMantPartArancelariasCab.frx":04E0
         FormatStyle(6)  =   "frmMantPartArancelariasCab.frx":0598
         ImageCount      =   0
         PrinterProperties=   "frmMantPartArancelariasCab.frx":0678
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2475
      TabIndex        =   7
      Top             =   4440
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantPartArancelariasCab.frx":0850
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmMantPartArancelariasCab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim sTipo As String
Public sNum_Partida_Arancelaria As String

Public Sub HABILITA_DATOS()
    If sTipo = "I" Then
        Me.txtNum_Partida_Arancelaria.Enabled = True
    Else
        Me.txtNum_Partida_Arancelaria.Enabled = False
    End If

    
    Me.txtDes_Partida.Enabled = True
       
End Sub

Public Sub INHABILITA_DATOS()
    Me.txtNum_Partida_Arancelaria.Enabled = False
    Me.txtDes_Partida.Enabled = False
End Sub

Public Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
                
        If Trim(Me.txtNum_Partida_Arancelaria.Text) = "" Then
            VALIDA_DATOS = False
            Call MsgBox("Número de  Partida Arancelaria no puede estar vacia. Sirvase verficar", vbInformation, "Mensaje")
            Exit Function
        End If
                
        If Trim(Me.txtDes_Partida.Text) = "" Then
            VALIDA_DATOS = False
            Call MsgBox("La Descripción no puede estar vacia. Sirvase verficar", vbInformation, "Mensaje")
            Exit Function
        End If
        
End Function

Public Sub LIMPIAR_DATOS()
    Me.txtNum_Partida_Arancelaria.Text = ""
    Me.txtDes_Partida.Text = ""
End Sub

Public Sub CARGA_DATOS()
    If gexLista.RowCount > 0 Then
        Me.txtNum_Partida_Arancelaria.Text = gexLista.Value(gexLista.Columns("Num_Partida_ARANCELARIA").Index)
        Me.txtDes_Partida.Text = gexLista.Value(gexLista.Columns("DES_PARTIDA_ARANCELARIA").Index)
    End If
End Sub

Public Sub CARGA_GRID()
    
    strSQL = "EXEC UP_SEL_TG_Partida_Arancelaria  "

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
        
        
        
        strSQL = "EXEC UP_MAN_Tg_Partida_Arancelaria '" & sTipo & "','" & _
        Trim(Me.txtNum_Partida_Arancelaria.Text) & "','" & _
        Trim(Me.txtDes_Partida.Text) & "'"
        
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
        
        strSQL = "EXEC UP_MAN_Tg_Partida_Arancelaria '" & "D" & "','" & _
        Trim(Me.txtNum_Partida_Arancelaria.Text) & "','" & _
        Trim(Me.txtDes_Partida.Text) & "'"

    
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
    gexLista.Columns("NUM_PARTIDA_ARANCELARIA").Width = 1500
    gexLista.Columns("DES_PARTIDA_ARANCELARIA").Width = 6000

End Sub

Private Sub txtComposicion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        MantFunc1.SetFocus
    End If
End Sub
