VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Begin VB.Form frmMantLugaresEntrega 
   Caption         =   "Lugares de Entrega"
   ClientHeight    =   6828
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   4908
   LinkTopic       =   "Form1"
   ScaleHeight     =   6828
   ScaleWidth      =   4908
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraLista 
      Caption         =   "Lista:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   15
      TabIndex        =   8
      Top             =   -15
      Width           =   4905
      Begin GridEX20.GridEX gexLista 
         Height          =   2385
         Left            =   120
         TabIndex        =   9
         Top             =   255
         Width           =   4650
         _ExtentX        =   8192
         _ExtentY        =   4212
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   288
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmMantLugaresEntrega.frx":0000
         Column(2)       =   "frmMantLugaresEntrega.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmMantLugaresEntrega.frx":016C
         FormatStyle(2)  =   "frmMantLugaresEntrega.frx":02A4
         FormatStyle(3)  =   "frmMantLugaresEntrega.frx":0354
         FormatStyle(4)  =   "frmMantLugaresEntrega.frx":0408
         FormatStyle(5)  =   "frmMantLugaresEntrega.frx":04E0
         FormatStyle(6)  =   "frmMantLugaresEntrega.frx":0598
         ImageCount      =   0
         PrinterProperties=   "frmMantLugaresEntrega.frx":0678
      End
   End
   Begin VB.Frame fraDetalle 
      Caption         =   "Detalle :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3360
      Left            =   15
      TabIndex        =   1
      Top             =   2760
      Width           =   4905
      Begin VB.TextBox TxtLinea7 
         Height          =   285
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   17
         Top             =   2568
         Width           =   3525
      End
      Begin VB.TextBox TxtLinea6 
         Height          =   285
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   16
         Top             =   2240
         Width           =   3525
      End
      Begin VB.TextBox txtDesPais 
         Height          =   285
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   14
         Top             =   2960
         Width           =   2880
      End
      Begin VB.TextBox txtCodPais 
         Height          =   285
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   13
         Top             =   2960
         Width           =   645
      End
      Begin VB.TextBox TxtLinea5 
         Height          =   285
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   12
         Top             =   1920
         Width           =   3525
      End
      Begin VB.TextBox TxtLinea4 
         Height          =   285
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   11
         Top             =   1635
         Width           =   3525
      End
      Begin VB.TextBox txtLinea3 
         Height          =   285
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   6
         Top             =   1305
         Width           =   3525
      End
      Begin VB.TextBox txtLinea2 
         Height          =   285
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   5
         Top             =   975
         Width           =   3540
      End
      Begin VB.TextBox txtsecuencia 
         Height          =   285
         Left            =   1155
         MaxLength       =   2
         TabIndex        =   2
         Top             =   210
         Width           =   630
      End
      Begin VB.TextBox txtLinea1 
         Height          =   285
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   4
         Top             =   645
         Width           =   3555
      End
      Begin VB.Label Label4 
         Caption         =   "Pais :"
         Height          =   252
         Left            =   240
         TabIndex        =   15
         Top             =   3000
         Width           =   612
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "(Número)"
         Height          =   195
         Left            =   1920
         TabIndex        =   10
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Secuencia :"
         Height          =   195
         Left            =   165
         TabIndex        =   7
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Dirección :"
         Height          =   195
         Left            =   165
         TabIndex        =   3
         Top             =   615
         Width           =   765
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   600
      TabIndex        =   0
      Top             =   6240
      Width           =   3576
      _ExtentX        =   6287
      _ExtentY        =   953
      Custom          =   $"frmMantLugaresEntrega.frx":0850
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmMantLugaresEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim sTipo As String
Public sCod_Cliente As String

Public Codigo As String
Public Descripcion As String

Public Sub HABILITA_DATOS()
    If sTipo <> "U" Then
        Me.txtsecuencia.Enabled = True
    End If
    Me.txtLinea1.Enabled = True
    Me.txtLinea2.Enabled = True
    Me.txtLinea3.Enabled = True
    Me.TxtLinea4.Enabled = True
    Me.TxtLinea5.Enabled = True
    Me.TxtLinea5.Enabled = True
    Me.TxtLinea6.Enabled = True
    Me.TxtLinea7.Enabled = True
    Me.txtCodPais.Enabled = True
    Me.txtDesPais.Enabled = True
    
End Sub

Public Sub INHABILITA_DATOS()
    Me.txtsecuencia.Enabled = False
    Me.txtLinea1.Enabled = False
    Me.txtLinea2.Enabled = False
    Me.txtLinea3.Enabled = False
    Me.TxtLinea4.Enabled = False
    Me.TxtLinea5.Enabled = False
    Me.TxtLinea6.Enabled = False
    Me.TxtLinea7.Enabled = False
    Me.txtCodPais.Enabled = False
    Me.txtDesPais.Enabled = False
    
End Sub

Public Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
    If sTipo <> "D" Then
        If Trim(Me.txtsecuencia.Text) = "" Then
            VALIDA_DATOS = False
            Call MsgBox("El código de Lugar de Entrega de Cliente no puede estar vacio. Sirvase verificar", vbInformation, "Mensaje")
            Exit Function
        End If
        
                
        strSQL = "SELECT count(*) FROM Tg_Cliente_LugEnt WHERE Cod_Cliente = '" & sCod_Cliente & "' AND Secuencia = " & Trim(Me.txtsecuencia.Text) & ""
        If sTipo = "I" And Val(DevuelveCampo(strSQL, cCONNECT)) > 0 Then
            VALIDA_DATOS = False
            Call MsgBox("El Lugar de Entrega de Cliente Ingresado ya se encuentra registrado. Sirvase verificar", vbInformation, "Mensaje")
            Exit Function
        End If
        
        If Trim(Me.txtLinea1.Text) = "" Then
            VALIDA_DATOS = False
            Call MsgBox("La Descripción del Lugar de Entrega de Cliente no puede estar vacia. Sirvase verficar", vbInformation, "Mensaje")
            Exit Function
        End If
        
    Else
        strSQL = "SELECT * FROM CF_PAKINGLIST WHERE Cod_Cliente ='" & sCod_Cliente & "' AND  Num_SecLugarEntrega = " & Trim(Me.txtsecuencia.Text) & ""
        If Val(DevuelveCampo(strSQL, cCONNECT)) > 0 Then
            VALIDA_DATOS = False
            Call MsgBox("El Lugar de Entrega de Cliente no puede ser eliminado por que posee registros relacionados. Sirvase verficar", vbInformation, "Mensaje")
            Exit Function
        End If
    End If
End Function

Public Sub LIMPIAR_DATOS()
    Me.txtsecuencia.Text = ""
    Me.txtLinea1.Text = ""
    Me.txtLinea2.Text = ""
    Me.txtLinea3.Text = ""
    Me.TxtLinea4.Text = ""
    Me.TxtLinea5.Text = ""
    Me.TxtLinea6.Text = ""
    Me.TxtLinea7.Text = ""
    Me.txtCodPais = ""
    Me.txtDesPais = ""
    
End Sub

Public Sub CARGA_DATOS()
    If gexLista.RowCount > 0 Then
        Me.txtsecuencia.Text = gexLista.Value(gexLista.Columns("secuencia").Index)
        Me.txtLinea1.Text = gexLista.Value(gexLista.Columns("Linea1").Index)
        Me.txtLinea2.Text = gexLista.Value(gexLista.Columns("Linea2").Index)
        Me.txtLinea3.Text = gexLista.Value(gexLista.Columns("Linea3").Index)
        Me.TxtLinea4.Text = IIf(IsNull(gexLista.Value(gexLista.Columns("Linea4").Index)), "", gexLista.Value(gexLista.Columns("Linea4").Index))
        Me.TxtLinea5.Text = gexLista.Value(gexLista.Columns("Linea5").Index)
        
        Me.txtCodPais = gexLista.Value(gexLista.Columns("cod_pais").Index)
        Me.txtDesPais = gexLista.Value(gexLista.Columns("descripcion").Index)
        
    End If
End Sub

Public Sub CARGA_GRID()
    
    strSQL = "EXEC UP_SEL_Tg_Cliente_LugEnt  '" & sCod_Cliente & "'"

    Set gexLista.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
    SetGeneralGridEX gexLista, 0, 1
    

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
        
        strSQL = "EXEC UP_MAN_Tg_Cliente_LugEnt '" & sCod_Cliente & "', '" & _
        sTipo & "'," & _
        Trim(Me.txtsecuencia.Text) & ",'" & _
        Trim(Me.txtLinea1.Text) & "','" & _
        Trim(Me.txtLinea2.Text) & "','" & _
        Trim(Me.txtLinea3.Text) & "','" & _
        Trim(Me.TxtLinea4.Text) & "','" & _
        Trim(Me.TxtLinea5.Text) & "','" & _
        Trim(Me.TxtLinea6.Text) & "','" & _
        Trim(Me.TxtLinea7.Text) & "','" & _
        Me.txtCodPais & "'"
        
    Con.Execute strSQL

    Con.CommitTrans
    Dim amensaje As New clsMessages
    amensaje.Codigo = MESSAGECODE.kMESSAGE_INF_PROCESS_SATISFACTO
    Mensaje MESSAGECODE.kMESSAGE_INF_PROCESS_SATISFACTO
    
    LIMPIAR_DATOS

    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler err, "Salvar_Datos"
End Sub

Private Sub ELIMINAR_DATOS()
    Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
    Con.ConnectionString = cCONNECT
    Con.Open
    
    Con.BeginTrans
        
        strSQL = "EXEC UP_MAN_Tg_Cliente_LugEnt  '" & sCod_Cliente & "', '" & _
        sTipo & "'," & _
        Trim(Me.txtsecuencia.Text) & ",'" & _
        Trim(Me.txtLinea1.Text) & "','" & _
        Trim(Me.txtLinea2.Text) & "','" & _
        Trim(Me.txtLinea3.Text) & "','" & _
        Trim(Me.TxtLinea4.Text) & "','" & _
        Trim(Me.TxtLinea5.Text) & "','" & _
        Trim(Me.TxtLinea6.Text) & "','" & _
        Trim(Me.TxtLinea7.Text) & "','" & _
        Trim(Me.txtCodPais) & "'"

    
    Con.Execute strSQL
   
    Con.CommitTrans
    Dim amensaje As New clsMessages
    amensaje.Codigo = MESSAGECODE.kMESSAGE_INF_DATA_DELETE
    Mensaje MESSAGECODE.kMESSAGE_INF_DATA_DELETE

    Exit Sub
    
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler err, "Eliminar_Datos"
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
            Me.txtsecuencia.SetFocus
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
    gexLista.Columns("Secuencia").Caption = "Secuencia"
    gexLista.Columns("Secuencia").Width = "1000"
    gexLista.Columns("Linea1").Caption = "Direccion1"
    gexLista.Columns("Linea1").Width = "2100"
    gexLista.Columns("Linea2").Caption = "Direccion1"
    gexLista.Columns("Linea2").Width = "2100"
    gexLista.Columns("Linea3").Caption = "Direccion1"
    gexLista.Columns("Linea3").Width = "2100"
    gexLista.Columns("Linea4").Caption = "Direccion1"
    gexLista.Columns("Linea4").Width = "2100"
    gexLista.Columns("Linea5").Caption = "Direccion1"
    gexLista.Columns("Linea5").Width = "2100"
End Sub

Private Sub txtCodPais_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Me.txtCodPais.Text) = "" Then
            Call Me.BuscaPais(3)
        Else
            Call Me.BuscaPais(1)
        End If
    End If
End Sub

Private Sub txtDesPais_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      If Trim(txtDesPais.Text) = "" Then
        Call Me.BuscaPais(3)
      Else
         Call Me.BuscaPais(2)
      End If
    End If
End Sub

Public Sub BuscaPais(opcion As Integer)
Dim sField As String, iRows As Long
Dim rstAux As ADODB.Recordset
    
    strSQL = " SELECT RTRIM(Cod_Pais) as 'Codigo' , RTRIM(Descripcion) AS 'Descripción' FROM cn_paises WHERE "
    txtCodPais = Trim(txtCodPais)
    txtDesPais = Trim(txtDesPais)

    Select Case opcion
    Case 1: strSQL = strSQL & " Cod_Pais   like '%" & Trim(txtCodPais.Text) & "%'  "
    Case 2: strSQL = strSQL & " Descripcion  like '%" & Trim(txtDesPais.Text) & "%' "
    Case 3: strSQL = " SELECT RTRIM(Cod_Pais) as 'Codigo' , RTRIM(Descripcion) AS 'Descripción' FROM cn_paises  "
    End Select
    
    
    txtCodPais = ""
    txtDesPais = ""
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = strSQL
        .Cargar_Datos
        
        Codigo = ""
        Descripcion = ""
        
        iRows = .DGridLista.RowCount
        Set rstAux = .DGridLista.ADORecordset
        If .DGridLista.RowCount > 1 Then
            .Show vbModal
        ElseIf .DGridLista.RowCount = 1 Then
            Codigo = .DGridLista.Value(.DGridLista.Columns("CODIGO").Index)
            Descripcion = .DGridLista.Value(.DGridLista.Columns("DESCRIPCIóN").Index)
        End If
        
        If Codigo <> "" Then
            txtCodPais = RTrim(Codigo)
            txtDesPais = RTrim(Descripcion)
            Me.MantFunc1.SetFocus
        End If
        
            
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
End Sub


