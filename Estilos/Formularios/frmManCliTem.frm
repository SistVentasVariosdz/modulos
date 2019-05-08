VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form frmManCliTem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Colores de la Temporada"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   5715
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   525
      Left            =   3000
      TabIndex        =   18
      Top             =   4995
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   926
      Custom          =   $"frmManCliTem.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1190
      ControlHeigth   =   495
      ControlSeparator=   110
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   5025
      Width           =   2055
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "frmManCliTem.frx":00B1
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Anterior"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   975
         Picture         =   "frmManCliTem.frx":0223
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Siguiente"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   0
         Picture         =   "frmManCliTem.frx":0395
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Primero"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1470
         Picture         =   "frmManCliTem.frx":0507
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Ultimo"
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.Frame fraDetalles 
      Caption         =   "Detalles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   0
      TabIndex        =   6
      Top             =   3060
      Width           =   5670
      Begin VB.TextBox txtCod_Estampado 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1320
         Width           =   1470
      End
      Begin VB.TextBox txtDes_Estampado 
         Height          =   285
         Left            =   2640
         MaxLength       =   20
         TabIndex        =   19
         Top             =   1305
         Width           =   2625
      End
      Begin VB.TextBox txtNom_TemCli 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3210
         TabIndex        =   13
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtAbr_Cliente 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtNom_ColCli 
         Height          =   285
         Left            =   1080
         MaxLength       =   150
         TabIndex        =   1
         Top             =   960
         Width           =   4185
      End
      Begin VB.TextBox txtCod_ColCli 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   0
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estampado"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1365
         Width           =   795
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Temporada"
         Height          =   195
         Left            =   2190
         TabIndex        =   17
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   290
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1005
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   645
         Width           =   495
      End
   End
   Begin VB.Frame FraLista 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3030
      Left            =   0
      TabIndex        =   4
      Top             =   30
      Width           =   5670
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2685
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5430
         _ExtentX        =   9578
         _ExtentY        =   4736
         _Version        =   393216
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "Cod_ColCli"
            Caption         =   "Código"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Nom_ColCli"
            Caption         =   "Nombre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Cod_Estampado"
            Caption         =   "Estampado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Des_Estampado"
            Caption         =   "Descrip. Estampado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            SizeMode        =   1
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   1544.882
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3360.189
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
         EndProperty
      End
   End
   Begin Mantenimientos.MantFunc MFEstCli 
      Height          =   540
      Left            =   930
      TabIndex        =   3
      Top             =   5625
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmManCliTem.frx":0679
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmManCliTem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public Codigo, Descripcion As String
Dim Opcion As Integer
Dim sTipo As String
Dim strSQL As String
Dim Rs_Lista As ADODB.Recordset
Public varCod_Cliente, varCod_TemCli As String

Private Sub cmdFirst_Click()
    If Not Rs_Lista.BOF Then
        Rs_Lista.MoveFirst
    End If
End Sub

Private Sub cmdLast_Click()
    If Not Rs_Lista.EOF Then
        Rs_Lista.MoveLast
    End If
End Sub

Private Sub cmdNext_Click()
    If Not Rs_Lista.EOF Then
        Rs_Lista.MoveNext
        If Rs_Lista.EOF Then
            Rs_Lista.MoveLast
        End If
    End If
End Sub

Private Sub cmdPrevious_Click()
    If Not Rs_Lista.BOF Then
        Rs_Lista.MovePrevious
        If Rs_Lista.BOF Then
            Rs_Lista.MoveFirst
        End If
    End If
End Sub

Public Sub RECARGA_LISTA()
    Set Rs_Lista = Nothing
    Call CARGA_LISTA
End Sub

Public Sub CARGA_LISTA()
    Dim strSQL As String
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.ActiveConnection = cCONNECT
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    'Esta cadena es la que nos devolvera los items segun la seleccion establecida
    strSQL = "EXEC UP_SEL_COLCLITEM '" & varCod_Cliente & "','" & varCod_TemCli & "'"
    Rs_Lista.Open strSQL
    Set DGridLista.DataSource = Rs_Lista

    If Rs_Lista.RecordCount > 0 Then
        Call Carga_Datos
        HabilitaMant Me.MFEstCli, "ADICIONAR/MODIFICAR/ELIMINAR"
    Else
        HabilitaMant Me.MFEstCli, "ADICIONAR"
        Call LIMPIA_DATOS
    End If
End Sub

Public Sub Carga_Datos()
    If Rs_Lista.RecordCount > 0 Then
        txtCod_ColCli = Trim(Rs_Lista("Cod_ColCli").Value)
        txtNom_ColCli = Trim(Rs_Lista("Nom_ColCli").Value)
        txtCod_Estampado = Trim(Rs_Lista("Cod_Estampado").Value)
        txtDes_Estampado = Trim(Rs_Lista("Des_Estampado").Value)
    End If
End Sub
Public Sub HABILITA_DATOS()
    txtCod_ColCli.Enabled = True
    txtNom_ColCli.Enabled = True
    txtCod_Estampado.Enabled = True
    txtDes_Estampado.Enabled = True
End Sub
Public Sub DESABILITA_DATOS()
    txtCod_ColCli.Enabled = False
    txtNom_ColCli.Enabled = False
    txtCod_Estampado.Enabled = False
    txtDes_Estampado.Enabled = False
End Sub

Public Sub LIMPIA_DATOS()
    txtCod_ColCli.Text = ""
    txtNom_ColCli.Text = ""
    txtCod_Estampado.Text = ""
    txtDes_Estampado.Text = ""
End Sub

Public Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
    If sTipo = "I" Then
        If Trim(txtCod_ColCli.Text) = "" Then
            Call MsgBox("Sirvase ingresar un codigo de Color. Sirvase verificar", vbExclamation)
            txtCod_ColCli.Text = ""
            txtCod_ColCli.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
        If Trim(txtNom_ColCli.Text) = "" Then
            Call MsgBox("La descripción no puede estar vacia. Sirvase verificar", vbExclamation)
            txtNom_ColCli.Text = ""
            txtNom_ColCli.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
        strSQL = "SELECT * FROM TG_COLCLITEM WHERE Cod_Cliente='" & varCod_Cliente & "' AND Cod_TemCli='" & varCod_TemCli & "' AND Cod_ColCli='" & txtCod_ColCli.Text & "'"
        If DevuelveCampo(strSQL, cCONNECT) <> "" Then
            Call MsgBox("El código ingresado ya existe. Sirvase verificar", vbCritical)
            txtCod_ColCli.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
    End If
    If sTipo = "U" Then
        If Trim(txtNom_ColCli.Text) = "" Then
            Call MsgBox("La descripción no puede estar vacia. Sirvase verificar", vbExclamation)
            txtNom_ColCli.Text = ""
            txtNom_ColCli.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
    End If
   
    
End Function

Public Sub SALVAR_DATOS()
    Dim con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    con.ConnectionString = cCONNECT
    con.Open
    
    con.BeginTrans
       
        'Esta es la sentencia que realizara el salvado de datos
        strSQL = "UP_MAN_COLCLITEM " & _
        sTipo & ",'" & _
        varCod_Cliente & "','" & _
        varCod_TemCli & "','" & _
        Trim(txtCod_ColCli.Text) & "','" & _
        Trim(txtNom_ColCli.Text) & "','" & Trim(txtCod_Estampado.Text) & "'"
        
        con.Execute strSQL
        
    con.CommitTrans
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
    Informa "", amensaje
    Call DESABILITA_DATOS
    Call LIMPIA_DATOS

    Exit Sub
Salvar_DatosErr:
    con.RollbackTrans
    Set con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub

Public Sub ELIMINAR_DATOS()
    Dim con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
    
'    Strsql = "SELECT Cod_EstCli FROM tg_estcliest WHERE Cod_Cliente='" & varCod_Cliente & "' AND Cod_TemCli='" & varCod_TemCli & "' AND Cod_EstCli='" & txtCod_EstCli.Text & "'"
'
'    If DevuelveCampo(Strsql, cCONNECT) <> "" Then
'        MsgBox ("No se puede eliminar el Registro por que posee registros relacionados")
'        Exit Sub
'    End If
    
    con.ConnectionString = cCONNECT
    con.Open
    con.BeginTrans
           
        'Esta es la sentencia que realiza la eliminacion del Registro
        strSQL = "UP_MAN_COLCLITEM " & _
        sTipo & ",'" & _
        varCod_Cliente & "','" & _
        varCod_TemCli & "','" & _
        Trim(txtCod_ColCli.Text) & "','" & _
        Trim(txtNom_ColCli.Text) & "',''"
        
        con.Execute strSQL
    
    con.CommitTrans
    
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_DELETE
    Informa "", amensaje

    LIMPIA_DATOS
Exit Sub
Eliminar_DatosErr:
    con.RollbackTrans
    Set con = Nothing
    ErrorHandler Err, "Eliminar_Datos"
End Sub

Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Rs_Lista.State <> 1 Then
        Exit Sub
    End If
    If Not Rs_Lista.BOF And Not Rs_Lista.EOF Then
        Call Carga_Datos
    End If
End Sub

Private Sub Form_Load()
On Error GoTo hand
    Call FormSet(Me)
    Call DESABILITA_DATOS
    Call FormateaGrid(DGridLista)
    'MFEstCli.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
        
Exit Sub
hand:
ErrorHandler Err, "Form_Load()"
End Sub


Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "CAMBIO"
                With frmCambioColor
                    .vCod_Cliente = Me.varCod_Cliente
                    .vcod_TemCli = Me.varCod_TemCli
                    .vCod_ColCli = Rs_Lista("Cod_ColCli").Value
                    .Show 1
                End With
                CARGA_LISTA
        Case "OTRA"
            Load FrmCopiarColoresTemporada
            FrmCopiarColoresTemporada.varCod_Cliente = Me.varCod_Cliente
            FrmCopiarColoresTemporada.varCod_TemCli_origen = Me.varCod_TemCli
            FrmCopiarColoresTemporada.CARGA_TEMPORADA
            FrmCopiarColoresTemporada.Show 1
            Set FrmCopiarColoresTemporada = Nothing
            CARGA_LISTA
    End Select
End Sub

Private Sub MFEstCli_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Eliminar As Integer
    Dim varCod_ColCli As String
    Select Case ActionName
        Case "ADICIONAR"
            varCod_ColCli = Trim(txtCod_ColCli.Text)
            sTipo = "I"
            LIMPIA_DATOS
            HABILITA_DATOS
            HabilitaMant Me.MFEstCli, "GRABAR/DESHACER"
            DGridLista.Enabled = False
            txtCod_ColCli.SetFocus
        Case "MODIFICAR"
            sTipo = "U"
            HABILITA_DATOS
            txtCod_ColCli.Enabled = False
            txtNom_ColCli.SetFocus
            HabilitaMant Me.MFEstCli, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "ELIMINAR"
            sTipo = "D"
            Eliminar = MsgBox("Usted desea eliminar el registro seleccionado", vbExclamation + vbYesNo)
            If Eliminar = vbYes Then
                Call ELIMINAR_DATOS
                Call RECARGA_LISTA
            Else
                Exit Sub
            End If
        Case "GRABAR"
            If VALIDA_DATOS Then
                varCod_ColCli = Trim(txtCod_ColCli.Text)
                Call SALVAR_DATOS
                Call RECARGA_LISTA
                HabilitaMant Me.MFEstCli, "ADICIONAR/MODIFICAR/ELIMINAR"
                DGridLista.Enabled = True
                Call Carga_Datos
                If sTipo = "I" Then
                    Call MFEstCli_ActionClick(0, 1, "ADICIONAR")
                Else
                    sTipo = ""
                    Call BuscaCampo(Rs_Lista, "Cod_ColCli", varCod_ColCli)
                End If
            End If
        Case "DESHACER"
            DESABILITA_DATOS
            sTipo = ""
            LIMPIA_DATOS
            Call Carga_Datos
            Call BuscaCampo(Rs_Lista, "Cod_ColCli", varCod_ColCli)
            HabilitaMant Me.MFEstCli, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridLista.Enabled = True
        Case "SALIR"
            sTipo = ""
            Unload Me
    End Select
End Sub

Private Sub txtCod_ColCli_LostFocus()
    txtNom_ColCli.Text = txtCod_ColCli.Text
End Sub

Private Sub txtCod_Estampado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If RTrim(txtCod_Estampado.Text) = "" Then
            Buscar_Estampados 2
        Else
            Buscar_Estampados 1
        End If
    End If
End Sub

Private Sub Buscar_Estampados(iOpcion As Integer)
    Dim oTipo As New frmBusqGeneral
    Dim Rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    
    If iOpcion = 1 Then
        oTipo.sQuery = "SELECT Cod_Estampado as Código, Des_Estampado as Descripción FROM TG_ESTCLITEM_ESTAMPADOS WHERE COD_CLIENTE = '" & varCod_Cliente & "' AND COD_TEMCLI = '" & varCod_TemCli & "' AND COD_ESTAMPADO = '" & txtCod_Estampado.Text & "' ORDER BY Cod_Estampado"
    End If
    If iOpcion = 2 Then
        oTipo.sQuery = "SELECT Cod_Estampado as Código, Des_Estampado as Descripción FROM TG_ESTCLITEM_ESTAMPADOS WHERE COD_CLIENTE = '" & varCod_Cliente & "' AND COD_TEMCLI = '" & varCod_TemCli & "' ORDER BY Des_Estampado"
    End If
    
    If iOpcion = 3 Then
        oTipo.sQuery = "SELECT Cod_Estampado as Código, Des_Estampado as Descripción FROM TG_ESTCLITEM_ESTAMPADOS WHERE COD_CLIENTE = '" & varCod_Cliente & "' AND COD_TEMCLI = '" & varCod_TemCli & "' AND DES_ESTAMPADO LIKE '%" & txtDes_Estampado & "%'ORDER BY Des_Estampado"
    End If
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtCod_Estampado.Text = Codigo
        txtDes_Estampado.Text = Descripcion
        MFEstCli.SetFocus
    End If
    Set oTipo = Nothing
    Set Rs = Nothing
End Sub
