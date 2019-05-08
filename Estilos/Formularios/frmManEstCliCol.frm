VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManEstCliCol 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Estilo Cliente"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   4320
      Width           =   2055
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1470
         Picture         =   "frmManEstCliCol.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Ultimo"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   0
         Picture         =   "frmManEstCliCol.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Primero"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   975
         Picture         =   "frmManEstCliCol.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Siguiente"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "frmManEstCliCol.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Anterior"
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
      Height          =   1395
      Left            =   50
      TabIndex        =   14
      Top             =   2850
      Width           =   5940
      Begin VB.CommandButton cmdBuscaColTem 
         Caption         =   "..."
         Height          =   300
         Left            =   1950
         TabIndex        =   5
         Top             =   960
         Width           =   315
      End
      Begin VB.TextBox txtCod_ColCli 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtCod_EstCli 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   3
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txtNom_ColCli 
         Height          =   285
         Left            =   2265
         MaxLength       =   20
         TabIndex        =   6
         Top             =   960
         Width           =   3480
      End
      Begin VB.TextBox txtAbr_Cliente 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtNom_TemCli 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label label2 
         AutoSize        =   -1  'True
         Caption         =   "Estilo"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   650
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1005
         Width           =   495
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Temporada"
         Height          =   195
         Left            =   2400
         TabIndex        =   15
         Top             =   285
         Width           =   810
      End
   End
   Begin VB.Frame FraListado 
      Caption         =   "Listado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2805
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   5940
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2475
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   4366
         _Version        =   393216
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Cod_ColCli"
            Caption         =   "Código"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Nom_ColCli"
            Caption         =   "Descripción"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               ColumnWidth     =   1844.787
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3344.882
            EndProperty
         EndProperty
      End
   End
   Begin Mantenimientos.MantFunc MFEstCli 
      Height          =   540
      Left            =   2280
      TabIndex        =   7
      Top             =   4320
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmManEstCliCol.frx":05C8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmManEstCliCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public varCod_Cliente, varCod_TemCli, varCod_EstCli As String
Dim Rs_Lista As ADODB.Recordset
Public Codigo, Descripcion As String
Dim StrSQL As String
Dim sTipo As String

Private Sub cmdBuscaColTem_Click()
    Call BUSCA_COLORCLITEMP(3)
End Sub

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

    Dim StrSQL As String
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.ActiveConnection = cCONNECT
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    'Esta cadena es la que nos devolvera los items segun la seleccion establecida
    StrSQL = "EXEC UP_SEL_ESTCLICOL '" & varCod_Cliente & "','" & varCod_TemCli & "','" & varCod_EstCli & "'"
    Rs_Lista.Open StrSQL
    Set DGridLista.DataSource = Rs_Lista

    If Rs_Lista.RecordCount > 0 Then
        HabilitaMant Me.MFEstCli, "ADICIONAR/MODIFICAR/ELIMINAR"
    Else
        HabilitaMant Me.MFEstCli, "ADICIONAR"
        Call LIMPIA_DATOS
    End If
    
End Sub

Public Sub Carga_Datos()
    If Rs_Lista.RecordCount > 0 Then
              
        txtCod_ColCli.Text = Trim(Rs_Lista("Cod_ColCli").Value)
        txtNom_ColCli.Text = Trim(Rs_Lista("Nom_ColCli").Value)
        
    End If
End Sub
Public Sub HABILITA_DATOS()
    txtCod_ColCli.Enabled = True
    txtNom_ColCli.Enabled = True
    cmdBuscaColTem.Enabled = True
End Sub
Public Sub DESABILITA_DATOS()
    txtCod_ColCli.Enabled = False
    txtNom_ColCli.Enabled = False
    cmdBuscaColTem.Enabled = False
End Sub

Public Sub LIMPIA_DATOS()
    txtCod_ColCli.Text = ""
    txtNom_ColCli.Text = ""
End Sub

Public Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
    If sTipo = "I" Then
        If Trim(txtCod_ColCli.Text) = "" Then
            Call MsgBox("Sirvase ingresar un codigo de Color", vbExclamation)
            VALIDA_DATOS = False
            Exit Function
        End If
        StrSQL = "SELECT * FROM TG_ESTCLICOL WHERE Cod_Cliente='" & varCod_Cliente & "' AND Cod_TemCli='" & varCod_TemCli & "' AND Cod_EstCli='" & varCod_EstCli & "' AND Cod_ColCli='" & Trim(txtCod_ColCli.Text) & "'"
        If DevuelveCampo(StrSQL, cCONNECT) <> "" Then
            Call MsgBox("El código ingresado ya existe. Sirvase verificar", vbCritical)
            txtCod_ColCli.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
        
    End If
End Function

Public Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Con.ConnectionString = cCONNECT
    Con.Open
    
    Con.BeginTrans
       
        'Esta es la sentencia que realizara el salvado de datos
        StrSQL = "EXEC UP_MAN_ESTCLICOL " & _
        sTipo & ",'" & _
        varCod_Cliente & "','" & _
        varCod_TemCli & "','" & _
        varCod_EstCli & "','" & _
        Trim(txtCod_ColCli.Text) & "','" & _
        Trim(txtNom_ColCli.Text) & "'"
        
        Con.Execute StrSQL
        
    Con.CommitTrans
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
    Informa "", amensaje
    Call DESABILITA_DATOS
    Call LIMPIA_DATOS

    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub

Public Sub ELIMINAR_DATOS()
    Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
    
    StrSQL = "SELECT Cod_ColCli FROM tg_estclicolpre WHERE Cod_Cliente='" & varCod_Cliente & "' AND Cod_TemCli='" & varCod_TemCli & "' AND Cod_EstCli='" & varCod_EstCli & "' AND Cod_ColCli='" & Trim(txtCod_ColCli.Text) & "'"

    If DevuelveCampo(StrSQL, cCONNECT) <> "" Then
        MsgBox ("No se puede eliminar el Registro por que posee registros relacionados")
        Exit Sub
    End If
    
    Con.ConnectionString = cCONNECT
    Con.Open
    Con.BeginTrans
           
        'Esta es la sentencia que realiza la eliminacion del Registro
        StrSQL = "EXEC UP_MAN_ESTCLICOL " & _
        sTipo & ",'" & _
        varCod_Cliente & "','" & _
        varCod_TemCli & "','" & _
        varCod_EstCli & "','" & _
        Trim(txtCod_ColCli.Text) & "','" & _
        Trim(txtNom_ColCli.Text) & "'"
        
        Con.Execute StrSQL
    
    Con.CommitTrans
    
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_DELETE
    Informa "", amensaje

    LIMPIA_DATOS
Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
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
    Call FormSet(Me)
    Call DESABILITA_DATOS
    Call FormateaGrid(DGridLista)
    MFEstCli.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
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
            Eliminar = MsgBox("Desea usted eliminar el registro seleccionado?", vbExclamation + vbYesNo)
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
                    Call BuscaCampo(Rs_Lista, "Cod_ColCli", varCod_ColCli)
                    sTipo = ""
                End If
            End If
        Case "DESHACER"
            DESABILITA_DATOS
            sTipo = ""
            LIMPIA_DATOS
            Call Carga_Datos
            HabilitaMant Me.MFEstCli, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridLista.Enabled = True
        Case "SALIR"
            sTipo = ""
            Unload Me
    End Select
End Sub

Sub BUSCA_COLORCLITEMP(tipo As Integer)
    Dim varProvTip_Presentacion As String

    Select Case tipo
        Case 1:
                
                StrSQL = "SELECT Nom_ColCli FROM TG_COLCLITEM WHERE Cod_Cliente = '" & varCod_Cliente & "' AND Cod_TemCli = '" & varCod_TemCli & "' AND Cod_ColCli='" & Trim(txtCod_ColCli.Text) & "'"
                txtNom_ColCli.Text = Trim(DevuelveCampo(StrSQL, cCONNECT))
                'Strsql = "SELECT " & CodigoTabla & " FROM " & NombreTabla & " WHERE " & DesTabla & " = '" & txtDescripcion.Text & "'"
                'txtCodigo.Text = Trim(DevuelveCampo(Strsql, cCONNECT))
        Case 2, 3:
        
                Dim oTipo As New frmBusqGeneral
                Dim Rs As New ADODB.Recordset
                Set oTipo.oParent = Me
                
                If tipo = 2 Then
                    oTipo.sQuery = "SELECT Cod_ColCli AS CODIGO, Nom_ColCli AS DESCRIPCION FROM TG_COLCLITEM WHERE Cod_Cliente = '" & varCod_Cliente & "' AND Cod_TemCli = '" & varCod_TemCli & "' AND Nom_ColCli LIKE '" & Trim(txtNom_ColCli.Text) & "%'"
                Else
                    oTipo.sQuery = "SELECT Cod_ColCli AS CODIGO, Nom_ColCli AS DESCRIPCION FROM TG_COLCLITEM WHERE Cod_Cliente = '" & varCod_Cliente & "' AND Cod_TemCli = '" & varCod_TemCli & "'"
                End If
                
                oTipo.Cargar_Datos
                oTipo.Show 1
                If Codigo <> "" Then
                    txtCod_ColCli.Text = Trim(Codigo)
                    txtNom_ColCli.Text = Trim(Descripcion)
                    Codigo = ""
                    Descripcion = ""
                    'cboCod_ProTex.SetFocus
                End If
                Set oTipo = Nothing
                Set Rs = Nothing
    End Select
    
    'Call CARGA_COMBOS
    
End Sub

Private Sub txtCod_ColCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call BUSCA_COLORCLITEMP(1)
    End If
End Sub

Private Sub txtNom_ColCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call BUSCA_COLORCLITEMP(2)
    End If
End Sub
