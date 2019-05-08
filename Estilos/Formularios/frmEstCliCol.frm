VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmEstCliCol 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estilo Cliente Color"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   7995
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEstPro 
      Caption         =   "Estilos de Presentación Propios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   12
      Top             =   3840
      Width           =   7935
      Begin VB.Frame fraAdicEstProPre 
         Caption         =   "Adicionar Estilo de Presentación Propio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   13
         Top             =   1860
         Visible         =   0   'False
         Width           =   7695
         Begin VB.CommandButton cmdNuePresent 
            Caption         =   "&Nuevo"
            Height          =   315
            Left            =   4440
            TabIndex        =   26
            Top             =   600
            Width           =   855
         End
         Begin VB.ComboBox cboCod_EstPro 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   240
            Width           =   3735
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   495
            Left            =   6480
            TabIndex        =   18
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtCod_Present 
            Height          =   285
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   17
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton cmdBusca_Present 
            Caption         =   "..."
            Height          =   285
            Left            =   2520
            TabIndex        =   16
            Top             =   600
            Width           =   285
         End
         Begin VB.TextBox txtDes_Present 
            Height          =   285
            Left            =   2760
            MaxLength       =   50
            TabIndex        =   15
            Top             =   600
            Width           =   1695
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "A&ceptar"
            Height          =   495
            Left            =   5390
            TabIndex        =   14
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cod. Presentación"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Width           =   1305
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Estilo Propio"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   870
         End
      End
      Begin FunctionsButtons.FunctButt FBEstPro 
         Height          =   510
         Left            =   240
         TabIndex        =   20
         Top             =   2040
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   900
         Custom          =   $"frmEstCliCol.frx":0000
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin MSDataGridLib.DataGrid DG_EstPro 
         Height          =   1575
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   2778
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "Cod_EstPro"
            Caption         =   "Cod. Est. Propio"
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
            DataField       =   "Des_EstPro"
            Caption         =   "Descripción Est. Propio"
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
         BeginProperty Column02 
            DataField       =   "Cod_Present"
            Caption         =   "Cod. Propio"
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
         BeginProperty Column03 
            DataField       =   "Des_Present"
            Caption         =   "Descripción Propia"
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
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2310.236
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2580.095
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   495
         Left            =   6240
         TabIndex        =   22
         Top             =   2040
         Width           =   1095
      End
   End
   Begin VB.Frame fraEstCli 
      Caption         =   "Color de Estilos del Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      TabIndex        =   9
      Top             =   1080
      Width           =   7935
      Begin FunctionsButtons.FunctButt FBEstCli 
         Height          =   510
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   900
         Custom          =   $"frmEstCliCol.frx":00C6
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin MSDataGridLib.DataGrid DG_EstCli 
         Height          =   1815
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   3201
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         Enabled         =   0   'False
         ColumnHeaders   =   -1  'True
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
               ColumnWidth     =   1649.764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   5520.189
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraBuscar 
      Caption         =   "Buscar por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.TextBox txtDes_EstCli 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   4
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtCod_EstCli 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtAbr_Cliente 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtCod_TemCli 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtDes_Cliente 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtNom_TemCli 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Estilo"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   280
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Temporada"
         Height          =   195
         Left            =   3360
         TabIndex        =   7
         Top             =   285
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmEstCliCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public varCod_Cliente, varCod_TemCli, varCod_EstCli As String
Dim Rs_EstCliCol As ADODB.Recordset
Dim Rs_EstCliColPre As ADODB.Recordset
Public Codigo, Descripcion As String
Dim tEstado As String
Dim StrSQL As String

Public Sub CARGA_ESTCLICOL()
    Set Rs_EstCliCol = New ADODB.Recordset
    Rs_EstCliCol.ActiveConnection = cCONNECT
    Rs_EstCliCol.CursorType = adOpenStatic
    Rs_EstCliCol.CursorLocation = adUseClient
    Rs_EstCliCol.LockType = adLockReadOnly
        
    'Esta cadena es la que nos devolvera los items segun la seleccion establecida
    StrSQL = "EXEC UP_SEL_ESTCLICOL '" & varCod_Cliente & "','" & varCod_TemCli & "','" & varCod_EstCli & "'"
    
    Rs_EstCliCol.Open StrSQL
    Set DG_EstCli.DataSource = Rs_EstCliCol
    If Rs_EstCliCol.RecordCount > 0 Then
        Call CARGA_ESTCLICOLPRE
        HabilitaMant Me.FBEstCli, "ADICIONAR/MODIFICAR/ELIMINAR/CAMBIO"
        DG_EstCli.Enabled = True
    Else
        HabilitaMant Me.FBEstCli, "ADICIONAR"
        HabilitaMant Me.FBEstPro, ""
        Set Rs_EstCliColPre = Nothing
        Set DG_EstPro.DataSource = Nothing
        DG_EstPro.Refresh
    End If
End Sub

Public Sub CARGA_ESTCLICOLPRE()
    Set Rs_EstCliColPre = New ADODB.Recordset
    Rs_EstCliColPre.ActiveConnection = cCONNECT
    Rs_EstCliColPre.CursorType = adOpenStatic
    Rs_EstCliColPre.CursorLocation = adUseClient
    Rs_EstCliColPre.LockType = adLockReadOnly
    
    'Esta cadena es la que nos devolvera los items segun la seleccion establecida
    StrSQL = "EXEC UP_SEL_ESTILOSPREPROPIOS '" & varCod_Cliente & "','" & varCod_TemCli & "','" & varCod_EstCli & "','" & Rs_EstCliCol("Cod_ColCli").Value & "'"
   
    Rs_EstCliColPre.Open StrSQL
    Set DG_EstPro.DataSource = Rs_EstCliColPre
    If Rs_EstCliColPre.RecordCount > 0 Then
        'Call CARGA_ESTPRO
        HabilitaMant Me.FBEstPro, "AGREGAR/SUPRIMIR/MODIFICAR"
    Else
        'Set rs_EstPro = Nothing
        HabilitaMant Me.FBEstPro, "AGREGAR"
    End If
End Sub

Private Function VALIDA_ANADE_ESTCLICOLPRE() As Boolean
    VALIDA_ANADE_ESTCLICOLPRE = True
    If cboCod_EstPro.Text = "" Then
        Call MsgBox("Sirvase seleccionar un Estilo Propio", vbInformation)
        cboCod_EstPro.SetFocus
        VALIDA_ANADE_ESTCLICOLPRE = False
        Exit Function
    End If
    
    StrSQL = "SELECT COUNT(Cod_Present) From TG_ESTCLICOLPRE Where Cod_Cliente='" & _
    varCod_Cliente & "' And Cod_TemCli='" & _
    varCod_TemCli & "' And Cod_EstCli='" & _
    varCod_EstCli & "' And Cod_ColCli='" & _
    Trim(Rs_EstCliCol("Cod_ColCli").Value) & "' And Cod_EstPro='" & _
    Trim(Mid(cboCod_EstPro.Text, 1, 5)) & "'"
    
    If DevuelveCampo(StrSQL, cCONNECT) > 0 Then
        Call MsgBox("El registro ya existe. Sirvase verificar", vbExclamation)
        VALIDA_ANADE_ESTCLICOLPRE = False
        Exit Function
    End If
   
End Function

Private Function VALIDA_MODIFICA_ESTCLICOLPRE() As Boolean
    
    VALIDA_MODIFICA_ESTCLICOLPRE = True
    If txtCod_Present.Text = "" Then
        VALIDA_MODIFICA_ESTCLICOLPRE = False
        MsgBox "Sirvase seleccionar la presentación", vbExclamation
        Exit Function
    End If
    
End Function

Private Sub ANADE_ESTCLICOLPRE()
    On Error GoTo Salvar_DatosErr

    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Con.ConnectionString = cCONNECT
    Con.Open
    
    Con.BeginTrans
        
        'Esta cadena es la que nos devolvera los items segun la seleccion establecida
        StrSQL = "EXEC UP_MAN_ESTCLICOLPRE '" & _
        "I" & "','" & _
        varCod_Cliente & "','" & _
        varCod_TemCli & "','" & _
        varCod_EstCli & "','" & _
        Trim(Rs_EstCliCol("Cod_ColCli").Value) & "','" & _
        Trim(Mid(cboCod_EstPro.Text, 1, 5)) & "','" & _
        Trim(txtCod_Present.Text) & "'"
  
    Con.Execute StrSQL
        
    Con.CommitTrans
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    Call MsgBox("Ocurrio un error al añadir el Color de Estilo Propio", vbCritical)
End Sub

Private Sub MODIFICA_ESTCLICOLPRE()
    On Error GoTo Salvar_DatosErr

    Dim Con As New ADODB.Connection
    Dim sErr As String
    On Error GoTo Salvar_DatosErr
    Con.ConnectionString = cCONNECT
    Con.Open
    
    Con.BeginTrans
        
        'Esta cadena es la que nos devolvera los items segun la seleccion establecida
        StrSQL = "EXEC UP_MAN_ESTCLICOLPRE '" & _
        "U" & "','" & _
        varCod_Cliente & "','" & _
        varCod_TemCli & "','" & _
        varCod_EstCli & "','" & _
        Trim(Rs_EstCliCol("Cod_ColCli").Value) & "','" & _
        Trim(Mid(cboCod_EstPro.Text, 1, 5)) & "','" & _
        Trim(txtCod_Present.Text) & "'"
  
    Con.Execute StrSQL
        
    Con.CommitTrans
    Exit Sub
Salvar_DatosErr:
    sErr = Err.Description
    Con.RollbackTrans
    Set Con = Nothing
    Call MsgBox("Ocurrio un error al añadir el Color de Estilo Propio: " & Err.Description, vbCritical)
End Sub

Private Sub ELIMINA_ESTCLICOLPRE()
    On Error GoTo Salvar_DatosErr

    Dim Con As New ADODB.Connection
    Dim sErr As String
    On Error GoTo Salvar_DatosErr
    Con.ConnectionString = cCONNECT
    Con.Open
    
    Con.BeginTrans
        
        'Esta cadena es la que nos devolvera los items segun la seleccion establecida
        StrSQL = "EXEC UP_MAN_ESTCLICOLPRE '" & _
        "D" & "','" & _
        varCod_Cliente & "','" & _
        varCod_TemCli & "','" & _
        varCod_EstCli & "','" & _
        Trim(Rs_EstCliCol("Cod_ColCli").Value) & "','" & _
        Trim(Rs_EstCliColPre("Cod_EstPro").Value) & "','" & _
        Trim(Rs_EstCliColPre("Cod_Present").Value) & "'"
        
  
    Con.Execute StrSQL
        
    Con.CommitTrans
    Exit Sub
Salvar_DatosErr:
    sErr = Err.Description
    Con.RollbackTrans
    Set Con = Nothing
    Call MsgBox("Ocurrio un error al eliminar el Color de Estilo Propio: " & Err.Description, vbCritical)
End Sub

Public Sub ELIMINA_ESTCLICOL()
    Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
    
    'Strsql = "SELECT Cod_ColCli FROM tg_estclicolpre WHERE Cod_Cliente='" & varCod_Cliente & "' AND Cod_TemCli='" & varCod_TemCli & "' AND Cod_EstCli='" & varCod_EstCli & "' AND Cod_ColCli='" & Rs_EstCliCol("Cod_ColCli").Value & "'"

    'If DevuelveCampo(Strsql, cCONNECT) <> "" Then
    '    MsgBox ("No se puede eliminar el Registro por que posee registros relacionados")
    '    Exit Sub
    'End If
    
    Con.ConnectionString = cCONNECT
    Con.Open
    Con.BeginTrans
           
        'Esta es la sentencia que realiza la eliminacion del Registro
        StrSQL = "EXEC UP_MAN_ESTCLICOL " & _
        "D" & ",'" & _
        varCod_Cliente & "','" & _
        varCod_TemCli & "','" & _
        varCod_EstCli & "','" & _
        Rs_EstCliCol("Cod_ColCli").Value & "',''"
        
        Con.Execute StrSQL
    
    Con.CommitTrans
    
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_DELETE
    Informa "", amensaje

Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Eliminar_Datos"
End Sub

Public Sub LLENA_COMBOS()
    'Strsql = "SELECT Cod_EstPro, Des_EstPro from ES_EstPro WHERE Cod_EstPro IN(SELECT DISTINCT Cod_EstPro FROM TG_ESTCLIEST WHERE Cod_Cliente='" & _
    'varCod_Cliente & "' AND Cod_TemCli='" & _
    'varCod_TemCli & "' AND Cod_EstCli='" & _
    'varCod_EstCli & "')"
    StrSQL = "EXEC UP_SEL_ESTCLIPRE '" & varCod_Cliente & "','" & varCod_TemCli & "','" & varCod_EstCli & "','" & Rs_EstCliCol("Cod_ColCli").Value & "'"
    Call LlenaCombo(cboCod_EstPro, StrSQL, cCONNECT)
End Sub

Private Sub BUSCA_PRESENTACION()
    Dim oTipo As New frmBusqGeneral
    Dim Rs As New ADODB.Recordset
    Set oTipo.oParent = Me
       
    StrSQL = "SELECT Cod_Present as Código, Des_Present as Descripción FROM ES_ESTPROPRE WHERE Cod_EstPro='" & Mid(cboCod_EstPro.Text, 1, 5) & "'"
    oTipo.sQuery = StrSQL
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtCod_Present.Text = Trim(Codigo)
        txtDes_Present.Text = Trim(Descripcion)
    End If
    Set oTipo = Nothing
    Set Rs = Nothing
    
    'FBBuscar.SetFocus
End Sub

Private Sub cmdAceptar_Click()
    If tEstado = "A" Then
        If VALIDA_ANADE_ESTCLICOLPRE Then
            Call ANADE_ESTCLICOLPRE
        End If
    Else
        If tEstado = "" Then Exit Sub
        If VALIDA_MODIFICA_ESTCLICOLPRE Then
            Call MODIFICA_ESTCLICOLPRE
        End If
    End If
                
   Call CARGA_ESTCLICOLPRE
   fraAdicEstProPre.Visible = False

End Sub

Private Sub cmdBusca_Present_Click()
    Call BUSCA_PRESENTACION
End Sub

Private Sub cmdCancelar_Click()
    fraAdicEstProPre.Visible = False
End Sub

Private Sub cmdNuePresent_Click()
    Load FrmPresentaciones
    FrmPresentaciones.Codigo = Trim(Mid(cboCod_EstPro.Text, 1, 5))
    FrmPresentaciones.Accion "V", "", "", False
    FrmPresentaciones.Show 1
End Sub



Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub DG_EstCli_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Rs_EstCliCol.RecordCount > 0 Then
        Call CARGA_ESTCLICOLPRE
    End If
End Sub

Private Sub FBEstCli_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Eliminar As Integer
     Select Case ActionName
        Case "ADICIONAR"
            'If VALIDA_CLIENTETEMPORADA Then
                'Strsql = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
                
'                strSQL = "SELECT count(*) FROM TG_ESTCLICOL WHERE Cod_Cliente = '" & Me.varCod_Cliente & "' AND Cod_TemCli = '" & Me.varCod_TemCli & "' AND Cod_EstCli = '" & Me.varCod_EstCli & "'"
'                If DevuelveCampo(strSQL, cCONNECT) > 0 Then
'
'                    Load frmManEstCliCol
'                    frmManEstCliCol.varCod_Cliente = varCod_Cliente
'                    frmManEstCliCol.varCod_TemCli = varCod_TemCli
'                    frmManEstCliCol.varCod_EstCli = varCod_EstCli
'
'                    frmManEstCliCol.txtAbr_Cliente.Text = txtAbr_Cliente.Text
'                    frmManEstCliCol.txtNom_TemCli.Text = txtNom_TemCli.Text
'                    frmManEstCliCol.txtCod_EstCli.Text = txtCod_EstCli.Text
'                    frmManEstCliCol.CARGA_LISTA
'                    frmManEstCliCol.Carga_Datos
'                    frmManEstCliCol.Show 1
'
'                Else
                    Load frmMantEstCliColAll
                    frmMantEstCliColAll.varCod_Cliente = Me.varCod_Cliente
                    frmMantEstCliColAll.varCod_EstCli = Me.varCod_EstCli
                    frmMantEstCliColAll.varCod_TemCli = Me.varCod_TemCli
                    frmMantEstCliColAll.CARGA_COLORES
                    frmMantEstCliColAll.Show 1
                    
'                End If
                
                Call CARGA_ESTCLICOL
                    
            'End If
         Case "MODIFICAR"
            'If VALIDA_CLIENTETEMPORADA Then
                'Strsql = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
                Load frmManEstCliCol
                frmManEstCliCol.varCod_Cliente = varCod_Cliente
                frmManEstCliCol.varCod_TemCli = varCod_TemCli
                frmManEstCliCol.varCod_EstCli = varCod_EstCli
                
                frmManEstCliCol.txtAbr_Cliente.Text = txtAbr_Cliente.Text
                frmManEstCliCol.txtNom_TemCli.Text = txtNom_TemCli.Text
                frmManEstCliCol.txtCod_EstCli.Text = txtCod_EstCli.Text
                frmManEstCliCol.CARGA_LISTA
                frmManEstCliCol.Carga_Datos
                frmManEstCliCol.Show 1
                CARGA_ESTCLICOL
            'End If
        Case "ELIMINAR"
            If Not Rs_EstCliCol.EOF And Not Rs_EstCliCol.BOF Then
                Eliminar = MsgBox("Desea usted eliminar el registro seleccionado?", vbExclamation + vbYesNo)
                If Eliminar = vbYes Then
                    Call ELIMINA_ESTCLICOL
                    Call CARGA_ESTCLICOL
                Else
                    Exit Sub
                End If
            End If
'        Case "CAMBIO"
'                Dim Cod_ColCliTemp As String
'                With frmCambioColor
'                    .vCod_Cliente = Me.varCod_Cliente
'                    .vCod_EstCli = Me.varCod_EstCli
'                    .vCod_TemCli = Me.varCod_TemCli
'                    .vCod_ColCli = Rs_EstCliCol.Fields("Cod_ColCli").Value
'
'                    .txtAbr_Cliente = Me.txtAbr_Cliente.Text
'                    .txtDes_Cliente = Me.txtDes_Cliente.Text
'                    .txtCod_TemCli = Me.txtCod_TemCli.Text
'                    .txtNom_TemCli = Me.txtNom_TemCli.Text
'                    .txtCod_EstCli = Me.txtCod_EstCli
'                    .txtDes_EstCli = Me.txtDes_EstCli
'                    .txtCod_Color = Rs_EstCliCol.Fields("Cod_ColCli").Value
'
'                    Cod_ColCliTemp = Rs_EstCliCol.Fields("Cod_ColCli").Value
'                    .Show 1
'                End With
'                Call CARGA_ESTCLICOL
'                Call BuscaCampo(Rs_EstCliCol, "Cod_ColCli", Cod_ColCliTemp)
    End Select
End Sub

Private Sub FBEstPro_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Eliminar As Integer
    Select Case ActionName
         Case "AGREGAR"
                tEstado = "A"
                fraAdicEstProPre.Visible = True
                Call LLENA_COMBOS
                cboCod_EstPro.Enabled = True
                'cboCod_EstPro.ListIndex = -1
                If cboCod_EstPro.ListCount > 0 Then
                    cboCod_EstPro.ListIndex = 0
                End If
                txtCod_Present.Text = ""
                txtDes_Present.Text = ""
                txtCod_Present.SetFocus
                
         Case "SUPRIMIR"
                Eliminar = MsgBox("Desea usted eliminar el registro seleccionado?", vbExclamation + vbYesNo)
                If Eliminar = vbYes Then
                    Call ELIMINA_ESTCLICOLPRE
                    Call CARGA_ESTCLICOLPRE
                Else
                    Exit Sub
                End If
                tEstado = ""
        Case "MODIFICAR"
                Dim Sql As String
                tEstado = "M"
                fraAdicEstProPre.Visible = True
                Sql = "select Cod_estPro,Des_EstPro from Es_EstPro where cod_estpro='" & Rs_EstCliColPre.Fields("cod_estpro").Value & "'"
                'aqui llenar el combo con la busqueda dada,m mejor hacer esto en el rowchange
                Call LlenaCombo(cboCod_EstPro, Sql, cCONNECT)
                Call BuscaCombo(Rs_EstCliColPre.Fields("Cod_EstPro").Value, 2, cboCod_EstPro)
                If cboCod_EstPro.ListCount > 0 Then
                    cboCod_EstPro.ListIndex = 0
                End If
                cboCod_EstPro.Enabled = False
                txtCod_Present.Text = Rs_EstCliColPre.Fields("Cod_Present").Value
                txtDes_Present.Text = Rs_EstCliColPre.Fields("Des_Present").Value
                txtCod_Present.SetFocus
        
    End Select
End Sub

Private Sub Form_Load()
    Call FormSet(Me)
    Call FormateaGrid(DG_EstCli)
    Call FormateaGrid(DG_EstPro)
    Me.FBEstCli.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    Me.FBEstPro.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub

Private Sub txtCod_Present_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtAbr_Cliente.Text) = "" Then
            Call BUSCA_PRESENTACION
        Else
            StrSQL = "SELECT Des_Present FROM ES_ESTPROPRE WHERE Cod_EstPro='" & Mid(cboCod_EstPro.Text, 1, 5) & "' AND Cod_Present='" & Trim(txtCod_Present.Text) & "'"
            txtDes_Present.Text = DevuelveCampo(StrSQL, cCONNECT)
            cmdAceptar.SetFocus
        End If
    End If
End Sub

Private Sub txtDes_Present_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtDes_Present.Text) < 3 Then
            Call MsgBox("La descripción debe ser mayor a 2 caracteres. Sirvase verificar", vbInformation)
            txtDes_Present.SetFocus
            Exit Sub
        Else
            StrSQL = "SELECT Cod_Present FROM ES_ESTPROPRE WHERE Cod_EstPro='" & Mid(cboCod_EstPro.Text, 1, 5) & "' AND Des_Present LIKE '" & Trim(txtDes_Present.Text) & "%'"
            txtCod_Present.Text = DevuelveCampo(StrSQL, cCONNECT)
            cmdAceptar.SetFocus
        End If
    End If
End Sub
