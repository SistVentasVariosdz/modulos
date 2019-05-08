VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmGrupoPro 
   Caption         =   "Grupos"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   7995
   StartUpPosition =   2  'CenterScreen
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
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.TextBox txtAbrCliente 
         Height          =   285
         Left            =   1290
         TabIndex        =   2
         Top             =   270
         Width           =   915
      End
      Begin VB.TextBox txtNomCliente 
         Height          =   285
         Left            =   2505
         TabIndex        =   4
         Top             =   270
         Width           =   2565
      End
      Begin VB.CommandButton cmdBusCliente 
         Caption         =   "..."
         Height          =   285
         Left            =   2205
         TabIndex        =   3
         Tag             =   "..."
         Top             =   270
         Width           =   300
      End
      Begin FunctionsButtons.FunctButt FBBuscar 
         Height          =   495
         Left            =   5985
         TabIndex        =   5
         Top             =   180
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~True~True~&Buscar~0~0~1~~0~False~False~&Buscar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   435
         TabIndex        =   1
         Top             =   315
         Width           =   480
      End
   End
   Begin VB.Frame fraEstCli 
      Caption         =   "Grupos de Producción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   7935
      Begin VB.Frame fraGrupoPro 
         Caption         =   "Adicionar Grupo de Producción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   135
         TabIndex        =   21
         Top             =   2100
         Visible         =   0   'False
         Width           =   7695
         Begin VB.TextBox txtNom_Cliente 
            Height          =   285
            Left            =   2400
            TabIndex        =   28
            Top             =   240
            Width           =   2535
         End
         Begin VB.CommandButton cmdBuscaCliente 
            Caption         =   "..."
            Height          =   285
            Left            =   2160
            TabIndex        =   27
            Top             =   240
            Width           =   285
         End
         Begin VB.CommandButton cmdAceptarGP 
            Caption         =   "A&ceptar"
            Height          =   495
            Left            =   5280
            TabIndex        =   26
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancelarGP 
            Caption         =   "&Cancelar"
            Height          =   495
            Left            =   6375
            TabIndex        =   25
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtDes_GrupoPro 
            Height          =   285
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   24
            Top             =   600
            Width           =   2775
         End
         Begin VB.TextBox txtCod_GrupoPro 
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   23
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtAbr_Cliente 
            Height          =   285
            Left            =   1200
            TabIndex        =   22
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   240
            TabIndex        =   30
            Top             =   650
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Left            =   240
            TabIndex        =   29
            Top             =   290
            Width           =   480
         End
      End
      Begin MSDataGridLib.DataGrid DG_EstCli 
         Height          =   1815
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   3201
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Enabled         =   -1  'True
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "Cod_GrupoPro"
            Caption         =   "Código Grupo"
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
            DataField       =   "Des_GrupoPro"
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
         BeginProperty Column02 
            DataField       =   "Abr_Cliente"
            Caption         =   "Abr. Cliente"
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
            DataField       =   "Nom_Cliente"
            Caption         =   "Nombre Cliente"
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
               ColumnWidth     =   2340.284
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2640.189
            EndProperty
         EndProperty
      End
      Begin FunctionsButtons.FunctButt FBGrupoPro 
         Height          =   510
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   900
         Custom          =   $"frmGrupoPro.frx":0000
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1050
         ControlHeigth   =   490
         ControlSeparator=   50
      End
   End
   Begin VB.Frame fraEstPro 
      Caption         =   "Ordenes de Producción"
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
      TabIndex        =   9
      Top             =   4080
      Width           =   7935
      Begin VB.Frame fraOrdPro 
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
         Left            =   135
         TabIndex        =   10
         Top             =   1860
         Visible         =   0   'False
         Width           =   7695
         Begin VB.TextBox txtDes_EstPro 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2880
            TabIndex        =   31
            Top             =   600
            Width           =   2415
         End
         Begin VB.ComboBox cboCod_Fabrica 
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   240
            Width           =   3735
         End
         Begin VB.CommandButton cmdCancelarOP 
            Caption         =   "&Cancelar"
            Height          =   495
            Left            =   6480
            TabIndex        =   14
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtCod_OrdPro 
            Height          =   285
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   13
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton cmdBusca_OrdPro 
            Caption         =   "..."
            Height          =   285
            Left            =   2640
            TabIndex        =   12
            Top             =   600
            Width           =   285
         End
         Begin VB.CommandButton cmdAceptarOP 
            Caption         =   "A&ceptar"
            Height          =   495
            Left            =   5390
            TabIndex        =   11
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ord. Producción :"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   1245
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fabrica :"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   615
         End
      End
      Begin MSDataGridLib.DataGrid DG_EstPro 
         Height          =   1575
         Left            =   120
         TabIndex        =   17
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "cod_ordpro"
            Caption         =   "O/P"
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
            DataField       =   "cod_estpro"
            Caption         =   "Cod. Est Propio"
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
            DataField       =   "Des_EstPro"
            Caption         =   "Desc. Estilo Propio"
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
            DataField       =   "fec_despachoact"
            Caption         =   "Fecha Despacho"
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
         BeginProperty Column04 
            DataField       =   "cod_estcli"
            Caption         =   "Estilo Cliente"
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
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1409.953
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2399.811
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1530.142
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1200.189
            EndProperty
         EndProperty
      End
      Begin FunctionsButtons.FunctButt FBOrdPro 
         Height          =   510
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   "0~0~AGREGAR~True~True~&Agregar~0~0~1~~0~False~False~&Agregar~~1~0~SUPRIMIR~True~True~&Suprimir~0~0~2~~0~False~False~&Suprimir~"
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   495
         Left            =   6270
         TabIndex        =   18
         Top             =   1920
         Width           =   1095
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   7560
      Top             =   3750
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmGrupoPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs_GrupoPro As ADODB.Recordset
Dim Rs_OrdPro As ADODB.Recordset
'Dim varCod_Cliente As String
Public Codigo, Descripcion As String
Dim sTipo As String
Dim StrSQL As String

Dim varEfectuoBusqueda

Public varOpcionAvances As String

Public Sub CARGA_GRUPOPRO()
    Set Rs_GrupoPro = New ADODB.Recordset
    Rs_GrupoPro.ActiveConnection = cCONNECT
    Rs_GrupoPro.CursorType = adOpenStatic
    Rs_GrupoPro.CursorLocation = adUseClient
    Rs_GrupoPro.LockType = adLockReadOnly
        
    StrSQL = "SELECT COD_CLIENTE FROM TG_CLIENTE WHERE ABR_CLIENTE='" & txtAbrCliente.Text & "'"
        
    'Esta cadena es la que nos devolvera los grupos de produccion
    StrSQL = "EXEC UP_SEL_GRUPOPRO '" & DevuelveCampo(StrSQL, cCONNECT) & "'"

    Rs_GrupoPro.Open StrSQL
    Set DG_EstCli.DataSource = Rs_GrupoPro
    If Rs_GrupoPro.RecordCount > 0 Then
        Call CARGA_ORDPRO
        'HabilitaMant Me.FBGrupoPro, "ADICIONAR/MODIFICAR/ELIMINAR/IMPRIMIR/AVANCES/REVISION"
        DG_EstCli.Enabled = True
    Else
        'HabilitaMant Me.FBGrupoPro, "ADICIONAR"
        'HabilitaMant Me.FBEstPro, ""
        Set Rs_OrdPro = Nothing
        Set DG_EstPro.DataSource = Nothing
        DG_EstPro.Refresh
    End If
End Sub

Public Sub CARGA_ORDPRO()
    Set Rs_OrdPro = New ADODB.Recordset
    Rs_OrdPro.ActiveConnection = cCONNECT
    Rs_OrdPro.CursorType = adOpenStatic
    Rs_OrdPro.CursorLocation = adUseClient
    Rs_OrdPro.LockType = adLockReadOnly
    
    'Esta cadena es la que nos devolvera los items segun la seleccion establecida
    StrSQL = "EXEC UP_SEL_ORDGRUPOPRODUCCION '" & Rs_GrupoPro("Cod_GrupoPro").Value & "'"
   
    Rs_OrdPro.Open StrSQL
    Set DG_EstPro.DataSource = Rs_OrdPro
    If Rs_OrdPro.RecordCount > 0 Then
        HabilitaMant Me.FBOrdPro, "AGREGAR/SUPRIMIR/REQAVIOS"
    Else
        HabilitaMant Me.FBOrdPro, "AGREGAR"
    End If
End Sub

Private Function VALIDA_ANADE_GRUPOPRO() As Boolean
    VALIDA_ANADE_GRUPOPRO = True
    If txtAbr_Cliente.Text = "" Then
        Call MsgBox("Sirvase seleccionar un Cliente", vbInformation)
        txtAbr_Cliente.SetFocus
        VALIDA_ANADE_GRUPOPRO = False
        Exit Function
    End If
    If txtCod_GrupoPro.Text = "" Then
        Call MsgBox("El código de grupo no puede estar vacio. Sirvase verificar", vbInformation)
        txtCod_GrupoPro.SetFocus
        VALIDA_ANADE_GRUPOPRO = False
        Exit Function
    End If
    If txtDes_GrupoPro.Text = "" Then
        Call MsgBox("La descripción del grupo no puede estar vacia. Sirvase verificar", vbInformation)
        txtDes_GrupoPro.SetFocus
        VALIDA_ANADE_GRUPOPRO = False
        Exit Function
    End If
    StrSQL = "SELECT COUNT(COD_CLIENTE) FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
    If DevuelveCampo(StrSQL, cCONNECT) = 0 Then
        Call MsgBox("El cliente ingresado no se encuentra registrado. Sirvase verificar", vbInformation)
        txtAbr_Cliente.SetFocus
        VALIDA_ANADE_GRUPOPRO = False
        Exit Function
    End If
End Function

Public Function VALIDA_ANADE_ORDPRO() As Boolean
    VALIDA_ANADE_ORDPRO = True
    If cboCod_Fabrica.Text = "" Then
        Call MsgBox("Sirvase seleccionar una Fabrica", vbInformation)
        cboCod_Fabrica.SetFocus
        VALIDA_ANADE_ORDPRO = False
        Exit Function
    End If
    
    If Trim(txtCod_OrdPro.Text) = "" Then
        Call MsgBox("La orden de produccion no puede estar vacia. Sirvase verificar", vbInformation)
        txtCod_OrdPro.SetFocus
        VALIDA_ANADE_ORDPRO = False
        Exit Function
    End If
    
    StrSQL = "SELECT  COUNT(cod_ordpro) From Es_ordpro WHERE Cod_fabrica ='" & _
    Right(cboCod_Fabrica.Text, 3) & "' AND Cod_Cliente ='" & _
    Rs_GrupoPro("Cod_Cliente").Value & "' AND Cod_GrupoPro='' AND cod_ordpro='" & _
    Trim(txtCod_OrdPro.Text) & "'"
    
    If DevuelveCampo(StrSQL, cCONNECT) = 0 Then
        Call MsgBox("La orden de producción ya se encuentra ingresada. Sirvase verificar", vbInformation)
        txtCod_OrdPro.SetFocus
        VALIDA_ANADE_ORDPRO = False
        Exit Function
    End If
End Function

Public Sub ANADE_ORDPRO()
    On Error GoTo Salvar_DatosErr

    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Con.ConnectionString = cCONNECT
    Con.Open
    
    Con.BeginTrans
        
        'Esta cadena es la que nos devolvera los items segun la seleccion establecida
        StrSQL = "EXEC UP_UPD_ORDPRO '" & _
        sTipo & "','" & _
        Right(cboCod_Fabrica.Text, 3) & "','" & _
        Trim(txtCod_OrdPro.Text) & "','" & _
        Trim(Rs_GrupoPro("Cod_GrupoPro").Value) & "'"

    Con.Execute StrSQL
        
    Con.CommitTrans
    
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
    Informa "", amensaje
    
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    Call MsgBox("Ocurrio un error al añadir el Color de Estilo Propio", vbCritical)
End Sub

Private Sub ELIMINA_ORDPRO()
    On Error GoTo Salvar_DatosErr

    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Con.ConnectionString = cCONNECT
    Con.Open
    
    Con.BeginTrans
        
        'Esta cadena es la que nos devolvera los items segun la seleccion establecida
        StrSQL = "EXEC UP_UPD_ORDPRO '" & _
        "D" & "','" & _
        Rs_OrdPro("Cod_Fabrica").Value & "','" & _
        Rs_OrdPro("Cod_OrdPro").Value & "','" & _
        Rs_OrdPro("Cod_GrupoPro").Value & "'"
        
  
    Con.Execute StrSQL
        
    Con.CommitTrans
    
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_DELETE
    Informa "", amensaje
    
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    Call MsgBox("Ocurrio un error al eliminar el Color de Estilo Propio", vbCritical)
End Sub

Private Sub ANADE_GRUPOPRO()
    On Error GoTo Salvar_DatosErr

    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Con.ConnectionString = cCONNECT
    Con.Open
    
    Con.BeginTrans
        StrSQL = "SELECT COD_CLIENTE FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
        
        'Esta cadena es la que nos devolvera los items segun la seleccion establecida
        StrSQL = "EXEC UP_MAN_GRUPOPRO '" & _
        sTipo & "','" & _
        DevuelveCampo(StrSQL, cCONNECT) & "','" & _
        Trim(txtCod_GrupoPro.Text) & "','" & _
        Trim(txtDes_GrupoPro.Text) & "'"
  
    Con.Execute StrSQL
        
    Con.CommitTrans
    
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
    Informa "", amensaje
  
    
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    Call MsgBox("Ocurrio un error al añadir el Color de Estilo Propio", vbCritical)
End Sub

Public Sub ELIMINA_GRUPOPRO()
    Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
    
    StrSQL = "SELECT COUNT(cod_ordpro) FROM es_ordpro WHERE Cod_GrupoPro='" & Rs_GrupoPro("Cod_GrupoPro").Value & "'"
    If DevuelveCampo(StrSQL, cCONNECT) > 0 Then
        MsgBox ("No se puede eliminar el Registro por que posee registros relacionados")
        Exit Sub
    End If
    
    Con.ConnectionString = cCONNECT
    Con.Open
    Con.BeginTrans
        
        'Strsql = "SELECT COD_CLIENTE FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
        
        StrSQL = "EXEC UP_MAN_GRUPOPRO '" & _
        sTipo & "','" & _
        Rs_GrupoPro("Cod_Cliente").Value & "','" & _
        Rs_GrupoPro("Cod_GrupoPro").Value & "','" & _
        Rs_GrupoPro("Des_GrupoPro").Value & "'"
        
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


Public Sub BUSCA_ORDPRO()
    StrSQL = "SELECT"
End Sub

Public Sub LLENA_COMBOS()
   
    StrSQL = "SELECT LEFT(Abr_Fabrica + '     ',7)  + Nom_Fabrica + SPACE(100) + Cod_Fabrica FROM TG_FABRICA ORDER BY Abr_Fabrica"
    Call LlenaCombo(cboCod_Fabrica, StrSQL, cCONNECT)
End Sub



Private Sub cmdAceptarGP_Click()
    If VALIDA_ANADE_GRUPOPRO Then
        Call ANADE_GRUPOPRO
        Call CARGA_GRUPOPRO
        sTipo = ""
        fraGrupoPro.Visible = False
    End If
End Sub

Private Sub cmdAceptarOP_Click()
    If VALIDA_ANADE_ORDPRO Then
        Call ANADE_ORDPRO
        Call CARGA_ORDPRO
        sTipo = ""
        fraOrdPro.Visible = False
    End If
End Sub

Private Sub cmdBusca_OrdPro_Click()
    Dim oTipo As New frmBusqGeneralGrupo
    'Dim Busqueda As New frmBusqGeneral
    Dim Rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    
    oTipo.DGridLista.Columns.Add (2)
    
    StrSQL = "SELECT  cod_ordpro AS Código, Des_EstPro as Descripción, cod_purord as PO From Es_ordpro OP, Es_EstPro EP WHERE OP.Cod_EstPro = EP.Cod_EstPro AND Cod_fabrica ='" & _
    Right(cboCod_Fabrica.Text, 3) & "' AND Cod_Cliente ='" & _
    Rs_GrupoPro("Cod_Cliente").Value & "' AND Cod_GrupoPro=''"
    
    oTipo.sQuery = StrSQL
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtCod_OrdPro.Text = Codigo
        txtDes_EstPro.Text = Descripcion
        cmdAceptarOP.SetFocus
    End If
    Set oTipo = Nothing
    Set Rs = Nothing
End Sub

Private Sub cmdBuscaCliente_Click()
    Dim oTipo As New frmBusqGeneral
    Dim Rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.sQuery = "SELECT Abr_Cliente as Código, nom_cliente as Descripción FROM TG_Cliente ORDER BY Abr_Cliente"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtAbr_Cliente.Text = Codigo
        txtNom_Cliente.Text = Descripcion
        txtCod_GrupoPro.SetFocus
    End If
    Set oTipo = Nothing
    Set Rs = Nothing
End Sub

Private Sub cmdBusCliente_Click()
    Dim oTipo As New frmBusqGeneral
    Dim Rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.sQuery = "SELECT Abr_Cliente as Código, nom_cliente as Descripción FROM TG_Cliente ORDER BY Abr_Cliente"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtAbrCliente.Text = Codigo
        txtNomCliente.Text = Descripcion
        FBBuscar.SetFocus
    End If
    Set oTipo = Nothing
    Set Rs = Nothing
End Sub


Private Sub cmdCancelarGP_Click()
    fraGrupoPro.Visible = False
    sTipo = ""
End Sub

Private Sub cmdCancelarOP_Click()
    sTipo = ""
    fraOrdPro.Visible = False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub DG_EstCli_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Rs_GrupoPro.RecordCount > 0 Then
        Call CARGA_ORDPRO
    End If
End Sub



Private Sub FBBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Call CARGA_GRUPOPRO
End Sub

Private Sub FBGrupoPro_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Eliminar As Integer
    Select Case ActionName
         Case "ADICIONAR"
                fraGrupoPro.Visible = True
                txtAbr_Cliente.Text = Me.txtAbrCliente
                txtNom_Cliente.Text = Me.txtNomCliente
                txtCod_GrupoPro.Text = ""
                txtDes_GrupoPro.Text = ""
                'Habilitamos los campos
                txtAbr_Cliente.Enabled = True
                txtNom_Cliente.Enabled = True
                txtCod_GrupoPro.Enabled = True
                txtDes_GrupoPro.Enabled = True
                txtCod_GrupoPro.SetFocus
                sTipo = "I"
         Case "MODIFICAR"
         
                If Rs_GrupoPro.State = 0 Then
                    MsgBox "No se ha seleccionado ningun registro. Sirvase verificar", vbInformation, "Mensaje"
                    Exit Sub
                End If
         
                fraGrupoPro.Visible = True
                StrSQL = "SELECT Abr_Cliente FROM TG_CLIENTE WHERE COD_CLIENTE='" & Trim(Rs_GrupoPro("Cod_Cliente").Value) & "'"
                txtAbr_Cliente.Text = DevuelveCampo(StrSQL, cCONNECT)
                Call txtAbr_Cliente_KeyPress(13)
                txtCod_GrupoPro.Text = Rs_GrupoPro("Cod_GrupoPro").Value
                txtDes_GrupoPro.Text = Rs_GrupoPro("Des_GrupoPro").Value
                'Habilitamos los campos
                txtAbr_Cliente.Enabled = False
                txtNom_Cliente.Enabled = False
                txtCod_GrupoPro.Enabled = False
                txtDes_GrupoPro.Enabled = True
                txtDes_GrupoPro.SetFocus
                sTipo = "U"
         Case "ELIMINAR"
         
                If Rs_GrupoPro.State = 0 Then
                    MsgBox "No se ha seleccionado ningun registro. Sirvase verificar", vbInformation, "Mensaje"
                    Exit Sub
                End If
         
                If Rs_GrupoPro.RecordCount = 0 Then
                    MsgBox "No existen registros para acceder a esta opción.", vbInformation, "Mensaje"
                    Exit Sub
                End If
         
                sTipo = "D"
                Eliminar = MsgBox("Desea usted eliminar el registro seleccionado?", vbExclamation + vbYesNo)
                If Eliminar = vbYes Then
                    Call ELIMINA_GRUPOPRO
                    Call CARGA_GRUPOPRO
                Else
                    Exit Sub
                End If
        Case "IMPRIMIR"
        
                If Rs_GrupoPro.State = 0 Then
                    MsgBox "No se ha seleccionado ningun registro. Sirvase verificar", vbInformation, "Mensaje"
                    Exit Sub
                End If
                If Rs_GrupoPro.RecordCount = 0 Then
                    MsgBox "No existen registros para acceder a esta opción.", vbInformation, "Mensaje"
                    Exit Sub
                End If
        
            GeneraReportes
        Case "REVISION"
                If Rs_GrupoPro.State = 0 Then
                    MsgBox "No se ha seleccionado ningun registro. Sirvase verificar", vbInformation, "Mensaje"
                    Exit Sub
                End If
                If Rs_GrupoPro.RecordCount = 0 Then
                    MsgBox "No existen registros para acceder a esta opción.", vbInformation, "Mensaje"
                    Exit Sub
                End If
        
            Call REPORTE
        Case "AVANCES"
                If Rs_GrupoPro.State = 0 Then
                    MsgBox "No se ha seleccionado ningun registro. Sirvase verificar", vbInformation, "Mensaje"
                    Exit Sub
                End If
                If Rs_GrupoPro.RecordCount = 0 Then
                    MsgBox "No existen registros para acceder a esta opción.", vbInformation, "Mensaje"
                    Exit Sub
                End If
        
            Load FrmOpcionAvance
            Set FrmOpcionAvance.oParent = Me
            FrmOpcionAvance.varCod_GrupoPro = Rs_GrupoPro("Cod_GrupoPro").Value
            FrmOpcionAvance.Show 1
            
            If Me.varOpcionAvances <> "" Then
                Load frmVerAvance
                frmVerAvance.Caption = "Avance del Grupo : " & Rs_GrupoPro("Cod_GrupoPro").Value & " - " & Rs_GrupoPro("Des_GrupoPro").Value
                frmVerAvance.StrOpcAvance = varOpcionAvances
                frmVerAvance.CARGA_GRID
                frmVerAvance.Show 1
            End If
       Case "REQAVIOS"
            RepReqAvios

    End Select
End Sub

Sub GeneraReportes()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String
    
    Ruta = vRuta & "\OPGRUPO.xlt"

    Set oo = CreateObject("excel.application")
    oo.workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.run "Reporte", DG_EstCli.Columns("código grupo"), cCONNECT, vemp, DG_EstCli.Columns("nombre cliente")
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler Err, "GeneraReportes"
    Set oo = Nothing
End Sub


Private Sub FBOrdPro_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Eliminar As Integer
    Select Case ActionName
         Case "AGREGAR"
                fraOrdPro.Visible = True
                sTipo = "I"
                If cboCod_Fabrica.ListCount > 0 Then
                    cboCod_Fabrica.ListIndex = 0
                End If
                txtCod_OrdPro.Text = ""
                txtDes_EstPro.Text = ""
                
         Case "SUPRIMIR"
                sTipo = "D"
                Eliminar = MsgBox("Desea usted eliminar el registro seleccionado?", vbExclamation + vbYesNo)
                If Eliminar = vbYes Then
                    Call ELIMINA_ORDPRO
                    Call CARGA_ORDPRO
                Else
                    Exit Sub
                End If
    End Select
End Sub

Sub RepReqAvios()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String
    
    If Rs_OrdPro.EOF And Rs_OrdPro.BOF Then Exit Sub
    
    'Ruta = "C:\Archivos de programa\Gestion de Pedidos\RequerimientoAvios.xlt"
    Ruta = vRuta & "\RequerimientoAvios.xlt"

    Set oo = CreateObject("excel.application")
    oo.workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.run "Reporte", cCONNECT, Rs_GrupoPro("Cod_GrupoPro").Value
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler Err, "RepReqAvios"
    Set oo = Nothing
End Sub

Private Sub Form_Load()

    Set Rs_GrupoPro = New ADODB.Recordset
    Set Rs_OrdPro = New ADODB.Recordset

    Call FormateaGrid(DG_EstCli)
    Call FormateaGrid(DG_EstPro)
    'Call CARGA_GRUPOPRO
    Call LLENA_COMBOS
    Me.FBGrupoPro.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    Me.FBOrdPro.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub txtAbr_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtAbr_Cliente.Text) = "" Then
            cmdBuscaCliente_Click
        Else
            StrSQL = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE Abr_Cliente ='" & Trim(txtAbr_Cliente.Text) & "'"
            txtNom_Cliente.Text = DevuelveCampo(StrSQL, cCONNECT)
            txtCod_GrupoPro.SetFocus
        End If
    End If
End Sub

Private Sub txtAbrCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtAbrCliente.Text) = "" Then
            txtAbrCliente.Text = ""
            txtAbrCliente.SetFocus
        Else
            txtAbrCliente.Text = UCase(txtAbrCliente.Text)
            StrSQL = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE Abr_Cliente ='" & Trim(txtAbrCliente.Text) & "'"
            txtNomCliente.Text = DevuelveCampo(StrSQL, cCONNECT)
            FBBuscar.SetFocus
        End If
    End If
End Sub


Private Sub txtCod_OrdPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_OrdPro.Text) = "" Then
            cmdBusca_OrdPro_Click
        Else
            StrSQL = "SELECT Des_EstPro From Es_ordpro OP, Es_EstPro EP WHERE OP.Cod_EstPro = EP.Cod_EstPro AND Cod_fabrica ='" & _
            Right(cboCod_Fabrica.Text, 3) & "' AND Cod_Cliente ='" & _
            Rs_GrupoPro("Cod_Cliente").Value & "' AND Cod_GrupoPro='' AND cod_ordpro='" & _
            txtCod_OrdPro.Text & "'"
            txtDes_EstPro.Text = DevuelveCampo(StrSQL, cCONNECT)
            cmdAceptarOP.SetFocus
        End If
    End If
End Sub

Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtNom_Cliente.Text) > 4 Then
            StrSQL = "SELECT Abr_Cliente FROM TG_CLIENTE WHERE Nom_Cliente LIKE '" & Trim(txtNom_Cliente.Text) & "%'"
            txtAbr_Cliente.Text = DevuelveCampo(StrSQL, cCONNECT)
            If Trim(txtAbr_Cliente.Text) <> "" Then
                txtAbr_Cliente_KeyPress (13)
            End If
        Else
            MsgBox ("El Texto Ingresado debe contar con un mínimo de 5 caracteres")
        End If
    End If
End Sub

Private Sub txtNomCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtNomCliente.Text) > 4 Then
            StrSQL = "SELECT Abr_Cliente FROM TG_CLIENTE WHERE Nom_Cliente LIKE '" & Trim(txtNomCliente.Text) & "%'"
            txtAbrCliente.Text = DevuelveCampo(StrSQL, cCONNECT)
            FBBuscar.SetFocus
            'If Trim(txtAbrCliente.Text) <> "" Then
            '    txtAbrCliente_KeyPress (13)
            'End If
        Else
            MsgBox ("El Texto Ingresado debe contar con un mínimo de 5 caracteres")
        End If
    End If
End Sub

Public Sub REPORTE()
On Error GoTo ErrorImpresion
    Dim oo As Object
    Set oo = CreateObject("excel.application")
    'oo.Workbooks.Open App.Path & "\RptRevisionColores.xlt"
    oo.workbooks.Open vRuta & "\RptRevisionColores.xlt"
    oo.Visible = True
    oo.run "REPORTE", Rs_GrupoPro("Cod_GrupoPro").Value, Rs_GrupoPro("Cod_GrupoPro").Value & " - " & Rs_GrupoPro("Des_GrupoPro").Value, cCONNECT
    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte de Revisión de Colores " & Err.Description, vbCritical, "Impresion"
End Sub
