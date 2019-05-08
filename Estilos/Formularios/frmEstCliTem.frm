VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmEstCliTem 
   Caption         =   "Estilos del Cliente por Temporada"
   ClientHeight    =   7920
   ClientLeft      =   300
   ClientTop       =   690
   ClientWidth     =   13830
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   13830
   Begin VB.Frame fraImpFecMaxAPro 
      Height          =   1635
      Left            =   10620
      TabIndex        =   39
      Top             =   4920
      Visible         =   0   'False
      Width           =   2655
      Begin VB.OptionButton optFecApro 
         Caption         =   "Todos"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   660
         Width           =   2445
      End
      Begin VB.OptionButton optFecApro 
         Caption         =   "Solo cliente temporada"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   330
         Value           =   -1  'True
         Width           =   2445
      End
      Begin FunctionsButtons.FunctButt FunctButt3 
         Height          =   510
         Left            =   60
         TabIndex        =   42
         Top             =   990
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmEstCliTem.frx":0000
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
   End
   Begin VB.Frame fraEstPro 
      Caption         =   "Estilos Propios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3150
      Left            =   0
      TabIndex        =   9
      Top             =   4560
      Width           =   13695
      Begin VB.CommandButton CmdImprimirDistribucionPOs 
         Caption         =   "Imprimir Distribución de POs por Est/Version/Dest/Fec"
         Height          =   615
         Left            =   7500
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2250
         Width           =   1695
      End
      Begin VB.CommandButton CmdImprimirDistOPsColor 
         Caption         =   "Imprimir Distribución de POs por Estilo/ Version/Dest/Fec/Col"
         Height          =   615
         Left            =   5670
         TabIndex        =   37
         Top             =   2250
         Width           =   1815
      End
      Begin VB.CommandButton cmdImprimirFecMaxAprobacion 
         Caption         =   "Imprimir Fecha Max Aprobacion"
         Height          =   600
         Left            =   12360
         TabIndex        =   36
         Top             =   2250
         Width           =   1020
      End
      Begin VB.CommandButton cmdEstiloCliTemp 
         Caption         =   "Imprime Estilo Cliente Temporada"
         Height          =   615
         Left            =   9210
         TabIndex        =   28
         Top             =   2250
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Telas Colores - Muestras"
         Height          =   600
         Left            =   10440
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2250
         Width           =   1020
      End
      Begin VB.CommandButton cmdImprimirEstilosPo 
         Caption         =   "Imprimir Estilos PO"
         Height          =   600
         Left            =   11460
         TabIndex        =   30
         Top             =   2250
         Width           =   900
      End
      Begin VB.Frame fraClaPo 
         BackColor       =   &H80000000&
         Height          =   1545
         Left            =   3645
         TabIndex        =   31
         Top             =   120
         Visible         =   0   'False
         Width           =   5925
         Begin FunctionsButtons.FunctButt FunctButt2 
            Height          =   510
            Left            =   1830
            TabIndex        =   34
            Top             =   885
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   900
            Custom          =   $"frmEstCliTem.frx":008D
            Orientacion     =   0
            Style           =   0
            Language        =   0
            TypeImageList   =   0
            ControlWidth    =   1155
            ControlHeigth   =   490
            ControlSeparator=   110
         End
         Begin VB.TextBox txtDes_ClaPurOrd 
            Height          =   300
            Left            =   2025
            TabIndex        =   33
            Top             =   375
            Width           =   3630
         End
         Begin VB.TextBox TxtCod_ClaPurOrd 
            Height          =   285
            Left            =   1140
            TabIndex        =   32
            Text            =   "PO"
            Top             =   375
            Width           =   825
         End
         Begin VB.Label Label5 
            Caption         =   "Clase PO"
            Height          =   300
            Left            =   180
            TabIndex        =   35
            Top             =   420
            Width           =   975
         End
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   1470
         Left            =   11280
         TabIndex        =   27
         Top             =   195
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   2593
         Custom          =   $"frmEstCliTem.frx":011A
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1200
         ControlHeigth   =   450
         ControlSeparator=   50
      End
      Begin GridEX20.GridEX DG_EstPro 
         Height          =   1680
         Left            =   150
         TabIndex        =   26
         Top             =   300
         Width           =   11100
         _ExtentX        =   19579
         _ExtentY        =   2963
         Version         =   "2.0"
         HoldSortSettings=   -1  'True
         DefaultGroupMode=   1
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "frmEstCliTem.frx":0247
         Column(2)       =   "frmEstCliTem.frx":030F
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmEstCliTem.frx":03B3
         FormatStyle(2)  =   "frmEstCliTem.frx":04EB
         FormatStyle(3)  =   "frmEstCliTem.frx":059B
         FormatStyle(4)  =   "frmEstCliTem.frx":064F
         FormatStyle(5)  =   "frmEstCliTem.frx":0727
         FormatStyle(6)  =   "frmEstCliTem.frx":07DF
         FormatStyle(7)  =   "frmEstCliTem.frx":08BF
         FormatStyle(8)  =   "frmEstCliTem.frx":096B
         ImageCount      =   0
         PrinterProperties=   "frmEstCliTem.frx":0A1B
      End
      Begin VB.Frame fraAdicEstPro 
         Caption         =   "Adicionar Estilo Propio"
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
         TabIndex        =   12
         Top             =   2040
         Visible         =   0   'False
         Width           =   13470
         Begin VB.CommandButton cmdNuePropio 
            Caption         =   "&Nuevo"
            Height          =   315
            Left            =   4650
            TabIndex        =   24
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtNum_Veces 
            Height          =   285
            Left            =   1200
            TabIndex        =   17
            Text            =   "1"
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "A&ceptar"
            Height          =   315
            Left            =   2340
            TabIndex        =   18
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtDes_estpro 
            Height          =   285
            Left            =   2400
            TabIndex        =   16
            Top             =   240
            Width           =   2235
         End
         Begin VB.CommandButton cmdBusca_EstPro 
            Caption         =   "..."
            Height          =   285
            Left            =   2160
            TabIndex        =   15
            Top             =   240
            Width           =   285
         End
         Begin VB.TextBox txtCod_EstPro 
            Height          =   285
            Left            =   1200
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   315
            Left            =   3420
            TabIndex        =   20
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Veces :"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   600
            Width           =   885
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Estilo Propio"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   870
         End
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   495
         Left            =   12600
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin FunctionsButtons.FunctButt FBEstPro 
         Height          =   510
         Left            =   255
         TabIndex        =   11
         Top             =   2130
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   900
         Custom          =   $"frmEstCliTem.frx":0BF3
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
   End
   Begin VB.Frame fraEstCli 
      Caption         =   "Estilos Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3645
      Left            =   0
      TabIndex        =   8
      Top             =   915
      Width           =   13710
      Begin FunctionsButtons.FunctButt FBEstCli 
         Height          =   510
         Left            =   90
         TabIndex        =   10
         Top             =   3100
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   900
         Custom          =   $"frmEstCliTem.frx":0CD7
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1080
         ControlHeigth   =   490
         ControlSeparator=   30
      End
      Begin GridEX20.GridEX DG_EstCli 
         Height          =   2850
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   13500
         _ExtentX        =   23813
         _ExtentY        =   5027
         Version         =   "2.0"
         HoldSortSettings=   -1  'True
         DefaultGroupMode=   1
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "frmEstCliTem.frx":10E4
         Column(2)       =   "frmEstCliTem.frx":11AC
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmEstCliTem.frx":1250
         FormatStyle(2)  =   "frmEstCliTem.frx":1388
         FormatStyle(3)  =   "frmEstCliTem.frx":1438
         FormatStyle(4)  =   "frmEstCliTem.frx":14EC
         FormatStyle(5)  =   "frmEstCliTem.frx":15C4
         FormatStyle(6)  =   "frmEstCliTem.frx":167C
         FormatStyle(7)  =   "frmEstCliTem.frx":175C
         FormatStyle(8)  =   "frmEstCliTem.frx":1808
         ImageCount      =   0
         PrinterProperties=   "frmEstCliTem.frx":18B8
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
      Height          =   855
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   13650
      Begin VB.CommandButton cmdBusCliente 
         Caption         =   "..."
         Height          =   285
         Left            =   1440
         TabIndex        =   23
         Tag             =   "..."
         Top             =   300
         Width           =   300
      End
      Begin VB.TextBox txtNom_TemCli 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6360
         TabIndex        =   19
         Top             =   300
         Width           =   2205
      End
      Begin VB.TextBox txtDes_Cliente 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   300
         Width           =   2325
      End
      Begin VB.CommandButton cmdBusca_Temporada 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   6120
         TabIndex        =   6
         Top             =   300
         Width           =   300
      End
      Begin VB.TextBox txtCod_TemCli 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   5
         Top             =   300
         Width           =   495
      End
      Begin VB.TextBox txtAbr_Cliente 
         Height          =   285
         Left            =   795
         TabIndex        =   2
         Top             =   285
         Width           =   615
      End
      Begin FunctionsButtons.FunctButt FBBuscar 
         Height          =   495
         Left            =   10320
         TabIndex        =   7
         Top             =   210
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Temporada"
         Height          =   195
         Left            =   4680
         TabIndex        =   4
         Top             =   345
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   345
         Width           =   480
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   9900
      Top             =   660
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmEstCliTem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrSQL As String
'Public rs_EstCli As ADODB.Recordset
'Dim rs_EstPro As ADODB.Recordset
Public Codigo As String, Descripcion As String

Public Estilo As String
Attribute Estilo.VB_VarUserMemId = 1073938435
Public Desc As String
Attribute Desc.VB_VarUserMemId = 1073938436
Public Valor As String
Attribute Valor.VB_VarUserMemId = 1073938437
Dim sTipo As String
Attribute sTipo.VB_VarUserMemId = 1073938438

Public varEst_Cot As Boolean
Attribute varEst_Cot.VB_VarUserMemId = 1073938439
Public varNumCot As Integer
Attribute varNumCot.VB_VarUserMemId = 1073938440
Public varObs As String
Attribute varObs.VB_VarUserMemId = 1073938441
Dim vCod_Cliente As String
Attribute vCod_Cliente.VB_VarUserMemId = 1073938442
Dim i As Long
Attribute i.VB_VarUserMemId = 1073938443
Dim vMensaje As Variant
Attribute vMensaje.VB_VarUserMemId = 1073938444
Dim sCodPO As String
Attribute sCodPO.VB_VarUserMemId = 1073938445

Private Function VALIDA_DATOSESTPRO() As Boolean
    VALIDA_DATOSESTPRO = True

    If sTipo <> "D" Then

    Else

        StrSQL = "SELECT COUNT(*) FROM TG_ESTCLICOLPRE" & _
                 "WHERE   Cod_Cliente = '" & DG_EstCli.Value(DG_EstCli.Columns("Cod_Cliente").Index) & "' AND " & _
                 "Cod_TemCli  = '" & DG_EstCli.Value(DG_EstCli.Columns("cod_temcli").Index) & "' AND " & _
                 "Cod_EstCli  = '" & DG_EstCli.Value(DG_EstCli.Columns("cod_estcli").Index) & "' AND " & _
                 "Cod_EstPro  = '" & DG_EstPro.Value(DG_EstPro.Columns("Cod_EstPro").Index) & "'"

        If DevuelveCampo(StrSQL, cCONNECT) <> 0 Then
            VALIDA_DATOSESTPRO = False
            Call MsgBox("No se puede eliminar por que posee colores " & Chr(13) & "asignados. Sirvase verificar", vbInformation)
            Exit Function
        End If
    End If
End Function

Public Sub CARGA_ESTCLI()
    vCod_Cliente = DevuelveCampo("SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'", cCONNECT)

    StrSQL = "EXEC UP_SEL_ESTILOCLIENTE '" & vCod_Cliente & "','" & txtCod_TemCli.Text & "'"
    Set DG_EstCli.ADORecordset = CargarRecordSetDesconectado(StrSQL, cCONNECT)
    If DG_EstCli.RowCount > 0 Then
        Call CARGA_ESTPRO
        HabilitaMant Me.FBEstCli, "ADICIONAR/MODIFICAR/ELIMINAR/COLORES/IMPRIMIR/TEMPORADA/COTIZACIONES/CAMBIOESTILO/COPIARESTILO/ESTAMPADOS/DATOSCOMPLEMENTARIOS"
    Else
        'Set DG_EstPro.DataSource = Nothing
        HabilitaMant Me.FBEstCli, "ADICIONAR/TEMPORADA/COPIARESTILO"
        HabilitaMant Me.FBEstPro, ""
        DG_EstPro.Refresh
    End If

    Call formato_grid

    StrSQL = "SELECT Num_Solicitud_Cons FROM TG_TEMCLI WHERE Cod_Cliente = '" & DevuelveCampo(StrSQL, cCONNECT) & "' AND Cod_TemCli = '" & txtCod_TemCli.Text & "'"
End Sub

Private Sub CARGA_ESTPRO()
    StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"

    StrSQL = "EXEC UP_SEL_ESTILOSPROPIOS '" & DevuelveCampo(StrSQL, cCONNECT) & "','" & txtCod_TemCli.Text & "','" & DG_EstCli.Value(DG_EstCli.Columns("cod_estcli").Index) & "'"
    Set DG_EstPro.ADORecordset = CargarRecordSetDesconectado(StrSQL, cCONNECT)

    Call formato_grid_EstPro

    If DG_EstPro.RowCount > 0 Then
        HabilitaMant Me.FBEstPro, "AGREGAR/MODIFICAR/SUPRIMIR"
    Else
        HabilitaMant Me.FBEstPro, "AGREGAR"
    End If
End Sub

Private Sub ANADE_ESTPRO()

    On Error GoTo Salvar_DatosErr

    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Con.ConnectionString = cCONNECT
    Con.Open

    Con.BeginTrans

    'Esta cadena es la que nos devolvera los items segun la seleccion establecida
    StrSQL = "EXEC UP_MAN_ESTCLIEST '" & _
             sTipo & "','" & _
             DG_EstCli.Value(DG_EstCli.Columns("Cod_Cliente").Index) & "','" & _
             DG_EstCli.Value(DG_EstCli.Columns("Cod_temcli").Index) & "','" & _
             DG_EstCli.Value(DG_EstCli.Columns("Cod_estcli").Index) & "','" & _
             txtCod_EstPro.Text & "'," & _
             txtNum_Veces.Text

    Con.Execute StrSQL

    Con.CommitTrans
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "ANADE_ESTPRO"
    'Call MsgBox("Ocurrio un error al añadir el Estilo Propio", vbCritical)
End Sub

Private Sub ELIMINA_ESTPRO()

    On Error GoTo Salvar_DatosErr

    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Con.ConnectionString = cCONNECT
    Con.Open

    Con.BeginTrans

    'Esta cadena es la que nos devolvera los items segun la seleccion establecida
    StrSQL = "EXEC UP_MAN_ESTCLIEST '" & _
             "D" & "','" & _
             DG_EstCli.Value(DG_EstCli.Columns("Cod_Cliente").Index) & "','" & _
             DG_EstCli.Value(DG_EstCli.Columns("Cod_temcli").Index) & "','" & _
             DG_EstCli.Value(DG_EstCli.Columns("Cod_estcli").Index) & "','" & _
             DG_EstPro.Value(DG_EstPro.Columns("Cod_EstPro").Index) & "',''"


    Con.Execute StrSQL

    Con.CommitTrans
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "ELIMINA_ESTPRO"
    'Call MsgBox("Ocurrio un error al eliminar el Estilo Propio", vbCritical)
End Sub


Private Function VALIDA_CLIENTETEMPORADA() As Boolean
    VALIDA_CLIENTETEMPORADA = True

    If txtAbr_Cliente.Text = "" Or txtCod_TemCli.Text = "" Then
        Call MsgBox("Sirvase seleccionar un cliente y una temporada", vbExclamation)
        VALIDA_CLIENTETEMPORADA = False
        Exit Function
    Else

        StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
        StrSQL = "SELECT Cod_TemCli FROM TG_TemCli WHERE Cod_Cliente='" & DevuelveCampo(StrSQL, cCONNECT) & "' AND Cod_TemCli='" & txtCod_TemCli & "'"

        If DevuelveCampo(StrSQL, cCONNECT) = "" Then
            Call MsgBox("Los datos ingresados no son validos. Sirvase verificar", vbExclamation)
            VALIDA_CLIENTETEMPORADA = False
            Exit Function
        End If
    End If

    If sTipo = "COLORES" Then
        If DG_EstCli.Value(DG_EstCli.Columns("Num_EstProRea").Index) <> DG_EstCli.Value(DG_EstCli.Columns("Num_EstProAsg").Index) Then
            Call MsgBox("Estilo cliente sin estilo propio asignado. Sirvase verificar", vbInformation)
            VALIDA_CLIENTETEMPORADA = False
        End If
    End If

End Function

Private Sub BUSCA_TEMPORADA()
    Dim oTipo As New frmBusqGeneral
    Dim rs As New ADODB.Recordset
    Set oTipo.oParent = Me

    StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
    oTipo.sQuery = "SELECT  Cod_TemCli as Código, Nom_TemCli as Descripción FROM TG_TemCli WHERE Cod_Cliente='" & DevuelveCampo(StrSQL, cCONNECT) & "'"

    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtCod_TemCli.Text = Codigo
        txtNom_TemCli.Text = Descripcion
    End If
    Set oTipo = Nothing
    Set rs = Nothing

    FBBuscar.SetFocus
End Sub

Private Sub BUSCA_ESTPRO()
    Dim oTipo As New frmBusqGeneral
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Set oTipo.oParent = Me

    If Trim(txtCod_EstPro.Text) <> "" Then
        StrSQL = "SELECT  Cod_EstPro as Código, Des_EstPro as Descripción FROM ES_EstPro WHERE Cod_EstPro LIKE '" & txtCod_EstPro.Text & "%'"
    Else
        If Len(txtDes_estpro.Text) < 5 Then
            MsgBox ("El Texto Ingresado debe contar con un mínimo de 5 caracteres")
            txtDes_estpro.SetFocus
            Exit Sub
        End If
        StrSQL = "SELECT  Cod_EstPro as Código, Des_EstPro as Descripción FROM ES_EstPro WHERE Des_EstPro LIKE '" & txtDes_estpro.Text & "%'"
    End If

    oTipo.sQuery = StrSQL
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtCod_EstPro.Text = Trim(Codigo)
        txtDes_estpro.Text = Trim(Descripcion)
    End If
    Set oTipo = Nothing
    Set rs = Nothing

    FBBuscar.SetFocus
End Sub


Private Sub ELIMINAESTCLI()
    Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr

    StrSQL = "SELECT Cod_EstCli FROM tg_estcliest WHERE Cod_Cliente='" & DG_EstCli.Value(DG_EstCli.Columns("cod_cliente").Index) & "' AND Cod_TemCli='" & DG_EstCli.Value(DG_EstCli.Columns("cod_temcli").Index) & "' AND Cod_EstCli='" & DG_EstCli.Value(DG_EstCli.Columns("cod_estcli").Index) & "'"

    If DevuelveCampo(StrSQL, cCONNECT) <> "" Then
        MsgBox ("No se puede eliminar el Registro por que posee registros relacionados")
        Exit Sub
    End If

    Con.ConnectionString = cCONNECT
    Con.Open
    Con.BeginTrans

    'Esta es la sentencia que realiza la eliminacion del Registro
    StrSQL = "UP_MAN_ESTCLITEM " & _
             "D" & ",'" & _
             DG_EstCli.Value(DG_EstCli.Columns("cod_cliente").Index) & "','" & _
             DG_EstCli.Value(DG_EstCli.Columns("cod_temcli").Index) & "','" & _
             DG_EstCli.Value(DG_EstCli.Columns("cod_estcli").Index) & "','" & _
             DG_EstCli.Value(DG_EstCli.Columns("des_estcli").Index) & "'," & _
             DG_EstCli.Value(DG_EstCli.Columns("num_estprorea").Index) & ",'','',''"

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

Private Function VALIDA_ANADE_ESTPRO() As Boolean

    VALIDA_ANADE_ESTPRO = True
    If sTipo = "I" Then
        'Aqui entra solo si es Insert
        If Trim(txtCod_EstPro.Text) = "" Then
            Call MsgBox("Sirvase ingresar un Estilo propio", vbInformation)
            VALIDA_ANADE_ESTPRO = False
            Exit Function
        Else
            StrSQL = "SELECT Cod_EstPro FROM ES_EstPro WHERE Cod_EstPro='" & txtCod_EstPro.Text & "'"
            If DevuelveCampo(StrSQL, cCONNECT) = "" Then
                Call MsgBox("El Estilo propio no existe. Sirvase verificar", vbCritical)
                VALIDA_ANADE_ESTPRO = False
                Exit Function
            End If
        End If

        StrSQL = "SELECT Cod_EstPro FROM Tg_EstCliEst WHERE " & _
                 "Cod_Cliente='" & DG_EstCli.Value(DG_EstCli.Columns("cod_cliente").Index) & "' AND " & _
                 "Cod_TemCli='" & DG_EstCli.Value(DG_EstCli.Columns("cod_temcli").Index) & "' AND " & _
                 "Cod_EstCli='" & DG_EstCli.Value(DG_EstCli.Columns("cod_estcli").Index) & "' AND " & _
                 "Cod_EstPro='" & txtCod_EstPro.Text & "'"
        If DevuelveCampo(StrSQL, cCONNECT) <> "" Then
            Call MsgBox("Este registro ya se encuentra ingresado. Sirvase verificar", vbCritical)
            VALIDA_ANADE_ESTPRO = False
            Exit Function
        End If
    Else
        'Aqui entra si es update
    End If

    'Esta validacion es generica
    If Val(txtNum_Veces.Text) < 1 Then
        Call MsgBox("El campo Numero de Veces debe ser mayor o igual a 1. Sirvase verificar", vbExclamation)
        VALIDA_ANADE_ESTPRO = False
        Exit Function
    End If

End Function

Private Sub CmdAceptar_Click()
    If VALIDA_ANADE_ESTPRO Then
        i = DG_EstCli.Row
        'Valor = DG_EstCli.Columns(1)
        Call ANADE_ESTPRO
        Call CARGA_ESTCLI
        BuscaCampo DG_EstCli.ADORecordset, "cod_estcli", Valor
        Call CARGA_ESTPRO
        fraAdicEstPro.Visible = False
        DG_EstCli.Enabled = True
        DG_EstCli.Row = i
    End If
End Sub

Private Sub cmdBusca_EstPro_Click()
    Call BUSCA_ESTPRO
End Sub

Private Sub cmdBusca_Temporada_Click()
    Call BUSCA_TEMPORADA
    FBBuscar.SetFocus
End Sub

Private Sub cmdBusCliente_Click()
    Dim oTipo As New frmBusqGeneral
    Dim rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.sQuery = "SELECT Abr_Cliente as Código, nom_cliente as Descripción FROM TG_Cliente ORDER BY Abr_Cliente"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtAbr_Cliente.Text = Codigo
        txtDes_Cliente.Text = Descripcion
        txtCod_TemCli.Enabled = True
        txtNom_TemCli.Enabled = True
        cmdBusca_Temporada.Enabled = True
        txtCod_TemCli.SetFocus
        Codigo = ""
    End If
    Set oTipo = Nothing
    Set rs = Nothing
End Sub

Private Sub CmdCancelar_Click()
    fraAdicEstPro.Visible = False
    DG_EstCli.Enabled = True
End Sub

Private Sub cmdEstiloCliTemp_Click()

    On Error GoTo ErrorImpresion

    Dim oo As Object, lvSql As String, lvFiltro As String
    Set oo = CreateObject("excel.application")
    oo.workbooks.Open vRuta & "\StatusStyle.xlt"
    oo.Visible = True
    oo.DisplayAlerts = False

    oo.run "reporte", "Es_muestra_status_estilos_clientes_por_temporada_protos '" & vCod_Cliente & "','" & txtCod_TemCli.Text & "'", "tg_Up_Estatus_Style_Client", cCONNECT
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox Err.Description, vbCritical, "Impresion"
End Sub

Private Sub CmdImprimirDistOPsColor_Click()
    On Error GoTo ErrorImpresion

    Dim oo As Object
    Dim StrSQL, Cadena, CodCli As String
    If Trim(txtAbr_Cliente.Text) = "" Or Trim(txtNom_TemCli.Text) = "" Then
        MsgBox "Selecciones un cliente y una temprada...", vbInformation, "Imprimir"
        Exit Sub
    End If

    Set oo = CreateObject("excel.application")
    oo.workbooks.Open vRuta & "\RptDistribucionOpsColor.xlt"
    oo.Visible = True
    oo.DisplayAlerts = False
    StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
    'Cadena = "SELECT  Cod_TemCli as Código, Nom_TemCli as Descripción FROM TG_TemCli WHERE Cod_Cliente='" & DevuelveCampo(StrSQL, cCONNECT) & "'"

    CodCli = DevuelveCampo(StrSQL, cCONNECT)

    oo.run "reporte", "ES_ENCUENTRA_MATRIZ_DISTRIBUCION_PO_CLIENTE_TEMPORADA_COLOR '" & CodCli & "','" & txtCod_TemCli.Text & "'", txtDes_Cliente, txtNom_TemCli, cCONNECT
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox Err.Description, vbCritical, "Impresion"

End Sub

Private Sub CmdImprimirDistribucionPOs_Click()
    On Error GoTo ErrorImpresion

    Dim oo As Object
    Dim StrSQL, Cadena, CodCli As String
    If Trim(txtAbr_Cliente.Text) = "" Or Trim(txtNom_TemCli.Text) = "" Then
        MsgBox "Selecciones un cliente y una temprada...", vbInformation, "Imprimir"
        Exit Sub
    End If

    Set oo = CreateObject("excel.application")
    oo.workbooks.Open vRuta & "\RptDistribucionOps.xlt"
    oo.Visible = True
    oo.DisplayAlerts = False
    StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
    'Cadena = "SELECT  Cod_TemCli as Código, Nom_TemCli as Descripción FROM TG_TemCli WHERE Cod_Cliente='" & DevuelveCampo(StrSQL, cCONNECT) & "'"

    CodCli = DevuelveCampo(StrSQL, cCONNECT)

    oo.run "reporte", "ES_ENCUENTRA_MATRIZ_DISTRIBUCION_PO_CLIENTE_TEMPORADA '" & CodCli & "','" & txtCod_TemCli.Text & "'", txtDes_Cliente, txtNom_TemCli, cCONNECT
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox Err.Description, vbCritical, "Impresion"

End Sub

Private Sub cmdImprimirEstilosPo_Click()
    fraClaPo.Visible = True
    TxtCod_ClaPurOrd.SetFocus
End Sub

Private Sub cmdImprimirFecMaxAprobacion_Click()
    fraImpFecMaxAPro.Visible = True
    optFecApro(0).SetFocus
End Sub
Private Sub FunctButt3_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim CodCli As String
    Dim adoRs As ADODB.Recordset
    On Error GoTo Errox
    Select Case ActionName
    Case "ACEPT"
        If optFecApro(0).Value Then
            StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
            'Cadena = "SELECT  Cod_TemCli as Código, Nom_TemCli as Descripción FROM TG_TemCli WHERE Cod_Cliente='" & DevuelveCampo(StrSQL, cCONNECT) & "'"

            CodCli = DevuelveCampo(StrSQL, cCONNECT)
            Set adoRs = CargarRecordSetDesconectado("GERENCIA_COMERCIAL_SEGUIMIENTO_FECHAS_APROBACIONES '" & CodCli & "','" & txtCod_TemCli.Text & "','S'", cCONNECT)
        Else
            Set adoRs = CargarRecordSetDesconectado("GERENCIA_COMERCIAL_SEGUIMIENTO_FECHAS_APROBACIONES '','','T'", cCONNECT)
        End If
        Dim oo As Object
        Dim Ruta As String
        Ruta = ""

        Ruta = vRuta & "\GERENCIA_COMERCIAL_SEGUIMIENTO_FECHAS_APROBACIONES.XLT"

        Set oo = CreateObject("excel.application")
        oo.workbooks.Open Ruta
        oo.Visible = True
        oo.DisplayAlerts = False

        oo.run "Reporte", adoRs, Trim(txtDes_Cliente.Text), Trim(txtNom_TemCli.Text), IIf(optFecApro(0).Value, "S", "T")
        Set oo = Nothing
        fraImpFecMaxAPro.Visible = False
        Exit Sub
    Case "CANCE"
        fraImpFecMaxAPro.Visible = False
        Exit Sub
    End Select
Errox:
    ErrorHandler Err, "Reporte"

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()



    Dim oo As Object



    On Error GoTo ErrorImpresion



    If Trim(vCod_Cliente) = "" Then

        MsgBox "Debe seleccionar Cliente/Temporada"

        Exit Sub

    End If



    sCodPO = ""



    BUSCAPO



    If sCodPO = "" Then Exit Sub



    Set oo = CreateObject("excel.application")

    oo.workbooks.Open vRuta & "\RptMuestras.XLT"

    oo.Visible = True

    oo.DisplayAlerts = False

    oo.run "reporte", "es_muestra_matriz_muestras_estilos_colores_po '" & vCod_Cliente & "','" & txtCod_TemCli.Text & "','" & sCodPO & "'", cCONNECT, Trim(txtAbr_Cliente) & "-" & Trim(txtDes_Cliente), Trim(txtCod_TemCli) & "-" & Trim(txtNom_TemCli), "M1", sCodPO

    Set oo = Nothing



    Exit Sub



ErrorImpresion:



    Set oo = Nothing

    MsgBox Err.Description, vbCritical, "Impresion"



End Sub



Private Sub BUSCAPO()

    On Error GoTo Fin

    Dim rstAux As ADODB.Recordset


    StrSQL = "ES_MUESTRA_POS_MUESTRA_POR_CLIENTE_TEMPORADA '" & vCod_Cliente & "','" & txtCod_TemCli.Text & "'"

    With frmBusqGeneral
        Set .oParent = Me
        .sQuery = StrSQL
        .Cargar_Datos

        Codigo = ".."
        .Show vbModal
        sCodPO = Codigo

    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing

    Exit Sub
    Resume
Fin:
    On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
           "Búsqueda "
End Sub


Private Sub DG_EstCli_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    If DG_EstCli.RowCount > 0 Then
        Call CARGA_ESTPRO
    End If
End Sub

Private Sub FBBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    DG_EstCli.Enabled = True
    Call CARGA_ESTCLI
End Sub

Private Sub FBEstCli_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    On Error GoTo hand
    Dim elimina As Integer
    sTipo = ""
    Select Case ActionName
    Case "ADICIONAR"
        If VALIDA_CLIENTETEMPORADA Then
            StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
            Load FrmManEstCliTem
            Set FrmManEstCliTem.oParent = Me
            FrmManEstCliTem.varCod_Cliente = DevuelveCampo(StrSQL, cCONNECT)
            FrmManEstCliTem.varCod_TemCli = txtCod_TemCli.Text
            FrmManEstCliTem.txtAbr_Cliente.Text = txtAbr_Cliente.Text
            FrmManEstCliTem.txtNom_TemCli.Text = txtNom_TemCli.Text
            FrmManEstCliTem.varCod_EstCli = ""
            FrmManEstCliTem.CARGA_LISTA
            FrmManEstCliTem.Carga_Datos
            FrmManEstCliTem.MFEstCli_ActionClick 0, 0, "ADICIONAR"
            FrmManEstCliTem.Show 1
            CARGA_ESTCLI
            'Call Asigna_EP
        End If
    Case "MODIFICAR"
        If VALIDA_CLIENTETEMPORADA Then
            i = DG_EstCli.Row
            StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
            Set FrmManEstCliTem.oParent = Me
            Load FrmManEstCliTem
            FrmManEstCliTem.varCod_Cliente = DevuelveCampo(StrSQL, cCONNECT)
            FrmManEstCliTem.varCod_TemCli = txtCod_TemCli.Text
            FrmManEstCliTem.txtAbr_Cliente.Text = txtAbr_Cliente.Text
            FrmManEstCliTem.txtNom_TemCli.Text = txtNom_TemCli.Text
            FrmManEstCliTem.varCod_EstCli = DG_EstCli.Value(DG_EstCli.Columns("cod_estcli").Index)
            FrmManEstCliTem.CARGA_LISTA
            FrmManEstCliTem.Carga_Datos
            FrmManEstCliTem.MFEstCli_ActionClick 0, 0, "MODIFICAR"
            FrmManEstCliTem.Show 1
            CARGA_ESTCLI
            DG_EstCli.Row = i
            'BuscaCampo DG_EstCli.ADORecordset, "cod_estcli", Valor

        End If
    Case "ELIMINAR"
        If VALIDA_CLIENTETEMPORADA Then
            If DG_EstCli.RowCount = 0 Then Exit Sub
            i = DG_EstCli.Row
            elimina = MsgBox("Desea usted eliminar el registro seleccionado?", vbExclamation + vbYesNo)
            If elimina = vbNo Then Exit Sub
            Call ELIMINAESTCLI
            CARGA_ESTCLI
            DG_EstCli.Row = i
        End If
    Case "COLORES"
        sTipo = "COLORES"
        If VALIDA_CLIENTETEMPORADA Then
            StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
            Load frmEstCliCol
            frmEstCliCol.varCod_Cliente = DevuelveCampo(StrSQL, cCONNECT)
            frmEstCliCol.varCod_TemCli = txtCod_TemCli.Text
            frmEstCliCol.varCod_EstCli = Trim(DG_EstCli.Value(DG_EstCli.Columns("Cod_EstCli").Index))
            frmEstCliCol.txtAbr_Cliente = txtAbr_Cliente.Text
            frmEstCliCol.txtDes_Cliente = txtDes_Cliente.Text
            frmEstCliCol.txtCod_TemCli = txtCod_TemCli.Text
            frmEstCliCol.txtNom_TemCli = txtNom_TemCli.Text
            frmEstCliCol.TxtCod_EstCli = Trim(DG_EstCli.Value(DG_EstCli.Columns("cod_estcli").Index))
            frmEstCliCol.TxtDes_EstCli = Trim(DG_EstCli.Value(DG_EstCli.Columns("des_EstCli").Index))
            'frmEstCliCol.LLENA_COMBOS
            frmEstCliCol.CARGA_ESTCLICOL
            frmEstCliCol.Show 1

        End If
    Case "IMPRIMIR"
        If DG_EstCli.Value(DG_EstCli.Columns("num_solicitud_cons").Index) = 0 Then
            MsgBox "Estilo sin Cotizacion", vbInformation, "Estilo Cliente"
        Else
            varNumCot = DG_EstCli.Value(DG_EstCli.Columns("num_solicitud_cons").Index)
            varObs = DG_EstCli.Value(DG_EstCli.Columns("observaciones").Index)
            'GeneraReportes
            Load frmRepEstCliTem
            frmRepEstCliTem.varAbr_Cliente = Me.txtAbr_Cliente.Text
            frmRepEstCliTem.varCod_TemCli = Me.txtCod_TemCli.Text
            frmRepEstCliTem.varDes_Cliente = Me.txtDes_Cliente.Text
            frmRepEstCliTem.varNom_TemCli = Me.txtNom_TemCli.Text
            frmRepEstCliTem.varNumCot = Me.varNumCot
            frmRepEstCliTem.varCod_EstCli = DG_EstCli.Value(DG_EstCli.Columns("cod_estcli").Index)
            If DG_EstPro.RowCount > 0 Then frmRepEstCliTem.varCod_EstPro = DG_EstPro.Value(DG_EstPro.Columns("Cod_EstPro").Index)
            frmRepEstCliTem.varObs = Me.varObs
            frmRepEstCliTem.Show 1

        End If
    Case "TEMPORADA"

        StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"

        Load frmManCliTem
        frmManCliTem.varCod_Cliente = DevuelveCampo(StrSQL, cCONNECT)
        frmManCliTem.varCod_TemCli = txtCod_TemCli.Text
        frmManCliTem.txtAbr_Cliente = txtAbr_Cliente.Text
        frmManCliTem.txtNom_TemCli = txtNom_TemCli.Text
        frmManCliTem.CARGA_LISTA
        frmManCliTem.Show 1
    Case "COTIZACION"
        Dim sMessage As Integer

        StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"

        Load frmSeleccionCotizacion
        With frmSeleccionCotizacion
            .varAbr_Cliente = Me.txtAbr_Cliente.Text
            .varCod_TemCli = Me.txtCod_TemCli.Text
            .varDes_Cliente = Me.txtDes_Cliente.Text
            .varNom_TemCli = Me.txtNom_TemCli.Text
            .varCod_EstCli = DG_EstCli.Value(DG_EstCli.Columns("cod_estcli").Index)
            If DG_EstPro.RowCount > 0 Then .varCod_EstPro = DG_EstPro.Value(DG_EstPro.Columns("Cod_EstPro").Index)
            frmRepEstCliTem.varObs = Me.varObs

            .varCod_Cliente = DevuelveCampo(StrSQL, cCONNECT)
            .CARGA_DATA
        End With
        varEst_Cot = True
        frmSeleccionCotizacion.Show 1
        CARGA_ESTCLI
    Case "CAMBIOESTILO"
        With frmCambioEstilo
            StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"

            .vmarCodCliente = DevuelveCampo(StrSQL, cCONNECT)
            .vmarAbrcliente = Me.txtAbr_Cliente
            .vmarNomCliente = Me.txtDes_Cliente
            .vmarCodTem = txtCod_TemCli.Text
            .vmarCodEstCli = DG_EstCli.Value(DG_EstCli.Columns("cod_estcli").Index)
            .vmarDesEstCli = DG_EstCli.Value(DG_EstCli.Columns("des_estcli").Index)
            .Show 1
            CARGA_ESTCLI
        End With
    Case "COPIARESTILO"
        StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
        Load frmCopiarEstilo
        frmCopiarEstilo.varCod_Cliente = DevuelveCampo(StrSQL, cCONNECT)
        frmCopiarEstilo.varCod_TemCli_origen = Me.txtCod_TemCli.Text
        frmCopiarEstilo.CARGA_TEMPORADA
        frmCopiarEstilo.Show 1

        Set frmCopiarEstilo = Nothing

        Call CARGA_ESTCLI
    Case "COPIAR"
        StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
        Load FrmCopiarEstiloNew
        FrmCopiarEstiloNew.varCod_Cliente = DevuelveCampo(StrSQL, cCONNECT)
        FrmCopiarEstiloNew.varCod_TemCli_origen = Me.txtCod_TemCli.Text
        FrmCopiarEstiloNew.varCod_EstCli = DG_EstCli.Value(DG_EstCli.Columns("cod_Estcli").Index)
        FrmCopiarEstiloNew.Show 1

        Set frmCopiarEstilo = Nothing

        Call CARGA_ESTCLI
    Case "ESTAMPADOS"
        Load frmEstCliTemp_Estampados
        StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
        frmEstCliTemp_Estampados.varCod_Cliente = DevuelveCampo(StrSQL, cCONNECT)

        frmEstCliTemp_Estampados.varCod_Cliente = DevuelveCampo(StrSQL, cCONNECT)
        frmEstCliTemp_Estampados.varCod_TemCli = txtCod_TemCli.Text

        frmEstCliTemp_Estampados.CARGA_LISTA
        frmEstCliTemp_Estampados.Carga_Datos
        frmEstCliTemp_Estampados.Show vbModal
        Set frmEstCliTemp_Estampados = Nothing
        Call CARGA_ESTCLI
    Case "DATOSCOMPLEMENTARIOS"
        If DG_EstCli.RowCount = 0 Then Exit Sub
        StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
        Load frmDatosComplementarios
        frmDatosComplementarios.sCod_Cliente = DevuelveCampo(StrSQL, cCONNECT)
        frmDatosComplementarios.sCod_EstCli = Trim(DG_EstCli.Value(DG_EstCli.Columns("cod_estcli").Index))
        frmDatosComplementarios.CargarDatos
        frmDatosComplementarios.Show vbModal
        Set frmDatosComplementarios = Nothing

    End Select
    Exit Sub
hand:
    ErrorHandler Err, "FBEstCli_ActionClick"

End Sub

Sub GeneraReportes()
    On Error GoTo hand
    Dim oo As Object
    Dim Ruta As String
    Dim Usu As String
    Dim StrSQL As String
    StrSQL = "select tip_fabrica from tg_control"
    If DevuelveCampo(StrSQL, cCONNECT) = 1 Then
        '    Ruta = App.Path & "\prototipo.xlt"
        Ruta = vRuta & "\prototipo.xlt"
    Else
        '    Ruta = App.Path & "\prototipoD.xlt"
        Ruta = vRuta & "\prototipoD.xlt"
    End If
    StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
    Set oo = CreateObject("excel.application")
    oo.workbooks.Open Ruta
    '    oo.workbooks.Add Ruta
    oo.Visible = False
    oo.DisplayAlerts = False
    oo.run "Reporte", CStr(DevuelveCampo(StrSQL, cCONNECT)), Me.txtCod_TemCli, varNumCot, cCONNECT, vemp, txtDes_Cliente, txtNom_TemCli, varObs, vusu
    '    oo.quit
    Set oo = Nothing
    Exit Sub
hand:
    ErrorHandler Err, "GeneraReportes"
    Set oo = Nothing
End Sub

Private Sub FBEstPro_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "AGREGAR"
        sTipo = "I"
        If DG_EstCli.Value(DG_EstCli.Columns("Num_EstProAsg").Index) < DG_EstCli.Value(DG_EstCli.Columns("Num_EstProRea").Index) Then
            DG_EstCli.Enabled = False
            Call Asigna_EP
            DG_EstCli.Enabled = True
            Call CARGA_ESTCLI
            '                    fraAdicEstPro.Visible = True
            '                    txtCod_EstPro.Enabled = True
            '                    txtDes_estpro.Enabled = True
            '                    txtCod_EstPro.Text = ""
            '                    txtDes_estpro.Text = ""
            '                    txtNum_Veces.Text = "1"
            '                    txtCod_EstPro.SetFocus

        Else
            Call MsgBox("No se pueden seguir añadiendo mas registros. Verifique el Nro de Estilos Reales", vbExclamation)
            Exit Sub
        End If
    Case "MODIFICAR"
        sTipo = "U"
        If DG_EstPro.RowCount = 0 Then Exit Sub
        DG_EstCli.Enabled = False
        fraAdicEstPro.Visible = True
        txtCod_EstPro.Text = DG_EstPro.Value(DG_EstPro.Columns("Cod_EstPro").Index)
        txtDes_estpro.Text = DG_EstPro.Value(DG_EstPro.Columns("des_EstPro").Index)
        txtCod_EstPro.Enabled = False
        txtDes_estpro.Enabled = False
        txtNum_Veces.Text = DG_EstPro.Value(DG_EstPro.Columns("num_veces").Index)
        txtNum_Veces.SetFocus

    Case "SUPRIMIR"
        sTipo = "I"
        If VALIDA_DATOSESTPRO = True Then
            i = DG_EstCli.Row
            'Valor = DG_EstCli.Value(DG_EstCli.Columns(1).Index)
            Call ELIMINA_ESTPRO
            Call CARGA_ESTCLI
            Call CARGA_ESTPRO
            DG_EstCli.Row = i
            'BuscaCampo DG_EstCli.ADORecordset, "cod_estcli", Valor
        End If
    End Select
End Sub

Private Sub Form_Load()
'Call FormateaGrid(DG_EstCli)
'Call FormateaGrid(DG_EstPro)
    'Me.FBEstCli.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    'Me.FBEstPro.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "VERSION"
        If DG_EstPro.RowCount = 0 Then Exit Sub
        Load FrmShowVersiones
        Set FrmShowVersiones.oParent = Me
        StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
        FrmShowVersiones.vCod_Cliente = DevuelveCampo(StrSQL, cCONNECT)
        FrmShowVersiones.vCod_TemCli = txtCod_TemCli.Text
        FrmShowVersiones.vCod_EstCli = DG_EstCli.Value(DG_EstCli.Columns("cod_estcli").Index)
        FrmShowVersiones.vCod_estPro = DG_EstPro.Value(DG_EstPro.Columns("Cod_EstPro").Index)
        FrmShowVersiones.CARGA_GRID
        FrmShowVersiones.Show vbModal
        Set FrmShowVersiones = Nothing
        Call CARGA_ESTPRO
    Case "WIPPROTOS"
        If DG_EstPro.RowCount = 0 Then Exit Sub
        Call WipProtos

    Case "ITERACION"
        If DG_EstPro.RowCount = 0 Then Exit Sub
        vMensaje = MsgBox("¿Esta seguro de crear una nueva iteración?", vbYesNo)
        If vMensaje = vbNo Then Exit Sub
        Call GeneraIteracion
    End Select
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "ACEPT"
        If TxtCod_ClaPurOrd.Text <> "" And txtAbr_Cliente.Text <> "" And txtCod_TemCli <> "" Then
            Imprimir_Reporte
        End If
    Case "CANCE"
        TxtCod_ClaPurOrd = "PO"
        txtDes_ClaPurOrd = ""
        fraClaPo.Visible = False
    End Select
End Sub


Sub Imprimir_Reporte()

    On Error GoTo ErrorImpresion

    Dim cod_cliente As String

    cod_cliente = DevuelveCampo("SELECT cod_cliente FROM TG_CLIENTE WHERE Abr_Cliente ='" & txtAbr_Cliente & "' ", cCONNECT)

    Dim AUXRS As ADODB.Recordset
    Set AUXRS = GetRecordset(cCONNECT, " ES_MUESTRA_DETALLE_APROBACIONES_POR_ESTILO_CLIENTE_PO  '" & Trim(cod_cliente) & "','" & Trim(txtCod_TemCli.Text) & "','" & TxtCod_ClaPurOrd.Text & "' ")

    Dim oo As Object

    Set oo = CreateObject("excel.application")
    oo.workbooks.Open vRuta & "\RPTESTILOS_PO.XLT"

    oo.Visible = True
    oo.DisplayAlerts = False
    oo.run "REPORTE", AUXRS, txtNom_TemCli.Text
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler Err, "Reporte"
End Sub



Private Sub txtAbr_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtAbr_Cliente.Text) = "" Then
            cmdBusCliente_Click
        Else
            StrSQL = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE Abr_Cliente LIKE '" & Trim(txtAbr_Cliente.Text) & "%'"
            txtDes_Cliente.Text = DevuelveCampo(StrSQL, cCONNECT)
            txtCod_TemCli.Enabled = True
            txtNom_TemCli.Enabled = True
            cmdBusca_Temporada.Enabled = True
            txtCod_TemCli.SetFocus

            HabilitaMant Me.FBEstCli, ""
            HabilitaMant Me.FBEstPro, ""

        End If
    End If
End Sub

Private Sub TxtCod_ClaPurOrd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BuscaClasePO (1)
        FunctButt2.SetFocus
    End If

End Sub

Private Sub txtDes_ClaPurOrd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BuscaClasePO (2)
        FunctButt2.SetFocus
    End If

End Sub


Private Sub BuscaClasePO(Opcion As Integer)
    Dim sField As String, iRows As Long
    Dim rstAux As ADODB.Recordset

    StrSQL = "Select Cod_ClaPurOrd, Des_ClaPurOrd From TG_CLAPURORD WHERE   "
    TxtCod_ClaPurOrd = Trim(TxtCod_ClaPurOrd)
    txtDes_ClaPurOrd = Trim(txtDes_ClaPurOrd)
    sField = TxtCod_ClaPurOrd
    Select Case Opcion
    Case 1: StrSQL = StrSQL & "Cod_ClaPurOrd like '%" & TxtCod_ClaPurOrd & "%'"
    Case 2: StrSQL = StrSQL & "Des_ClaPurOrd like '%" & txtDes_ClaPurOrd & "%'"
    End Select

    TxtCod_ClaPurOrd = ""
    txtDes_ClaPurOrd = ""
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = StrSQL
        .Caption = "Seleccionar - Clase PO"
        .Cargar_Datos

        Codigo = ""
        Descripcion = ""

        iRows = .DGridLista.RowCount
        Set rstAux = .DGridLista.ADORecordset
        If .DGridLista.RowCount > 1 Then
            .Show vbModal
        ElseIf .DGridLista.RowCount = 1 Then
            Codigo = .DGridLista.Value(.DGridLista.Columns("Cod_ClaPurOrd").Index)
            Descripcion = .DGridLista.Value(.DGridLista.Columns("Des_ClaPurOrd").Index)
        End If

        If Codigo <> "" Then
            TxtCod_ClaPurOrd = RTrim(Codigo)
            txtDes_ClaPurOrd = RTrim(Descripcion)
        End If
    End With
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
End Sub
Private Sub txtCod_EstPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_EstPro.Text) = "" Then
            'cmdBusCliente_Click
        Else
            txtCod_EstPro.Text = Right("00000" & Trim(txtCod_EstPro.Text), 5)
            StrSQL = "SELECT  Des_EstPro FROM ES_EstPro WHERE Cod_EstPro='" & txtCod_EstPro.Text & "'"
            txtDes_estpro.Text = Trim(DevuelveCampo(StrSQL, cCONNECT))
            txtNum_Veces.SetFocus
        End If
    End If
End Sub

Private Sub txtCod_TemCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_TemCli.Text) = "" Then
            Call BUSCA_TEMPORADA
        Else
            StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
            StrSQL = "SELECT Nom_TemCli FROM TG_TemCli WHERE Cod_Cliente='" & DevuelveCampo(StrSQL, cCONNECT) & "' AND Cod_TemCli='" & txtCod_TemCli.Text & "'"
            txtNom_TemCli.Text = DevuelveCampo(StrSQL, cCONNECT)

            HabilitaMant Me.FBEstCli, ""
            HabilitaMant Me.FBEstPro, ""

            FBBuscar.SetFocus
        End If
    End If
End Sub

Private Sub TxtDes_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtDes_Cliente) > 4 Then
            StrSQL = "SELECT Abr_Cliente FROM TG_CLIENTE WHERE Nom_Cliente LIKE '" & Trim(txtDes_Cliente.Text) & "%'"
            txtAbr_Cliente.Text = DevuelveCampo(StrSQL, cCONNECT)
            StrSQL = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
            txtDes_Cliente.Text = DevuelveCampo(StrSQL, cCONNECT)
            txtCod_TemCli.Enabled = True
            txtNom_TemCli.Enabled = True
            cmdBusca_Temporada.Enabled = True
            txtCod_TemCli.SetFocus

            HabilitaMant Me.FBEstCli, ""
            HabilitaMant Me.FBEstPro, ""

        Else
            MsgBox ("El Texto Ingresado debe contar con un mínimo de 5 caracteres")
            txtDes_Cliente.SetFocus
        End If
    End If
End Sub

Private Sub txtDes_estpro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtDes_estpro) > 4 Then
            StrSQL = "SELECT Cod_EstPro FROM ES_EstPro WHERE Des_EstPro LIKE '" & Trim(txtDes_estpro.Text) & "%'"
            txtCod_EstPro.Text = Trim(DevuelveCampo(StrSQL, cCONNECT))
            StrSQL = "SELECT  Des_EstPro FROM ES_EstPro WHERE Cod_EstPro='" & txtCod_EstPro.Text & "'"
            txtDes_estpro.Text = Trim(DevuelveCampo(StrSQL, cCONNECT))
            'txtCod_TemCli.SetFocus
        Else
            MsgBox ("El Texto Ingresado debe contar con un mínimo de 5 caracteres")
            txtDes_estpro.SetFocus
        End If
    End If
End Sub

Private Sub txtNum_Veces_(KeyAscii As Integer)
    Call SoloNumeros(txtNum_Veces, KeyAscii, False, 0, 3)
End Sub

Private Sub txtNum_Veces_LostFocus()
    If Trim(txtNum_Veces.Text) = "" Then
        txtNum_Veces.Text = 1
    End If
End Sub

Private Sub cmdNuePropio_Click()
'    Load FrmEstProp
'    FrmEstProp.Show 1
    Set FrmIngresoEstilo.Papa = Me
    FrmIngresoEstilo.Descripcion = DG_EstCli.Value(DG_EstCli.Columns("des_estcli").Index)
    FrmIngresoEstilo.Estilo = DG_EstCli.Value(DG_EstCli.Columns("cod_estcli").Index)
    FrmIngresoEstilo.Show 1

    '    Me.txtCod_EstPro = Me.Estilo
    '    Me.txtDes_EstPro = Me.Desc
    'Call LLENA_COMBOS
End Sub


Sub formato_grid()
    DG_EstCli.Columns("COD_ESTCLI").Width = 2000
    DG_EstCli.Columns("DES_ESTCLI").Width = 4250
    DG_EstCli.Columns("FLG_STATUS").Width = 700
    DG_EstCli.Columns("COMENTARIO").Width = 3000
    DG_EstCli.Columns("NUM_ESTPROREA").Width = 1100
    DG_EstCli.Columns("NUM_ESTPROASG").Width = 1100
    DG_EstCli.Columns("NUM_SOLICITUD_CONS").Width = 1000

    DG_EstCli.Columns("COD_ESTCLI").Caption = "Código"
    DG_EstCli.Columns("DES_ESTCLI").Caption = "Descripción"
    DG_EstCli.Columns("FLG_STATUS").Caption = "Estado"
    DG_EstCli.Columns("NUM_ESTPROREA").Caption = "Est.Reales"
    DG_EstCli.Columns("NUM_ESTPROASG").Caption = "Est.Asignados"
    DG_EstCli.Columns("NUM_SOLICITUD_CONS").Caption = "#Cotizacion"

    DG_EstCli.Columns("COD_CLIENTE").Width = 0
    DG_EstCli.Columns("COD_TEMCLI").Width = 0
    DG_EstCli.Columns("OBSERVACIONES").Width = 0

End Sub

Sub formato_grid_EstPro()
    DG_EstPro.Columns("cod_estpro").Width = 900
    DG_EstPro.Columns("des_estpro").Width = 4300
    DG_EstPro.Columns("num_veces").Width = 600
    DG_EstPro.Columns("Version_Costeo").Width = 1200
    DG_EstPro.Columns("Des_Version").Width = 3000
    DG_EstPro.Columns("ult_iteracion").Width = 1000

    DG_EstPro.Columns("cod_estpro").Caption = "Est.Propio"
    DG_EstPro.Columns("des_estpro").Caption = "Descripcion"
    DG_EstPro.Columns("num_veces").Caption = "Veces"
End Sub

Sub Asigna_EP()
    Load FrmAddEstiloPropio
    StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
    FrmAddEstiloPropio.vCod_Cliente = DevuelveCampo(StrSQL, cCONNECT)
    FrmAddEstiloPropio.txtAbr_Cliente = DevuelveCampo("select abr_cliente from tg_cliente where cod_cliente='" & FrmAddEstiloPropio.vCod_Cliente & "'", cCONNECT)
    FrmAddEstiloPropio.txtNom_Cliente = DevuelveCampo("select nom_cliente from tg_cliente where cod_cliente='" & FrmAddEstiloPropio.vCod_Cliente & "'", cCONNECT)
    FrmAddEstiloPropio.vCod_TemCli = txtCod_TemCli.Text
    FrmAddEstiloPropio.txtCod_TemCli = txtCod_TemCli.Text
    FrmAddEstiloPropio.TxtDes_TemCli = txtNom_TemCli
    FrmAddEstiloPropio.vCod_EstCli = DG_EstCli.Value(DG_EstCli.Columns("cod_estcli").Index)
    FrmAddEstiloPropio.TxtCod_EstCli.Text = DG_EstCli.Value(DG_EstCli.Columns("cod_estcli").Index)
    FrmAddEstiloPropio.TxtDes_EstCli.Text = DevuelveCampo("select Des_EstCli from tg_estclitem where cod_cliente ='" & FrmAddEstiloPropio.vCod_Cliente & "' and cod_temcli='" & txtCod_TemCli.Text & "' and cod_estcli='" & DG_EstCli.Value(DG_EstCli.Columns("cod_estcli").Index) & "'", cCONNECT)
    FrmAddEstiloPropio.sDes_Estilo = DevuelveCampo("select Des_EstCli from tg_estclitem where cod_cliente ='" & FrmAddEstiloPropio.vCod_Cliente & "' and cod_temcli='" & txtCod_TemCli.Text & "' and cod_estcli='" & DG_EstCli.Value(DG_EstCli.Columns("cod_estcli").Index) & "'", cCONNECT)
    FrmAddEstiloPropio.vDes_Tela = DevuelveCampo("select Des_Tela from tg_estclitem where cod_cliente ='" & FrmAddEstiloPropio.vCod_Cliente & "' and cod_temcli='" & txtCod_TemCli.Text & "' and cod_estcli='" & DG_EstCli.Value(DG_EstCli.Columns("cod_estcli").Index) & "'", cCONNECT)
    FrmAddEstiloPropio.Show vbModal
    Set FrmAddEstiloPropio = Nothing
End Sub

Sub WipProtos()
    StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
    vCod_Cliente = DevuelveCampo(StrSQL, cCONNECT)

    Load FrmShowWipProtos
    FrmShowWipProtos.vCod_Cliente = vCod_Cliente
    FrmShowWipProtos.vCod_TemCli = txtCod_TemCli.Text
    FrmShowWipProtos.vCod_EstCli = DG_EstCli.Value(DG_EstCli.Columns("cod_estcli").Index)
    FrmShowWipProtos.vCod_estPro = DG_EstPro.Value(DG_EstPro.Columns("cod_estpro").Index)
    FrmShowWipProtos.Cliente = UCase(Trim(txtAbr_Cliente.Text)) & "-" & Trim(txtDes_Cliente.Text)
    FrmShowWipProtos.Temporada = UCase(txtCod_TemCli.Text) & "-" & Trim(txtCod_TemCli.Text)
    FrmShowWipProtos.varNumCot = DG_EstCli.Value(DG_EstCli.Columns("NUM_SOLICITUD_CONS").Index)
    FrmShowWipProtos.varObs = DG_EstCli.Value(DG_EstCli.Columns("observaciones").Index)
    FrmShowWipProtos.varDes_Cliente = Me.txtDes_Cliente.Text
    FrmShowWipProtos.varNom_TemCli = Me.txtNom_TemCli.Text
    FrmShowWipProtos.CARGA_GRID
    FrmShowWipProtos.Show vbModal
    Set FrmShowWipProtos = Nothing
    Call CARGA_ESTPRO
End Sub

Sub GeneraIteracion()
    On Error GoTo errGeneraIteracion
    StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
    vCod_Cliente = DevuelveCampo(StrSQL, cCONNECT)

    StrSQL = "Es_Actualiza_Version_Costeo_Estilo '" & vCod_Cliente & "','" & txtCod_TemCli.Text & "','" & DG_EstCli.Value(DG_EstCli.Columns("cod_estcli").Index) & "','" & DG_EstPro.Value(DG_EstPro.Columns("Cod_EstPro").Index) & "','" & DG_EstPro.Value(DG_EstPro.Columns("version_costeo").Index) & "','I',0"
    Call ExecuteSQL(cCONNECT, StrSQL)

    Call CARGA_ESTPRO

    Exit Sub
errGeneraIteracion:
    MsgBox Err.Description, vbCritical, "Genera Iteracion"
End Sub
