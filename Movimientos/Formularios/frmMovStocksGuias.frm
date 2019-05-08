VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMovStocksGuias 
   Caption         =   "Movimiento de Stocks de Guias"
   ClientHeight    =   8295
   ClientLeft      =   2865
   ClientTop       =   1740
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   9405
   Begin VB.Frame fraFlechas 
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   675
      TabIndex        =   56
      Top             =   7515
      Width           =   2310
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1485
         Picture         =   "frmMovStocksGuias.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Ultimo"
         Top             =   45
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   1005
         Picture         =   "frmMovStocksGuias.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Siguiente"
         Top             =   45
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   525
         Picture         =   "frmMovStocksGuias.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Anterior"
         Top             =   45
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   45
         Picture         =   "frmMovStocksGuias.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Primero"
         Top             =   45
         Width           =   495
      End
   End
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   9195
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   495
         Left            =   7800
         TabIndex        =   15
         Top             =   315
         Width           =   1170
      End
      Begin VB.OptionButton optGuia 
         Caption         =   "Guia"
         Height          =   285
         Left            =   6240
         TabIndex        =   5
         Top             =   195
         Width           =   630
      End
      Begin VB.OptionButton optMov 
         Caption         =   "Movimiento"
         Height          =   195
         Left            =   4845
         TabIndex        =   4
         Top             =   225
         Width           =   1125
      End
      Begin VB.OptionButton optFecha 
         Caption         =   "Fecha"
         Height          =   195
         Left            =   3735
         TabIndex        =   3
         Top             =   225
         Value           =   -1  'True
         Width           =   750
      End
      Begin VB.ComboBox cboCod_Almacen 
         Height          =   315
         Left            =   1065
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   375
         Width           =   1620
      End
      Begin VB.Frame fraFecha 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3180
         TabIndex        =   6
         Top             =   420
         Width           =   4100
         Begin MSComCtl2.DTPicker dtpFecMovStk 
            Height          =   330
            Left            =   1995
            TabIndex        =   8
            Top             =   150
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   582
            _Version        =   393216
            Format          =   23592961
            CurrentDate     =   37579
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Mov.:"
            Height          =   195
            Left            =   405
            TabIndex        =   7
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.Frame fraGuia 
         Caption         =   "Guia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3180
         TabIndex        =   9
         Top             =   420
         Width           =   4100
         Begin VB.TextBox txtNumGuia 
            Height          =   285
            Left            =   2000
            MaxLength       =   15
            TabIndex        =   11
            Top             =   180
            Width           =   1800
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Guia :"
            Height          =   195
            Left            =   390
            TabIndex        =   10
            Top             =   255
            Width           =   765
         End
      End
      Begin VB.Frame fraMov 
         Caption         =   "Movimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3180
         TabIndex        =   12
         Top             =   420
         Width           =   4100
         Begin VB.TextBox txtNumMovStk 
            Height          =   285
            Left            =   2000
            MaxLength       =   6
            TabIndex        =   14
            Top             =   180
            Width           =   1800
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Movimiento :"
            Height          =   195
            Left            =   405
            TabIndex        =   13
            Top             =   225
            Width           =   1245
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Almacen :"
         Height          =   195
         Left            =   165
         TabIndex        =   1
         Top             =   450
         Width           =   705
      End
   End
   Begin VB.Frame fraOpciones 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   7845
      TabIndex        =   18
      Top             =   1080
      Width           =   1410
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   1110
         Left            =   75
         TabIndex        =   19
         Top             =   915
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   1958
         Custom          =   $"frmMovStocksGuias.frx":05C8
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   4035
      TabIndex        =   55
      Top             =   7545
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMovStocksGuias.frx":066C
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      Left            =   45
      TabIndex        =   20
      Top             =   4155
      Width           =   9240
      Begin VB.TextBox txtNum_MovStk 
         Height          =   315
         Left            =   45
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   2895
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtOrden_Compra 
         Height          =   285
         Left            =   1305
         MaxLength       =   12
         TabIndex        =   49
         Top             =   1560
         Width           =   1305
      End
      Begin VB.TextBox txtPedido 
         Height          =   285
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   47
         Top             =   1230
         Width           =   2340
      End
      Begin VB.TextBox txtReferencia 
         Height          =   285
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   45
         Top             =   900
         Width           =   2340
      End
      Begin VB.TextBox txtLinea2 
         Height          =   1150
         Left            =   5910
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   53
         Top             =   1920
         Width           =   3100
      End
      Begin VB.TextBox txtLinea1 
         Height          =   1150
         Left            =   1305
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Top             =   1920
         Width           =   3150
      End
      Begin VB.OptionButton optOtros 
         Caption         =   "Otros"
         Height          =   195
         Left            =   7740
         TabIndex        =   27
         Top             =   225
         Width           =   750
      End
      Begin VB.OptionButton optCliente 
         Caption         =   "Cliente"
         Height          =   195
         Left            =   6195
         TabIndex        =   26
         Top             =   210
         Width           =   870
      End
      Begin VB.OptionButton optProveedor 
         Caption         =   "Proveedor"
         Height          =   195
         Left            =   4305
         TabIndex        =   25
         Top             =   210
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.TextBox txtNum_Guia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1305
         TabIndex        =   24
         Top             =   570
         Width           =   1800
      End
      Begin MSComCtl2.DTPicker dtpFec_MovStk 
         Height          =   300
         Left            =   1305
         TabIndex        =   22
         Top             =   225
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         _Version        =   393216
         Format          =   23592961
         CurrentDate     =   37579
      End
      Begin VB.Frame fraProveedor 
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   3720
         TabIndex        =   28
         Top             =   480
         Width           =   5310
         Begin VB.TextBox txtCod_Proveedor 
            Height          =   300
            Left            =   1200
            MaxLength       =   12
            TabIndex        =   30
            Top             =   360
            Width           =   1200
         End
         Begin VB.TextBox txtDes_Proveedor 
            Height          =   300
            Left            =   2385
            MaxLength       =   50
            TabIndex        =   31
            Top             =   360
            Width           =   2805
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor :"
            Height          =   195
            Left            =   225
            TabIndex        =   29
            Top             =   405
            Width           =   825
         End
      End
      Begin VB.Frame fraCliente 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   3735
         TabIndex        =   32
         Top             =   420
         Visible         =   0   'False
         Width           =   5310
         Begin VB.CommandButton cmdBuscaCliente 
            Caption         =   "..."
            Height          =   330
            Left            =   1935
            TabIndex        =   35
            Top             =   210
            Width           =   330
         End
         Begin VB.TextBox txtNom_Cliente 
            Height          =   300
            Left            =   2265
            MaxLength       =   50
            TabIndex        =   36
            Top             =   225
            Width           =   2910
         End
         Begin VB.TextBox txtAbr_Cliente 
            Height          =   300
            Left            =   1200
            MaxLength       =   5
            TabIndex        =   34
            Top             =   210
            Width           =   735
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Left            =   225
            TabIndex        =   33
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame fraDestinatario 
         Caption         =   "Destinatario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3735
         TabIndex        =   37
         Top             =   420
         Visible         =   0   'False
         Width           =   5310
         Begin VB.TextBox txtRuc_Destinatario 
            Height          =   285
            Left            =   1200
            MaxLength       =   14
            TabIndex        =   43
            Top             =   840
            Width           =   1620
         End
         Begin VB.TextBox txtDom_Destinatario 
            Height          =   300
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   41
            Top             =   510
            Width           =   3990
         End
         Begin VB.TextBox txtDestinatario 
            Height          =   285
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   39
            Top             =   180
            Width           =   3990
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "R.U.C. :"
            Height          =   195
            Left            =   225
            TabIndex        =   42
            Top             =   915
            Width           =   570
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Dirección :"
            Height          =   195
            Left            =   225
            TabIndex        =   40
            Top             =   615
            Width           =   765
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nombre :"
            Height          =   195
            Left            =   225
            TabIndex        =   38
            Top             =   300
            Width           =   645
         End
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Orden Compra :"
         Height          =   195
         Left            =   135
         TabIndex        =   48
         Top             =   1650
         Width           =   1110
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Pedido :"
         Height          =   195
         Left            =   135
         TabIndex        =   46
         Top             =   1335
         Width           =   585
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Referencia :"
         Height          =   195
         Left            =   135
         TabIndex        =   44
         Top             =   1005
         Width           =   870
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Observacion 1:"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   2085
         Width           =   1080
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Observación 2:"
         Height          =   195
         Left            =   4680
         TabIndex        =   52
         Top             =   2025
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Guia"
         Height          =   195
         Left            =   135
         TabIndex        =   23
         Top             =   645
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Mov."
         Height          =   195
         Left            =   150
         TabIndex        =   21
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Frame fraLista 
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
      Height          =   3075
      Left            =   60
      TabIndex        =   16
      Top             =   1080
      Width           =   7740
      Begin GridEX20.GridEX gexLista 
         Height          =   2700
         Left            =   105
         TabIndex        =   17
         Top             =   240
         Width           =   7545
         _ExtentX        =   13309
         _ExtentY        =   4763
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmMovStocksGuias.frx":0812
         Column(2)       =   "frmMovStocksGuias.frx":08DA
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmMovStocksGuias.frx":097E
         FormatStyle(2)  =   "frmMovStocksGuias.frx":0AB6
         FormatStyle(3)  =   "frmMovStocksGuias.frx":0B66
         FormatStyle(4)  =   "frmMovStocksGuias.frx":0C1A
         FormatStyle(5)  =   "frmMovStocksGuias.frx":0CF2
         FormatStyle(6)  =   "frmMovStocksGuias.frx":0DAA
         ImageCount      =   0
         PrinterProperties=   "frmMovStocksGuias.frx":0E8A
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   0
      Top             =   7560
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmMovStocksGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim varOpcion As Integer
Dim varBusqueda As String
Dim sTipo As String

Public varCod_Cliente_Tex As String
Public Codigo As String, Descripcion As String
Public Paso As String

Public Sub BUSCA_PROVEEDOR(Tipo As Integer)
    Select Case Tipo
        Case 1:
                    strSQL = "SELECT Des_Proveedor as 'Descripción' FROM tx_proveedor WHERE Cod_Proveedor = '" & Trim(Me.txtCod_Proveedor.Text) & "'"
                    Me.txtDes_Proveedor.Text = Trim(DevuelveCampo(strSQL, cConnect))
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim Rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.sQuery = "SELECT Cod_Proveedor AS 'Código', Des_Proveedor as 'Descripción' FROM tx_proveedor WHERE Des_Proveedor LIKE '%" & Trim(Me.txtDes_Proveedor.Text) & "%'"
                    Else
                        oTipo.sQuery = "SELECT Cod_Proveedor AS 'Código', Des_Proveedor as 'Descripción' FROM tx_proveedor "
                    End If
                    
                    oTipo.CARGAR_DATOS
                    oTipo.Show 1
                    If Codigo <> "" Then
                        Me.txtCod_Proveedor.Text = Trim(Codigo)
                        Me.txtDes_Proveedor.Text = Trim(Descripcion)
                        Codigo = "": Descripcion = ""
                    End If
                    Set oTipo = Nothing
                    Set Rs = Nothing
                    
    End Select
    txtReferencia.SetFocus
End Sub

Public Sub BUSCA_CLIENTE(Tipo As Integer)
    Select Case Tipo
        Case 1:
                    strSQL = "SELECT Nom_Cliente as 'Descripción' FROM tx_cliente WHERE Abr_Cliente ='" & Trim(Me.txtAbr_Cliente.Text) & "'"
                    Me.txtNom_Cliente.Text = Trim(DevuelveCampo(strSQL, cConnect))
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim Rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.sQuery = "SELECT Abr_Cliente AS 'Código', Nom_Cliente as 'Descripción' FROM tx_cliente WHERE Nom_Cliente LIKE '%" & Trim(Me.txtNom_Cliente.Text) & "%'"
                    Else
                        oTipo.sQuery = "SELECT Abr_Cliente AS 'Código', Nom_Cliente as 'Descripción' FROM tx_cliente ORDER BY Abr_Cliente "
                    End If
                    
                    oTipo.CARGAR_DATOS
                    oTipo.Show 1
                    If Codigo <> "" Then
                        Me.txtAbr_Cliente.Text = Trim(Codigo)
                        Me.txtNom_Cliente.Text = Trim(Descripcion)
                        Codigo = "": Descripcion = ""
                    End If
                    Set oTipo = Nothing
                    Set Rs = Nothing
                    
    End Select
    txtReferencia.SetFocus
End Sub


Sub CARGA_GRID()
    
    'Esta cadena es para devolver el Codigo de Cliente
    strSQL = "EXEC UP_SEL_LGMOVISTKGUI " & CStr(varOpcion) & ",'" & Right(Me.cboCod_Almacen, 2) & "','" & Me.dtpFecMovStk.Value & "','" & Trim(Me.txtNumMovStk.Text) & "','" & Trim(Me.txtNumGuia.Text) & "'"
    
    Set gexLista.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    
    SetGeneralGridEX gexLista, 0, 1
    
    If Me.gexLista.RowCount = 0 Then
        varBusqueda = ""
    End If
    
    Call Me.gexLista.Find(2, jgexEqual, varBusqueda)
    
    Call Configurar_Grid
    
    If Me.gexLista.RowCount > 0 Then
        gexLista.Enabled = True
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Call CARGA_DATOS
    Else
        gexLista.Enabled = False
        HabilitaMant Me.MantFunc1, "ADICIONAR"
        Call LIMPIAR_DATOS
    End If
    
    Call Me.INHABILITA_DATOS
    
End Sub

Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    Dim Rs As New ADODB.Recordset
    On Error GoTo Salvar_DatosErr
    Dim strSQL As String
    
    Con.ConnectionString = cConnect
    Con.Open
    
        Con.BeginTrans
    
        strSQL = "EXEC UP_MAN_LGMOVISTKGUI '" & _
        sTipo & "','" & _
        Right(Me.cboCod_Almacen, 2) & "','" & _
        Trim(Me.txtNum_MovStk.Text) & "','" & _
        Me.dtpFec_MovStk.Value & "','" & _
        vusu & "','" & _
        IIf(Me.optProveedor.Value, Trim(Me.txtCod_Proveedor), "") & "'," & _
        "0" & ",'" & _
        IIf(Me.optCliente.Value, Me.varCod_Cliente_Tex, "") & "','" & _
        Trim(Me.txtNum_Guia.Text) & "','" & _
        IIf(Me.optOtros.Value, Trim(Me.txtDestinatario), "") & "','" & _
        Trim(Me.txtDom_Destinatario) & "','" & _
        Trim(Me.txtRuc_Destinatario) & "','" & _
        Trim(Me.txtLinea1) & "','" & _
        Trim(Me.txtLinea2) & "','" & _
        Trim(Me.txtReferencia.Text) & "','" & _
        Trim(Me.txtPedido.Text) & "','" & _
        Trim(Me.txtOrden_Compra.Text) & "'"
        
        Con.Execute strSQL
       
        Con.CommitTrans
        
'        Dim amensaje As New clsMessages
'        amensaje.Codigo = CodeMsg.kMSG_INF_DATA_SAVE
'        Informa "", amensaje
'        Mensaje kMSG_INF_DATA_SAVE
        
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler err, "Salvar_Datos"
End Sub

Sub ELIMINAR_DATOS()
 Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
   
    Con.ConnectionString = cConnect
    Con.Open
    Con.BeginTrans
       
        strSQL = "EXEC UP_MAN_LGMOVISTKGUI '" & _
        sTipo & "','" & _
        Right(Me.cboCod_Almacen, 2) & "','" & _
        Trim(Me.txtNum_MovStk.Text) & "','" & _
        Me.dtpFec_MovStk.Value & "','" & _
        vusu & "','" & _
        Trim(Me.txtCod_Proveedor) & "'," & _
        "0" & ",'" & _
        Me.varCod_Cliente_Tex & "','" & _
        Trim(Me.txtNum_Guia.Text) & "','" & _
        Trim(Me.txtDestinatario) & "','" & _
        Trim(Me.txtDom_Destinatario) & "','" & _
        Trim(Me.txtRuc_Destinatario) & "','" & _
        Trim(Me.txtLinea1) & "','" & _
        Trim(Me.txtLinea2) & "','" & _
        Trim(Me.txtReferencia.Text) & "','" & _
        Trim(Me.txtPedido.Text) & "','" & _
        Trim(Me.txtOrden_Compra.Text) & "'"
        
        Con.Execute strSQL
    
    Con.CommitTrans
    
'    Dim amensaje As New clsMessages
'    amensaje.Codigo = CodeMsg.kMSG_INF_DATA_DELETE
'    Informa "", amensaje
'    Mensaje kMSG_INF_DATA_DELETE
    
Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler err, "Eliminar_Datos"

End Sub

Public Sub CARGA_DATOS()
    If Me.gexLista.RowCount > 0 Then
       
        If gexLista.Value(gexLista.Columns("Fec_MovStk").Index) <> "" Then
            Me.dtpFec_MovStk.Value = gexLista.Value(gexLista.Columns("Fec_MovStk").Index)
        End If
        Me.txtNum_MovStk.Text = Trim(gexLista.Value(gexLista.Columns("Num_MovStk").Index))
        Me.txtNum_Guia.Text = Trim(gexLista.Value(gexLista.Columns("Num_Guia").Index))
        Me.txtLinea1.Text = Trim(gexLista.Value(gexLista.Columns("Linea1").Index))
        Me.txtLinea2.Text = Trim(gexLista.Value(gexLista.Columns("Linea2").Index))
        
        Me.txtAbr_Cliente.Text = Trim(gexLista.Value(gexLista.Columns("Abr_Cliente").Index))
        Me.txtNom_Cliente.Text = Trim(gexLista.Value(gexLista.Columns("Nom_Cliente").Index))
        Me.txtCod_Proveedor.Text = Trim(gexLista.Value(gexLista.Columns("Cod_Proveedor").Index))
        Me.txtDes_Proveedor.Text = Trim(gexLista.Value(gexLista.Columns("Des_Proveedor").Index))
        Me.txtDestinatario.Text = Trim(gexLista.Value(gexLista.Columns("Destinatario").Index))
        Me.txtDom_Destinatario.Text = Trim(gexLista.Value(gexLista.Columns("Dom_Destinatario").Index))
        Me.txtRuc_Destinatario.Text = Trim(gexLista.Value(gexLista.Columns("Ruc_Destinatario").Index))
       
        Me.txtReferencia.Text = Trim(gexLista.Value(gexLista.Columns("Referencia").Index))
        Me.txtPedido.Text = Trim(gexLista.Value(gexLista.Columns("Pedido").Index))
        Me.txtOrden_Compra.Text = Trim(gexLista.Value(gexLista.Columns("Orden_Compra").Index))
       
        If Trim(gexLista.Value(gexLista.Columns("Cod_Proveedor").Index)) <> "" Then
            Me.optProveedor.Value = True
        Else
            If Trim(gexLista.Value(gexLista.Columns("Destinatario").Index)) <> "" Then
                Me.optOtros.Value = True
            Else
                Me.optCliente.Value = True
            End If
        End If
        
        'varBusqueda = Trim(gexLista.Value(gexLista.Columns("Num_MovStk").Index))
        
    End If
End Sub

Public Sub CARGA_COMBOS()
    strSQL = "Select a.Nom_Almacen+space(100)+ a.Cod_Almacen from lg_almacen_guias a, lg_segalm_guias b  where a.cod_almacen=b.cod_almacen and b.cod_usuario='" & vusu & "' order by 1"
    Call LlenaCombo(Me.cboCod_Almacen, strSQL, cConnect)
End Sub

Public Sub LIMPIAR_DATOS()
    Me.txtNum_MovStk.Text = ""
    Me.txtAbr_Cliente.Text = ""
    Me.txtNom_Cliente.Text = ""
    Me.txtLinea1.Text = ""
    Me.txtLinea2.Text = ""
    Me.txtCod_Proveedor.Text = ""
    Me.txtDes_Proveedor.Text = ""
    Me.txtDestinatario.Text = ""
    Me.txtDom_Destinatario.Text = ""
    Me.txtRuc_Destinatario.Text = ""
    Me.txtNum_Guia.Text = ""
    
    If optFecha.Value Then
        Me.dtpFec_MovStk.Value = Me.dtpFecMovStk.Value
    Else
        Me.dtpFec_MovStk.Value = Date
    End If
    
    Me.txtReferencia.Text = ""
    Me.txtPedido.Text = ""
    Me.txtOrden_Compra.Text = ""
    
    Me.optProveedor.Value = True
End Sub

Public Sub HABILITA_DATOS()
    Me.txtAbr_Cliente.Enabled = True
    Me.txtNom_Cliente.Enabled = True
    Me.cmdBuscaCliente.Enabled = True
    Me.txtLinea1.Enabled = True
    Me.txtLinea2.Enabled = True
    Me.txtCod_Proveedor.Enabled = True
    Me.txtDes_Proveedor.Enabled = True
    Me.txtDestinatario.Enabled = True
    Me.txtDom_Destinatario.Enabled = True
    Me.txtRuc_Destinatario.Enabled = True
    'Me.txtNum_Guia.Enabled = True
    Me.dtpFec_MovStk.Enabled = True
    
    Me.optCliente.Enabled = True
    Me.optProveedor.Enabled = True
    Me.optOtros.Enabled = True
    
    Me.txtReferencia.Enabled = True
    Me.txtPedido.Enabled = True
    Me.txtOrden_Compra.Enabled = True
    
    Me.fraLista.Enabled = False
    Me.fraFiltro.Enabled = False
    Me.fraOpciones.Enabled = False
    Me.fraFlechas.Enabled = False
    
End Sub

Public Sub INHABILITA_DATOS()
    Me.txtAbr_Cliente.Enabled = False
    Me.txtNom_Cliente.Enabled = False
    Me.cmdBuscaCliente.Enabled = False
    Me.txtLinea1.Enabled = False
    Me.txtLinea2.Enabled = False
    Me.txtCod_Proveedor.Enabled = False
    Me.txtDes_Proveedor.Enabled = False
    Me.txtDestinatario.Enabled = False
    Me.txtDom_Destinatario.Enabled = False
    Me.txtRuc_Destinatario.Enabled = False
    Me.txtNum_Guia.Enabled = False
    Me.dtpFec_MovStk.Enabled = False

    Me.optCliente.Enabled = False
    Me.optProveedor.Enabled = False
    Me.optOtros.Enabled = False
    
    Me.txtReferencia.Enabled = False
    Me.txtPedido.Enabled = False
    Me.txtOrden_Compra.Enabled = False
    
    Me.fraLista.Enabled = True
    Me.fraFiltro.Enabled = True
    Me.fraOpciones.Enabled = True
    Me.fraFlechas.Enabled = True
End Sub

Public Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
    If sTipo <> "D" Then
    
        If sTipo = "I" Then
            If Trim(Me.cboCod_Almacen.Text) = "" Then
                VALIDA_DATOS = False
                MsgBox "El Almacen no puede estar vacio. Sirvase verificar", vbInformation, "Mensaje"
                Me.cboCod_Almacen.SetFocus
                Exit Function
            End If
        End If
    
        If Me.optCliente.Value Then
            If Trim(Me.txtAbr_Cliente.Text) = "" Then
                VALIDA_DATOS = False
                MsgBox "El código de cliente no puede estar vacio. Sirvase verificar", vbInformation, "Mensaje"
                Me.txtAbr_Cliente.SetFocus
                Exit Function
            End If
            
            strSQL = "SELECT COUNT(*) FROM  tx_cliente WHERE Abr_Cliente = '" & Trim(Me.txtAbr_Cliente.Text) & "'"
            If DevuelveCampo(strSQL, cConnect) = 0 Then
                VALIDA_DATOS = False
                MsgBox "El código de cliente ingresado no existe. Sirvase verificar", vbInformation, "Mensaje"
                Me.txtAbr_Cliente.SetFocus
                Exit Function
            End If
            
            strSQL = "SELECT Cod_Cliente_Tex FROM  tx_cliente WHERE Abr_Cliente = '" & Trim(Me.txtAbr_Cliente.Text) & "'"
            Me.varCod_Cliente_Tex = DevuelveCampo(strSQL, cConnect)
        Else
            If Me.optOtros.Value Then
                If Trim(Me.txtDestinatario.Text) = "" Then
                    VALIDA_DATOS = False
                    MsgBox "El código de destinatario no puede estar vacio. Sirvase verificar", vbInformation, "Mensaje"
                    Me.txtDestinatario.SetFocus
                    Exit Function
                End If
            
            Else
                If Trim(Me.txtCod_Proveedor.Text) = "" Then
                    VALIDA_DATOS = False
                    MsgBox "El código de proveedor no puede estar vacio. Sirvase verificar", vbInformation, "Mensaje"
                    Me.txtCod_Proveedor.SetFocus
                    Exit Function
                End If
                
                strSQL = "SELECT COUNT(*) FROM  tx_proveedor WHERE Cod_Proveedor = '" & Trim(Me.txtCod_Proveedor.Text) & "'"
                If DevuelveCampo(strSQL, cConnect) = 0 Then
                    VALIDA_DATOS = False
                    MsgBox "El código de proveedor ingresado no existe. Sirvase verificar", vbInformation, "Mensaje"
                    Me.txtCod_Proveedor.SetFocus
                    Exit Function
                End If
            
            End If
        End If
    
    Else
    
        strSQL = "SELECT COUNT(*) FROM LG_MOVISTK_GUI_DET WHERE COD_ALMACEN = '" & gexLista.Value(gexLista.Columns("Cod_Almacen").Index) & "' and Num_MovStk = '" & Trim(Me.txtNum_MovStk.Text) & "'"
        If DevuelveCampo(strSQL, cConnect) > 0 Then
            VALIDA_DATOS = False
            MsgBox "El registro no se puede eliminar por que posee registros relacionados. Sirvase verificar", vbInformation, "Mensaje"
            Exit Function
        End If
        
        
    End If
End Function

Private Sub cmdBuscaCliente_Click()
    Call Me.BUSCA_CLIENTE(3)
End Sub

Private Sub CmdBuscar_Click()
    Call CARGA_GRID
End Sub

Private Sub cmdFirst_Click()
    gexLista.MoveFirst
End Sub

Private Sub cmdLast_Click()
    gexLista.MoveLast
End Sub

Private Sub cmdNext_Click()
    gexLista.MoveNext
End Sub

Private Sub cmdPrevious_Click()
    gexLista.MovePrevious
End Sub

Private Sub Form_Load()
    Call CARGA_COMBOS
    varOpcion = 0
    Call INHABILITA_DATOS
    Call Me.CARGA_GRID
    dtpFecMovStk.Value = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "DETALLE"
                        Load frmMovStocksGuiasDet
                        frmMovStocksGuiasDet.varCOD_ALMACEN = gexLista.Value(gexLista.Columns("Cod_Almacen").Index)
                        frmMovStocksGuiasDet.varNUM_MOVSTK = gexLista.Value(gexLista.Columns("Num_MovStk").Index)
                        'frmMovStocksGuiasDet.varNum_Secuencia = gexLista.Value(gexLista.Columns("Num_Secuencia").Index)
                        
                        If Trim(gexLista.Value(gexLista.Columns("Ser_Guia_Propia").Index) & gexLista.Value(gexLista.Columns("Nro_Guia_Propia").Index)) <> "" Then
                            frmMovStocksGuiasDet.varBloqueado = True
                        Else
                            frmMovStocksGuiasDet.varBloqueado = False
                        End If
                        
                        frmMovStocksGuiasDet.CARGA_GRID
                        frmMovStocksGuiasDet.Show 1
        Case "GUIA"
'                        If Trim(txtNum_Guia.Text) <> "" Then
'                            MsgBox "Guia ya fue impresa", vbInformation, Me.Caption
'                            Exit Sub
'                        End If
                        varBusqueda = Trim(gexLista.Value(gexLista.Columns("Num_MovStk").Index))
                        Call Reporte_Guia
                        Call CARGA_GRID
    End Select
End Sub

Private Sub gexLista_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    Call Me.CARGA_DATOS
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo hand


    Dim eliminar As Integer
    Dim vRow As Long
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            Call LIMPIAR_DATOS
            Call HABILITA_DATOS
            Me.dtpFec_MovStk.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Case "MODIFICAR"
        
'            If (Trim(gexLista.Value(gexLista.Columns("Ser_Guia_Propia").Index) & gexLista.Value(gexLista.Columns("Nro_Guia_Propia").Index)) <> "") And ActionName <> "SALIR" Then
'                MsgBox "No se puede modificar. Los registros estan bloqueados", vbInformation, "Mensaje"
'                Exit Sub
'            End If
            
            varBusqueda = Trim(gexLista.Value(gexLista.Columns("Num_MovStk").Index))
            
            sTipo = "U"
            Call HABILITA_DATOS
            Me.dtpFec_MovStk.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Case "ELIMINAR"
        
            If (Trim(gexLista.Value(gexLista.Columns("Ser_Guia_Propia").Index) & gexLista.Value(gexLista.Columns("Nro_Guia_Propia").Index)) <> "") And ActionName <> "SALIR" Then
                MsgBox "No se puede eliminar. Los registros estan bloqueados", vbInformation, "Mensaje"
                Exit Sub
            End If
        
            eliminar = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Combinación-Detalle")
            If eliminar = vbYes Then
                sTipo = "D"
                If VALIDA_DATOS Then
                    Call ELIMINAR_DATOS
                    varBusqueda = ""
                    Call CARGA_GRID
                    sTipo = ""
                End If
            End If
        Case "GRABAR"
            If VALIDA_DATOS Then
                Call SALVAR_DATOS
                Call CARGA_GRID
                If sTipo = "I" Then
                    Me.gexLista.MoveLast
                End If
                Call INHABILITA_DATOS
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                sTipo = ""
            End If
        Case "DESHACER"
            Call LIMPIAR_DATOS
            Call CARGA_DATOS
            Call INHABILITA_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            sTipo = ""
         Case "SALIR"
            Unload Me
    End Select
Exit Sub
hand:
ErrorHandler err, "MantFunc1_ActionClick"
End Sub

Private Sub optFecha_Click()
    Me.fraFecha.Visible = True
    Me.fraMov.Visible = False
    Me.fraGuia.Visible = False
    
    varOpcion = 0
    
End Sub

Private Sub OptGuia_Click()
    Me.fraFecha.Visible = False
    Me.fraMov.Visible = False
    Me.fraGuia.Visible = True
    
    varOpcion = 2
End Sub

Private Sub optMov_Click()
    Me.fraFecha.Visible = False
    Me.fraMov.Visible = True
    Me.fraGuia.Visible = False
    
    varOpcion = 1
End Sub

Sub Reporte_Guia()
On Error GoTo hand
Dim Rs_Lista As ADODB.Recordset, vResp As String

Dim varMensaje As Integer

varMensaje = MsgBox("¿Es transportada por el mismo?", vbYesNo, "Guia de Remision")

Load frmDatosAdicionales

If varMensaje = vbYes Then
    vResp = "S"
    frmDatosAdicionales.TxtPlaca = Trim(gexLista.Value(gexLista.Columns("Num_Placa").Index))
    If Me.optOtros.Value Then
        With frmDatosAdicionales
            .TxtTransportista = Trim(gexLista.Value(gexLista.Columns("Destinatario").Index))
            .TxtDomicilio = Trim(gexLista.Value(gexLista.Columns("Dom_Destinatario").Index))
            .TxtRuc = Trim(gexLista.Value(gexLista.Columns("Ruc_Destinatario").Index))
            
            .varOpt = "2"
        End With
    Else
        If Me.optCliente.Value Then
            strSQL = "SELECT Nom_Cliente AS 'NOMBRE', Lug_Entrega AS 'DIRECCION', Num_Ruc AS 'RUC' FROM tx_cliente WHERE Cod_Cliente_Tex = '" & Trim(gexLista.Value(gexLista.Columns("Cod_Cliente").Index)) & "'"
            '.CodProveedor = Trim(gexLista.Value(gexLista.Columns("Cod_Cliente_Tex").Index))
        Else
            strSQL = "SELECT Des_Proveedor AS 'NOMBRE', Dom_Proveedor AS 'DIRECCION', Num_Ruc AS 'RUC',* FROM tx_proveedor WHERE Cod_Proveedor = '" & Trim(gexLista.Value(gexLista.Columns("Cod_Proveedor").Index)) & "'"
            '.CodProveedor = Trim(gexLista.Value(gexLista.Columns("Cod_Proveedor").Index))
        End If
        
        Set Rs_Lista = New ADODB.Recordset
        Rs_Lista.Open strSQL, cConnect
        If Not Rs_Lista.EOF Then
            With frmDatosAdicionales
                .TxtTransportista = Rs_Lista("NOMBRE").Value
                .TxtDomicilio = Rs_Lista("DIRECCION").Value
                .TxtRuc = Rs_Lista("RUC").Value
                '.TxtSec_Transportista.Enabled = False
                '.TxtNom_Transportista.Enabled = False
                '.CmdTransportista.Enabled = False
                
                .varOpt = IIf(Me.optProveedor.Value, "0" & Trim(gexLista.Value(gexLista.Columns("Cod_Proveedor").Index)), "1" & Trim(gexLista.Value(gexLista.Columns("Cod_Cliente").Index)))
            End With
        End If
        Rs_Lista.Close
        Set Rs_Lista = Nothing
     End If
Else
    vResp = "N"
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.Open "select * from seguridad..seg_Empresas where cod_empresa='" & vemp1 & "'", cConnect, adOpenStatic
    If Not Rs_Lista.EOF Then
        With frmDatosAdicionales
            .TxtTransportista = Rs_Lista!Des_Empresa
            .TxtDomicilio = Rs_Lista!Direccion
            .TxtRuc = Rs_Lista!Num_Ruc
            'Aqui pasaremos esos datos
            If Me.optOtros.Value Then
                .varOpt = "2"
            Else
                .varOpt = IIf(Me.optProveedor.Value, "0" & Trim(gexLista.Value(gexLista.Columns("Cod_Proveedor").Index)), "1" & Trim(gexLista.Value(gexLista.Columns("Cod_Cliente").Index)))
            End If
           
        End With
    End If
    Rs_Lista.Close

End If


With frmDatosAdicionales
    .CodAlmacen = Trim(gexLista.Value(gexLista.Columns("Cod_Almacen").Index))
    .NumMovStk = Trim(gexLista.Value(gexLista.Columns("Num_MovStk").Index))
    .Ser_OrdComp = ""
    .Cod_OrdComp = Trim(gexLista.Value(gexLista.Columns("Orden_Compra").Index))
    .varPedido = Trim(gexLista.Value(gexLista.Columns("Pedido").Index))
    .varReferencia = Trim(gexLista.Value(gexLista.Columns("Referencia").Index))
    .varMoviStk_Guia = True
    .vRespuesta = vResp
    
    Call LeeNroGuia(gexLista.Value(gexLista.Columns("Num_Guia").Index), .TxtSerie, .TxtNumero)
    .Show 1
End With

Set frmDatosAdicionales = Nothing

Exit Sub
Resume
hand:
    ErrorHandler err, "GeneraReportes"
End Sub

Public Function LeeNroGuia(ByVal varCadena, ByRef TxtSerie As TextBox, ByRef TxtNumero As TextBox) As String
    Dim NroPos As Integer
    Dim NroCarac As Integer
    Dim varSerie As String, varNumero As String
    Dim varResult As String
    
    varCadena = Trim(varCadena)
    
    NroCarac = Len(varCadena)
    NroPos = InStr(1, varCadena, "-")
    
    varResult = ""
    
    If NroCarac <= 0 Then
        TxtSerie.Text = ""
        TxtNumero.Text = ""
        Exit Function
    Else
        If NroPos > 0 Then
            varResult = Mid(varCadena, 1, NroPos - 1) & Right(varCadena, NroCarac - NroPos)
        End If
        If Len(varResult) > 3 Then
            TxtSerie.Text = Mid(varResult, 1, 3)
            TxtNumero.Text = Right(varResult, Len(varResult) - 3)
        Else
            TxtSerie.Text = varResult
            TxtNumero.Text = ""
        End If
    End If
    LeeNroGuia = varResult
End Function

Private Sub optCliente_Click()
    Me.fraCliente.Visible = True
    Me.fraDestinatario.Visible = False
    Me.fraProveedor.Visible = False
    If Me.txtAbr_Cliente.Enabled = True Then
        txtAbr_Cliente.SetFocus
    End If
End Sub

Private Sub optOtros_Click()
    Me.fraCliente.Visible = False
    Me.fraDestinatario.Visible = True
    Me.fraProveedor.Visible = False
    If Me.txtDestinatario.Enabled = True Then
        txtDestinatario.SetFocus
    End If
End Sub

Private Sub optProveedor_Click()
    Me.fraCliente.Visible = False
    Me.fraDestinatario.Visible = False
    Me.fraProveedor.Visible = True
    
    If Me.txtCod_Proveedor.Enabled = True Then
        txtCod_Proveedor.SetFocus
    End If
End Sub



Public Sub Configurar_Grid()
    Me.gexLista.Columns("Cod_Almacen").Visible = False
    Me.gexLista.Columns("Fec_Creacion").Visible = False
    Me.gexLista.Columns("Cod_Usuario").Visible = False
    Me.gexLista.Columns("Cod_Proveedor").Visible = False
    Me.gexLista.Columns("PROVEEDOR").Visible = False
    Me.gexLista.Columns("UltSec").Visible = False
    Me.gexLista.Columns("Cod_Cliente").Visible = False
    Me.gexLista.Columns("Abr_Cliente").Visible = False
    Me.gexLista.Columns("Nom_Cliente").Visible = False
    Me.gexLista.Columns("Dom_Destinatario").Visible = False
    Me.gexLista.Columns("Ruc_Destinatario").Visible = False
    Me.gexLista.Columns("Nom_Transportista").Visible = False
    Me.gexLista.Columns("Dom_Transportista").Visible = False
    Me.gexLista.Columns("Ruc_Transportista").Visible = False
    Me.gexLista.Columns("Num_Placa").Visible = False
    Me.gexLista.Columns("Ser_Guia_Propia").Visible = False
    Me.gexLista.Columns("Nro_Guia_Propia").Visible = False
    Me.gexLista.Columns("cod_mottra").Visible = False

    Me.gexLista.Columns("Referencia").Visible = False
    Me.gexLista.Columns("Pedido").Visible = False
    Me.gexLista.Columns("Orden_Compra").Visible = False


    Me.gexLista.Columns("Fec_MovStk").Caption = "F. Mov"
    Me.gexLista.Columns("Fec_MovStk").Width = 1000
    
    Me.gexLista.Columns("Num_MovStk").Caption = "Nro. Mov"
    Me.gexLista.Columns("Num_MovStk").Width = 1000
    
    Me.gexLista.Columns("Des_Proveedor").Caption = "Proveedor"
    Me.gexLista.Columns("Des_Proveedor").Width = 1600
    Me.gexLista.Columns("CLIENTE").Caption = "Cliente"
    Me.gexLista.Columns("CLIENTE").Width = 1600
    Me.gexLista.Columns("Destinatario").Caption = "Destinatario"
    Me.gexLista.Columns("Destinatario").Width = 1600
    Me.gexLista.Columns("Num_Guia").Caption = "Num. Guia"
    Me.gexLista.Columns("Num_Guia").Width = 1300
    Me.gexLista.Columns("Linea1").Caption = "Obs. 1"
    Me.gexLista.Columns("Linea1").Width = 1200
    Me.gexLista.Columns("Linea2").Caption = "Obs. 2"
    Me.gexLista.Columns("Linea2").Width = 1200
    
    
    Me.gexLista.Columns("Referencia").Caption = "Referencia"
    Me.gexLista.Columns("Referencia").Width = 1200
    Me.gexLista.Columns("Pedido").Caption = "Pedido"
    Me.gexLista.Columns("Pedido").Width = 1200
    Me.gexLista.Columns("Orden_Compra").Caption = "Orden Compra"
    Me.gexLista.Columns("Orden_Compra").Width = 1200
    
End Sub

Private Sub optProveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtCod_Proveedor.SetFocus
    End If
End Sub

Private Sub txtAbr_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Me.BUSCA_CLIENTE(1)
    End If
End Sub

Private Sub TxtCod_Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCod_Proveedor = Right("00000000000" & Trim(txtCod_Proveedor.Text), 12)
        Call Me.BUSCA_PROVEEDOR(1)
    End If
End Sub

Private Sub txtDes_Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Me.BUSCA_PROVEEDOR(2)
    End If
End Sub

Private Sub txtDestinatario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDom_Destinatario.SetFocus
    End If
End Sub

Private Sub txtDom_Destinatario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtRuc_Destinatario.SetFocus
    End If
End Sub

Private Sub txtLinea1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtLinea2.SetFocus
    End If
End Sub

Private Sub txtLinea2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        MantFunc1.SetFocus
    End If
End Sub

Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Me.BUSCA_CLIENTE(2)
    End If
End Sub

Private Sub txtNum_Guia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.optProveedor.SetFocus
    End If
End Sub

Private Sub txtOrden_Compra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtLinea1.SetFocus
    End If
End Sub

Private Sub txtPedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtOrden_Compra.SetFocus
    End If
End Sub

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPedido.SetFocus
    End If
End Sub

Private Sub txtRuc_Destinatario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtReferencia.SetFocus
    End If
End Sub
