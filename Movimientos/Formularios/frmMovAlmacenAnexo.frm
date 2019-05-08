VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMovAlmacenAnexo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generacion Consulta Partida"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   510
      Left            =   2265
      TabIndex        =   29
      Top             =   4530
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   510
      Left            =   4290
      TabIndex        =   30
      Top             =   4530
      Width           =   1455
   End
   Begin VB.Frame fraLista 
      Caption         =   "Lista Colores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   60
      TabIndex        =   11
      Top             =   1065
      Width           =   8025
      Begin GridEX20.GridEX gexLista 
         Height          =   2265
         Left            =   120
         TabIndex        =   12
         Top             =   195
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   3995
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ContScroll      =   -1  'True
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmMovAlmacenAnexo.frx":0000
         Column(2)       =   "frmMovAlmacenAnexo.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmMovAlmacenAnexo.frx":016C
         FormatStyle(2)  =   "frmMovAlmacenAnexo.frx":02A4
         FormatStyle(3)  =   "frmMovAlmacenAnexo.frx":0354
         FormatStyle(4)  =   "frmMovAlmacenAnexo.frx":0408
         FormatStyle(5)  =   "frmMovAlmacenAnexo.frx":04E0
         FormatStyle(6)  =   "frmMovAlmacenAnexo.frx":0598
         ImageCount      =   0
         PrinterProperties=   "frmMovAlmacenAnexo.frx":0678
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
      Height          =   930
      Left            =   45
      TabIndex        =   0
      Top             =   120
      Width           =   8040
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   495
         Left            =   6615
         TabIndex        =   10
         Top             =   240
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
      Begin VB.OptionButton optOrdPro 
         Caption         =   "OP"
         Height          =   150
         Left            =   225
         TabIndex        =   31
         Top             =   285
         Value           =   -1  'True
         Width           =   930
      End
      Begin VB.OptionButton optGrupo 
         Caption         =   "Grupo"
         Height          =   150
         Left            =   225
         TabIndex        =   1
         Top             =   540
         Width           =   885
      End
      Begin VB.Frame fraOP 
         Caption         =   "OP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1605
         TabIndex        =   6
         Top             =   165
         Width           =   4770
         Begin VB.TextBox txtDes_estpro 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1695
            TabIndex        =   9
            Top             =   195
            Width           =   2505
         End
         Begin VB.TextBox txtcod_ordpro 
            Height          =   285
            Left            =   915
            TabIndex        =   8
            Top             =   195
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "OP"
            Height          =   195
            Left            =   495
            TabIndex        =   7
            Top             =   240
            Width           =   225
         End
      End
      Begin VB.Frame fraGrupo 
         Caption         =   "Grupo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1605
         TabIndex        =   2
         Top             =   165
         Visible         =   0   'False
         Width           =   4770
         Begin VB.TextBox txtDes_Grupo 
            Height          =   285
            Left            =   1680
            TabIndex        =   5
            Top             =   195
            Width           =   1950
         End
         Begin VB.TextBox txtCod_GrupoTex 
            Height          =   285
            Left            =   900
            TabIndex        =   4
            Top             =   195
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Grupo :"
            Height          =   195
            Left            =   150
            TabIndex        =   3
            Top             =   270
            Width           =   525
         End
      End
   End
   Begin VB.Frame fraEnvios 
      Caption         =   "Envios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   45
      TabIndex        =   17
      Top             =   3645
      Width           =   8000
      Begin VB.OptionButton opt2doEnvio 
         Caption         =   "2do Envio"
         Height          =   150
         Left            =   195
         TabIndex        =   23
         Top             =   480
         Width           =   1185
      End
      Begin VB.OptionButton opt1erEnvio 
         Caption         =   "1er Envio"
         Height          =   150
         Left            =   195
         TabIndex        =   18
         Top             =   255
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.Frame fra2doEnvio 
         Caption         =   "2do Envio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1620
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   4875
         Begin VB.TextBox txtCod_TipOrdTra2da 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1020
            TabIndex        =   21
            Text            =   "TI"
            Top             =   180
            Width           =   615
         End
         Begin VB.TextBox txtCod_Ordtra2do 
            Height          =   285
            Left            =   1635
            TabIndex        =   22
            Top             =   165
            Width           =   1680
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Partida :"
            Height          =   195
            Left            =   270
            TabIndex        =   20
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.Frame fra1eroEnvio 
         Caption         =   "1er Envio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1620
         TabIndex        =   24
         Top             =   120
         Width           =   4875
         Begin VB.TextBox txtCod_TipOrdTra1er 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1020
            TabIndex        =   26
            Text            =   "TI"
            Top             =   180
            Width           =   615
         End
         Begin VB.CommandButton cmdAddPartida 
            Caption         =   "&Añade Partida"
            Height          =   360
            Left            =   3120
            TabIndex        =   28
            Top             =   150
            Width           =   1290
         End
         Begin VB.TextBox txtCod_Ordtra1er 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1620
            TabIndex        =   27
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Partida :"
            Height          =   195
            Left            =   285
            TabIndex        =   25
            Top             =   255
            Width           =   585
         End
      End
   End
   Begin VB.Frame fraPartidas 
      Caption         =   "Partidas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   45
      TabIndex        =   13
      Top             =   3645
      Width           =   8000
      Begin VB.TextBox txtCod_TipOrdTraPar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1605
         TabIndex        =   15
         Text            =   "TI"
         Top             =   285
         Width           =   615
      End
      Begin VB.TextBox txtCod_OrdtraPar 
         Height          =   285
         Left            =   2235
         TabIndex        =   16
         Top             =   285
         Width           =   1845
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Partida :"
         Height          =   195
         Left            =   315
         TabIndex        =   14
         Top             =   360
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmMovAlmacenAnexo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String

Public varCod_ClaOrdComp As String
Public varCod_Clamov As String
Public varCod_Grupo As String
Public varCod_Fabrica As String
Public varTip_Item As String
Public varServicio As String

Public oParent As Object
Public CODIGO As String, DESCRIPCION As String
Public varSalidaCorrecta As Boolean

Public varCancelar As Boolean 'Es cuando cancelamos en el frmaddtx_ordtra_ordenes

Public Sub BUSCA_GRUPO(Tipo As Integer)
    Select Case Tipo
    Case 1:
        strSQL = "SELECT Des_Grupo as 'Descripción' FROM ES_GRUPOTEX WHERE Cod_GrupoTex = '" & Trim(Me.txtCod_GrupoTex.Text) & "' ORDER BY Cod_GrupoTex"
        txtDes_Grupo.Text = Trim(DevuelveCampo(strSQL, cConnect))
        
        'txtCod_TemCli.SetFocus
    Case 2, 3:
        Dim oTipo As New frmBusqGeneral2
        Dim rs As New ADODB.Recordset
        Set oTipo.oParent = Me
        
        If Tipo = 2 Then
            oTipo.sQuery = "SELECT Cod_GrupoTex as 'Código', Des_Grupo as 'Descripción' FROM ES_GRUPOTEX WHERE Des_Grupo LIKE '%" & Trim(Me.txtDes_Grupo.Text) & "%' ORDER BY Cod_GrupoTex"
        Else
            oTipo.sQuery = "SELECT Cod_GrupoTex as 'Código', Des_Grupo as 'Descripción' FROM ES_GRUPOTEX ORDER BY Cod_GrupoTex"
        End If
        
        oTipo.Cargar_Datos
        oTipo.Show 1
        If CODIGO <> "" Then
            txtCod_GrupoTex.Text = Trim(CODIGO)
            txtDes_Grupo.Text = Trim(DESCRIPCION)
            CODIGO = "": DESCRIPCION = ""
            'txtCod_TemCli.SetFocus
        End If
        Set oTipo = Nothing
        Set rs = Nothing
    End Select
    FunctButt1.SetFocus
End Sub

Public Sub BUSCA_OP(Tipo As Integer)
    Select Case Tipo
        Case 1:
                    strSQL = "SELECT cod_estpro FROM ES_ORDPRO WHERE COD_ORDPRO='" & Trim(Me.txtCod_Ordpro.Text) & "'"
                    strSQL = DevuelveCampo(strSQL, cConnect)
                    strSQL = "SELECT Des_estpro FROM ES_ESTPRO WHERE COD_ESTPRO = '" & strSQL & "'"
                    Me.txtDes_estpro.Text = Trim(DevuelveCampo(strSQL, cConnect))
                    
        Case 2, 3:
'                    Dim oTipo As New frmBusqGeneral2
'                    Dim rs As New ADODB.Recordset
'                    Set oTipo.oParent = Me
'
'                    If Tipo = 2 Then
'                        oTipo.sQuery = "SELECT Abr_Fabrica AS 'Código', Nom_Fabrica as 'Descripción' FROM TG_FABRICA WHERE Nom_Fabrica LIKE '%" & Trim(Me.txtNom_Fabrica.Text) & "%' ORDER BY Abr_Fabrica"
'                    Else
'                        oTipo.sQuery = "SELECT Abr_Fabrica AS 'Código', Nom_Fabrica as 'Descripción' FROM TG_FABRICA ORDER BY "
'                    End If
'
'                    oTipo.CARGAR_DATOS
'                    oTipo.Show 1
'                    If Codigo <> "" Then
'                        Me.txtAbr_Fabrica.Text = Trim(Codigo)
'                        Me.txtNom_Fabrica.Text = Trim(Descripcion)
'                        Codigo = "": Descripcion = ""
'                        'txtCod_TemCli.SetFocus
'                    End If
'                    Set oTipo = Nothing
'                    Set rs = Nothing
                    
    End Select
    FunctButt1.SetFocus
End Sub

Public Sub BUSCA_PAQUETES(Tipo As Integer, ByVal Ubic As Integer)
Dim oTipo As New frmBusqPartidas
Dim rs As New ADODB.Recordset
    
    Set oTipo.oParent = Me
    oTipo.sCod_TipOrdTra = txtCod_TipOrdTra2da
    
    If gexLista.RowCount = 0 Then Exit Sub
    
    If txtCod_TipOrdTra2da = "TJ" Then
        oTipo.sQuery = "EXEC TJ_MUESTRA_OTS_SEGUN_GRUPO_PROVEEDOR_TELA '" & _
        Trim(Me.txtCod_GrupoTex.Text) & "', '" & gexLista.Value(gexLista.Columns _
        ("Cod_Proveedor").Index) & "', '" & gexLista.Value(gexLista.Columns _
        ("Cod_Item").Index) & "', '" & gexLista.Value(gexLista.Columns _
        ("Cod_Comb").Index) & "', '" & gexLista.Value(gexLista.Columns _
        ("Cod_Medida").Index) & "'"
    Else
        oTipo.sQuery = "EXEC sM_TRAE_PARTIDAS_POR_GRUPO_COLOR_PROVEEDOR '" & _
        Trim(Me.txtCod_GrupoTex.Text) & "', '" & gexLista.Value(gexLista.Columns _
        ("Cod_Color").Index) & "', '" & gexLista.Value(gexLista.Columns _
        ("Cod_Proveedor").Index) & "', '" & varCod_Fabrica & "', '" & _
        Trim(Me.txtCod_Ordpro.Text) & "', '" & varTip_Item & "'"
    End If
'    Select Case Tipo
'        Case 1:
'                    oTipo.sQuery = "EXEC UP_SEL_PARTIDASDESPACHADAS '" & Tipo & "','" & Trim(Me.txtCod_GrupoTex.Text) & "','','','" & gexLista.Value(gexLista.Columns("Cod_Color").Index) & "',"
'        Case 2:
'                    oTipo.sQuery = "EXEC UP_SEL_PARTIDASDESPACHADAS '" & Tipo & "','','" & Me.varCod_Fabrica & "','" & Trim(Me.txtcod_ordpro.Text) & "','" & gexLista.Value(gexLista.Columns("Cod_Color").Index) & "'"
'    End Select
    
    oTipo.Cargar_Datos
    
    oTipo.Show 1
    If CODIGO <> "" Then
        
        Select Case Ubic
            Case 1:
                    Me.txtCod_Ordtra1er.Text = Trim(DESCRIPCION)
            Case 2:
                    Me.txtCod_Ordtra2do.Text = Trim(DESCRIPCION)
            Case 3:
                    Me.txtCod_OrdtraPar.Text = Trim(DESCRIPCION)
        End Select
        
        CODIGO = "": DESCRIPCION = ""
    End If
    Set oTipo = Nothing
    Set rs = Nothing

End Sub

Sub CARGA_GRID()
    Me.varCod_Grupo = Me.txtCod_GrupoTex
    
    'Esta cadena es para devolver el Codigo de Cliente
    
    If txtCod_TipOrdTra2da = "TJ" Then
        strSQL = "EXEC TJ_MUESTRA_ORDENES_SERVICIO_GRUPO '" & varCod_Grupo & "'"
    Else
        strSQL = "EXEC UP_SEL_COLORES_LG_ORDCOMPITEMREQ '" & Me.varCod_ClaOrdComp & _
        "','" & varCod_Grupo & "', '" & Me.varCod_Fabrica & "', '" & _
        Me.txtCod_Ordpro.Text & "'"
    End If
    
    Set gexLista.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    
    SetGeneralGridEX gexLista, 0, 1
    
'    If gexLista.RowCount > 0 Then
'        HabilitaMant Me.FunctButt1, "ADICIONAR/MODIFICAR/ELIMINAR/IMPRIMIR/CAMBIOESTADO/BLOQUES/IMPRIMIROPE"
'    Else
'        HabilitaMant Me.FunctButt1, "ADICIONAR"
'    End If
    
    Call Configurar_Grid
    
    'Esto es para la 2da parte
    If Me.varCod_Clamov = "E" Then
        Me.fraPartidas.Visible = True
        Me.fraEnvios.Visible = False
    Else
        Me.fraPartidas.Visible = False
        Me.fraEnvios.Visible = True
    End If

End Sub

Public Sub ANADE_PARTIDA()
    Dim Con As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    On Error GoTo Salvar_DatosErr
    Dim strSQL As String
    
    Con.ConnectionString = cConnect
    Con.Open
    
        Con.BeginTrans
        
        strSQL = "EXEC UP_INSERTA_PARTIDA '" & _
        "TI" & "','" & _
        gexLista.Value(gexLista.Columns("Cod_Proveedor").Index) & "','" & _
        Me.varCod_Grupo & "','" & _
        gexLista.Value(gexLista.Columns("Cod_Color").Index) & "','" & _
        gexLista.Value(gexLista.Columns("Ser_OrdComp").Index) & "','" & _
        gexLista.Value(gexLista.Columns("Cod_OrdComp").Index) & "','" & _
        Me.varCod_Fabrica & "','" & _
        Trim(Me.txtCod_Ordpro.Text) & "'"
        
        Me.txtCod_Ordtra1er.Text = DevuelveCampo(strSQL, cConnect)
        
        'Con.Execute Strsql
        
        oParent.TxtObservaciones = "SERVICIO DE TENIDO : " & gexLista.Value(gexLista.Columns("Des_Color").Index)
        Con.CommitTrans
        
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler err, "Salvar_Datos"
End Sub

Public Function VALIDA_DATOS() As Boolean
Dim vCod_OrdTra As String
    VALIDA_DATOS = True
    
    If Me.gexLista.RowCount = 0 Then
        VALIDA_DATOS = False
        MsgBox "No existe ningún color seleccionado. Sirvase verificar", vbInformation, "Mensaje"
        Exit Function
    End If
    If Me.varCod_Clamov = "E" Then
        vCod_OrdTra = txtCod_OrdtraPar
    Else
        If Me.opt1erEnvio Then
            vCod_OrdTra = txtCod_Ordtra1er
        Else
            vCod_OrdTra = txtCod_Ordtra2do
        End If
    End If
    If txtCod_TipOrdTra2da <> "TJ" Then
        strSQL = "EXEC SM_VALIDA_PARTIDA_EN_SALIDA_TELA_CRUDA '" & _
        Me.varCod_Grupo & "', '" & gexLista.Value(gexLista.Columns _
        ("Cod_Color").Index) & "', '" & gexLista.Value(gexLista.Columns _
        ("Cod_Proveedor").Index) & "', '" & Me.varCod_Fabrica & "', '" & _
        Trim(Me.txtCod_Ordpro.Text) & "', '" & Trim(vCod_OrdTra) & "', '" & varTip_Item & "'"
        If CargarRecordSetDesconectado(strSQL, cConnect).RecordCount = 0 Then
            VALIDA_DATOS = False
            MsgBox "La partida seleccionada no existe. Sirvase verificar", _
            vbInformation, "Mensaje"
        End If
    End If
End Function
''''revisa22222
Private Sub cmdAceptar_Click()
Dim sCartaCol As String, sProcesos As String
    
    If VALIDA_DATOS Then
        
        oParent.Txtproveedor = gexLista.Value(gexLista.Columns("Cod_Proveedor").Index)
        oParent.TxtDetalle = gexLista.Value(gexLista.Columns("Des_Proveedor").Index)
        
        If Me.varCod_Clamov = "E" Then
            oParent.txtCod_TipOrdTra = Me.txtCod_TipOrdTraPar.Text
            oParent.txtCod_OrdTra = Me.txtCod_OrdtraPar.Text
        Else
            If Me.opt1erEnvio Then
                oParent.txtCod_TipOrdTra = Me.txtCod_TipOrdTra1er.Text
                oParent.txtCod_OrdTra = Me.txtCod_Ordtra1er.Text
            Else
                oParent.txtCod_TipOrdTra = Me.txtCod_TipOrdTra2da.Text
                oParent.txtCod_OrdTra = Me.txtCod_Ordtra2do.Text
            End If
        End If
        
        If txtCod_TipOrdTra2da <> "TJ" Then
            'oParent.varCod_color = gexLista.Value(gexLista.Columns("Cod_Color").Index)
            oParent.txtDes_Color = gexLista.Value(gexLista.Columns("Des_Color").Index)
        End If
        
        strSQL = "SELECT dbo.TX_Obtiene_Color_Proveedor_Partida ('" & varTip_Item & "', '" & txtCod_Ordtra2do & "') AS CartaCol"
        sCartaCol = Trim(DevuelveCampo(strSQL, cConnect))
        
        If sCartaCol = "" Then
            strSQL = "SELECT dbo.TX_Obtiene_Color_Proveedor_Partida ('" & varTip_Item & "', '" & txtCod_Ordtra1er & "') AS CartaCol"
            sCartaCol = Trim(DevuelveCampo(strSQL, cConnect))
        End If
        
        strSQL = "SELECT dbo.uf_ProcesosOC('" & _
                 gexLista.Value(gexLista.Columns("Ser_OrdComp").Index) & _
                 "', '" & gexLista.Value(gexLista.Columns("Cod_OrdComp").Index) & "')"
        sProcesos = DevuelveCampo(strSQL, cConnect)
        If sProcesos <> "" Then
            oParent.TxtObservaciones = "SERVICIO DE " & sProcesos
        End If
        
        If txtCod_TipOrdTra2da = "TJ" Then
            oParent.TxtObservaciones = oParent.TxtObservaciones & " " & Trim(gexLista.Value(gexLista _
            .Columns("Cod_Item").Index)) & " - " & Trim(gexLista.Value(gexLista _
            .Columns("Des_Tela").Index)) & " / OT: " & txtCod_Ordtra2do
        Else
            oParent.TxtObservaciones = oParent.TxtObservaciones & Trim(gexLista.Value(gexLista.Columns("Des_Color").Index)) & _
            ", Carta Color : " & sCartaCol & " / OT: " & txtCod_Ordtra2do
        End If
        Call oParent.CARGA_ORDCOMP
        Call BuscaCombo(gexLista.Value(gexLista.Columns("Ser_OrdComp").Index) + "-" + gexLista.Value(gexLista.Columns("Cod_OrdComp").Index), 1, oParent.CmbOrdComp)
        
        varSalidaCorrecta = True
        Unload Me
    End If
End Sub

Private Sub cmdAddPartida_Click()
    If Me.gexLista.RowCount = 0 Then
        MsgBox "No xiste ningun color seleccionado. Sirvase verificar", vbInformation, "Mensaje"
        Exit Sub
    End If
    
    varCancelar = False
    
    'Aqui trataremos en ingresar en el Tx_OrdTraOrdenes
    Load frmAddTX_Ordtra_Ordenes
    
    strSQL = "SELECT Flg_Requerimiento FROM LG_CLAORDCOMP " & _
             "WHERE Cod_ClaOrdComp = '" & varCod_ClaOrdComp & "'"
    frmAddTX_Ordtra_Ordenes.Flg_Requerimiento = DevuelveCampo(strSQL, cConnect)
    frmAddTX_Ordtra_Ordenes.varSer_OrdComp = gexLista.Value(gexLista.Columns("Ser_OrdComp").Index)
    frmAddTX_Ordtra_Ordenes.varCod_OrdComp = gexLista.Value(gexLista.Columns("Cod_OrdComp").Index)
    frmAddTX_Ordtra_Ordenes.CARGA_DATOS
    
    frmAddTX_Ordtra_Ordenes.txtSER_ORDCOMP.Text = gexLista.Value(gexLista.Columns("Ser_OrdComp").Index)
    frmAddTX_Ordtra_Ordenes.txtCOD_ORDCOMP.Text = gexLista.Value(gexLista.Columns("Cod_OrdComp").Index)
    
    Set frmAddTX_Ordtra_Ordenes.oParent = Me
    
    frmAddTX_Ordtra_Ordenes.Show 1
    
    Set frmAddTX_Ordtra_Ordenes = Nothing
    
    If varCancelar = False Then
        'Call Me.ANADE_PARTIDA
        Call cmdAceptar_Click
    Else
        Call cmdCancelar_Click
    End If
End Sub

Private Sub cmdCancelar_Click()
    'Call oParent.MantFunc1_ActionClick(5, 0, "DESHACER")
    Unload Me
End Sub

Private Sub Form_Load()
    varSalidaCorrecta = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If varSalidaCorrecta = False Then
        cmdCancelar_Click
    End If
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Call Me.CARGA_GRID
End Sub

Private Sub opt1erEnvio_Click()
    Me.fra1eroEnvio.Visible = True
    Me.fra2doEnvio.Visible = False
    cmdAddPartida.SetFocus
End Sub

Private Sub opt2doEnvio_Click()
    Me.fra1eroEnvio.Visible = False
    Me.fra2doEnvio.Visible = True
    'Me.txtCod_Ordtra2do.SetFocus
End Sub

Private Sub optGrupo_Click()
    Me.fraGrupo.Visible = True
    Me.fraOP.Visible = False
    Me.txtCod_Ordpro.Text = ""
    SendKeys "{TAB}"
    'txtCod_GrupoTex.SetFocus
End Sub

Private Sub optOrdPro_Click()
    Me.fraGrupo.Visible = False
    Me.fraOP.Visible = True
    'Me.txtAbr_Fabrica.Text = ""
    'Me.txtNom_Fabrica.Text = ""
    txtCod_Ordpro.SetFocus
End Sub

Private Sub txtCod_GrupoTex_Change()
   
    If CODIGO = Me.txtCod_GrupoTex Then
        Exit Sub
    End If
   
    Load frmBuscaGrupo
    Set frmBuscaGrupo.oParent = Me
    frmBuscaGrupo.varTipo = "1"
    frmBuscaGrupo.txtCod_GrupoTex = Me.txtCod_GrupoTex.Text
    frmBuscaGrupo.CARGA_GRID
    frmBuscaGrupo.Show 1
    
    Set frmBuscaGrupo = Nothing
    
    If Trim(CODIGO) <> "" Then
        Me.txtDes_Grupo.Text = DESCRIPCION
        Me.txtCod_GrupoTex.Text = CODIGO
    End If
    CODIGO = ""
    DESCRIPCION = ""
    FunctButt1.SetFocus
End Sub

Private Sub txtCod_GrupoTex_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Me.BUSCA_GRUPO(1)
    End If
End Sub

Private Sub txtcod_ordpro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtCod_Ordpro = Right("00000" & Trim(txtCod_Ordpro.Text), 5)
        Call Me.BUSCA_OP(1)
    End If
End Sub

Private Sub txtCod_Ordtra2do_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_Ordtra2do.Text) = "" Then
            Call Me.BUSCA_PAQUETES(IIf(Me.optGrupo.Value, 1, 2), 2)
        End If
    End If
End Sub

Private Sub txtCod_OrdtraPar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_OrdtraPar.Text) = "" Then
            Call Me.BUSCA_PAQUETES(IIf(Me.optGrupo.Value, 1, 2), 3)
        End If
    End If
End Sub

Private Sub txtDes_Grupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Me.BUSCA_GRUPO(2)
    End If
End Sub

Public Sub Configurar_Grid()
    With gexLista
        If txtCod_TipOrdTra2da = "TJ" Then
            .Columns("SER_ORDCOMP").Width = 450
            .Columns("COD_ORDCOMP").Width = 750
            .Columns("DES_PROVEEDOR").Width = 3000
            .Columns("COD_ITEM").Width = 750
            .Columns("DES_TELA").Width = 3960
            .Columns("COD_COMB").Width = 615
            .Columns("DES_COMB").Width = 720
            .Columns("COD_MEDIDA").Width = 360
            .Columns("DES_MEDIDA").Width = 540
            .Columns("COD_PROVEEDOR").Width = 90
            
            .Columns("SER_ORDCOMP").Caption = "Serie"
            .Columns("COD_ORDCOMP").Caption = "Ord.Comp."
            .Columns("DES_PROVEEDOR").Caption = "Proveedor"
            .Columns("COD_ITEM").Caption = "Item"
            .Columns("DES_TELA").Caption = "Tela"
            .Columns("COD_COMB").Caption = "Cod.Comb"
            .Columns("DES_COMB").Caption = "Comb."
            .Columns("COD_MEDIDA").Caption = "Med"
            .Columns("DES_MEDIDA").Caption = "Desc."
            .Columns("COD_PROVEEDOR").Visible = False
            .Columns("Cod_GrupoTex").Visible = False
        Else
            .Columns("SER_ORDCOMP").Caption = "Serie O/C"
            .Columns("COD_ORDCOMP").Caption = "Nro.O/C"
            .Columns("ORDCOMP").Caption = "Nro.O/C"
            .Columns("COD_COLOR").Caption = "Cod.Color"
            .Columns("DES_COLOR").Caption = "Des.Color"
            .Columns("COLOR").Caption = "Color"
            .Columns("PROVEEDOR").Caption = "Proveedor"
            
            .Columns("SER_ORDCOMP").Visible = False
            .Columns("COD_ORDCOMP").Visible = False
            .Columns("ORDCOMP").Width = 2000
            .Columns("COD_COLOR").Visible = False
            .Columns("DES_COLOR").Visible = False
            .Columns("COLOR").Width = 2000
            .Columns("COD_PROVEEDOR").Visible = False
            .Columns("DES_PROVEEDOR").Visible = False
            .Columns("PROVEEDOR").Width = 2500
            
        End If
    End With
End Sub

