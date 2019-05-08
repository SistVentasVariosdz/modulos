VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMovAlmacenAnexoTemp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generacion Consulta Partida"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
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
      Left            =   30
      TabIndex        =   4
      Top             =   0
      Width           =   8040
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
         TabIndex        =   12
         Top             =   165
         Width           =   4770
         Begin VB.TextBox txtCod_GrupoTex 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   900
            TabIndex        =   14
            Top             =   195
            Width           =   915
         End
         Begin VB.TextBox txtDes_Grupo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   13
            Top             =   195
            Width           =   2820
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Grupo :"
            Height          =   195
            Left            =   150
            TabIndex        =   15
            Top             =   270
            Width           =   525
         End
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
         TabIndex        =   8
         Top             =   165
         Visible         =   0   'False
         Width           =   4770
         Begin VB.TextBox txtcod_ordpro 
            Height          =   285
            Left            =   915
            TabIndex        =   10
            Top             =   195
            Width           =   780
         End
         Begin VB.TextBox txtDes_estpro 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1695
            TabIndex        =   9
            Top             =   195
            Width           =   2505
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "OP"
            Height          =   195
            Left            =   495
            TabIndex        =   11
            Top             =   240
            Width           =   225
         End
      End
      Begin VB.OptionButton optGrupo 
         Caption         =   "Grupo"
         Height          =   150
         Left            =   225
         TabIndex        =   7
         Top             =   540
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optOrdPro 
         Caption         =   "OP"
         Height          =   150
         Left            =   225
         TabIndex        =   6
         Top             =   285
         Width           =   930
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   495
         Left            =   6615
         TabIndex        =   5
         Top             =   240
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
      Left            =   45
      TabIndex        =   2
      Top             =   945
      Width           =   8025
      Begin GridEX20.GridEX gexLista 
         Height          =   2265
         Left            =   135
         TabIndex        =   3
         Top             =   195
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   3995
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
         Column(1)       =   "frmMovAlmacenAnexoTemp.frx":0000
         Column(2)       =   "frmMovAlmacenAnexoTemp.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmMovAlmacenAnexoTemp.frx":016C
         FormatStyle(2)  =   "frmMovAlmacenAnexoTemp.frx":02A4
         FormatStyle(3)  =   "frmMovAlmacenAnexoTemp.frx":0354
         FormatStyle(4)  =   "frmMovAlmacenAnexoTemp.frx":0408
         FormatStyle(5)  =   "frmMovAlmacenAnexoTemp.frx":04E0
         FormatStyle(6)  =   "frmMovAlmacenAnexoTemp.frx":0598
         ImageCount      =   0
         PrinterProperties=   "frmMovAlmacenAnexoTemp.frx":0678
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   510
      Left            =   4275
      TabIndex        =   1
      Top             =   4380
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   510
      Left            =   2250
      TabIndex        =   0
      Top             =   4365
      Width           =   1455
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
      Left            =   30
      TabIndex        =   16
      Top             =   3525
      Width           =   8000
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
         TabIndex        =   19
         Top             =   120
         Width           =   4875
         Begin VB.TextBox txtCod_Ordtra1er 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1620
            TabIndex        =   22
            Top             =   180
            Width           =   1020
         End
         Begin VB.CommandButton cmdAddPartida 
            Caption         =   "&Añade Partida"
            Height          =   360
            Left            =   3135
            TabIndex        =   21
            Top             =   135
            Width           =   1290
         End
         Begin VB.TextBox txtCod_TipOrdTra1er 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1020
            TabIndex        =   20
            Text            =   "TI"
            Top             =   180
            Width           =   615
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Partida :"
            Height          =   195
            Left            =   285
            TabIndex        =   23
            Top             =   255
            Width           =   585
         End
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
      Begin VB.OptionButton opt2doEnvio 
         Caption         =   "2do Envio"
         Enabled         =   0   'False
         Height          =   150
         Left            =   195
         TabIndex        =   17
         Top             =   480
         Width           =   1185
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
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   4875
         Begin VB.TextBox txtCod_Ordtra2do 
            Height          =   285
            Left            =   1635
            TabIndex        =   26
            Top             =   165
            Width           =   1680
         End
         Begin VB.TextBox txtCod_TipOrdTra2da 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1020
            TabIndex        =   25
            Text            =   "TI"
            Top             =   180
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Partida :"
            Height          =   195
            Left            =   270
            TabIndex        =   27
            Top             =   240
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frmMovAlmacenAnexoTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Strsql As String

Public varCod_ClaOrdComp As String
Public varCod_Clamov As String
Public varCod_Grupo As String
Public varCod_Fabrica As String

Public oParent As Object
Public Codigo As String, Descripcion As String
Public varSalidaCorrecta As Boolean

Public varCod_Almacen As String
Public varNum_MovStk As String

Public varSer_OrdComp As String
Public varCod_OrdComp As String

Public Sub BUSCA_GRUPO(Tipo As Integer)
    Select Case Tipo
        Case 1:
                    Strsql = "SELECT Des_Grupo as 'Descripción' FROM ES_GRUPOTEX WHERE Cod_GrupoTex = '" & Trim(Me.txtCod_GrupoTex.Text) & "' ORDER BY Cod_GrupoTex"
                    txtDes_Grupo.Text = Trim(DevuelveCampo(Strsql, cConnect))
                    
                    'txtCod_TemCli.SetFocus
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral2
                    Dim Rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.sQuery = "SELECT Cod_GrupoTex as 'Código', Des_Grupo as 'Descripción' FROM ES_GRUPOTEX WHERE Des_Grupo LIKE '%" & Trim(Me.txtDes_Grupo.Text) & "%' ORDER BY Cod_GrupoTex"
                    Else
                        oTipo.sQuery = "SELECT Cod_GrupoTex as 'Código', Des_Grupo as 'Descripción' FROM ES_GRUPOTEX ORDER BY Cod_GrupoTex"
                    End If
                    
                    oTipo.CARGAR_DATOS
                    oTipo.Show 1
                    If Codigo <> "" Then
                        txtCod_GrupoTex.Text = Trim(Codigo)
                        txtDes_Grupo.Text = Trim(Descripcion)
                        Codigo = "": Descripcion = ""
                        'txtCod_TemCli.SetFocus
                    End If
                    Set oTipo = Nothing
                    Set Rs = Nothing
                    
    End Select
    'FunctButt1.SetFocus
End Sub

Public Sub BUSCA_OP(Tipo As Integer)
    Select Case Tipo
        Case 1:
                    Strsql = "SELECT cod_estpro FROM ES_ORDPRO WHERE COD_ORDPRO='" & Trim(Me.txtcod_ordpro.Text) & "' AND Cod_GrupoTex = '" & Me.varCod_Grupo & "'"
                    Strsql = DevuelveCampo(Strsql, cConnect)
                    Strsql = "SELECT Des_estpro FROM ES_ESTPRO WHERE COD_ESTPRO = '" & Strsql & "'"
                    Me.txtDes_estpro.Text = Trim(DevuelveCampo(Strsql, cConnect))
                    
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
    Dim Rs As New ADODB.Recordset
    Set oTipo.oParent = Me

    Select Case Tipo
        Case 1:
                    oTipo.sQuery = "EXEC UP_SEL_PARTIDASDESPACHADAS '" & Tipo & "','" & Trim(Me.txtCod_GrupoTex.Text) & "','','','" & gexLista.Value(gexLista.Columns("Cod_Color").Index) & "'"
        Case 2:
                    oTipo.sQuery = "EXEC UP_SEL_PARTIDASDESPACHADAS '" & Tipo & "','','" & Me.varCod_Fabrica & "','" & Trim(Me.txtcod_ordpro.Text) & "','" & gexLista.Value(gexLista.Columns("Cod_Color").Index) & "'"
    End Select
    
    oTipo.CARGAR_DATOS
    oTipo.Show 1
    If Codigo <> "" Then
        
        Select Case Ubic
            Case 1:
                    Me.txtCod_Ordtra1er.Text = Trim(Descripcion)
            Case 2:
                    Me.txtCod_Ordtra2do.Text = Trim(Descripcion)
        End Select
        
        Codigo = "": Descripcion = ""
    End If
    Set oTipo = Nothing
    Set Rs = Nothing

End Sub


Sub CARGA_GRID()
    
    If Me.optGrupo.Value Then
        Me.varCod_Grupo = Me.txtCod_GrupoTex
    Else
        Strsql = "SELECT COUNT(*) FROM ES_ORDPRO WHERE cod_fabrica = '" & Me.varCod_Fabrica & "' AND cod_ordpro = '" & Me.txtcod_ordpro.Text & "' AND Cod_GrupoTex = '" & Me.varCod_Grupo & "'"
        If DevuelveCampo(Strsql, cConnect) = 0 Then
            MsgBox "La O/P ingresada no pertenece al grupo textil definido. Sirvase verificar", vbInformation, "Mensaje"
            txtcod_ordpro.SetFocus
            'Me.gexLista.Delete
            Exit Sub
        End If
        
    End If
    
    
    'Esta cadena es para devolver el Codigo de Cliente
    Strsql = "EXEC UP_SEL_COLORES_LG_ORDCOMPITEMREQ '" & Me.varCod_ClaOrdComp & "','" & Me.varCod_Grupo & "','" & Me.varSer_OrdComp & "','" & Me.varCod_OrdComp & "'"
    
    Set gexLista.ADORecordset = CargarRecordSetDesconectado(Strsql, cConnect)
    
    SetGeneralGridEX gexLista, 0, 1
    
'    If gexLista.RowCount > 0 Then
'        HabilitaMant Me.FunctButt1, "ADICIONAR/MODIFICAR/ELIMINAR/IMPRIMIR/CAMBIOESTADO/BLOQUES/IMPRIMIROPE"
'    Else
'        HabilitaMant Me.FunctButt1, "ADICIONAR"
'    End If

    Call CONFIGURAR_GRID
    

End Sub

Sub ANADE_PARTIDA()
    Dim Con As New ADODB.Connection
    Dim Rs As New ADODB.Recordset
    On Error GoTo Salvar_DatosErr
    Dim Strsql As String
    
    Con.ConnectionString = cConnect
    Con.Open
    
        Con.BeginTrans

        Strsql = "EXEC UP_INSERTA_PARTIDA_TEMPORAL '" & _
        "TI" & "','" & _
        gexLista.Value(gexLista.Columns("Cod_Proveedor").Index) & "','" & _
        Me.varCod_Grupo & "','" & _
        gexLista.Value(gexLista.Columns("Cod_Color").Index) & "','" & _
        gexLista.Value(gexLista.Columns("Ser_OrdComp").Index) & "','" & _
        gexLista.Value(gexLista.Columns("Cod_OrdComp").Index) & "','" & _
        Me.varCod_Fabrica & "','" & _
        Trim(Me.txtcod_ordpro.Text) & "','" & _
        Me.varCod_Almacen & "','" & _
        Me.varNum_MovStk & "'"
        
        Me.txtCod_Ordtra1er.Text = DevuelveCampo(Strsql, cConnect)
        
        'Con.Execute Strsql
       
        Con.CommitTrans
        
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub

Sub SALVAR_DATOS()
    'Esta funcion se llamara principalmente cuando sea 2da partida
    Dim Con As New ADODB.Connection
    Dim Rs As New ADODB.Recordset
    On Error GoTo Salvar_DatosErr
    Dim Strsql As String
    
    Con.ConnectionString = cConnect
    Con.Open
    
        Con.BeginTrans

        'Strsql = "EXEC UP_INSERTA_PARTIDA '" & _
        "TI" & "','" & _
        gexLista.Value(gexLista.Columns("Cod_Proveedor").Index) & "','" & _
        Me.varCod_Grupo & "','" & _
        gexLista.Value(gexLista.Columns("Cod_Color").Index) & "','" & _
        gexLista.Value(gexLista.Columns("Ser_OrdComp").Index) & "','" & _
        gexLista.Value(gexLista.Columns("Cod_OrdComp").Index) & "','" & _
        Me.varCod_Fabrica & "','" & _
        Trim(Me.txtcod_ordpro.Text) & "'"
        
        Con.Execute Strsql
       
        Con.CommitTrans
        
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub


Public Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True

    If Me.gexLista.RowCount = 0 Then
        VALIDA_DATOS = False
        MsgBox "No existe ningún color seleccionado. Sirvase verificar", vbInformation, "Mensaje"
        Exit Function
    End If
    
    
        
    
        If Me.opt1erEnvio Then
        
            If Trim(Me.txtCod_Ordtra1er.Text) = "" Then
                VALIDA_DATOS = False
                MsgBox "El código de partida no puede estar vacia. Sirvase verificar", vbInformation, "Mensaje"
                Exit Function
            End If
        
            'Si el filtro es por Grupo
            If Me.optGrupo.Value Then
                Strsql = "SELECT COUNT(*) FROM TX_ORDTRA WHERE Flg_Status = 'O' AND Cod_Color = '" & gexLista.Value(gexLista.Columns("Cod_Color").Index) & "' AND Cod_GrupoTex = '" & Me.varCod_Grupo & "' AND Cod_Ordtra = '" & Trim(Me.txtCod_Ordtra1er.Text) & "'"
                If DevuelveCampo(Strsql, cConnect) = 0 Then
                    VALIDA_DATOS = False
                    MsgBox "La partida seleccionada no existe. Sirvase verificar", vbInformation, "Mensaje"
                    Exit Function
                End If
            Else
                'Si el filtro es por Cod_Ordpro
                Strsql = "SELECT COUNT(*) FROM TX_ORDTRA WHERE Flg_Status = 'O' AND Cod_Color = '" & gexLista.Value(gexLista.Columns("Cod_Color").Index) & "' AND cod_fabrica = '" & Me.varCod_Fabrica & "' AND cod_ordpro = '" & Trim(Me.txtcod_ordpro.Text) & "' AND Cod_Ordtra = '" & Trim(Me.txtCod_Ordtra1er.Text) & "'"
                If DevuelveCampo(Strsql, cConnect) = 0 Then
                    VALIDA_DATOS = False
                    MsgBox "La partida seleccionada no existe. Sirvase verificar", vbInformation, "Mensaje"
                    Exit Function
                End If
            End If
            
        Else
        
            If Trim(Me.txtCod_Ordtra2do.Text) = "" Then
                VALIDA_DATOS = False
                MsgBox "El código de partida no puede estar vacio. Sirvase verificar", vbInformation, "Mensaje"
                Exit Function
            End If
        
            'Si el filtro es por Grupo
            If Me.optGrupo.Value Then
                Strsql = "SELECT COUNT(*) FROM TX_ORDTRA WHERE Flg_Status = 'O' AND Cod_Color = '" & gexLista.Value(gexLista.Columns("Cod_Color").Index) & "' AND Cod_GrupoTex = '" & Me.varCod_Grupo & "' AND Cod_Ordtra = '" & Trim(Me.txtCod_Ordtra2do.Text) & "'"
                If DevuelveCampo(Strsql, cConnect) = 0 Then
                    VALIDA_DATOS = False
                    MsgBox "La partida seleccionada no existe. Sirvase verificar", vbInformation, "Mensaje"
                    Exit Function
                End If
            Else
                'Si el filtro es por Cod_Ordpro
                Strsql = "SELECT COUNT(*) FROM TX_ORDTRA WHERE Flg_Status = 'O' AND Cod_Color = '" & gexLista.Value(gexLista.Columns("Cod_Color").Index) & "' AND cod_fabrica = '" & Me.varCod_Fabrica & "' AND cod_ordpro = '" & Trim(Me.txtcod_ordpro.Text) & "' AND Cod_Ordtra = '" & Trim(Me.txtCod_Ordtra2do.Text) & "'"
                If DevuelveCampo(Strsql, cConnect) = 0 Then
                    VALIDA_DATOS = False
                    MsgBox "La partida seleccionada no existe. Sirvase verificar", vbInformation, "Mensaje"
                    Exit Function
                End If
            End If
        
        End If

End Function

Private Sub cmdAceptar_Click()
    If VALIDA_DATOS Then
        'Aqui efectuaremos too la actualizacion de datos
        
       If Me.opt1erEnvio Then
            oParent.txtCod_TipOrdTra = Me.txtCod_TipOrdTra1er.Text
            oParent.txtCod_Ordtra = Me.txtCod_Ordtra1er.Text
        Else
            oParent.txtCod_TipOrdTra = Me.txtCod_TipOrdTra2da.Text
            oParent.txtCod_Ordtra = Me.txtCod_Ordtra2do.Text
        End If
      
        oParent.varCod_color = gexLista.Value(gexLista.Columns("Cod_Color").Index)
        oParent.txtDes_Color = gexLista.Value(gexLista.Columns("Des_Color").Index)
        
        
        If Me.opt2doEnvio.Value Then
            Me.SALVAR_DATOS
        End If
        varSalidaCorrecta = True
        Unload Me
    End If
End Sub

Private Sub cmdAddPartida_Click()
    If Me.gexLista.RowCount = 0 Then
        MsgBox "No xiste ningun color seleccionado. Sirvase verificar", vbInformation, "Mensaje"
    End If
    
    'Validamos si existe el me.txtcod_ordpro
    If Me.optOrdPro Then
        Strsql = "SELECT COUNT(*) FROM ES_ORDPRO WHERE cod_fabrica = '" & Me.varCod_Fabrica & "' AND cod_ordpro = '" & Me.txtcod_ordpro.Text & "' AND Cod_GrupoTex = '" & Me.varCod_Grupo & "'"
        If DevuelveCampo(Strsql, cConnect) = 0 Then
            MsgBox "La O/P ingresada no pertenece al grupo textil definido. Sirvase verificar", vbInformation, "Mensaje"
            txtcod_ordpro.SetFocus
            Exit Sub
        End If
    End If
    
    Call Me.ANADE_PARTIDA
    Call cmdAceptar_Click
End Sub

Private Sub CmdCancelar_Click()
'    Call oParent.MantFunc1_ActionClick(5, 0, "DESHACER")
    Unload Me
End Sub

Private Sub Form_Load()
    varSalidaCorrecta = False
    Call Me.CARGA_GRID
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If varSalidaCorrecta = False Then
        CmdCancelar_Click
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
    Me.txtCod_Ordtra2do.SetFocus
End Sub

Private Sub optGrupo_Click()
    Me.fraGrupo.Visible = True
    Me.fraOP.Visible = False
    Me.txtcod_ordpro.Text = ""
    'txtCod_GrupoTex.SetFocus
End Sub

Private Sub optOrdPro_Click()
    Me.fraGrupo.Visible = False
    Me.fraOP.Visible = True
    'Me.txtAbr_Fabrica.Text = ""
    'Me.txtNom_Fabrica.Text = ""
    txtcod_ordpro.SetFocus
End Sub

Private Sub txtCod_GrupoTex_Change()
   
'    If Codigo = Me.txtCod_GrupoTex Then
'        Exit Sub
'    End If
'
'    Load frmBuscaGrupo
'    Set frmBuscaGrupo.oParent = Me
'    frmBuscaGrupo.varTipo = "1"
'    frmBuscaGrupo.txtCod_GrupoTex = Me.txtCod_GrupoTex.Text
'    frmBuscaGrupo.CARGA_GRID
'    frmBuscaGrupo.Show 1
'
'    Set frmBuscaGrupo = Nothing
'
'    If Trim(Codigo) <> "" Then
'        Me.txtDes_Grupo.Text = Descripcion
'        Me.txtCod_GrupoTex.Text = Codigo
'    End If
'    Codigo = ""
'    Descripcion = ""
'    FunctButt1.SetFocus
End Sub

Private Sub txtCod_GrupoTex_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Me.BUSCA_GRUPO(1)
    End If
End Sub

Private Sub txtcod_ordpro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtcod_ordpro = Right("00000" & Trim(txtcod_ordpro.Text), 5)
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

Private Sub txtDes_Grupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Me.BUSCA_GRUPO(2)
    End If
End Sub

Public Sub CONFIGURAR_GRID()
    Me.gexLista.Columns("COD_PROVEEDOR").Visible = False
    Me.gexLista.Columns("DES_PROVEEDOR").Visible = False
    Me.gexLista.Columns("SER_ORDCOMP").Visible = False
    Me.gexLista.Columns("COD_ORDCOMP").Visible = False
    Me.gexLista.Columns("COD_COLOR").Visible = False
    Me.gexLista.Columns("DES_COLOR").Visible = False
    
    Me.gexLista.Columns("ORDCOMP").Caption = "Orden Compra"
    Me.gexLista.Columns("ORDCOMP").Width = 1200
    Me.gexLista.Columns("COLOR").Caption = "Color"
    Me.gexLista.Columns("COLOR").Width = 2500
    Me.gexLista.Columns("PROVEEDOR").Caption = "Proveedor"
    Me.gexLista.Columns("PROVEEDOR").Width = 3500
    

End Sub



