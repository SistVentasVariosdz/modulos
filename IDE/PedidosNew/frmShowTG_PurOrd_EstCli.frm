VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmShowTG_PurOrd_EstCli 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ESTILOS CLIENTE - GRAFICOS"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   9105
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   5625
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   435
         Left            =   4530
         TabIndex        =   7
         Top             =   8580
         Width           =   1065
      End
      Begin VB.TextBox txtPO 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   660
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         Top             =   150
         Width           =   4845
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "PO"
         Height          =   195
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Value           =   -1  'True
         Width           =   705
      End
      Begin GridEX20.GridEX GEXListaGrafico 
         Height          =   3495
         Left            =   60
         TabIndex        =   2
         Top             =   4890
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   6165
         Version         =   "2.0"
         RecordNavigator =   -1  'True
         RecordNavigatorString=   "Registros:|de"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         GridLineStyle   =   2
         RowHeight       =   30
         MethodHoldFields=   -1  'True
         GroupByBoxVisible=   0   'False
         BackColorHeader =   -2147483626
         HeaderFontName  =   "Verdana"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   6.75
         HeaderFontWeight=   700
         FontName        =   "MS Sans Serif"
         GridLines       =   2
         ColumnHeaderHeight=   465
         IntProp1        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmShowTG_PurOrd_EstCli.frx":0000
         Column(2)       =   "frmShowTG_PurOrd_EstCli.frx":00C8
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmShowTG_PurOrd_EstCli.frx":016C
         FormatStyle(2)  =   "frmShowTG_PurOrd_EstCli.frx":0294
         FormatStyle(3)  =   "frmShowTG_PurOrd_EstCli.frx":0344
         FormatStyle(4)  =   "frmShowTG_PurOrd_EstCli.frx":03F8
         FormatStyle(5)  =   "frmShowTG_PurOrd_EstCli.frx":04D0
         FormatStyle(6)  =   "frmShowTG_PurOrd_EstCli.frx":0588
         FormatStyle(7)  =   "frmShowTG_PurOrd_EstCli.frx":0668
         ImageCount      =   0
         PrinterProperties=   "frmShowTG_PurOrd_EstCli.frx":0774
      End
      Begin GridEX20.GridEX GEXListaEstiloCliente 
         Height          =   3675
         Left            =   60
         TabIndex        =   3
         Top             =   870
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   6482
         Version         =   "2.0"
         RecordNavigator =   -1  'True
         RecordNavigatorString=   "Registros:|de"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         GridLineStyle   =   2
         RowHeight       =   30
         MethodHoldFields=   -1  'True
         GroupByBoxVisible=   0   'False
         BackColorHeader =   -2147483626
         HeaderFontName  =   "Verdana"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   6.75
         HeaderFontWeight=   700
         FontName        =   "MS Sans Serif"
         GridLines       =   2
         ColumnHeaderHeight=   465
         IntProp1        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmShowTG_PurOrd_EstCli.frx":094C
         Column(2)       =   "frmShowTG_PurOrd_EstCli.frx":0A14
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmShowTG_PurOrd_EstCli.frx":0AB8
         FormatStyle(2)  =   "frmShowTG_PurOrd_EstCli.frx":0BE0
         FormatStyle(3)  =   "frmShowTG_PurOrd_EstCli.frx":0C90
         FormatStyle(4)  =   "frmShowTG_PurOrd_EstCli.frx":0D44
         FormatStyle(5)  =   "frmShowTG_PurOrd_EstCli.frx":0E1C
         FormatStyle(6)  =   "frmShowTG_PurOrd_EstCli.frx":0ED4
         FormatStyle(7)  =   "frmShowTG_PurOrd_EstCli.frx":0FB4
         ImageCount      =   0
         PrinterProperties=   "frmShowTG_PurOrd_EstCli.frx":10C0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00CCCCCC&
         Caption         =   "GRÁFICOS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   4650
         Width           =   795
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00CCCCCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00666666&
         Height          =   315
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   4590
         Width           =   5475
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00CCCCCC&
         Caption         =   "ESTILO CLIENTE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   630
         Width           =   1230
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00CCCCCC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00666666&
         Height          =   315
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   570
         Width           =   5475
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00666666&
         Height          =   8445
         Left            =   30
         Top             =   30
         Width           =   5565
      End
   End
End
Attribute VB_Name = "frmShowTG_PurOrd_EstCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public strCodCliente As String

Private strSql       As String

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Public Sub CARGA_GRID_ESTILO_CLIENTE()

    On Error GoTo Err_Buscar
    
    'TG_MUESTRA_LOTEST_GRAFICOS

    strSql = "EXEC TG_MUESTRA_LOTEST_GRAFICOS " & vbNewLine
    strSql = strSql & "@Cod_PurOrd    ='" & txtPO.Text & "'" & vbNewLine
    strSql = strSql & ",@Cod_Cliente   ='" & strCodCliente & "'" & vbNewLine
    Set GEXListaEstiloCliente.ADORecordset = CargarRecordSetDesconectado(strSql, cCONNECT)

    Dim oFrm As New Frm_Toolbar

    oFrm.CambiarContenedor Me, "ESTILOS CLIENTE"
    Set oFrm = Nothing
       
    FORMATO_GRILLA_ESTILO_CLIENTE
    GEXListaEstiloCliente.Row = -1

    Exit Sub

Err_Buscar:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub

Private Sub FORMATO_GRILLA_ESTILO_CLIENTE()

    With GEXListaEstiloCliente

        For n = 1 To .Columns.count
            .Columns.ItemByPosition(n).EditType = jgexEditNone
            .Columns.ItemByPosition(n).HeaderAlignment = jgexAlignCenter
            .Columns.ItemByPosition(n).Caption = UCase(.Columns.ItemByPosition(n).Caption)
            .Columns.ItemByPosition(n).WordWrap = True
            .Columns.ItemByPosition(n).Visible = False
            '.Columns.ItemByPosition(n).Width = Me.Width / 3
        Next

    End With
    
    With GEXListaEstiloCliente.Columns("Cod_EstCli")
        .Caption = "ESTILO CLIENTE"
        .Visible = True
        .Width = 4000
        .TextAlignment = jgexAlignLeft
        .ColPosition = 1
    End With

    With GEXListaEstiloCliente.Columns("SEL")
        .Caption = ""
        .Visible = True
        .Width = 800
        .TextAlignment = jgexAlignCenter
        .CellStyle = "HyperLink"
        .ColPosition = 2
    End With
    
    '   Dim ObjProveedor As GridEX20.JSGroup
    '   Set ObjProveedor = gexLista.Groups.Add(gexLista.Columns("Ruc_Proveedor").Index, jgexSortAscending)
    '
    '   Dim ObjTipoDocumento As GridEX20.JSGroup
    '   Set ObjTipoDocumento = gexLista.Groups.Add(gexLista.Columns("Tipo_Documento").Index, jgexSortAscending)
   
End Sub

Public Sub CARGA_GRID_GRAFICO()

    On Error GoTo Err_Buscar
    
    'TG_MUESTRA_LOTEST_GRAFICOS

    strSql = "EXEC TG_MUESTRA_COLOR_GRAFICOS " & vbNewLine
    strSql = strSql & "@Cod_PurOrd    ='" & txtPO.Text & "'" & vbNewLine
    strSql = strSql & ",@Cod_Cliente   ='" & strCodCliente & "'" & vbNewLine
    strSql = strSql & ",@Cod_EstCli    ='" & GEXListaEstiloCliente.value(GEXListaEstiloCliente.Columns("Cod_EstCli").Index) & "'" & vbNewLine
    Set GEXListaGrafico.ADORecordset = CargarRecordSetDesconectado(strSql, cCONNECT)

    Dim oFrm As New Frm_Toolbar

    oFrm.CambiarContenedor Me, "GRAFICO"
    Set oFrm = Nothing
       
    FORMATO_GRILLA_GRAFICO
    GEXListaGrafico.Row = -1

    Exit Sub

Err_Buscar:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub

Private Sub FORMATO_GRILLA_GRAFICO()

    With GEXListaGrafico

        For n = 1 To .Columns.count
            .Columns.ItemByPosition(n).EditType = jgexEditNone
            .Columns.ItemByPosition(n).HeaderAlignment = jgexAlignCenter
            .Columns.ItemByPosition(n).Caption = UCase(.Columns.ItemByPosition(n).Caption)
            .Columns.ItemByPosition(n).WordWrap = True
            .Columns.ItemByPosition(n).Visible = False
            '  .Columns.ItemByPosition(n).Width = Me.Width / 6
        Next

    End With

    With GEXListaGrafico.Columns("Cod_Grafico")
        .Caption = "GRÁFICO"
        .Visible = True
        .Width = 1000
        .TextAlignment = jgexAlignCenter
        .ColPosition = 1
    End With

    With GEXListaGrafico.Columns("Cod_ColCli")
        .Caption = "COD. COLOR"
        .Visible = True
        .Width = 1000
        .TextAlignment = jgexAlignLeft
        .ColPosition = 2
    End With
    
    With GEXListaGrafico.Columns("Nom_ColCli")
        .Caption = "NOM. COLOR"
        .Visible = True
        .Width = 1500
        .TextAlignment = jgexAlignLeft
        .ColPosition = 3
    End With
    
    With GEXListaGrafico.Columns("Color_Segun_Cliente")
        .Caption = "NOM. COLOR CLIENTE"
        .Visible = True
        .Width = 1500
        .TextAlignment = jgexAlignLeft
        .ColPosition = 4
    End With

    Dim ObjCod_EstCli As GridEX20.JSGroup

    Set ObjCod_EstCli = GEXListaGrafico.Groups.Add(GEXListaGrafico.Columns("Cod_EstCli").Index, jgexSortAscending)
    '
    '   Dim ObjTipoDocumento As GridEX20.JSGroup
    '   Set ObjTipoDocumento = gexLista.Groups.Add(gexLista.Columns("Tipo_Documento").Index, jgexSortAscending)
   
End Sub

Private Sub GEXListaEstiloCliente_Click()

    If (GEXListaEstiloCliente.RowCount > 0) Then

        ' If (GEXListaEstiloCliente.GroupRowLevel(GEXListaEstiloCliente.Row) = 0) Then
        Select Case GEXListaEstiloCliente.col

            Case GEXListaEstiloCliente.Columns("SEL").ColPosition

                With frmShowTG_PurOrd_EstCli_Grafico
                    .txtPO.Text = txtPO
                    .txtEstiloCliente.Text = GEXListaEstiloCliente.value(GEXListaEstiloCliente.Columns("Cod_EstCli").Index)
                    .strCodCliente = strCodCliente
                    .TRANSAC_DATOS_GRAFICO ("C")
                    .Show 1
                    CARGA_GRID_GRAFICO
                End With

                ' Case GEXLista.Columns("VERINFORMACION").ColPosition
        End Select

        ' End If
    End If

End Sub

Private Sub GEXListaEstiloCliente_SelectionChange()
    Call CARGA_GRID_GRAFICO
End Sub
