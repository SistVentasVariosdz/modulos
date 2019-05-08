VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmKardexTelaca 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Kardex - Tela Acabada"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   11520
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   6360
      TabIndex        =   40
      Top             =   7560
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   900
      Custom          =   $"FrmKardexTelaca.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6105
      Left            =   0
      TabIndex        =   13
      Top             =   1275
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10769
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16761024
      TabCaption(0)   =   "Resultados de la busqueda"
      TabPicture(0)   =   "FrmKardexTelaca.frx":0150
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Grilla1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "grilla2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin MSDataGridLib.DataGrid grilla2 
         Height          =   2505
         Left            =   120
         TabIndex        =   30
         Top             =   3390
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   4419
         _Version        =   393216
         BackColor       =   12648384
         HeadLines       =   1
         RowHeight       =   15
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
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFC0C0&
         Height          =   1335
         Left            =   90
         TabIndex        =   15
         Top             =   1995
         Width           =   11205
         Begin VB.TextBox TxtObs 
            Enabled         =   0   'False
            Height          =   315
            Left            =   8760
            TabIndex        =   39
            Top             =   930
            Width           =   2355
         End
         Begin VB.TextBox TxtColor 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4320
            TabIndex        =   31
            Top             =   900
            Width           =   3375
         End
         Begin VB.TextBox TxtCalidad 
            Enabled         =   0   'False
            Height          =   315
            Left            =   990
            TabIndex        =   29
            Top             =   900
            Width           =   1455
         End
         Begin VB.TextBox TxtCombinacion 
            Enabled         =   0   'False
            Height          =   315
            Left            =   8760
            TabIndex        =   28
            Top             =   570
            Width           =   2355
         End
         Begin VB.TextBox TxtProveedor 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4320
            TabIndex        =   27
            Top             =   540
            Width           =   3375
         End
         Begin VB.TextBox TxtStock 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   990
            TabIndex        =   26
            Top             =   540
            Width           =   1455
         End
         Begin VB.TextBox TxtTalla 
            Enabled         =   0   'False
            Height          =   315
            Left            =   8760
            TabIndex        =   25
            Top             =   210
            Width           =   2355
         End
         Begin VB.TextBox TxtTela2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4320
            TabIndex        =   24
            Top             =   180
            Width           =   3375
         End
         Begin VB.TextBox TxtLote2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   990
            TabIndex        =   23
            Top             =   180
            Width           =   2415
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Observaciones:"
            Height          =   255
            Left            =   7710
            TabIndex        =   38
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Color:"
            Height          =   255
            Index           =   7
            Left            =   3510
            TabIndex        =   32
            Top             =   930
            Width           =   1065
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Calidad:"
            Height          =   255
            Index           =   6
            Left            =   60
            TabIndex        =   22
            Top             =   930
            Width           =   1065
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Combinacion:"
            Height          =   255
            Index           =   5
            Left            =   7710
            TabIndex        =   21
            Top             =   600
            Width           =   1065
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Proveedor:"
            Height          =   255
            Index           =   4
            Left            =   3510
            TabIndex        =   20
            Top             =   570
            Width           =   1065
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Stock Actual:"
            Height          =   255
            Index           =   3
            Left            =   60
            TabIndex        =   19
            Top             =   600
            Width           =   1065
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Medida:"
            Height          =   255
            Index           =   2
            Left            =   7710
            TabIndex        =   18
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Tela:"
            Height          =   255
            Index           =   1
            Left            =   3510
            TabIndex        =   17
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Lote:"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   16
            Top             =   210
            Width           =   1065
         End
      End
      Begin MSDataGridLib.DataGrid Grilla1 
         Height          =   1635
         Left            =   90
         TabIndex        =   14
         Top             =   360
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   2884
         _Version        =   393216
         BackColor       =   12648384
         HeadLines       =   1
         RowHeight       =   15
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
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Buscar"
      Height          =   525
      Left            =   10080
      TabIndex        =   12
      Top             =   375
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Conusulta de Partidas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   11400
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   3480
         TabIndex        =   41
         Top             =   120
         Width           =   2055
      End
      Begin VB.OptionButton optGrupo 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Por Grupo"
         Height          =   195
         Left            =   4320
         TabIndex        =   33
         Top             =   195
         Width           =   1080
      End
      Begin VB.OptionButton OptTela 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Por Tela"
         Height          =   195
         Left            =   5670
         TabIndex        =   6
         Top             =   195
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton OptLote 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Por Lote"
         Height          =   195
         Left            =   7140
         TabIndex        =   5
         Top             =   210
         Width           =   975
      End
      Begin VB.ComboBox CmbAlmacen 
         Height          =   315
         Left            =   945
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   510
         Width           =   2355
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Height          =   660
         Left            =   3465
         TabIndex        =   3
         Top             =   420
         Width           =   5400
         Begin VB.TextBox TxtDes_Tela 
            Height          =   300
            Left            =   1650
            TabIndex        =   9
            Top             =   225
            Width           =   2145
         End
         Begin VB.TextBox TxtCod_Tela 
            Height          =   300
            Left            =   750
            TabIndex        =   8
            Top             =   225
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Tela:"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   7
            Top             =   285
            Width           =   360
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         Height          =   660
         Left            =   3480
         TabIndex        =   4
         Top             =   420
         Width           =   5400
         Begin VB.TextBox TxtLote 
            Height          =   300
            Left            =   750
            TabIndex        =   11
            Top             =   225
            Width           =   3060
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Lote:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   285
            Width           =   360
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFC0C0&
         Height          =   660
         Left            =   3465
         TabIndex        =   34
         Top             =   420
         Width           =   5400
         Begin VB.TextBox txtDesGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1680
            TabIndex        =   37
            Top             =   225
            Width           =   2220
         End
         Begin VB.TextBox txtCodGrupoTex 
            Height          =   315
            Left            =   750
            MaxLength       =   8
            TabIndex        =   36
            Top             =   225
            Width           =   930
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Grupo:"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   285
            Width           =   480
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Almacen:"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   2
         Top             =   555
         Width           =   660
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   300
      Top             =   810
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmKardexTelaca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Paso As Boolean
Public CODIGO As String
Public DESCRIPCION As String
Dim Reg As ADODB.Recordset
Dim Reg2 As New ADODB.Recordset
Dim strSQL As String
Sub Buscar1()
On Error GoTo hand
Set Reg = Nothing
Set Reg = New ADODB.Recordset
Reg.CursorLocation = adUseClient

If OptTela Or optGrupo Then
    If OptTela Then
        Reg.Open "SM_BUSCA_TELTEN_PORTELA '" & Right(Me.CmbAlmacen, 2) & "' ,'" & Me.TxtCod_Tela & "'", cConnect
    Else
        Set Me.Grilla1.DataSource = Nothing
        ' Reg.Open "EXEC UP_SEL_TOTALORDPROREQ_TEXTIL '" & Trim(Me.txtCodGrupoTex.Text) & "',4", cConnect
        
         Reg.Open "EXEC sm_muestra_telas_grupo_stock '" & Trim(Me.txtCodGrupoTex.Text) & "'", cConnect
        
        
        
    End If
Else
    Reg.Open "SM_BUSCA_TELTEN_PORLOTE '" & Right(Me.CmbAlmacen, 2) & "' ,'" & Me.TxtLote & "'", cConnect
End If

Set Me.Grilla1.DataSource = Reg

If optGrupo Then
    Grilla1.Columns("Cod_Tela").Visible = False
    Grilla1.Columns("Des_Tela").Visible = False
    Grilla1.Columns("Cod_Comb").Visible = False
    Grilla1.Columns("Des_Comb").Visible = False
    Grilla1.Columns("Cod_color").Visible = False
    Grilla1.Columns("Des_Color").Visible = False
    Grilla1.Columns("Cod_Medida").Visible = False
Else
    Grilla1.Columns("cod_color").Visible = False
    Grilla1.Columns("Cod_Comb").Visible = False
    Grilla1.Columns("Cod_tela").Visible = False
    Grilla1.Columns("Cod_Talla").Visible = False
    Grilla1.Columns("Cod_Proveedor").Visible = False
End If
Exit Sub
hand:
ErrorHandler err, "Buscar1"
End Sub
Sub Buscar2()
On Error GoTo hand
Set Reg2 = Nothing
Reg2.CursorLocation = adUseClient

If Not optGrupo Then
    Reg2.Open "SM_BUSCA_MOVIMIENTOS_LOTE_TELTEN '" & Right(Me.CmbAlmacen, 2) & "' ,'" & Reg("lote") & "','" & Reg("Cod_Proveedor") & "','" & Reg("cod_tela") & "','" & Reg("cod_comb") & "','" & Reg("cod_color") & "','" & Reg("cod_talla") & "','" & Reg("calidad") & "','" & "" & "'", cConnect
Else
    Reg2.Open "SM_BUSCA_MOVIMIENTOS_LOTE_TELTEN '" & Right(Me.CmbAlmacen, 2) & "' ,'" & "" & "','" & "" & "','" & Reg("cod_tela") & "','" & Reg("cod_comb") & "','" & Reg("cod_color") & "','" & Reg("cod_talla") & "','" & "" & "','" & Trim(Me.txtCodGrupoTex.Text) & "'", cConnect
End If

Set grilla2.DataSource = Reg2
Exit Sub
hand:
ErrorHandler err, "Buscar2"
End Sub



Private Sub Command1_Click()
Buscar1
End Sub



Private Sub Form_Load()
'cCONNECT = "Provider=sqloledb;Server=SERVIDOR;Database=lives;UID=sa;pwd=;"
'cSEGURIDAD = "Provider=sqloledb;Server=servidor;Database=seguridad;UID=sa;pwd=;"

FormateaGrid Grilla1
FormateaGrid grilla2
Buscar1
'Buscar2
LlenaCombo CmbAlmacen, "Select Nom_Almacen+space(100)+ Cod_Almacen from lg_almacen  where tip_item='T' and tip_presentacion='T' order by 1", cConnect
OptTela_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ORDCORTE"
            If Reg.EOF And Reg.BOF Then Exit Sub
            If Me.OptLote Then
                With frmOrdCorte
                    .sCod_Almacen = Right(Me.CmbAlmacen, 2)
                    .sCod_OrdProv = Grilla1.Columns("lote")
                    .sCod_Tela = Grilla1.Columns("cod_tela")
                    .scod_color = Grilla1.Columns("cod_color")
                    .sCod_Combo = Grilla1.Columns("cod_comb")
                    .sCod_Calidad = Grilla1.Columns("Calidad")
                    .sCod_Medida = Grilla1.Columns("cod_talla")
                    .SCOD_PROVEEDOR = Grilla1.Columns("Cod_Proveedor")
                    .CARGA_GRID
                    .Show 1
                End With
            Else
                MsgBox "Opcion valida solo para busqueda por Lote", vbInformation, "Kardex Tela Acabada"
            End If
        Case "IMPRIMIR"
                Call Reporte
        Case "PARTIDAS"
            FrmPartidasProgramadas.Show 1
        Case "SALIR"
            Unload Me
    End Select
End Sub

Private Sub Grilla1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo hand
'If Reg.RecordCount > 0 And Grilla1.Columns.Count > 2 Then
If Not Reg.EOF And Not Reg.BOF And Grilla1.Columns.Count > 2 Then

    If optGrupo Then
    Me.TxtTela2 = Grilla1.Columns("tela")
    Me.TxtTalla = Grilla1.Columns("medida")
        Me.TxtStock = Grilla1.Columns("Stock")
        Me.TxtCombinacion = Grilla1.Columns("combinacion")
        Me.TxtColor = Grilla1.Columns("color")
    Else
        Me.TxtLote2 = Grilla1.Columns("lote")
        Me.TxtTela2 = Grilla1.Columns("tela")
        Me.TxtTalla = Grilla1.Columns("medida")
        Me.TxtStock = Grilla1.Columns("Stock")
        Me.TxtProveedor = Grilla1.Columns("Proveedor")
        Me.TxtCombinacion = Grilla1.Columns("combinacion")
        Me.TxtCalidad = Grilla1.Columns("calidad")
        Me.TxtColor = Grilla1.Columns("color")
        Me.TxtObs = Grilla1.Columns("Observacion")
    End If
    Buscar2
End If
Exit Sub
hand:
ErrorHandler err, "Grilla1_RowColChange"
End Sub


Private Sub optGrupo_Click()
If optGrupo Then
    Frame2.Visible = False
    Me.TxtCod_Tela = ""
    Me.TxtDes_Tela = ""
    Frame2.Visible = False
    
    Frame3.Enabled = False
    TxtLote = ""
    Frame3.Visible = False
    
    Frame5.Enabled = True
    Me.txtCodGrupoTex.Text = ""
    Me.txtDesGrupo.Text = ""
    Frame5.Visible = True
    
End If
End Sub

Private Sub OptLote_Click()
If OptLote Then
    Frame2.Enabled = False
    Me.TxtCod_Tela = ""
    Me.TxtDes_Tela = ""
    Frame2.Visible = False
    
    Frame3.Enabled = True
    TxtLote = ""
    Frame3.Visible = True
    
    Frame5.Enabled = False
    Me.txtCodGrupoTex.Text = ""
    Me.txtDesGrupo.Text = ""
    Frame5.Visible = False
End If

End Sub

Private Sub OptTela_Click()
If OptTela Then
    Frame2.Enabled = True
    Me.TxtCod_Tela = ""
    Me.TxtDes_Tela = ""
    Frame2.Visible = True
    
    Frame3.Enabled = False
    TxtLote = ""
    Frame3.Visible = False
    
    Frame5.Enabled = False
    Me.txtCodGrupoTex.Text = ""
    Me.txtDesGrupo.Text = ""
    Frame5.Visible = False
    
End If
End Sub


Private Sub TxtCod_Tela_KeyPress(KeyAscii As Integer)
On Error GoTo hand
Dim Temp As String
If KeyAscii = 13 And Len(Trim(TxtCod_Tela)) >= 3 Then
    TxtCod_Tela = Left(TxtCod_Tela, 2) & Trim(DevuelveCampo("Select dbo.uf_devuelvecodigo(6," & Mid(TxtCod_Tela, 3) & ")", cConnect))
    Me.TxtDes_Tela = DevuelveCampo("select des_tela from tx_tela where cod_tela='" & TxtCod_Tela & "'", cConnect)
End If
Exit Sub
hand:
ErrorHandler err, "TxtCod_Tela"

End Sub

Private Sub txtCodGrupoTex_Change()

    If Trim(CODIGO) <> "" Or Not optGrupo Then
        Exit Sub
    End If
    
    Load frmBuscaGrupo
    Set frmBuscaGrupo.oParent = Me
    frmBuscaGrupo.varTipo = "1"
    frmBuscaGrupo.txtCod_GrupoTex = Me.txtCodGrupoTex.Text
    frmBuscaGrupo.CARGA_GRID
    frmBuscaGrupo.Show 1
    
    Set frmBuscaGrupo = Nothing
    
    If Trim(CODIGO) <> "" Then
        Me.txtCodGrupoTex.Text = CODIGO
        Me.txtDesGrupo.Text = DESCRIPCION
    End If
    CODIGO = ""
    DESCRIPCION = ""

End Sub

Private Sub txtCodGrupoTex_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If Trim(txtCodGrupoTex.Text) = "" Then
'            'cmdBusCliente_Click
'        Else
'            Strsql = "SELECT Des_Grupo FROM ES_GRUPOTEX WHERE Cod_GrupoTex ='" & Trim(txtCodGrupoTex.Text) & "'"
'            txtDesGrupo.Text = Trim(DevuelveCampo(Strsql, cCONNECT))
'            'FunctBuscar.SetFocus
'        End If
'    End If
End Sub

Private Sub TxtDes_Tela_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Set frmBusqGeneral.oParent = Me
    frmBusqGeneral.sQuery = "select Cod_tela AS Codigo,des_tela as Descripcion from tx_tela where des_tela like '%" & TxtDes_Tela & "%'"
    frmBusqGeneral.Cargar_Datos
    frmBusqGeneral.Show 1
    TxtCod_Tela = CODIGO
    TxtDes_Tela = DESCRIPCION
    
    CODIGO = ""
    DESCRIPCION = ""
    
End If
End Sub

Private Sub txtDesGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim oTipo As New frmBusqGeneral
        Dim rs As New ADODB.Recordset
        Set oTipo.oParent = Me
        
        strSQL = "SELECT Cod_GrupoTex as Código , Des_Grupo as Descripción FROM ES_GRUPOTEX WHERE Des_Grupo  LIKE '" & Trim(txtDesGrupo.Text) & "%'"
    
        oTipo.sQuery = strSQL
        oTipo.Cargar_Datos
        oTipo.Show 1
        If CODIGO <> "" Then
            Me.txtCodGrupoTex.Text = Trim(CODIGO)
            Me.txtDesGrupo.Text = Trim(DESCRIPCION)
            'FunctBuscar.SetFocus
            CODIGO = ""
            DESCRIPCION = ""
        End If
        Set oTipo = Nothing
        Set rs = Nothing
    End If
End Sub

Private Sub txtLote_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Buscar1
End Sub


Public Sub Reporte()
On Error GoTo ErrorImpresion
    Dim oo As Object
    strSQL = "select ruta_logo from seguridad..seg_empresas where cod_Empresa='" & vemp1 & "'"
    
    Set oo = CreateObject("excel.application")
    'oo.Workbooks.Open App.Path & "\RptMovStockFecha.xlt"
    'oo.Workbooks.Open vRuta & "\RptMovStockFecha.xlt"
    oo.Workbooks.Open vRuta & "\Kardex-Tela.xlt"
    oo.Visible = True
    
    oo.Run "REPORTE", Trim(Me.TxtTela2.Text), Trim(Me.TxtColor.Text), Trim(Me.TxtCalidad.Text), Trim(Me.TxtCombinacion.Text), Trim(Me.TxtTalla.Text), Val(Me.TxtStock.Text), Reg2, DevuelveCampo(strSQL, cConnect), cConnect
    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte  " & err.Description, vbCritical, "Impresion"
End Sub

