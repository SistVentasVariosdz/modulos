VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmKardexHilTen 
   Caption         =   "Hilado Color "
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Buscar"
      Height          =   525
      Left            =   9210
      TabIndex        =   16
      Top             =   600
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Consulta de Partidas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   90
      TabIndex        =   17
      Top             =   90
      Width           =   10590
      Begin VB.OptionButton OptLote 
         Caption         =   "Por Lote"
         Height          =   195
         Left            =   7365
         TabIndex        =   33
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OptTela 
         Caption         =   "Por Hil."
         Height          =   195
         Left            =   5895
         TabIndex        =   32
         Top             =   225
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optGrupo 
         Caption         =   "Por Grupo"
         Height          =   195
         Left            =   4545
         TabIndex        =   31
         Top             =   225
         Width           =   1080
      End
      Begin VB.ComboBox CmbAlmacen 
         Height          =   315
         Left            =   945
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   510
         Width           =   2355
      End
      Begin VB.Frame Frame2 
         Height          =   660
         Left            =   3465
         TabIndex        =   27
         Top             =   420
         Width           =   5400
         Begin VB.TextBox TxtDes_Tela 
            Height          =   300
            Left            =   1650
            TabIndex        =   29
            Top             =   225
            Width           =   2145
         End
         Begin VB.TextBox TxtCod_Tela 
            Height          =   300
            Left            =   750
            TabIndex        =   28
            Top             =   225
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Hilado"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   30
            Top             =   285
            Width           =   450
         End
      End
      Begin VB.Frame Frame3 
         Height          =   660
         Left            =   3465
         TabIndex        =   24
         Top             =   420
         Width           =   5400
         Begin VB.TextBox TxtLote 
            Height          =   300
            Left            =   750
            TabIndex        =   25
            Top             =   225
            Width           =   3060
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Lote:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   285
            Width           =   360
         End
      End
      Begin VB.Frame Frame5 
         Height          =   660
         Left            =   3465
         TabIndex        =   20
         Top             =   420
         Width           =   5400
         Begin VB.TextBox txtDesGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1680
            TabIndex        =   23
            Top             =   240
            Width           =   2220
         End
         Begin VB.TextBox txtCodGrupoTex 
            Height          =   315
            Left            =   750
            MaxLength       =   8
            TabIndex        =   22
            Top             =   240
            Width           =   930
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Grupo:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   300
            Width           =   480
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Almacen:"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   19
         Top             =   555
         Width           =   660
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   60
      TabIndex        =   0
      Top             =   1410
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Resultados de la busqueda"
      TabPicture(0)   =   "FrmKardexHilTen.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Grilla1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "grilla2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame Frame4 
         Height          =   1035
         Left            =   90
         TabIndex        =   2
         Top             =   2130
         Width           =   10425
         Begin VB.TextBox TxtLote2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1050
            TabIndex        =   8
            Top             =   180
            Width           =   1785
         End
         Begin VB.TextBox TxtTela2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4350
            TabIndex        =   7
            Top             =   240
            Width           =   1785
         End
         Begin VB.TextBox TxtStock 
            Enabled         =   0   'False
            Height          =   315
            Left            =   8160
            TabIndex        =   6
            Top             =   150
            Width           =   1785
         End
         Begin VB.TextBox TxtProveedor 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1050
            TabIndex        =   5
            Top             =   540
            Width           =   1785
         End
         Begin VB.TextBox TxtCalidad 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4350
            TabIndex        =   4
            Top             =   600
            Width           =   1785
         End
         Begin VB.TextBox TxtColor 
            Enabled         =   0   'False
            Height          =   315
            Left            =   8160
            TabIndex        =   3
            Top             =   510
            Width           =   1785
         End
         Begin VB.Label Label3 
            Caption         =   "Lote:"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   14
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label Label3 
            Caption         =   "Hilado:"
            Height          =   255
            Index           =   1
            Left            =   3360
            TabIndex        =   13
            Top             =   300
            Width           =   1065
         End
         Begin VB.Label Label3 
            Caption         =   "Stock Actual:"
            Height          =   255
            Index           =   3
            Left            =   7170
            TabIndex        =   12
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label Label3 
            Caption         =   "Proveedor:"
            Height          =   255
            Index           =   4
            Left            =   60
            TabIndex        =   11
            Top             =   570
            Width           =   1065
         End
         Begin VB.Label Label3 
            Caption         =   "Calidad:"
            Height          =   255
            Index           =   6
            Left            =   3360
            TabIndex        =   10
            Top             =   630
            Width           =   1065
         End
         Begin VB.Label Label3 
            Caption         =   "Color:"
            Height          =   255
            Index           =   7
            Left            =   7170
            TabIndex        =   9
            Top             =   540
            Width           =   1065
         End
      End
      Begin MSDataGridLib.DataGrid grilla2 
         Height          =   2295
         Left            =   90
         TabIndex        =   1
         Top             =   3270
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   4048
         _Version        =   393216
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
      Begin MSDataGridLib.DataGrid Grilla1 
         Height          =   1755
         Left            =   90
         TabIndex        =   15
         Top             =   360
         Width           =   10425
         _ExtentX        =   18389
         _ExtentY        =   3096
         _Version        =   393216
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
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   210
      Top             =   450
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmKardexHilTen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Paso As Boolean
Public Codigo As String
Public Descripcion As String
Dim Reg As New ADODB.Recordset
Dim Reg2 As New ADODB.Recordset

Dim Strsql As String

Sub Buscar1()
On Error GoTo hand
Set Reg = Nothing
Reg.CursorLocation = adUseClient

If OptTela Or optGrupo Then
    If OptTela Then
        Reg.Open "SM_BUSCA_HILTEN_PORHILO '" & Right(Me.CmbAlmacen, 2) & "' ,'" & Me.TxtCod_Tela & "'", cConnect
    Else
        Set Me.Grilla1.DataSource = Nothing
        Reg.Open "EXEC UP_SEL_TOTALORDPROREQ_TEXTIL '" & Trim(Me.txtCodGrupoTex.Text) & "',2", cConnect
    End If
Else
    Reg.Open "SM_BUSCA_HILTEN_PORLOTE '" & Right(Me.CmbAlmacen, 2) & "' ,'" & Me.TxtLote & "'", cConnect
End If

Set Me.Grilla1.DataSource = Reg

If optGrupo Then
    Grilla1.Columns("Cod_HilTel").Visible = False
    Grilla1.Columns("Des_hiltel").Visible = False
    Grilla1.Columns("Cod_color").Visible = False
    Grilla1.Columns("Des_Color").Visible = False
Else
    Grilla1.Columns("cod_color").Visible = False
    Grilla1.Columns("cod_hiltel").Visible = False
    Grilla1.Columns("cod_proveedor").Visible = False
End If


Exit Sub
hand:
ErrorHandler Err, "Buscar1"
End Sub
Sub Buscar2()
On Error GoTo hand
Set Reg2 = Nothing
Reg2.CursorLocation = adUseClient

'Reg2.Open "SM_BUSCA_MOVIMIENTOS_LOTE_HILTEN '" & Right(Me.CmbAlmacen, 2) & "' ,'" & Reg("lote") & "','" & Reg("cod_proveedor") & "','" & Reg("cod_hiltel") & "','" & Reg("cod_color") & "','" & Reg("calidad") & "'", cCONNECT

If Not optGrupo Then
    Reg2.Open "SM_BUSCA_MOVIMIENTOS_LOTE_HILTEN '" & Right(Me.CmbAlmacen, 2) & "' ,'" & Reg("lote") & "','" & Reg("cod_proveedor") & "','" & Reg("cod_hiltel") & "','" & Reg("cod_color") & "','" & Reg("calidad") & "'", cConnect
Else
    Reg2.Open "SM_BUSCA_MOVIMIENTOS_LOTE_HILTEN '" & Right(Me.CmbAlmacen, 2) & "' ,'" & "" & "','" & "" & "','" & Reg("cod_hiltel") & "','" & Reg("cod_color") & "','" & "" & "','" & Trim(Me.txtCodGrupoTex.Text) & "'", cConnect
End If

Set grilla2.DataSource = Reg2
Exit Sub
hand:
ErrorHandler Err, "Buscar2"
End Sub





Private Sub Command1_Click()
Buscar1
End Sub

Private Sub Form_Load()
'cCONNECT = "Provider=sqloledb;Server=SERVIDOR;Database=lives;UID=sa;pwd=;"
'cSEGURIDAD = "Provider=sqloledb;Server=servidor;Database=seguridad;UID=sa;pwd=;"

FormateaGrid Grilla1
FormateaGrid grilla2

LlenaCombo CmbAlmacen, "Select Nom_Almacen+space(100)+ Cod_Almacen from lg_almacen  where tip_item='H' and tip_presentacion='T' order by 1", cConnect
OptTela_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub Grilla1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo hand
If Reg.RecordCount > 0 And Grilla1.Columns.Count > 2 Then

    If optGrupo Then
        
    Else
        Me.TxtLote2 = Grilla1.Columns("lote")
        Me.TxtTela2 = Grilla1.Columns("hilado")
    '    Me.TxtTalla = Grilla1.Columns("talla")
        Me.TxtStock = Grilla1.Columns("Stock")
        Me.TxtProveedor = Grilla1.Columns("Proveedor")
    '    Me.TxtCombinacion = Grilla1.Columns("combinacion")
        Me.TxtCalidad = Grilla1.Columns("calidad")
        Me.TxtColor = Grilla1.Columns("color")
    End If
    
    Buscar2
End If
Exit Sub
hand:
ErrorHandler Err, "Grilla1_RowColChange"
End Sub


Private Sub optGrupo_Click()
If optGrupo Then
    Frame2.Visible = False
    Me.TxtCod_Tela = ""
    Me.TxtDes_Tela = ""
    
    Frame3.Visible = False
    TxtLote = ""
    
    Me.Frame5.Visible = True
    Me.txtCodGrupoTex.Text = ""
    Me.txtDesGrupo.Text = ""
    
End If
End Sub

Private Sub OptLote_Click()
If OptLote Then
    Frame2.Visible = False
    Me.TxtCod_Tela = ""
    Me.TxtDes_Tela = ""
    
    Frame3.Visible = True
    TxtLote = ""
    
    Me.Frame5.Visible = False
    Me.txtCodGrupoTex.Text = ""
    Me.txtDesGrupo.Text = ""
    
End If

End Sub

Private Sub OptTela_Click()
If OptTela Then
    Frame2.Visible = True
    Me.TxtCod_Tela = ""
    Me.TxtDes_Tela = ""
    
    Frame3.Visible = False
    TxtLote = ""
    
    Me.Frame5.Visible = False
    Me.txtCodGrupoTex.Text = ""
    Me.txtDesGrupo.Text = ""
    
End If
End Sub


Private Sub TxtCod_Tela_KeyPress(KeyAscii As Integer)
On Error GoTo hand
Dim Temp As String
If KeyAscii = 13 Then
    TxtCod_Tela = Trim(DevuelveCampo("Select dbo.uf_devuelvecodigo(3," & TxtCod_Tela & ")", cConnect))
    Me.TxtDes_Tela = DevuelveCampo("select des_hiltel from it_hilado where cod_hiltel='" & TxtCod_Tela & "'", cConnect)
Else
SoloNumeros ActiveControl, KeyAscii, False, 0, 3
End If
Exit Sub
hand:
ErrorHandler Err, "TxtCod_Tela"

End Sub

Private Sub txtCodGrupoTex_Change()

    If Trim(Codigo) <> "" Or Not optGrupo Then
        Exit Sub
    End If
    
    Load frmBuscaGrupo
    Set frmBuscaGrupo.oParent = Me
    frmBuscaGrupo.varTipo = "1"
    frmBuscaGrupo.txtCod_GrupoTex = Me.txtCodGrupoTex.Text
    frmBuscaGrupo.CARGA_GRID
    frmBuscaGrupo.Show 1
    
    Set frmBuscaGrupo = Nothing
    
    If Trim(Codigo) <> "" Then
        Me.txtCodGrupoTex.Text = Codigo
        Me.txtDesGrupo.Text = Descripcion
    End If
    Codigo = ""
    Descripcion = ""

End Sub

Private Sub TxtDes_Tela_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Set frmBusqGeneral.oParent = Me
    frmBusqGeneral.sQuery = "select cod_hiltel AS Codigo,des_hiltel as Descripcion from it_hilado where des_hiltel like '%" & TxtDes_Tela & "%'"
    frmBusqGeneral.CARGAR_DATOS
    frmBusqGeneral.Show 1
    TxtCod_Tela = Codigo
    TxtDes_Tela = Descripcion
End If
End Sub


Private Sub txtDesGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim oTipo As New frmBusqGeneral
        Dim rs As New ADODB.Recordset
        Set oTipo.oParent = Me
        
        Strsql = "SELECT Cod_GrupoTex as Código , Des_Grupo as Descripción FROM ES_GRUPOTEX WHERE Des_Grupo  LIKE '" & Trim(txtDesGrupo.Text) & "%'"
    
        oTipo.sQuery = Strsql
        oTipo.CARGAR_DATOS
        oTipo.Show 1
        If Codigo <> "" Then
            Me.txtCodGrupoTex.Text = Trim(Codigo)
            Me.txtDesGrupo.Text = Trim(Descripcion)
            'FunctBuscar.SetFocus
            Codigo = ""
            Descripcion = ""
        End If
        Set oTipo = Nothing
        Set rs = Nothing
    End If
End Sub
