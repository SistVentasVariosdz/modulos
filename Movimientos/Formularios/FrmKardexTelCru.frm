VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmKardexTelCru 
   Caption         =   "Kardex Tela Cruda"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   12855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Buscar"
      Height          =   525
      Left            =   11235
      TabIndex        =   18
      Top             =   525
      Width           =   1245
   End
   Begin VB.Frame Frame1 
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
      Left            =   105
      TabIndex        =   19
      Top             =   15
      Width           =   12585
      Begin VB.OptionButton Option1 
         Caption         =   "Por OC  Tej."
         Height          =   195
         Left            =   3600
         TabIndex        =   35
         Top             =   240
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton OptLote 
         Caption         =   "Por Lote"
         Height          =   195
         Left            =   7935
         TabIndex        =   27
         Top             =   210
         Width           =   975
      End
      Begin VB.OptionButton OptTela 
         Caption         =   "Por Tela"
         Height          =   195
         Left            =   5745
         TabIndex        =   26
         Top             =   195
         Width           =   915
      End
      Begin VB.ComboBox CmbAlmacen 
         Height          =   315
         Left            =   945
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   510
         Width           =   2355
      End
      Begin VB.Frame Frame2 
         Height          =   660
         Left            =   3465
         TabIndex        =   28
         Top             =   420
         Width           =   5400
         Begin VB.TextBox TxtCod_Tela 
            Height          =   300
            Left            =   750
            TabIndex        =   30
            Top             =   225
            Width           =   915
         End
         Begin VB.TextBox TxtDes_Tela 
            Height          =   300
            Left            =   1650
            TabIndex        =   29
            Top             =   225
            Width           =   2145
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tela:"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   31
            Top             =   285
            Width           =   360
         End
      End
      Begin VB.Frame Frame5 
         Height          =   660
         Left            =   3465
         TabIndex        =   22
         Top             =   420
         Width           =   5400
         Begin VB.TextBox txtCodGrupoTex 
            Height          =   315
            Left            =   750
            MaxLength       =   8
            TabIndex        =   24
            Top             =   210
            Width           =   930
         End
         Begin VB.TextBox txtDesGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1680
            TabIndex        =   25
            Top             =   210
            Width           =   2220
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Grupo:"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   270
            Width           =   480
         End
      End
      Begin VB.Frame Frame3 
         Height          =   660
         Left            =   3465
         TabIndex        =   32
         Top             =   420
         Width           =   5400
         Begin VB.TextBox TxtLote 
            Height          =   300
            Left            =   750
            TabIndex        =   33
            Top             =   225
            Width           =   3060
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Lote:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   34
            Top             =   285
            Width           =   360
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Almacen:"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   21
         Top             =   555
         Width           =   660
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7365
      Left            =   90
      TabIndex        =   0
      Top             =   1275
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   12991
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Resultados de la busqueda"
      TabPicture(0)   =   "FrmKardexTelCru.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Grilla1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "grilla2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame Frame4 
         Height          =   1275
         Left            =   90
         TabIndex        =   2
         Top             =   2130
         Width           =   12315
         Begin VB.TextBox TxtLote2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1050
            TabIndex        =   9
            Top             =   180
            Width           =   2295
         End
         Begin VB.TextBox TxtTela2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4350
            TabIndex        =   8
            Top             =   180
            Width           =   4095
         End
         Begin VB.TextBox TxtTalla 
            Enabled         =   0   'False
            Height          =   315
            Left            =   9930
            TabIndex        =   7
            Top             =   210
            Width           =   2295
         End
         Begin VB.TextBox TxtStock 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   1050
            TabIndex        =   6
            Top             =   510
            Width           =   1395
         End
         Begin VB.TextBox TxtProveedor 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4350
            TabIndex        =   5
            Top             =   510
            Width           =   4125
         End
         Begin VB.TextBox TxtCombinacion 
            Enabled         =   0   'False
            Height          =   315
            Left            =   9930
            TabIndex        =   4
            Top             =   540
            Width           =   2295
         End
         Begin VB.TextBox TxtCalidad 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1050
            TabIndex        =   3
            Top             =   840
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "Lote:"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   16
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label Label3 
            Caption         =   "Tela:"
            Height          =   255
            Index           =   1
            Left            =   3390
            TabIndex        =   15
            Top             =   240
            Width           =   2865
         End
         Begin VB.Label Label3 
            Caption         =   "Medida:"
            Height          =   255
            Index           =   2
            Left            =   8790
            TabIndex        =   14
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label3 
            Caption         =   "Stock Actual:"
            Height          =   255
            Index           =   3
            Left            =   60
            TabIndex        =   13
            Top             =   540
            Width           =   1065
         End
         Begin VB.Label Label3 
            Caption         =   "Proveedor:"
            Height          =   255
            Index           =   4
            Left            =   3390
            TabIndex        =   12
            Top             =   570
            Width           =   2865
         End
         Begin VB.Label Label3 
            Caption         =   "Combinacion:"
            Height          =   255
            Index           =   5
            Left            =   8790
            TabIndex        =   11
            Top             =   570
            Width           =   1065
         End
         Begin VB.Label Label3 
            Caption         =   "Calidad:"
            Height          =   255
            Index           =   6
            Left            =   90
            TabIndex        =   10
            Top             =   870
            Width           =   1065
         End
      End
      Begin MSDataGridLib.DataGrid grilla2 
         Height          =   3705
         Left            =   90
         TabIndex        =   1
         Top             =   3480
         Width           =   12285
         _ExtentX        =   21669
         _ExtentY        =   6535
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
         TabIndex        =   17
         Top             =   360
         Width           =   12315
         _ExtentX        =   21722
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
      Left            =   420
      Top             =   810
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmKardexTelCru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Paso As Boolean
Public CODIGO As String
Public DESCRIPCION As String
Dim Reg As New ADODB.Recordset
Dim Reg2 As New ADODB.Recordset
 
Dim strSQL As String

Sub Buscar1()
On Error GoTo hand

    Set Reg = Nothing
    Reg.CursorLocation = adUseClient

    If OptTela Then
       
        If OptTela Then
            Reg.Open "SM_BUSCA_TELCRU_TELA '" & Right(Me.CmbAlmacen, 2) & "' ,'" & Me.TxtCod_Tela & "'", cConnect
        Else
            Set Me.Grilla1.DataSource = Nothing
            Reg.Open "EXEC UP_SEL_TOTALORDPROREQ_TEXTIL '" & Trim(Me.txtCodGrupoTex.Text) & "',3", cConnect
        End If
        
    Else
        Reg.Open "SM_BUSCA_TELCRU_PORLOTE '" & Right(Me.CmbAlmacen, 2) & "' ,'" & Me.TxtLote & "'", cConnect
    End If

    Set Me.Grilla1.DataSource = Reg

   ' If optGrupo Then
   '     Grilla1.Columns("Cod_Tela").Visible = False
   '     Grilla1.Columns("Des_Tela").Visible = False
   '     Grilla1.Columns("Cod_Comb").Visible = False
   '     Grilla1.Columns("Des_Comb").Visible = False
   '     'Grilla1.Columns("Cod_color").Visible = False
   '     'Grilla1.Columns("Des_Color").Visible = False
   '     Grilla1.Columns("Cod_Medida").Visible = False
   ' Else
        'Grilla1.Columns("cod_color").Visible = False
        Grilla1.Columns("Cod_Comb").Visible = False
        Grilla1.Columns("Cod_tela").Visible = False
        Grilla1.Columns("Cod_Talla").Visible = False
    'End If
    
Exit Sub
hand:
ErrorHandler err, "Buscar1"
End Sub
Sub Buscar2()
On Error GoTo hand
Set Reg2 = Nothing
Reg2.CursorLocation = adUseClient

'If Not optGrupo Then
'    Reg2.Open "SM_BUSCA_MOVIMIENTOS_LOTE_TELCRU '" & Right(Me.CmbAlmacen, 2) & "' ,'" & Reg("lote") & "','" & Reg("proveedor") & "','" & Reg("cod_tela") & "','" & Reg("cod_comb") & "','" & Reg("Cod_Talla") & "','" & Reg("calidad") & "'", cConnect
'Else
    'Reg2.Open "SM_BUSCA_MOVIMIENTOS_LOTE_TELCRU '" & Right(Me.CmbAlmacen, 2) & "' ,'" & "" & "','" & "" & "','" & Reg("cod_tela") & "','" & Reg("cod_comb") & "','" & Reg("cod_talla") & "','" & "" & "','" & Trim(Me.txtCodGrupoTex.Text) & "'", cConnect
    Reg2.Open "SM_BUSCA_MOVIMIENTOS_LOTE_TELCRU '" & Right(Me.CmbAlmacen, 2) & "' ,'" & Trim(Reg("lote")) & "','" & Trim(Reg("proveedor")) & "','" & Reg("cod_tela") & "','" & Reg("cod_comb") & "','" & Reg("Cod_Talla") & "','" & Reg("calidad") & "'", cConnect
'End If

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
LlenaCombo CmbAlmacen, "Select Nom_Almacen+space(100)+ Cod_Almacen from lg_almacen  where tip_item='T' and tip_presentacion='C' order by 1", cConnect
OptTela_Click

Dim oFrm As New Frm_Toolbar
oFrm.CambiarContenedor Me
Set oFrm = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub Grilla1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo hand
'If Reg.RecordCount > 0 And Grilla1.Columns.Count > 2 Then
If Not Reg.EOF And Not Reg.BOF And Grilla1.Columns.Count > 2 Then

 
        Me.TxtLote2 = Grilla1.Columns("lote")
        Me.TxtTela2 = Grilla1.Columns("tela")
        Me.TxtTalla = Grilla1.Columns("Medida")
        Me.TxtStock = Grilla1.Columns("Stock")
        Me.Txtproveedor = Grilla1.Columns("Proveedor")
        Me.TxtCombinacion = Grilla1.Columns("combinacion")
        Me.TxtCalidad = Grilla1.Columns("calidad")

    
    Buscar2
End If
Exit Sub
hand:
ErrorHandler err, "Grilla1_RowColChange"
End Sub


'Private Sub optGrupo_Click()
'If optGrupo Then
'    Frame2.Enabled = False
'    Me.TxtCod_Tela = ""
'    Me.TxtDes_Tela = ""
'    Frame2.Visible = False
'
'    Frame3.Enabled = False
'    TxtLote = ""
'    Frame3.Visible = False
'
'    Frame5.Enabled = True
'    Me.txtCodGrupoTex.Text = ""
'    Me.txtDesGrupo.Text = ""
'    Frame5.Visible = True
'
'End If
'End Sub

Private Sub OptLote_Click()
If OptLote Then
    Frame2.Enabled = False
    Me.TxtCod_Tela = ""
    Me.txtDes_Tela = ""
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
    Me.txtDes_Tela = ""
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
    Me.txtDes_Tela = DevuelveCampo("select des_tela from tx_tela where cod_tela='" & TxtCod_Tela & "'", cConnect)

End If
Exit Sub
hand:
ErrorHandler err, "TxtCod_Tela"

End Sub

Private Sub txtCodGrupoTex_Change()
    If Trim(CODIGO) <> "" Then
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
    frmBusqGeneral.sQuery = "select Cod_tela AS Codigo,des_tela as Descripcion from tx_tela where des_tela like '%" & txtDes_Tela & "%'"
    frmBusqGeneral.Cargar_Datos
    frmBusqGeneral.Show 1
    TxtCod_Tela = CODIGO
    txtDes_Tela = DESCRIPCION
    
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
            txtCodGrupoTex.Text = Trim(CODIGO)
            txtDesGrupo.Text = Trim(DESCRIPCION)
            CODIGO = ""
            DESCRIPCION = ""
        End If
        Set oTipo = Nothing
        Set rs = Nothing
    End If
End Sub
