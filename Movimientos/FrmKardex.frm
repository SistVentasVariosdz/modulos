VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmKardex 
   Caption         =   "Kardex"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5025
      Left            =   30
      TabIndex        =   4
      Top             =   3240
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   8864
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Resultados de la Busqueda"
      TabPicture(0)   =   "FrmKardex.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame2 
         Height          =   1725
         Left            =   30
         TabIndex        =   7
         Top             =   360
         Width           =   11145
         Begin MSDataGridLib.DataGrid GridResult2 
            Height          =   1275
            Left            =   150
            TabIndex        =   8
            Top             =   270
            Width           =   10905
            _ExtentX        =   19235
            _ExtentY        =   2249
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   17
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Entradas y Stock"
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
      Begin VB.Frame Frame3 
         Height          =   2625
         Left            =   30
         TabIndex        =   5
         Top             =   2160
         Width           =   11145
         Begin MSDataGridLib.DataGrid GridResult3 
            Height          =   2235
            Left            =   150
            TabIndex        =   6
            Top             =   270
            Width           =   10905
            _ExtentX        =   19235
            _ExtentY        =   3942
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   17
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Movimiento de Stock-Items"
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
   End
   Begin VB.Frame Frame1 
      Height          =   3150
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   11205
      Begin VB.CheckBox chkSoloItemConStk 
         Caption         =   "Solo muestra items con Stock"
         Height          =   195
         Left            =   270
         TabIndex        =   29
         Top             =   1020
         Value           =   1  'Checked
         Width           =   3165
      End
      Begin VB.OptionButton optGrupo 
         Caption         =   "Por Grupo"
         Height          =   195
         Left            =   7590
         TabIndex        =   27
         Top             =   225
         Width           =   1065
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Por Item"
         Height          =   225
         Left            =   5670
         TabIndex        =   26
         Top             =   210
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   525
         Left            =   9540
         TabIndex        =   10
         Top             =   1515
         Width           =   1395
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Imprimir"
         Height          =   525
         Left            =   9540
         TabIndex        =   9
         Top             =   2385
         Width           =   1395
      End
      Begin VB.ComboBox CmbAlmacen 
         Height          =   315
         Left            =   915
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   540
         Width           =   2475
      End
      Begin MSDataGridLib.DataGrid GridResult1 
         Height          =   1545
         Left            =   150
         TabIndex        =   2
         Top             =   1500
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   2725
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   17
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin VB.Frame fraItem 
         Height          =   975
         Left            =   3570
         TabIndex        =   11
         Top             =   390
         Width           =   7455
         Begin VB.TextBox TxtDesc 
            Height          =   315
            Left            =   1500
            TabIndex        =   14
            Top             =   420
            Width           =   2865
         End
         Begin VB.TextBox TxtItem 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   570
            MaxLength       =   8
            TabIndex        =   13
            Top             =   420
            Width           =   945
         End
         Begin VB.CommandButton Command2 
            Caption         =   "..."
            Height          =   315
            Left            =   4380
            TabIndex        =   12
            Top             =   420
            Width           =   375
         End
         Begin MSComCtl2.DTPicker DtDesde 
            Height          =   315
            Left            =   5820
            TabIndex        =   15
            Top             =   195
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   71630849
            CurrentDate     =   37273
         End
         Begin MSComCtl2.DTPicker DtHasta 
            Height          =   315
            Left            =   5820
            TabIndex        =   16
            Top             =   555
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   71630849
            CurrentDate     =   37273
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hasta:"
            Height          =   195
            Index           =   1
            Left            =   5160
            TabIndex        =   19
            Top             =   615
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            Height          =   195
            Index           =   0
            Left            =   5160
            TabIndex        =   18
            Top             =   255
            Width           =   510
         End
         Begin VB.Label Etiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Item:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   105
            TabIndex        =   17
            Tag             =   "Hilado :"
            Top             =   480
            Width           =   330
         End
      End
      Begin VB.Frame fraGrupo 
         Height          =   975
         Left            =   3570
         TabIndex        =   20
         Top             =   390
         Width           =   7470
         Begin VB.TextBox txtDes_Item 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2265
            TabIndex        =   28
            Top             =   540
            Width           =   2775
         End
         Begin VB.TextBox txtCod_Item 
            Height          =   315
            Left            =   1020
            TabIndex        =   25
            Top             =   555
            Width           =   1215
         End
         Begin VB.TextBox txtDesGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2265
            TabIndex        =   22
            Top             =   195
            Width           =   2775
         End
         Begin VB.TextBox txtCodGrupoTex 
            Height          =   315
            Left            =   1020
            MaxLength       =   8
            TabIndex        =   21
            Top             =   195
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Item:"
            Height          =   195
            Left            =   225
            TabIndex        =   24
            Top             =   615
            Width           =   510
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Grupo:"
            Height          =   195
            Left            =   240
            TabIndex        =   23
            Top             =   255
            Width           =   480
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Almacen:"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   3
         Top             =   600
         Width           =   660
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   0
      Top             =   0
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Combinacion As String
Dim Color As String
Dim Talla As String
Dim Destino As String
Dim Estilo As String

Dim Filas As Integer
Public Codigo
Public Descripcion
Public Paso As Boolean
Dim Reg As New ADODB.Recordset
Dim Reg2 As New ADODB.Recordset
Dim Reg3 As New ADODB.Recordset

Sub Datos(Accion As String)
Set Reg = Nothing

Reg.CursorLocation = adUseClient

If optGrupo Then
    Reg.Open "UP_Kardex '1','" & Trim(Right(CmbAlmacen, 3)) & "','" & txtCod_Item.Text & "','" & Combinacion & "','" & _
    Color & "','" & Talla & "','" & Destino & "','" & IIf(IsNull(DtDesde.Value), "01/01/1900", DtDesde.Value) & "','" & IIf(IsNull(DtHasta.Value), "01/01/1900", DtHasta.Value) & "','" & Estilo & "','" & Me.txtCodGrupoTex.Text & "','" & IIf(chkSoloItemConStk.Value, "S", "N") & "'", cConnect
Else
   Reg.Open "UP_Kardex '1','" & Trim(Right(CmbAlmacen, 3)) & "','" & TxtItem.Text & "','" & Combinacion & "','" & _
    Color & "','" & Talla & "','" & Destino & "','" & IIf(IsNull(DtDesde.Value), "01/01/1900", DtDesde.Value) & "','" & IIf(IsNull(DtHasta.Value), "01/01/1900", DtHasta.Value) & "','" & Estilo & "','','" & IIf(chkSoloItemConStk.Value, "S", "N") & "'", cConnect
End If


Set GridResult1.DataSource = Reg

GridResult1.Columns("cod_color").Visible = False
GridResult1.Columns("cod_talla").Visible = False
GridResult1.Columns("cod_destino").Visible = False
GridResult1.Columns("cod_comb").Visible = False

End Sub


Sub grilla2()
Set Reg2 = Nothing

Reg2.CursorLocation = adUseClient

If optGrupo Then
    Reg2.Open "UP_Kardex '2','" & Trim(Right(CmbAlmacen, 3)) & "','" & txtCod_Item.Text & "','" & Combinacion & "','" & _
    Color & "','" & Talla & "','" & Destino & "','" & IIf(IsNull(DtDesde.Value), "01/01/1900", DtDesde.Value) & "','" & IIf(IsNull(DtHasta.Value), "01/01/1900", DtHasta.Value) & "','" & Estilo & "','" & Me.txtCodGrupoTex.Text & "','" & IIf(chkSoloItemConStk.Value, "S", "N") & "'", cConnect
Else
    Reg2.Open "UP_Kardex '2','" & Trim(Right(CmbAlmacen, 3)) & "','" & TxtItem.Text & "','" & Combinacion & "','" & _
    Color & "','" & Talla & "','" & Destino & "','" & IIf(IsNull(DtDesde.Value), "01/01/1900", DtDesde.Value) & "','" & IIf(IsNull(DtHasta.Value), "01/01/1900", DtHasta.Value) & "','" & Estilo & "','','" & IIf(chkSoloItemConStk.Value, "S", "N") & "'", cConnect
End If

Set GridResult2.DataSource = Reg2
End Sub

Sub Grilla3()
Set Reg3 = Nothing

Reg3.CursorLocation = adUseClient

If optGrupo Then
    Reg3.Open "UP_Kardex '3','" & Trim(Right(CmbAlmacen, 3)) & "','" & txtCod_Item.Text & "','" & Combinacion & "','" & _
    Color & "','" & Talla & "','" & Destino & "','" & IIf(IsNull(DtDesde.Value), "01/01/1900", DtDesde.Value) & "','" & IIf(IsNull(DtHasta.Value), "01/01/1900", DtHasta.Value) & "','" & Estilo & "','" & Me.txtCodGrupoTex.Text & "','" & IIf(chkSoloItemConStk.Value, "S", "N") & "'", cConnect
Else
    Reg3.Open "UP_Kardex '3','" & Trim(Right(CmbAlmacen, 3)) & "','" & TxtItem.Text & "','" & Combinacion & "','" & _
    Color & "','" & Talla & "','" & Destino & "','" & IIf(IsNull(DtDesde.Value), "01/01/1900", DtDesde.Value) & "','" & IIf(IsNull(DtHasta.Value), "01/01/1900", DtHasta.Value) & "','" & Estilo & "','','" & IIf(chkSoloItemConStk.Value, "S", "N") & "'", cConnect
End If

Set GridResult3.DataSource = Reg3
End Sub

Private Sub CmdBuscar_Click()
On Error GoTo hand
Datos "1"
grilla2
Grilla3
Exit Sub
hand:
ErrorHandler err, "CmdBuscar_Click"
End Sub

Private Sub Command1_Click()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String

    Ruta = vRuta & "\kardex.xlt"
    'Ruta = App.Path & "\kardex.xlt"
'    Usu = "Usuario : " & vusu
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
        
    oo.Run "reporte", Left(CmbAlmacen, 30), Me.TxtItem & "-" & Me.TxtDesc, DtDesde.Value, DtHasta.Value, GridResult1.Columns("combinacion").Text, GridResult1.Columns("color").Text, GridResult1.Columns("COD_TALLA").Text, GridResult1.Columns("destino").Text, Reg2, Reg3 ' cConnect
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub

Private Sub Command2_Click()
    Set frmBusqGeneral.oParent = Me
    frmBusqGeneral.sQuery = "select Cod_Item AS Codigo,des_item as Descripcion from lg_item "
    frmBusqGeneral.Cargar_Datos
    frmBusqGeneral.Show 1
    TxtDesc = Descripcion
    TxtItem = Codigo
    Codigo = ""
    Descripcion = ""

End Sub

Private Sub DtDesde_Click()
If Not IsNull(DtDesde.Value) Then
    DtDesde.Format = dtpShortDate
    DtDesde.Value = Date
Else
    DtDesde.Format = dtpCustom
    DtDesde.CustomFormat = " "
End If

End Sub


Private Sub DtHasta_Click()
If Not IsNull(DtHasta.Value) Then
    DtHasta.Format = dtpShortDate
    DtHasta.Value = Date
Else
    DtHasta.Format = dtpCustom
    DtHasta.CustomFormat = " "
End If

End Sub


Private Sub Form_Load()
On Error GoTo hand
'cCONNECT = "Provider=sqloledb;Server=SERVIDOR;Database=lives;UID=sa;pwd=;"
'cSEGURIDAD = "Provider=sqloledb;Server=servidor;Database=seguridad;UID=sa;pwd=;"

Datos "1"
grilla2
Grilla3
LlenaCombo CmbAlmacen, "Select Nom_Almacen+space(100)+Cod_Almacen from lg_almacen where Tip_Item='I' order by 1", cConnect
FormateaGrid GridResult1
FormateaGrid GridResult2
FormateaGrid GridResult3

DtHasta.Value = Date
DtDesde.Value = (Date - 15)
Exit Sub
hand:
ErrorHandler err, "Form_Load"

End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub GridResult1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo hand
If Reg.RecordCount > 0 And Me.GridResult1.Columns.Count > 2 Then
    Color = Trim(GridResult1.Columns("cod_color").Text)
    Talla = Trim(GridResult1.Columns("cod_talla").Text)
    Destino = Trim(GridResult1.Columns("cod_destino").Text)
    Combinacion = Trim(GridResult1.Columns("cod_comb").Text)
    Estilo = Trim(GridResult1.Columns("Est.Cli.").Text)
    grilla2
    Grilla3
Else
    Color = ""
    Talla = ""
    Destino = ""
    Combinacion = ""
    grilla2
    Grilla3
End If

Exit Sub
hand:
ErrorHandler err, "GridResult1_RowColChange"
End Sub


Private Sub optGrupo_Click()
    If optGrupo Then
        fraItem.Visible = False
        TxtItem.Text = ""
        TxtDesc.Text = ""
        
        fraGrupo.Visible = True
        Me.txtCodGrupoTex.Text = ""
        Me.txtDesGrupo.Text = ""
        Me.txtCod_Item.Text = ""
        Me.txtDes_Item.Text = ""
    End If
End Sub

Private Sub optItem_Click()
    If optItem Then
        fraItem.Visible = True
        TxtItem.Text = ""
        TxtDesc.Text = ""
        
        fraGrupo.Visible = False
        Me.txtCodGrupoTex.Text = ""
        Me.txtDesGrupo.Text = ""
        Me.txtCod_Item.Text = ""
        Me.txtDes_Item.Text = ""
    End If
End Sub

Private Sub txtCod_Item_Change()
    If Trim(Codigo) <> "" Or Not optGrupo Then
        Exit Sub
    End If
    
    Load frmBuscaItem
    Set frmBuscaItem.oParent = Me
    frmBuscaItem.varCod_Grupo = Me.txtCodGrupoTex.Text
    frmBuscaItem.txtCod_GrupoTex = Me.txtCod_Item.Text
    frmBuscaItem.CARGA_GRID
    frmBuscaItem.Show 1
    
    Set frmBuscaItem = Nothing
    
    If Trim(Codigo) <> "" Then
        Me.txtCod_Item.Text = Codigo
        Me.txtDes_Item.Text = Descripcion
    End If
    Codigo = ""
    Descripcion = ""
End Sub

Private Sub txtCodGrupoTex_Change()
    
    If Trim(Codigo) <> "" Or Not optGrupo Then
        Exit Sub
    End If
    
    Load frmBuscaGrupo
    Set frmBuscaGrupo.oParent = Me
    frmBuscaGrupo.varTipo = "0"
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

Private Sub TxtDesc_KeyPress(KeyAscii As Integer)
On Error GoTo hand
Dim Temp As String
If KeyAscii = 13 Then
    
        If DevuelveCampo("select count(*) from lg_item where des_item like'%" & TxtDesc & "%'", cConnect) > 1 Then
            Set frmBusqGeneral.oParent = Me
            frmBusqGeneral.sQuery = "select Cod_Item AS Codigo,des_item as Descripcion from lg_item where des_item like '" & TxtDesc & "%'"
            frmBusqGeneral.Cargar_Datos
            frmBusqGeneral.Show 1
            TxtDesc = Descripcion
            TxtItem = Codigo
            Temp = TxtItem
            
            Codigo = ""
            Descripcion = ""
            
        Else
            TxtItem = DevuelveCampo("select cod_item from lg_item where des_item like'" & TxtDesc & "%'", cConnect)
            Temp = TxtItem
        End If
                
End If
Exit Sub
hand:
ErrorHandler err, "TxtDesc_KeyPress"


End Sub


Private Sub TxtItem_KeyPress(KeyAscii As Integer)
On Error GoTo hand
Dim Temp As String

If KeyAscii = 13 Then
    If Len(Trim(TxtItem.Text)) < 3 Then
        MsgBox "El código a buscar debe tener 3 caracteres como mínimo", vbInformation
        Exit Sub
    End If
    Temp = Trim(DevuelveCampo("Select dbo.uf_devuelvecodigo(6," & IIf(Trim(TxtItem) = "", 0, Mid(TxtItem, 3)) & ")", cConnect))
    Temp = Left(TxtItem, 2) & Temp
    If DevuelveCampo("select count(*) from lg_item where cod_item ='" & Temp & "'", cConnect) > 0 Then
        Me.TxtDesc = DevuelveCampo("select Des_Item from lg_item where cod_item ='" & Temp & "'", cConnect)
        TxtItem = Temp
    Else
        MsgBox "Codigo no existe", vbInformation
        Me.TxtDesc = ""
        Exit Sub
    End If
                
End If
Exit Sub
hand:
ErrorHandler err, "TxtItem_KeyPress"

End Sub


