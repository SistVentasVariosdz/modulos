VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "mantenimientos.ocx"
Begin VB.Form FrmDetalleTelaCa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tela Acabada"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAsigna 
      Caption         =   "&Adiciona O/Corte"
      Height          =   525
      Left            =   6810
      TabIndex        =   38
      Top             =   5925
      Width           =   1335
   End
   Begin VB.CommandButton lote 
      Caption         =   "&Ingresar Lote"
      Height          =   525
      Left            =   8310
      TabIndex        =   25
      Top             =   5910
      Width           =   1245
   End
   Begin VB.Frame Fralista 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   60
      TabIndex        =   23
      Tag             =   "List"
      Top             =   60
      Width           =   9495
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2775
         Left            =   180
         TabIndex        =   24
         Top             =   345
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   4895
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
   End
   Begin VB.Frame Fradetalle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   60
      TabIndex        =   13
      Tag             =   "Detail"
      Top             =   3390
      Width           =   9510
      Begin VB.CommandButton cmdCapturarPeso 
         Caption         =   "Capturar Peso"
         Height          =   300
         Left            =   5535
         TabIndex        =   40
         Top             =   195
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton cmdRollos 
         Caption         =   "Ver &Rollos"
         Height          =   405
         Left            =   8430
         TabIndex        =   39
         Top             =   195
         Width           =   1020
      End
      Begin VB.Frame fraOP 
         Caption         =   "O/P"
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
         Left            =   5460
         TabIndex        =   35
         Top             =   1800
         Width           =   3915
         Begin VB.TextBox txtDes_estpro 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2040
            TabIndex        =   37
            Top             =   150
            Width           =   1680
         End
         Begin VB.TextBox txtcod_ordpro 
            Height          =   285
            Left            =   1320
            MaxLength       =   5
            TabIndex        =   6
            Top             =   150
            Width           =   720
         End
         Begin VB.Label Label6 
            Caption         =   "O/P"
            Height          =   210
            Left            =   120
            TabIndex        =   36
            Top             =   210
            Width           =   405
         End
      End
      Begin VB.TextBox txtCan_Movimiento_2daunimed 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6780
         TabIndex        =   3
         Text            =   "0"
         Top             =   525
         Width           =   945
      End
      Begin VB.TextBox TxtObs 
         Height          =   645
         Left            =   6780
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1170
         Width           =   2385
      End
      Begin VB.TextBox TxtBultos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         Left            =   6780
         TabIndex        =   4
         Text            =   "0"
         Top             =   825
         Width           =   945
      End
      Begin VB.TextBox TxtProveedor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   1140
         TabIndex        =   15
         Top             =   510
         Width           =   3315
      End
      Begin VB.TextBox TxtLote 
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
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   0
         Top             =   180
         Width           =   1935
      End
      Begin VB.TextBox TxtCantidad 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         Left            =   6780
         TabIndex        =   2
         Text            =   "0"
         Top             =   180
         Width           =   945
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
         Left            =   1140
         MaxLength       =   8
         TabIndex        =   1
         Top             =   840
         Width           =   945
      End
      Begin VB.TextBox TxtDesitem 
         Height          =   315
         Left            =   2100
         TabIndex        =   14
         Top             =   840
         Width           =   2355
      End
      Begin VB.Label lblCantidad1 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8010
         TabIndex        =   34
         Top             =   285
         Width           =   75
      End
      Begin VB.Label lblCantidad2 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8010
         TabIndex        =   33
         Top             =   615
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Medida:"
         Height          =   195
         Index           =   5
         Left            =   2370
         TabIndex        =   32
         Top             =   1920
         Width           =   570
      End
      Begin VB.Label Label5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3270
         TabIndex        =   31
         Top             =   1920
         Width           =   1800
      End
      Begin VB.Label Label4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1140
         TabIndex        =   30
         Top             =   1530
         Width           =   3405
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1140
         TabIndex        =   29
         Top             =   1230
         Width           =   3345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comb:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   28
         Top             =   1275
         Width           =   450
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1140
         TabIndex        =   27
         Top             =   1920
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Calidad:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   26
         Top             =   1920
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   9
         Left            =   5520
         TabIndex        =   22
         Top             =   1230
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad Bultos:"
         Height          =   195
         Index           =   8
         Left            =   5520
         TabIndex        =   21
         Top             =   915
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Top             =   585
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lote:"
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   19
         Top             =   255
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tela:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   945
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad Mov:"
         Height          =   195
         Index           =   7
         Left            =   5520
         TabIndex        =   17
         Top             =   255
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Color:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   1575
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   390
      TabIndex        =   8
      Top             =   5940
      Width           =   1965
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "FrmDetalleTelaCa.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Ultimo"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   0
         Picture         =   "FrmDetalleTelaCa.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Primero"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "FrmDetalleTelaCa.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Siguiente"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "FrmDetalleTelaCa.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Anterior"
         Top             =   0
         Width           =   495
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2445
      TabIndex        =   7
      Top             =   5880
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmDetalleTelaCa.frx":05C8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "FrmDetalleTelaCa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CODIGO
Public DESCRIPCION
Public Cod_color As String
Public NewLote As String

'Public lote As String
Public Paso As Boolean
Dim Tip_item As String
Dim Tip_presentacion As String
Public Cod_Calidad As String
Public Cod_Calidad_Envio As String
Dim Cant_Anterior As Double
Public Cod_Comb As String
Dim Des_Comb As String
Public Cod_Talla As String
Public Flg_Rollo As String

'CAMPOS HEREDADOS
Public Sec_OrdComp As String
Public Ser_OrdComp As String
Public Cod_OrdComp As String
Public Fec_MOVsTK As Date
Public Cod_TipMovi As String
Public Cod_ClaOrdComp As String
Public Cod_Almacen As String
Public Num_MovStk As String
Public Cod_OrdPro As String
Public Cod_TipOrdTra As String
Public Cod_OrdTra As String
Public Cod_Proveedor As String
Public Cod_TipOrdPro As String
Public Tip_PtMp As String
Public varUltCod_Item As String
Public varUltDes_Item As String
Public varUltCod_Color As String
Public varUltCod_Comb As String
Public varUltCod_Talla As String

Dim Reg As New ADODB.Recordset
Dim Estado As String
Dim Num_Secuencia As String
Public Cod_ClaMov As String

Public varValida_Factura As Boolean
Dim strSQL As String
Dim vTempLote As String

Public varCod_TipOrdTra As String
Public varCod_OrdTra As String
Public varCod_OrdPro As String

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
    'FunctButt1.SetFocus
End Sub

Sub Etiquetas()

End Sub

Sub Etiquetas1y4()
'Cod_Comb = DevuelveCampo("select cod_comb from lg_ordcompitem  where Sec_OrdComp='" & Sec_OrdComp & "' and Cod_Item='" & Me.TxtItem & "'", cCONNECT)
Label3.Caption = DevuelveCampo("select Des_Comb from tx_telacomb where Cod_Comb='" & Cod_Comb & "' and Cod_tela='" & Me.TxtItem & "'", cConnect)

Label4.Caption = DevuelveCampo("select a.des_color from lb_color a,lg_ordcompitem b where a.cod_color=b.cod_color and b.Cod_OrdComp='" & Cod_OrdComp & "' and Sec_OrdComp='" & Sec_OrdComp & "' and Cod_Item ='" & CODIGO & "'", cConnect)
Cod_color = DevuelveCampo("select a.cod_color from lb_color a,lg_ordcompitem b where a.cod_color=b.cod_color and b.Cod_OrdComp='" & Cod_OrdComp & "' and Sec_OrdComp='" & Sec_OrdComp & "' and Cod_Item ='" & CODIGO & "'", cConnect)

Label5.Caption = Cod_Talla

End Sub

Sub Etiquetas2y3()
    
Label3.Caption = DevuelveCampo(" select a.des_comb " & _
                                " from tx_telacomb a  " & _
                                " Where a.cod_comb='" & Cod_Comb & "' and cod_tela='" & CODIGO & "'", cConnect)
    

Label4.Caption = DevuelveCampo(" select a.des_color " & _
                                " from lg_stockstelten b,lb_color a  " & _
                                " Where a.Cod_color = b.Cod_Color And b.Cod_Almacen='" & Cod_Almacen & "' and " & _
                                " b.cod_tela='" & TxtItem & "' and Cod_TipOrdTra='" & _
                                Cod_TipOrdTra & "' and Cod_OrdTra='" & Cod_OrdTra & "' and a.cod_color='" & Cod_color & "'", cConnect)



'Cod_Talla = DevuelveCampo("select Cod_Talla  from lg_stockstelten  where Cod_Almacen='" & Cod_Almacen & "' and Cod_TipOrdTra='" & Cod_TipOrdTra & "' and Cod_OrdTra='" & Cod_OrdTra & "' and Cod_Tela='" & Me.TxtItem & "' and cod_comb='" & Cod_Comb & "'", cCONNECT)
Label5.Caption = Cod_Talla

End Sub

Sub ValidaHilo()
    Dim varOpcion As Boolean
On Error GoTo hand

    CODIGO = ""
    DESCRIPCION = ""

    varOpcion = False
    'Verificamos is haremos lo nuevo
    strSQL = "SELECT COUNT(*) FROM LG_ALMACEN WHERE Cod_Almacen = '" & Trim(Me.Cod_Almacen) & "' AND Tip_Item = 'T' AND Tip_Presentacion = 'T'"
    If DevuelveCampo(strSQL, cConnect) Then
        strSQL = "SELECT Cod_ClaOrdComp FROM LG_TIPOSMOV WHERE Cod_TipMov = '" & Trim(Cod_TipMovi) & "' AND Tip_Item = 'T' AND Flg_Partidas_Tinto = 'S'"
        strSQL = DevuelveCampo(strSQL, cConnect)
        If Trim(strSQL) <> "" Then
            strSQL = "SELECT COUNT(*) FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp = '" & strSQL & "' AND Tip_Item = 'T' AND Tip_Presentacion = 'T'"
            If DevuelveCampo(strSQL, cConnect) > 0 Then
                varOpcion = True
            End If
        End If
    End If


If varOpcion = True Then

'    varCod_TipOrdTra = ""
'    varCod_Ordtra = ""
'
'    If Reg.State <> 0 Then
'        If Reg.RecordCount > 0 Then
'            If IsNull(Reg("Cod_TipOrdTra").Value) Then
'                varCod_TipOrdTra = ""
'            Else
'                varCod_TipOrdTra = Reg("Cod_TipOrdTra").Value
'            End If
'
'            If IsNull(Reg("Cod_Ordtra").Value) Then
'                varCod_Ordtra = ""
'            Else
'                varCod_Ordtra = Reg("Cod_Ordtra").Value
'            End If
'        End If
'    End If
    
    Load frmBusqPartidasTelas
    frmBusqPartidasTelas.varCod_TipOrdTra = Me.varCod_TipOrdTra
    frmBusqPartidasTelas.varCod_OrdTra = Me.varCod_OrdTra
    frmBusqPartidasTelas.CARGA_GRID
    Set frmBusqPartidasTelas.oParent = Me
    frmBusqPartidasTelas.Show 1

Else
    Dim Temp
    Temp = Trim(DevuelveCampo("Select dbo.uf_devuelvecodigo(6," & IIf(Trim(TxtItem) = "", 0, Mid(TxtItem, 3)) & ")", cConnect))
    TxtItem = Left(TxtItem, 2) & Temp
    If Cod_ClaMov = "E" And Cod_ClaOrdComp <> "" Then
            Set frmBusqGeneral.oParent = Me
            If Tip_PtMp = "PT" Then
                frmBusqGeneral.sQuery = "UP_AyudaTelTen '1','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad_Envio & "','" & TxtLote & "','" & Cod_Proveedor & "'"
            Else
                frmBusqGeneral.sQuery = "UP_AyudaTelTen '3','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad_Envio & "','" & TxtLote & "','" & Cod_Proveedor & "'"
            End If
                                    
            frmBusqGeneral.Cargar_Datos
            frmBusqGeneral.Show 1
            If Tip_PtMp = "PT" Then
                If Paso = True Then
                    TxtItem = CODIGO
                    Sec_OrdComp = DESCRIPCION
                    TxtDesitem = DevuelveCampo("select des_tela  from tx_tela where cod_tela='" & CODIGO & "'", cConnect)
                End If
                Etiquetas1y4
            Else
                If Paso = True Then
                    TxtItem = CODIGO
                    TxtDesitem = DESCRIPCION
                End If
                    Sec_OrdComp = DevuelveCampo("select sec_ordcomp from lg_ordcompitem where Ser_OrdComp='" & Ser_OrdComp & "' and Cod_OrdComp='" & Cod_OrdComp & "' and Cod_Item='" & Me.TxtItem & "' AND cod_Comb='" & Me.Cod_Comb & "' and cod_talla='" & Me.Cod_Talla & "' and cod_color='" & Me.Cod_color & "'", cConnect)
                    Etiquetas2y3
            End If
    Else
        If Cod_ClaMov = "S" Then
            Set frmBusqGeneral.oParent = Me
            Cod_Proveedor = DevuelveCampo("select Cod_Proveedor from tx_ordtra where Cod_TipOrdTra='" & _
            Cod_TipOrdTra & "' and Cod_Ordtra='" & Cod_OrdTra & "' and Cod_OrdProv='" & Me.TxtLote & "'", cConnect)
            
            frmBusqGeneral.sQuery = "UP_AyudaTelTen '2','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad_Envio & "','" & TxtLote & "','" & Cod_Proveedor & "'"
            frmBusqGeneral.Cargar_Datos
            frmBusqGeneral.Show 1
            If Paso = True Then
                TxtItem = CODIGO
                TxtDesitem = DESCRIPCION
            End If
                Sec_OrdComp = DevuelveCampo("select sec_ordcomp from lg_ordcompitem where Ser_OrdComp='" & Ser_OrdComp & "' and Cod_OrdComp='" & Cod_OrdComp & "' and Cod_Item='" & Me.TxtItem & "' AND cod_Comb='" & Me.Cod_Comb & "' and cod_talla='" & Me.Cod_Talla & "' and cod_color='" & Me.Cod_color & "'", cConnect)
                Etiquetas2y3
        ElseIf Cod_ClaMov = "E" Then
            Set frmBusqGeneral.oParent = Me
            frmBusqGeneral.sQuery = "UP_AyudaTelTen '3','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad_Envio & "','" & TxtLote & "','" & Cod_Proveedor & "'"
            frmBusqGeneral.Cargar_Datos
            frmBusqGeneral.Show 1
            If Paso = True Then
                TxtItem = CODIGO
                TxtDesitem = DESCRIPCION
            End If
                Sec_OrdComp = DevuelveCampo("select sec_ordcomp from lg_ordcompitem where Ser_OrdComp='" & Ser_OrdComp & "' and Cod_OrdComp='" & Cod_OrdComp & "' and Cod_Item='" & Me.TxtItem & "'", cConnect)
                Etiquetas2y3
        ElseIf Cod_ClaOrdComp <> "" Then
            Set frmBusqGeneral.oParent = Me
            frmBusqGeneral.sQuery = "UP_AyudaTelTen '4','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad_Envio & "','" & TxtLote & "','" & Cod_Proveedor & "'"
            frmBusqGeneral.Cargar_Datos
            frmBusqGeneral.Show 1
            If Paso = True Then
                TxtItem = CODIGO
                Sec_OrdComp = DESCRIPCION
                TxtDesitem = DevuelveCampo("select des_hiltel from it_hilado where cod_hiltel='" & CODIGO & "'", cConnect)
            End If
            Etiquetas1y4
        End If
    
    End If
End If

    'Aqui cargamos las etiquetas
    strSQL = "SELECT Cod_UniMed FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
    Me.lblCantidad1.Caption = DevuelveCampo(strSQL, cConnect)
    strSQL = "SELECT Cod_UniMedCnf FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
    Me.lblCantidad2.Caption = DevuelveCampo(strSQL, cConnect)

Exit Sub
hand:
ErrorHandler err, "ValidaHilo"
End Sub
Public Sub Datos(Accion As String, EsAccion As Boolean)
On Error GoTo hand
Set Reg = Nothing
Reg.CursorLocation = adUseClient
If UCase(Accion) = "V" Then
    Reg.Open "UP_Lg_MoviStkTelACA '" & Accion & "','" & Cod_Almacen & "','" & Num_MovStk & "'", cConnect
Else
    Reg.Open "UP_ACT_STOCKSTELACA '" & Cod_Almacen & "','" & Num_MovStk & "','" & Accion & "','" & Num_Secuencia & "','" & _
            TxtLote & "','" & Cod_Proveedor & "','" & TxtItem & "','" & Cod_Comb & "','" & Cod_color & "','" & Cod_Talla & "'," & Me.TxtCantidad & "," & Me.txtBultos & ",'" & Me.TxtObs & "'," & _
            Cant_Anterior & ",'" & Sec_OrdComp & "'," & Me.txtCan_Movimiento_2daunimed.Text & ",'" & Me.txtCod_Ordpro.Text & "','" & vusu & "'", cConnect
End If
If EsAccion = False Then
    Set Me.DGridLista.DataSource = Reg
    DGridLista_RowColChange 0, 0
    Me.DGridLista.Columns("Cod_OrdTra").Visible = False
    Me.DGridLista.Columns("Cod_TipOrdTra").Visible = False
    Me.DGridLista.Columns("cod_color").Visible = False
    Me.DGridLista.Columns("cod_tela").Visible = False
    Me.DGridLista.Columns("Cod_Calidad").Visible = False
    Me.DGridLista.Columns("Cod_Comb").Visible = False
    Me.DGridLista.Columns("Cod_Talla").Visible = False
    Me.DGridLista.Columns("cod_proveedor").Visible = False
    
    Me.DGridLista.Columns("Can_Movimiento_2daunimed").Caption = "Cant Movimiento 2do"
    Me.DGridLista.Columns("Cod_OrdPro").Caption = "O/P"
End If
Exit Sub
hand:
ErrorHandler err, "Datos"
End Sub

Sub Habilita()
TxtItem.Enabled = True
TxtDesitem.Enabled = True
Me.TxtCantidad.Enabled = True
Me.txtCan_Movimiento_2daunimed.Enabled = True

Me.txtBultos.Enabled = True
Me.TxtLote.Enabled = True
Me.TxtObs.Enabled = True

Me.fraOP.Enabled = True
Me.txtCod_Ordpro.Enabled = True

End Sub
Sub Deshabilita()
TxtItem.Enabled = False
TxtDesitem.Enabled = False
Me.TxtCantidad.Enabled = False
Me.txtCan_Movimiento_2daunimed.Enabled = False

Me.txtBultos.Enabled = False
Me.TxtLote.Enabled = False
Me.TxtObs.Enabled = False

Me.txtCod_Ordpro.Enabled = False
Me.fraOP.Enabled = False

End Sub

Sub Limpia()
TxtItem = ""
TxtDesitem = ""
Me.TxtCantidad = "0"
Me.txtCan_Movimiento_2daunimed = "0"
Label3.Caption = ""
Label4.Caption = ""
Me.txtBultos = "0.00"
Me.TxtLote = ""
Me.TxtObs = ""

Me.txtCod_Ordpro.Text = ""
Me.txtDes_estpro.Text = ""

End Sub

Private Sub cmdAsigna_Click()
'    If Reg.RecordCount > 0 Then
    
        strSQL = "SELECT COUNT(*) FROM LG_TIPOSMOV WHERE Cod_TipMov = '" & Me.Cod_TipMovi & "' AND Cod_ClaMov = 'S' AND Cod_TipOrdPro = 'CO'"
    
        If DevuelveCampo(strSQL, cConnect) = 0 Then
            MsgBox "No se puede acceder a esta opción. Sirvase verificar", vbInformation, "Mensaje"
            Exit Sub
        End If
    
        Load frmAsignaOCorte
        frmAsignaOCorte.varCOD_ALMACEN = Me.Cod_Almacen
        frmAsignaOCorte.varNUM_MOVSTK = Me.Num_MovStk
        frmAsignaOCorte.Show 1
        Set frmAsignaOCorte = Nothing
        Datos "V", False
'    Else
'        MsgBox "No existen registros para acceder a esta opción", vbInformation, "Mensaje"
'    End If
End Sub

Private Sub cmdCapturarPeso_Click()
    TxtCantidad.Text = CapturaPeso
    If RTrim(TxtCantidad.Text) <> "0" Then
        txtBultos.Text = "1"
        
            MantFunc1.SetFocus
        
    End If
End Sub

Private Sub cmdFirst_Click()
    If Not Reg.BOF Then Reg.MoveFirst
End Sub

Private Sub cmdLast_Click()
    If Not Reg.EOF Then Reg.MoveLast
End Sub

Private Sub cmdNext_Click()
If Not Reg.EOF Then Reg.MoveNext
End Sub

Private Sub cmdPrevious_Click()
If Not Reg.BOF Then Reg.MovePrevious
End Sub


Private Sub cmdRollos_Click()
    Dim sSQl As String
    Dim sStrOrdTra As String
    
    If RTrim(TxtItem) = "" Then
        Estado = "DESHACER"
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Limpia
        Datos "V", False
        Deshabilita
    End If
    
    sSQl = "SM_DEVUELVEORDTRA '" & Cod_Almacen & "','" & Mid(Txtproveedor.Text, 1, 12) & "','" & TxtLote.Text & "','" & TxtItem.Text & "','" & Cod_Comb & "','" & Cod_color & "','" & Cod_Talla & "'"
    sStrOrdTra = DevuelveCampo(sSQl, cConnect)
        
    Load frmShowTX_Ordtra_ItemsRollos
    frmShowTX_Ordtra_ItemsRollos.sCod_TipOrdTra = Mid(sStrOrdTra, 1, 2)
    frmShowTX_Ordtra_ItemsRollos.Scod_ordtra = Mid(sStrOrdTra, 3, 8)
    frmShowTX_Ordtra_ItemsRollos.sNum_Secuencia = Mid(sStrOrdTra, 11, 4)
    frmShowTX_Ordtra_ItemsRollos.BUSCAR
    frmShowTX_Ordtra_ItemsRollos.Show vbModal
    Set frmShowTX_Ordtra_ItemsRollos = Nothing
End Sub

Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo hand
If Not Reg.EOF And Not Reg.BOF Then
    If Estado = "NUEVO" Then Exit Sub
    Me.txtBultos = Reg("Bultos")
    Me.TxtCantidad = Reg("Cant Movimiento")
    Me.TxtDesitem = Reg("tela")
    Me.TxtItem = Reg("cod_tela")
    Me.TxtLote = Reg("lote")
    Me.TxtObs = Reg("Observaciones")
    Me.Txtproveedor = Reg("Proveedor")
    Cod_OrdTra = Reg("Cod_OrdTra")
    Cant_Anterior = Reg("Cant Movimiento")
    Num_Secuencia = Reg("secuencia")
    Cod_Comb = Reg("Cod_Comb")
    Cod_Talla = Reg("Cod_Talla")
    Cod_Proveedor = Reg("cod_proveedor")
    
    Label5.Caption = Reg("medida")
    Label2.Caption = Reg("cod_calidad")
    Label3.Caption = Reg("combinacion")
    Label4.Caption = Reg("color")
    Cod_color = Reg("Cod_color")
    
    Me.txtCan_Movimiento_2daunimed.Text = Reg("Can_Movimiento_2daunimed").Value
    
    'Aqui cargamos la descripcion de la OP
    Me.txtCod_Ordpro.Text = Reg("Cod_OrdPro")
    Call Me.BUSCA_OP(1)
    
    'Aqui cargamos las etiquetas
    strSQL = "SELECT Cod_UniMed FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
    Me.lblCantidad1.Caption = DevuelveCampo(strSQL, cConnect)
    strSQL = "SELECT Cod_UniMedCnf FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
    Me.lblCantidad2.Caption = DevuelveCampo(strSQL, cConnect)
    
End If
Exit Sub
hand:
ErrorHandler err, "DGridLista_RowColChange"
End Sub


Private Sub Form_Load()
Label2.Caption = Cod_Calidad
Tip_item = DevuelveCampo("select tip_item from lg_almacen where cod_almacen='" & Me.Cod_Almacen & "'", cConnect)
Tip_presentacion = DevuelveCampo("select Tip_presentacion from lg_almacen where cod_almacen='" & Me.Cod_Almacen & "'", cConnect)
Cod_Calidad = DevuelveCampo("select isnull(Cod_Calidad,'') from lg_tiposmov where cod_tipmov='" & Me.Cod_TipMovi & "'", cConnect)
Cod_Calidad_Envio = DevuelveCampo("select isnull(Cod_Calidad_Envio,'') from lg_tiposmov where cod_tipmov='" & Me.Cod_TipMovi & "'", cConnect)

Cod_TipOrdTra = DevuelveCampo("select cod_tipordtra from Tx_TiposOrdTra where tip_item='" & Tip_item & "' and tip_presentacion='" & Tip_presentacion & "'", cConnect)

Me.Txtproveedor = DevuelveCampo("Select des_proveedor from lg_proveedor where cod_proveedor='" & Cod_Proveedor & "'", cConnect)
Limpia
Deshabilita
FormateaGrid Me.DGridLista
Datos "V", False
End Sub


Private Sub lote_Click()
Dim strSQL As String
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset
strSQL = "select cod_clamov,Tip_Accion,TIP_PTMP,cod_tipanx,Flg_Partidas_Tinto from lg_tiposmov where cod_tipmov='" & Cod_TipMovi & "'"

rs.Open strSQL, cConnect, adOpenStatic

If rs.RecordCount Then
    If rs("cod_clamov").Value = "E" And rs("Tip_ACcion").Value = "E" And Trim(rs("TIP_PTMP").Value) = "PT" And Trim(rs("cod_tipanx")) = "P" And Trim(rs("Flg_Partidas_Tinto")) <> "S" Then
        Set FRmLote.Padre = Me
        FRmLote.Cod_TipOrdTra = Me.Cod_TipOrdTra
        FRmLote.Cod_Proveedor = Me.Cod_Proveedor
        FRmLote.Grupo = DevuelveCampo("select cod_grupo from lg_ordcomp where ser_ordcomp='" & Ser_OrdComp & "' and cod_ordcomp='" & Cod_OrdComp & "'", cConnect)
        FRmLote.Show 1
        If Estado <> "NUEVO" Then
            'Estado = "NUEVO"
            Call MantFunc1_ActionClick(0, 0, "ADICIONAR")
        End If
        TxtLote.Text = NewLote
        NewLote = ""
    Else
        MsgBox "El Tipo de Movimiento no permite adicionar Lote", vbInformation, "Tela Acabada"
        NewLote = ""
    End If
End If
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo hand
Select Case ActionName
    Case "ADICIONAR"
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            Limpia
            Habilita
            Estado = "NUEVO"
            
            Me.txtCod_Ordpro.Text = Me.varCod_OrdPro
            Call Me.BUSCA_OP(1)
            
            If Not Reg.EOF And Not Reg.BOF Then
                Reg.MoveLast
                vTempLote = Trim(Reg("lote"))
                Reg.MoveFirst
                
                If IsNull(Reg("Cod_TipOrdTra").Value) Then
                    varCod_TipOrdTra = ""
                Else
                    varCod_TipOrdTra = Reg("Cod_TipOrdTra").Value
                End If
                
                If IsNull(Reg("Cod_Ordtra").Value) Then
                    varCod_OrdTra = ""
                Else
                    varCod_OrdTra = Reg("Cod_Ordtra").Value
                End If
                Me.TxtLote = vTempLote
                
                If Flg_Rollo = "*" Then
                    TxtItem.Text = varUltCod_Item
                    TxtDesitem.Text = varUltDes_Item
                    Cod_color = varUltCod_Color
                    Cod_Comb = varUltCod_Comb
                    Cod_Talla = varUltCod_Talla
                    
                    Etiquetas2y3
                    If cmdCapturarPeso.Visible Then
                        cmdCapturarPeso.SetFocus
                    End If
                Else
                    If TxtItem.Visible Then
                        TxtItem.SetFocus
                    End If
                End If
            Else
                If TxtLote.Visible Then
                    Me.TxtLote.SetFocus
                End If
            End If
            
    Case "MODIFICAR"
    
        If Me.varValida_Factura = False Then
            MsgBox "No se puede modificar el registro por que el Mov. Almacen posee una factura relacionada.", vbInformation, "Mensaje"
            Exit Sub
        End If

    
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Estado = "MODIFICAR"
        Deshabilita
        TxtCantidad.Enabled = True
        txtCan_Movimiento_2daunimed.Enabled = True
        Me.txtBultos.Enabled = True
        Me.TxtObs.Enabled = True
        TxtCantidad.SetFocus
    Case "ELIMINAR"
    
        If Me.varValida_Factura = False Then
            MsgBox "No se puede eliminar el registro por que el Mov. Almacen posee una factura relacionada.", vbInformation, "Mensaje"
            Exit Sub
        End If

        
        Datos "e", True
        Limpia
        Datos "v", False
        Estado = "ELIMINAR"
        Deshabilita
    Case "GRABAR"
        If Trim(TxtCantidad) = "" Or TxtCantidad = "0" Then MsgBox "Llene la cantidad", vbInformation: Exit Sub
        If Trim(txtBultos) = "" Then txtBultos = "0"
        If Trim(TxtItem) = "" Then MsgBox "Debe seleccionar un item", vbInformation: Exit Sub
        
        'Aqui haremos una validacion sobre cantidades
        If Val(Me.txtCan_Movimiento_2daunimed.Text) < 0 Then
            MsgBox "La 2da cantidad no puede ser menor que 0", vbInformation, "Mensaje"
            Me.txtCan_Movimiento_2daunimed.SetFocus
            Exit Sub
        End If
        
        If Estado = "NUEVO" Then
            Datos "i", True
        Else
            Datos "m", True
        End If
        Limpia
        Deshabilita
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Datos "V", False
        If Estado = "NUEVO" Then
            Call MantFunc1_ActionClick(0, 0, "ADICIONAR")
        Else
            Estado = "GRABAR"
        End If
    Case "DESHACER"
        Estado = "DESHACER"
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Limpia
        Datos "V", False
        Deshabilita
    Case "SALIR"
        Unload Me
End Select
Exit Sub
hand:
ErrorHandler err, "MantFunc1_ActionClick"
End Sub


Private Sub TxtBultos_GotFocus()
    txtBultos.SelStart = 0
    txtBultos.SelLength = Len(txtBultos.Text)
End Sub

Private Sub txtBultos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
    SoloNumeros txtBultos, KeyAscii, False, 0, 4
End Sub

Private Sub txtCan_Movimiento_2daunimed_GotFocus()
    Me.txtCan_Movimiento_2daunimed.SelStart = 0
    Me.txtCan_Movimiento_2daunimed.SelLength = Len(Me.txtCan_Movimiento_2daunimed.Text)
End Sub

Private Sub txtCan_Movimiento_2daunimed_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBultos.SetFocus
    Else
        Call SoloNumeros(Me.txtCan_Movimiento_2daunimed, KeyAscii, True, 3, 9)
    End If
End Sub

Private Sub txtCan_Movimiento_2daunimed_LostFocus()
    If Trim(Me.txtCan_Movimiento_2daunimed.Text) = "" Then
        Me.txtCan_Movimiento_2daunimed.Text = "0"
    End If
End Sub

Private Sub TxtCantidad_GotFocus()
    TxtCantidad.SelStart = 0
    TxtCantidad.SelLength = Len(TxtCantidad.Text)
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
    If Flg_Rollo = "*" Then
        If KeyAscii = vbKeyReturn Then
            txtBultos.Text = "1"
            MantFunc1.SetFocus
        End If
    Else
        If KeyAscii = 13 Then SendKeys "{tab}"
            SoloNumeros TxtCantidad, KeyAscii, True, 3, 6
    End If
End Sub

Private Sub txtcod_ordpro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtCod_Ordpro = Right("00000" & Trim(txtCod_Ordpro.Text), 5)
        Call Me.BUSCA_OP(1)
        MantFunc1.SetFocus
    End If
End Sub

Private Sub TxtDesitem_KeyPress(KeyAscii As Integer)
On Error GoTo hand
If KeyAscii = 13 Then
    ValidaHilo
    SendKeys "{tab}"
End If
Exit Sub
hand:



End Sub



Private Sub TxtItem_KeyPress(KeyAscii As Integer)
On Error GoTo hand
If KeyAscii = 13 Then
    ValidaHilo
    SendKeys "{tab}"
    If Flg_Rollo = "*" Then
        varUltCod_Item = TxtItem.Text
        varUltDes_Item = TxtDesitem.Text
        varUltCod_Color = Cod_color
        varUltCod_Comb = Cod_Comb
        varUltCod_Talla = Cod_Talla
        cmdCapturarPeso.SetFocus
    End If
End If

Exit Sub
hand:
    ErrorHandler err, "TxtItem"
End Sub




Private Sub txtLote_KeyPress(KeyAscii As Integer)
On Error GoTo hand
Dim varOpcion As Boolean

If KeyAscii = 13 Then

    varOpcion = False
    'Verificamos is haremos lo nuevo
    strSQL = "SELECT COUNT(*) FROM LG_ALMACEN WHERE Cod_Almacen = '" & Trim(Me.Cod_Almacen) & "' AND Tip_Item = 'T' AND Tip_Presentacion = 'T'"
    If DevuelveCampo(strSQL, cConnect) Then
        strSQL = "SELECT Cod_ClaOrdComp FROM LG_TIPOSMOV WHERE Cod_TipMov = '" & Trim(Cod_TipMovi) & "' AND Tip_Item = 'T' AND Flg_Partidas_Tinto = 'S'"
        strSQL = DevuelveCampo(strSQL, cConnect)
        If Trim(strSQL) <> "" Then
            strSQL = "SELECT COUNT(*) FROM LG_CLAORDCOMP WHERE Cod_ClaOrdComp = '" & strSQL & "' AND Tip_Item = 'T' AND Tip_Presentacion = 'T'"
            If DevuelveCampo(strSQL, cConnect) > 0 Then
                varOpcion = True
            End If
        End If
    End If

    If varOpcion Then
    
        varCod_TipOrdTra = ""
        varCod_OrdTra = ""
        
        If Reg.State <> 0 Then
            If Reg.RecordCount > 0 Then
                If IsNull(Reg("Cod_TipOrdTra").Value) Then
                    varCod_TipOrdTra = ""
                Else
                    varCod_TipOrdTra = Reg("Cod_TipOrdTra").Value
                End If
                
                If IsNull(Reg("Cod_Ordtra").Value) Then
                    varCod_OrdTra = ""
                Else
                    varCod_OrdTra = Reg("Cod_Ordtra").Value
                End If
            End If
        End If
    
    
        Load frmBusqPartidasLote
        frmBusqPartidasLote.varCod_OrdComp = Me.Cod_OrdComp
        frmBusqPartidasLote.varSer_OrdComp = Me.Ser_OrdComp
        frmBusqPartidasLote.CARGA_GRID
        Set frmBusqPartidasLote.oParent = Me
        frmBusqPartidasLote.Show 1
        'vTempLote = Me.TxtLote.Text
        Set frmBusqPartidasLote = Nothing
    
    Else
        If Cod_ClaMov = "E" And Cod_ClaOrdComp <> "" And Tip_PtMp = "PT" Then
            If DevuelveCampo("select  count(*) from   TX_ORDTRA  Where  COD_TIPORDTRA='" & Cod_TipOrdTra & "' and " & _
                            "   Cod_OrdProv like '%" & Me.TxtLote & "%' and cod_proveedor='" & Cod_Proveedor & "'", cConnect) <= 0 Then
                MsgBox "Este Lote no existe", vbInformation
            Else
            
                Set frmBusqGeneral.oParent = Me
                frmBusqGeneral.sQuery = " select  a.Cod_OrdProv as Orden,b.cod_Proveedor as [Cod Proveedor] ,b.des_proveedor as Descripcion " & _
                                        " from    TX_ORDTRA a,lg_proveedor b " & _
                                        " Where " & _
                                        " a.Cod_Proveedor=b.Cod_Proveedor and " & _
                                        " a.COD_TIPORDTRA='" & Cod_TipOrdTra & "' and " & _
                                        " a.Cod_OrdProv like '%" & Me.TxtLote & "%'"
                frmBusqGeneral.Cargar_Datos
                frmBusqGeneral.Show 1
                TxtLote = CODIGO
                Cod_Proveedor = DESCRIPCION
    '                                    " a.cod_proveedor='" & Cod_Proveedor & "' and "
            End If
        
        Else
            If DevuelveCampo("select  count(*) from   TX_ORDTRA  Where  COD_TIPORDTRA='" & Cod_TipOrdTra & "' and " & _
                            "   Cod_OrdProv like '%" & Me.TxtLote & "%'", cConnect) > 0 Then
                Set frmBusqGeneral.oParent = Me
                frmBusqGeneral.sQuery = " select  a.Cod_OrdProv as Orden,b.cod_proveedor as [Cod Prov],b.Des_Proveedor as Proveedor " & _
                                        " from    TX_ORDTRA a,lg_proveedor b " & _
                                        " Where " & _
                                        " a.Cod_Proveedor=b.Cod_Proveedor and " & _
                                        " a.COD_TIPORDTRA='" & Cod_TipOrdTra & "' and " & _
                                        " a.Cod_OrdProv like '%" & Me.TxtLote & "%'"
                frmBusqGeneral.Cargar_Datos
                frmBusqGeneral.Show 1
                TxtLote = CODIGO
                Cod_Proveedor = DESCRIPCION
            Else
                TxtLote = Trim(DevuelveCampo("select  Cod_OrdProv from   TX_ORDTRA  Where  COD_TIPORDTRA='" & Cod_TipOrdTra & "' and " & _
                            "   Cod_OrdProv like '%" & Me.TxtLote & "%'", cConnect))
                            
                Cod_Proveedor = Trim(DevuelveCampo("select  Cod_Proveedor  from   TX_ORDTRA  Where  COD_TIPORDTRA='" & Cod_TipOrdTra & "' and " & _
                            "   Cod_OrdProv like '%" & Me.TxtLote & "%'", cConnect))
            End If
            
        End If
        Cod_OrdTra = DevuelveCampo("select cod_ordtra from tx_ordtra where Cod_TipOrdTra='" & Cod_TipOrdTra & _
            "' and Cod_Proveedor='" & Cod_Proveedor & "' and Cod_OrdProv='" & TxtLote & "'", cConnect)
        Me.Txtproveedor = DevuelveCampo("Select des_proveedor from lg_proveedor where cod_proveedor='" & Cod_Proveedor & "'", cConnect)
        'SendKeys "{tab}"
        
    End If
    
End If
If KeyAscii = 13 Then SendKeys "{tab}"
Exit Sub
hand:
ErrorHandler err, "TxtLote_KeyPress"
End Sub

Private Sub TxtObs_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Len(TxtObs.Text) = 0 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Function CapturaPeso()
  On Error GoTo ControlErrores
    Dim sBuffer As String
    sBuffer = String(19, 0)
    If (Captura) Then
       CapturaPeso = Captura / 100
  Else
      MsgBox "Error en Lectura. Comuníquese con Sistemas", vbExclamation
   End If
  
  Exit Function
ControlErrores:
  CapturaPeso = -1
End Function

