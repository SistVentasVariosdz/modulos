VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "mantenimientos.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form FrmDetalleTelCru 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tela Cruda"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraValorizar 
      Caption         =   "Valorizar Transferencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3060
      TabIndex        =   34
      Top             =   1620
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox TxtDolares 
         Height          =   285
         Left            =   2160
         TabIndex        =   36
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TxtSoles 
         Height          =   285
         Left            =   2160
         TabIndex        =   35
         Top             =   360
         Width           =   975
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   510
         Left            =   600
         TabIndex        =   37
         Top             =   1200
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"FrmDetalleTelCru.frx":0000
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label4 
         Caption         =   "Importe Dolares"
         Height          =   255
         Left            =   600
         TabIndex        =   39
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Soles"
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   38
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton lote 
      Caption         =   "&Ingresar Lote"
      Height          =   525
      Left            =   8280
      TabIndex        =   27
      Top             =   6135
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
      Left            =   30
      TabIndex        =   25
      Tag             =   "List"
      Top             =   60
      Width           =   9495
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2775
         Left            =   180
         TabIndex        =   26
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
      Height          =   2670
      Left            =   60
      TabIndex        =   15
      Tag             =   "Detail"
      Top             =   3360
      Width           =   9450
      Begin VB.CommandButton CmdTransferir 
         Caption         =   "Transferir a"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   5520
         TabIndex        =   41
         Top             =   2160
         Width           =   2220
      End
      Begin VB.CommandButton cmdPesosBal 
         Caption         =   "..."
         Height          =   315
         Left            =   7800
         TabIndex        =   33
         ToolTipText     =   "Secuencia Pesos Balanza"
         Top             =   195
         Width           =   405
      End
      Begin VB.CommandButton cmdGetInfo 
         Height          =   285
         Left            =   3180
         Picture         =   "FrmDetalleTelCru.frx":0096
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Seleccionar Datos por Tela"
         Top             =   180
         Width           =   375
      End
      Begin VB.TextBox txtCan_Movimiento_2daunimed 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6780
         TabIndex        =   5
         Text            =   "0"
         Top             =   525
         Width           =   945
      End
      Begin VB.TextBox TxtObs 
         Height          =   645
         Left            =   6780
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
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
         TabIndex        =   7
         Text            =   "0"
         Top             =   855
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
         Left            =   1170
         TabIndex        =   1
         Top             =   510
         Width           =   3285
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
         Left            =   1170
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
         TabIndex        =   3
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
         Left            =   1185
         MaxLength       =   8
         TabIndex        =   2
         Top             =   840
         Width           =   945
      End
      Begin VB.TextBox TxtDesitem 
         Height          =   315
         Left            =   2130
         TabIndex        =   16
         Top             =   825
         Width           =   2325
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
         TabIndex        =   6
         Top             =   600
         Width           =   75
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
         TabIndex        =   4
         Top             =   270
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Medida:"
         Height          =   195
         Index           =   5
         Left            =   5550
         TabIndex        =   31
         Top             =   1875
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
         Left            =   6810
         TabIndex        =   30
         Top             =   1875
         Width           =   2355
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
         Height          =   240
         Left            =   1185
         TabIndex        =   29
         Top             =   1260
         Width           =   4185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comb:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   1260
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   9
         Left            =   5520
         TabIndex        =   24
         Top             =   1230
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad Bultos:"
         Height          =   195
         Index           =   8
         Left            =   5520
         TabIndex        =   23
         Top             =   870
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   22
         Top             =   585
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lote:"
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   21
         Top             =   255
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tela:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   945
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad Mov:"
         Height          =   195
         Index           =   7
         Left            =   5520
         TabIndex        =   19
         Top             =   255
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Calidad:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   1590
         Width           =   570
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
         Left            =   1170
         TabIndex        =   17
         Top             =   1590
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   2535
      TabIndex        =   10
      Top             =   6075
      Width           =   2115
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1530
         Picture         =   "FrmDetalleTelCru.frx":03A0
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ultimo"
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   1050
         Picture         =   "FrmDetalleTelCru.frx":0512
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Siguiente"
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   570
         Picture         =   "FrmDetalleTelCru.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Anterior"
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   90
         Picture         =   "FrmDetalleTelCru.frx":07F6
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Primero"
         Top             =   60
         Width           =   495
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   4620
      TabIndex        =   9
      Top             =   6105
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmDetalleTelCru.frx":0968
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   60
      TabIndex        =   40
      Top             =   6135
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmDetalleTelCru.frx":0B20
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmDetalleTelCru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CODIGO
Public DESCRIPCION

Public NewLote As String, bElijeDatos As Boolean

'Public lote As String
Public Paso As Boolean
Dim Tip_item As String
Dim Tip_presentacion As String
Public Cod_Calidad As String
Dim Cant_Anterior As Double
Public Cod_Comb As String
Dim Des_Comb As String
Public Cod_Talla As String

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
Public Tip_PtMp As String, sFlg_Ot_Tejeduria_Generada As String
Public Flg_Partidas_Tinto As String
Public Flg_Partida_Generada As String
Public Cod_TipOrdTra1 As String
Public Cod_OrdTra1 As String

Dim Reg As New ADODB.Recordset
Dim Estado As String
Dim Num_Secuencia As String
Dim Num_Secuencia_OrdTra_Tinto As String
Public Cod_ClaMov As String

Public varValida_Factura As Boolean
Dim strSQL As String
Dim varCod_TipFamTela As String

Sub Etiquetas2y3()
'Cod_Comb = DevuelveCampo("select cod_comb from lg_stockstelten  where Cod_Almacen='" & Cod_Almacen & "' and Cod_TipOrdTra='" & Cod_TipOrdTra & "' and Cod_OrdTra='" & Cod_OrdTra & "' and Cod_Tela='" & Me.TxtItem & "'", cCONNECT)
'Label3.Caption = DevuelveCampo("select Des_Comb from tx_telacomb where Cod_Comb='" & Cod_Comb & "' and Cod_tela='" & Me.TxtItem & "'", cCONNECT)
Label3.Caption = DevuelveCampo(" select a.des_comb " & _
                                " from tx_telacomb a  " & _
                                " Where a.cod_comb='" & Cod_Comb & "' and a.cod_tela='" & CODIGO & "'", cConnect)
'Cod_Talla = DevuelveCampo("select Cod_Talla  from lg_stockstelten  where Cod_Almacen='" & Cod_Almacen & "' and Cod_TipOrdTra='" & Cod_TipOrdTra & "' and Cod_OrdTra='" & Cod_OrdTra & "' and Cod_Tela='" & Me.TxtItem & "'", cCONNECT)
Label5.Caption = Cod_Talla
Label2.Caption = Cod_Calidad
End Sub

Sub ValidaHilo()
Dim Temp

CODIGO = ""
DESCRIPCION = ""

strSQL = "select flg_partidas_tinto from Lg_TiposMov where cod_tipmov='" & Cod_TipMovi & "'"
If Cod_ClaMov = "E" And Cod_ClaOrdComp <> "" And DevuelveCampo(strSQL, cConnect) = "N" Then
    Set frmBusqGeneral.oParent = Me
    frmBusqGeneral.sQuery = "UP_AyudaTellCru '1','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
    
    frmBusqGeneral.Cargar_Datos
    frmBusqGeneral.gexList.Columns("Codigo").Width = 795
    frmBusqGeneral.gexList.Columns("Secuencial").Width = 420
    frmBusqGeneral.gexList.Columns("Descripcion").Width = 2775
    frmBusqGeneral.gexList.Columns("Cant Comp").Width = 915
    frmBusqGeneral.gexList.Columns("Cant Recib.").Width = 975
    frmBusqGeneral.gexList.Columns("Combinacion").Width = 1050
    frmBusqGeneral.gexList.Columns("Talla").Width = 525
    frmBusqGeneral.Show 1
    If Paso = True Then
        TxtItem = CODIGO
        Sec_OrdComp = DESCRIPCION
        TxtDesitem = DevuelveCampo("select des_tela  from tx_tela where cod_tela='" & CODIGO & "'", cConnect)
    End If
Else
    If Cod_ClaMov = "S" Then
        Set frmBusqGeneral.oParent = Me
        frmBusqGeneral.sQuery = "UP_AyudaTellCru '2','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "','" & TxtItem.Text & "','" & Num_MovStk & "'"
        frmBusqGeneral.Cargar_Datos
        frmBusqGeneral.Show 1
        If Paso = True Then
            TxtItem = CODIGO
            TxtDesitem = DESCRIPCION
        End If
        Etiquetas2y3
    ElseIf Cod_ClaMov = "E" Then
        Set frmBusqGeneral.oParent = Me
        frmBusqGeneral.sQuery = "UP_AyudaTellCru '3','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "','" & TxtItem.Text & "','" & Num_MovStk & "'"
        frmBusqGeneral.Cargar_Datos
        frmBusqGeneral.Caption = "Seleccionar Tela"
        frmBusqGeneral.Show 1
        If Paso = True Then
            TxtItem = CODIGO
            TxtDesitem = DESCRIPCION
            BuscaComb CStr(CODIGO)
            strSQL = "DECLARE @COD_TIPFAMTELA AS CHAR(1) " & _
                             "EXEC SM_ENCUENTRA_TIPOFAMTELA '" & TxtItem & "', " & _
                             "@COD_TIPFAMTELA OUTPUT"
                    If DevuelveCampo(strSQL, cConnect) = "R" Then BuscaTalla
            Sec_OrdComp = DevuelveCampo("select sec_ordcomp from lg_ordcompitem where Ser_OrdComp='" & Ser_OrdComp & "' and Cod_OrdComp='" & Cod_OrdComp & "' and Cod_Item='" & Me.TxtItem & "' AND cod_Comb='" & Me.Cod_Comb & "' and cod_talla='" & Me.Cod_Talla & "'", cConnect)
        End If
        Etiquetas2y3
    ElseIf Cod_ClaOrdComp <> "" Then
        Set frmBusqGeneral.oParent = Me
        frmBusqGeneral.sQuery = "UP_AyudaTellCru '4','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
        frmBusqGeneral.Cargar_Datos
        frmBusqGeneral.Show 1
        Unload frmBusqGeneral
        If Paso = True Then
            TxtItem = CODIGO
            Sec_OrdComp = DESCRIPCION
            TxtDesitem = DevuelveCampo("select des_tela from tx_tela where cod_tela='" & CODIGO & "'", cConnect)
        End If
    End If
End If

'If Trim(Cod_Comb) = "" Then
'    'AHSP MODIFICO ESTA LINEA
'    'dECIA DevuelveCampo("select cod_comb from lg_ordcompitem  where Sec_OrdComp='" & Sec_OrdComp & "' and Cod_Item='" & Me.TxtItem & "'", cConnect)
'    Cod_Comb = DevuelveCampo("select cod_comb from lg_ordcompitem  where Ser_OrdComp='" & Ser_OrdComp & "' AND Sec_OrdComp='" & Sec_OrdComp & "' and Cod_OrdComp = '" & Me.Cod_OrdComp & "' and Cod_Item='" & Me.TxtItem & "'", cConnect)
'End If

Label3.Caption = DevuelveCampo("select Des_Comb from tx_telacomb where Cod_Comb='" & Cod_Comb & "' and Cod_tela='" & Me.TxtItem & "'", cConnect)

'Cod_Talla = DevuelveCampo("select Cod_Talla  from lg_ordcompitem  where Sec_OrdComp='" & Sec_OrdComp & "' and Cod_Item='" & Me.TxtItem & "'", cCONNECT)
Label5.Caption = Cod_Talla

'Aqui cargaremos los nuevos valores para los labels de cantidades
strSQL = "EXEC SM_ENCUENTRA_TIPOFAMTELA '" & Me.TxtItem.Text & "',''"
varCod_TipFamTela = DevuelveCampo(strSQL, cConnect)
If varCod_TipFamTela = "N" Then
    strSQL = "SELECT Cod_UniMed FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
    Me.lblCantidad1.Caption = DevuelveCampo(strSQL, cConnect)
    strSQL = "SELECT Cod_UniMedCnf FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
    Me.lblCantidad2.Caption = DevuelveCampo(strSQL, cConnect)
Else
    strSQL = "SELECT Cod_UniMedCnf FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
    Me.lblCantidad1.Caption = DevuelveCampo(strSQL, cConnect)
    strSQL = "SELECT Cod_UniMed FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
    Me.lblCantidad2.Caption = DevuelveCampo(strSQL, cConnect)
End If

End Sub

Private Sub BuscaComb(Cod_Item As String)
On Error GoTo ErrBCombo
Dim oBusq As New frmBusqGeneral
    
    strSQL = "SELECT Cod_Comb AS Codigo, Des_Comb AS Descripcion FROM TX_TELACOMB " & _
             "WHERE Cod_Tela = '" & Cod_Item & "'"
    CODIGO = ""
    Set oBusq.oParent = Me
    oBusq.sQuery = strSQL
    oBusq.Cargar_Datos
    oBusq.Caption = "Seleccionar Combinacion"
    oBusq.Show 1
    Unload oBusq
    Cod_Comb = ""
    If Paso = True Then Cod_Comb = CODIGO
Exit Sub
ErrBCombo:
    MsgBox err.Description, vbCritical + vbOKOnly, "Buscar Combinacion"
End Sub

Private Sub BuscaTalla()
On Error GoTo ErrBCombo
Dim oBusq As New frmBusqGeneral
    
    strSQL = "SELECT Cod_Talla AS Codigo, Cod_Talla AS Descripcion FROM TG_TALLA "
    CODIGO = ""
    Set oBusq.oParent = Me
    oBusq.sQuery = strSQL
    oBusq.Cargar_Datos
    oBusq.Caption = "Seleccionar Talla"
    oBusq.Show 1
    Unload oBusq
    Cod_Talla = ""
    If Paso = True Then Cod_Talla = CODIGO
Exit Sub
ErrBCombo:
    MsgBox err.Description, vbCritical + vbOKOnly, "Buscar Combinacion"
End Sub


Public Sub Datos(Accion As String, EsAccion As Boolean)
On Error GoTo hand
Set Reg = Nothing
Reg.CursorLocation = adUseClient

If UCase(Accion) = "V" Then
    Reg.Open "UP_Lg_MoviStkTelCru '" & Accion & "','" & Cod_Almacen & "','" & Num_MovStk & "'", cConnect
Else
    Reg.Open "UP_ACT_STOCKSTELCRU '" & Cod_Almacen & "','" & Num_MovStk & "','" & Accion & "','" & Num_Secuencia & "','" & _
            TxtLote & "','" & Cod_Proveedor & "','" & TxtItem & "','" & Cod_Comb & "','" & Cod_Talla & "'," & Me.TxtCantidad & "," & Me.txtBultos & ",'" & TxtObs.Text & "'," & _
            Cant_Anterior & ",'" & Sec_OrdComp & "'," & Me.txtCan_Movimiento_2daunimed.Text & ",'" & vusu & "', " & Num_Secuencia_OrdTra_Tinto, cConnect
End If

If EsAccion = False Then
    Set Me.DGridLista.DataSource = Reg
    DGridLista_RowColChange 0, 0
    Me.DGridLista.Columns("Cod_OrdTra").Visible = False
    Me.DGridLista.Columns("calidad").Visible = False
    Me.DGridLista.Columns("cod_tela").Visible = False
    Me.DGridLista.Columns("Cod_Comb").Visible = False
    Me.DGridLista.Columns("Cod_Talla").Visible = False
    Me.DGridLista.Columns("cod_proveedor").Visible = False
    Me.DGridLista.Columns("Can_Movimiento_2daunimed").Caption = "Cant Movimiento 2do"
    
End If
Exit Sub
hand:
ErrorHandler err, "Datos"
End Sub

Sub Habilita()
    TxtCantidad.Enabled = True
    txtCan_Movimiento_2daunimed.Enabled = True
    txtBultos.Enabled = True
    TxtObs.Enabled = True
    If Flg_Partida_Generada <> "S" Then
        TxtLote.Enabled = True
        TxtItem.Enabled = True
        TxtDesitem.Enabled = True
        TxtLote.SetFocus
    Else
        cmdGetInfo.Visible = True
        TxtCantidad.SetFocus
    End If
    cmdPesosBal.Enabled = True
End Sub

Sub Deshabilita()
TxtItem.Enabled = False
TxtDesitem.Enabled = False
Me.TxtCantidad.Enabled = False
txtCan_Movimiento_2daunimed.Enabled = False

Me.txtBultos.Enabled = False
Me.TxtLote.Enabled = False
Me.TxtObs.Enabled = False
cmdGetInfo.Visible = False
cmdPesosBal.Enabled = False
End Sub

Sub Limpia()
TxtItem = ""
TxtDesitem = ""
Me.TxtCantidad = "0"
Me.txtCan_Movimiento_2daunimed = "0"
Me.txtBultos = "0.00"
Me.TxtLote = ""
Me.TxtObs = ""
Label3.Caption = ""
Num_Secuencia_OrdTra_Tinto = "0"
End Sub

Private Sub cmdFirst_Click()
    If Not Reg.BOF Then Reg.MoveFirst
End Sub

Private Sub cmdGetInfo_Click()
    If Cod_ClaMov = "E" Then
        BuscPartidaGenEnt
    Else
        BuscPartidaGenSal
    End If
    TxtCantidad.SetFocus
End Sub

Private Sub BuscPartidaGenSal()
    With frmDetTelCruInfo
        '.Width = 10500
        .vCod_OrdTra = Cod_OrdTra1
        .vCod_TipOrdTra = Cod_TipOrdTra1
        .vCod_Almacen = Cod_Almacen
        .SM_AYUDA_ITEMS_DE_PARTIDA
        If .gexLotes.RowCount > 1 Then .Show vbModal
        '''If .gexLotes.RowCount > 0 Then .Show vbModal
        If .gexLotes.RowCount > 0 And Not .bCancel Then
            TxtItem = .gexLotes.Value(.gexLotes.Columns("COD_TELA").Index)
            TxtDesitem = Trim(.gexLotes.Value(.gexLotes.Columns("DES_TELA").Index))
            Cod_Comb = .gexLotes.Value(.gexLotes.Columns("COD_COMB").Index)
            Label3 = Trim(.gexLotes.Value(.gexLotes.Columns("DES_COMB").Index))
            Cod_Talla = .gexLotes.Value(.gexLotes.Columns("COD_TALLA").Index)
            Label5 = Cod_Talla
            Num_Secuencia_OrdTra_Tinto = .gexLotes.Value(.gexLotes.Columns("NUM_SECUENCIA").Index)
            Sec_OrdComp = DevuelveCampo("select sec_ordcomp from lg_ordcompitem where Ser_OrdComp='" & Ser_OrdComp & "' and Cod_OrdComp='" & Cod_OrdComp & "' and Cod_Item='" & Me.TxtItem & "' AND cod_Comb='" & Me.Cod_Comb & "' and cod_talla='" & Me.Cod_Talla & "'", cConnect)
            If Not .bCancelSec Then
                If TxtItem <> .lblCod_Tela Then
                    TxtItem = .lblCod_Tela
                    TxtDesitem = .lblDes_Tela
                    
                End If
                TxtLote = .lblCod_OrdProv
                Cod_Proveedor = .lblCod_Proveedor
                Txtproveedor = .lblDes_Proveedor
                Label2 = .lblCod_Calidad
                '.vStock
                'Cod_OrdTra
            End If
            'Aqui cargaremos los nuevos valores para los labels de cantidades
            strSQL = "EXEC SM_ENCUENTRA_TIPOFAMTELA '" & TxtItem.Text & "',''"
            varCod_TipFamTela = DevuelveCampo(strSQL, cConnect)
            If varCod_TipFamTela = "N" Then
                strSQL = "SELECT Cod_UniMed FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
                Me.lblCantidad1.Caption = DevuelveCampo(strSQL, cConnect)
                strSQL = "SELECT Cod_UniMedCnf FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
                Me.lblCantidad2.Caption = DevuelveCampo(strSQL, cConnect)
            Else
                strSQL = "SELECT Cod_UniMedCnf FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
                Me.lblCantidad1.Caption = DevuelveCampo(strSQL, cConnect)
                strSQL = "SELECT Cod_UniMed FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
                Me.lblCantidad2.Caption = DevuelveCampo(strSQL, cConnect)
            End If
        End If
    End With
    Unload frmDetTelCruInfo
End Sub

Private Sub BuscPartidaGenEnt()
    With frmDetTelCruEnt
        .vCod_OrdTra = Cod_OrdTra1
        .vCod_TipOrdTra = Cod_TipOrdTra1
        .vCod_Almacen = Cod_Almacen
        .SM_AYUDA_DEVOLUCION_TELA_CRUDA_DE_PARTIDAS
        If .gexLotes.RowCount > 1 Then .Show vbModal
        If .gexLotes.RowCount > 0 And Not .bCancel Then
            TxtItem = .gexLotes.Value(.gexLotes.Columns("COD_TELA").Index)
            TxtDesitem = Trim(.gexLotes.Value(.gexLotes.Columns("DES_TELA").Index))
            Cod_Comb = .gexLotes.Value(.gexLotes.Columns("COD_COMB").Index)
            Label3 = Trim(.gexLotes.Value(.gexLotes.Columns("DES_COMB").Index))
            Cod_Talla = .gexLotes.Value(.gexLotes.Columns("COD_MEDIDA").Index)
            Label5 = Cod_Talla
            Num_Secuencia_OrdTra_Tinto = .gexLotes.Value(.gexLotes.Columns("NUM_SECUENCIA").Index)
            Sec_OrdComp = DevuelveCampo("select sec_ordcomp from lg_ordcompitem where Ser_OrdComp='" & Ser_OrdComp & "' and Cod_OrdComp='" & Cod_OrdComp & "' and Cod_Item='" & Me.TxtItem & "' AND cod_Comb='" & Me.Cod_Comb & "' and cod_talla='" & Me.Cod_Talla & "'", cConnect)
            TxtLote = .gexLotes.Value(.gexLotes.Columns("COD_ORDPROV").Index)
            Cod_Proveedor = .gexLotes.Value(.gexLotes.Columns("COD_PROVEEDOR").Index)
            Txtproveedor = .gexLotes.Value(.gexLotes.Columns("DES_PROVEEDOR").Index)
            If DevuelveCampo("SELECT dbo.tx_obtiene_tipofamilia_tela('" & TxtItem & "')", cConnect) = "N" Then
                TxtCantidad = .gexLotes.Value(.gexLotes.Columns("KGS_ENVIADOS").Index)
                txtCan_Movimiento_2daunimed = .gexLotes.Value(.gexLotes.Columns("uni_enviados").Index)
            Else
                TxtCantidad = .gexLotes.Value(.gexLotes.Columns("uni_enviados").Index)
                txtCan_Movimiento_2daunimed = .gexLotes.Value(.gexLotes.Columns("KGS_ENVIADOS").Index)
            End If
            'Aqui cargaremos los nuevos valores para los labels de cantidades
            strSQL = "EXEC SM_ENCUENTRA_TIPOFAMTELA '" & TxtItem.Text & "',''"
            varCod_TipFamTela = DevuelveCampo(strSQL, cConnect)
            If varCod_TipFamTela = "N" Then
                strSQL = "SELECT Cod_UniMed FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
                Me.lblCantidad1.Caption = DevuelveCampo(strSQL, cConnect)
                strSQL = "SELECT Cod_UniMedCnf FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
                Me.lblCantidad2.Caption = DevuelveCampo(strSQL, cConnect)
            Else
                strSQL = "SELECT Cod_UniMedCnf FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
                Me.lblCantidad1.Caption = DevuelveCampo(strSQL, cConnect)
                strSQL = "SELECT Cod_UniMed FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
                Me.lblCantidad2.Caption = DevuelveCampo(strSQL, cConnect)
            End If
        End If
    End With
    Unload frmDetTelCruEnt
End Sub

Private Sub cmdLast_Click()
    If Not Reg.EOF Then Reg.MoveLast
End Sub

Private Sub cmdNext_Click()
If Not Reg.EOF Then Reg.MoveNext
End Sub

Private Sub cmdPesosBal_Click()
    frmPesosBal.Show vbModal
    TxtCantidad = frmPesosBal.lblTotal
    Unload frmPesosBal
End Sub

Private Sub cmdPrevious_Click()
If Not Reg.BOF Then Reg.MovePrevious
End Sub

Private Sub CmdTransferir_Click()
'    Load frmDetalleTelCruTransf
'
'    frmDetalleTelCruTransf.xlote = Me.TxtLote
'    frmDetalleTelCruTransf.xCod_Proveedor = Mid(TxtProveedor.Text, 1, 12)
'    frmDetalleTelCruTransf.xCod_Tela = TxtItem.Text
'    frmDetalleTelCruTransf.xCod_Comb = Mid(Label3, 1, 3)
'    frmDetalleTelCruTransf.xCod_Calidad = Mid(Label2, 1, 1)
'    If Mid(Label5, 1, 3) <> "-" Then
'       frmDetalleTelCruTransf.xCod_Medida = Mid(Label5, 1, 10)
'    Else
'       frmDetalleTelCruTransf.xCod_Medida = ""
'    End If
'
'    frmDetalleTelCruTransf.TxtLote = Me.TxtLote
'    frmDetalleTelCruTransf.TxtProveedor = Me.TxtProveedor
'    frmDetalleTelCruTransf.TxtItem = Me.TxtItem
'    frmDetalleTelCruTransf.LlenaDatos
'    frmDetalleTelCruTransf.Label2 = Label2
'    frmDetalleTelCruTransf.Label5 = Label5
'
'
'    frmDetalleTelCruTransf.xcod_almacen = Me.Cod_Almacen
'    frmDetalleTelCruTransf.xNum_MovStk = Me.Num_MovStk
'    frmDetalleTelCruTransf.Show vbModal
'
'    If frmDetalleTelCruTransf.bOk Then
'        Me.LoteaTransf = frmDetalleTelCruTransf.xlote
'        Me.Cod_ProveedoraTransf = frmDetalleTelCruTransf.xCod_Proveedor
'        Me.Cod_TelaaTransf = frmDetalleTelCruTransf.xCod_Tela
'        Me.Cod_CombaTransf = frmDetalleTelCruTransf.xCod_Comb
'        Me.Cod_CalidadaTransf = frmDetalleTelCruTransf.xCod_Calidad
'        Me.Cod_MedidaaTransf = frmDetalleTelCruTransf.xCod_Medida
'        bElijeDatos = True
'    Else
'        Me.LoteaTransf = ""
'        Me.Cod_ProveedoraTransf = ""
'        Me.Cod_TelaaTransf = ""
'        Me.Cod_CombaTransf = ""
'        Me.Cod_CalidadaTransf = ""
'        Me.Cod_MedidaaTransf = ""
'    End If
'
'    Set frmDetalleTelCruTransf = Nothing
    
End Sub

Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo hand
If Not Reg.EOF And Not Reg.BOF Then
    Me.txtBultos = Reg("Bultos")
    Me.TxtCantidad = Reg("Cant Movimiento")
    Me.TxtDesitem = Reg("Tela")
    Me.TxtItem = Reg("cod_tela")
    Me.TxtLote = Reg("lote")
    Me.TxtObs = Reg("Observaciones")
    Me.Txtproveedor = Reg("Proveedor")
    Cod_OrdTra = Reg("Cod_OrdTra")
    Cant_Anterior = Reg("Cant Movimiento")
    Num_Secuencia = Reg("secuencia")
    Cod_Proveedor = Reg("cod_proveedor")
    Label2.Caption = Reg("calidad")
    Label3.Caption = Reg("combinacion")
    Cod_Comb = Reg("Cod_Comb")
    Cod_Talla = Reg("Cod_Talla")
    Label5.Caption = Reg("medida")
    Me.txtCan_Movimiento_2daunimed.Text = Reg("Can_Movimiento_2daunimed").Value
    
    'Aqui cargaremos los nuevos valores para los labels de cantidades
    strSQL = "EXEC SM_ENCUENTRA_TIPOFAMTELA '" & Me.TxtItem.Text & "',''"
    varCod_TipFamTela = DevuelveCampo(strSQL, cConnect)
    If varCod_TipFamTela = "N" Then
        strSQL = "SELECT Cod_UniMed FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
        Me.lblCantidad1.Caption = DevuelveCampo(strSQL, cConnect)
        strSQL = "SELECT Cod_UniMedCnf FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
        Me.lblCantidad2.Caption = DevuelveCampo(strSQL, cConnect)
    Else
        strSQL = "SELECT Cod_UniMedCnf FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
        Me.lblCantidad1.Caption = DevuelveCampo(strSQL, cConnect)
        strSQL = "SELECT Cod_UniMed FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
        Me.lblCantidad2.Caption = DevuelveCampo(strSQL, cConnect)
    End If
    
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
strSQL = "select cod_clamov,Tip_Accion,TIP_PTMP,cod_tipordpro, cod_tipanx,Flg_Partidas_Tinto from lg_tiposmov where cod_tipmov='" & Cod_TipMovi & "'"

rs.Open strSQL, cConnect, adOpenStatic

If rs.RecordCount Then
    If rs("cod_clamov").Value = "E" And rs("Tip_ACcion").Value = "E" And (Trim(rs("TIP_PTMP").Value) = "PT" Or Trim(rs("Cod_TipOrdPro").Value) = "TR") And Trim(rs("cod_tipanx")) = "P" And Trim(rs("Flg_Partidas_Tinto")) <> "S" Then
        Set FRmLote.Padre = Me
        FRmLote.Cod_TipOrdTra = Me.Cod_TipOrdTra
        FRmLote.Cod_Proveedor = Me.Cod_Proveedor
        FRmLote.Grupo = DevuelveCampo("select cod_grupo from lg_ordcomp where ser_ordcomp='" & Ser_OrdComp & "' and cod_ordcomp='" & Cod_OrdComp & "'", cConnect)
        FRmLote.sSer_OrdComp = Me.Ser_OrdComp
        FRmLote.sCod_OrdComp = Me.Cod_OrdComp
        FRmLote.Show 1
        If Estado <> "NUEVO" Then
            'Estado = "NUEVO"
            Call MantFunc1_ActionClick(0, 0, "ADICIONAR")
        End If
        TxtLote.Text = NewLote
        NewLote = ""
    Else
        MsgBox "El Tipo de Movimiento no permite adicionar Lote", vbInformation, "Tela Cruda"
        NewLote = ""
    End If
End If

End Sub

'Private Sub lote_Click()
'If Cod_ClaMov = "E" And Cod_ClaOrdComp <> "" Then
'    Set FRmLote.Padre = Me
'    FRmLote.Cod_TipOrdTra = Me.Cod_TipOrdTra
'    FRmLote.Show 1
'End If
'End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo hand
Select Case ActionName
    Case "ADICIONAR"
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            Limpia
            Habilita
            Estado = "NUEVO"
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
        cmdPesosBal.Enabled = True
        TxtCantidad.SetFocus
    Case "ELIMINAR"
        If Me.varValida_Factura = False Then
            MsgBox "No se puede eliminar el registro por que el Mov. Almacen posee una factura relacionada.", vbInformation, "Mensaje"
            Exit Sub
        End If
        Datos "e", True
        Limpia
        Datos "v", False
        Deshabilita
    Case "GRABAR"
        If Trim(TxtCantidad) = "" Or TxtCantidad = "0" Then MsgBox "Llene la cantidad", vbInformation: Exit Sub
        'If Trim(TxtBultos) = "" Or TxtBultos  "0" Then MsgBox "Llene la cantidad de bultos", vbInformation: Exit Sub
        If Trim(txtBultos) = "" Then txtBultos = "0"
        If Trim(TxtItem) = "" Then MsgBox "Debe seleccionar un item", vbInformation: Exit Sub
        
        'Aqui haremos una validacion sobre cantidades
        strSQL = "EXEC SM_ENCUENTRA_TIPOFAMTELA '" & Me.TxtItem.Text & "',''"
        varCod_TipFamTela = DevuelveCampo(strSQL, cConnect)
        If varCod_TipFamTela = "N" Then
            If Val(Me.txtCan_Movimiento_2daunimed.Text) < 0 Then
                MsgBox "La 2da cantidad no puede ser menor que 0", vbInformation, "Mensaje"
                Me.txtCan_Movimiento_2daunimed.SetFocus
                Exit Sub
            End If
        Else
            If Val(Me.txtCan_Movimiento_2daunimed.Text) < 0 Then
                MsgBox "La cantidad no puede ser menor que 0", vbInformation, "Mensaje"
                Me.txtCan_Movimiento_2daunimed.SetFocus
                Exit Sub
            End If
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
    Case "DESHACER"
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
    If Trim(txtCan_Movimiento_2daunimed.Text) = "" Then
        txtCan_Movimiento_2daunimed.Text = "0"
    End If
End Sub

Private Sub TxtCantidad_GotFocus()
    TxtCantidad.SelStart = 0
    TxtCantidad.SelLength = Len(TxtCantidad.Text)
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
    SoloNumeros TxtCantidad, KeyAscii, True, 3, 6

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
End If

Exit Sub
hand:
    ErrorHandler err, "TxtItem"
End Sub

Private Sub txtLote_KeyPress(KeyAscii As Integer)
On Error GoTo ErrBusqLote
Dim rstAux As ADODB.Recordset, sErr As String
If KeyAscii = 13 Then
    If Cod_ClaMov = "E" And Cod_ClaOrdComp <> "" And Tip_PtMp = "PT" Then
        If DevuelveCampo("select  count(*) from   TX_ORDTRA  Where  COD_TIPORDTRA='" & Cod_TipOrdTra & "' and " & _
                        "   Cod_OrdProv like '%" & Me.TxtLote & "%' and cod_proveedor='" & Cod_Proveedor & "'", cConnect) <= 0 Then
            MsgBox "Este Lote no existe", vbInformation
        Else
            If sFlg_Ot_Tejeduria_Generada = "S" Then
                Set frmBusqGeneral.oParent = Me
                frmBusqGeneral.sQuery = "EXEC TJO_MUESTRA_OTS_GRUPO_PROVEEDOR_OC '" & _
                                        Cod_Almacen & "', '" & Num_MovStk & "'"
                frmBusqGeneral.Cargar_Datos
                frmBusqGeneral.Show 1
                TxtLote = Trim(DESCRIPCION)
                CODIGO = Trim(CODIGO)
                
                If CODIGO = "" Then Exit Sub
                
               ' If TxtLote = "" Then
                '    frmOtLote.Show vbModal
                 '   If frmOtLote.bCancel Then
                  '      Unload frmOtLote
                   '     Exit Sub
                    'End If
                   ' strSQL = "EXEC tjo_actualiza_lote_crudo_en_ot '" & Mid(Codigo, 4) & "', '" & _
                    '         frmOtLote.TxtLote & "'"
                  '  ExecuteSQL cConnect, strSQL
                   ' TxtLote = frmOtLote.TxtLote
                   ' frmOtLote.TxtLote = ""
                   ' Unload frmOtLote
               '  End If
            Else
                Set frmBusqGeneral.oParent = Me
                frmBusqGeneral.sQuery = " select  a.Cod_OrdProv as Orden, b.cod_Proveedor as [Cod Proveedor], b.des_proveedor as Descripcion " & _
                                        " from    TX_ORDTRA a,lg_proveedor b " & _
                                        " Where " & _
                                        " a.Cod_Proveedor=b.Cod_Proveedor and " & _
                                        " a.COD_TIPORDTRA='" & Cod_TipOrdTra & "' and " & _
                                        " a.cod_proveedor='" & Cod_Proveedor & "' and " & _
                                        " a.Cod_OrdProv like '%" & Me.TxtLote & "%'"
    '            frmBusqGeneral.sQuery = "EXEC SM_AYUDA_DEVOLUCION_TELA_CRUDA_DE_PARTIDAS '" & Cod_OrdTra1 & "'"
                frmBusqGeneral.Cargar_Datos
                frmBusqGeneral.Show 1
                TxtLote = CODIGO
                Cod_Proveedor = DESCRIPCION
            End If
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
        End If
    End If
End If
Cod_OrdTra = DevuelveCampo("select cod_ordtra from tx_ordtra where Cod_TipOrdTra='" & Cod_TipOrdTra & _
             "' and Cod_Proveedor='" & Cod_Proveedor & "' and Cod_OrdProv='" & TxtLote & "'", cConnect)
Me.Txtproveedor = DevuelveCampo("Select des_proveedor from lg_proveedor where cod_proveedor='" & Cod_Proveedor & "'", cConnect)
If KeyAscii = 13 Then SendKeys "{TAB}"
Exit Sub
ErrBusqLote:
    sErr = err.Description
    'Unload frmOtLote
    MsgBox sErr, vbCritical + vbOKOnly, "Busca Lote"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "VALORIZAR"
        FraValorizar.Visible = True
        TxtSoles.SetFocus
        
        TxtSoles.Text = DevuelveCampo("select imp_factura from lg_movistktelcru where cod_almacen ='" & Me.Cod_Almacen & "' and num_movstk='" & Me.Num_MovStk & "' and num_secuencia ='" & Reg("Secuencia") & "'", cConnect)
        TxtDolares.Text = DevuelveCampo("select imp_factura_dolares from lg_movistktelcru where cod_almacen ='" & Me.Cod_Almacen & "' and num_movstk='" & Me.Num_MovStk & "' and num_secuencia ='" & Reg("Secuencia") & "'", cConnect)
    Case "ROLLOS"
        
        frmVerDetRollos.sNum_Secuencia = Num_Secuencia
        frmVerDetRollos.sCod_Almacen = Cod_Almacen
        frmVerDetRollos.sNum_MovStk = Num_MovStk
        frmVerDetRollos.sCod_TipMov = Me.Cod_TipMovi
        frmVerDetRollos.Scod_ordtra = Cod_OrdPro
        frmVerDetRollos.Show 1
        Datos "V", False
        
        
        
    End Select
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "ACEPTAR"
        Call Grabar
        FraValorizar.Visible = False
    Case "CANCELAR"
        FraValorizar.Visible = False
    End Select
End Sub

Private Sub TxtDolares_GotFocus()
    SelectionText TxtDolares
End Sub

Private Sub TxtDolares_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        Call SoloNumeros(TxtDolares, KeyAscii, True, 2)
    End If
End Sub

Private Sub TxtObs_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub TxtSoles_GotFocus()
    SelectionText TxtSoles
End Sub

Private Sub TxtSoles_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        Call SoloNumeros(TxtSoles, KeyAscii, True, 2)
    End If
End Sub

Sub Grabar()
On Error GoTo errGrabar
Dim strSQL As String

If Trim(TxtSoles.Text) = "" Then
    TxtSoles.Text = "0"
End If

If Trim(TxtDolares.Text) = "" Then
    TxtDolares.Text = "0"
End If

strSQL = "EXEC LG_MovistkItem_Valoriza_Transferencia '" & Me.Cod_Almacen & "','" & Me.Num_MovStk & "','" & Reg("Secuencia") & "'," & _
        CDbl(TxtSoles.Text) & "," & CDbl(TxtDolares.Text)
ExecuteSQL cConnect, strSQL

TxtSoles.Text = ""
TxtDolares.Text = ""
Exit Sub
errGrabar:
    ErrorHandler err, "Grabar"
End Sub
