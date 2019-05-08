VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form FrmDetalleHilTel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Hilos Teñidos"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   9630
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
      Left            =   3105
      TabIndex        =   32
      Top             =   1440
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox TxtSoles 
         Height          =   285
         Left            =   2160
         TabIndex        =   33
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TxtDolares 
         Height          =   285
         Left            =   2160
         TabIndex        =   34
         Top             =   720
         Width           =   975
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   510
         Left            =   600
         TabIndex        =   35
         Top             =   1200
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"FrmDetalleHilTel.frx":0000
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Soles"
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   37
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Importe Dolares"
         Height          =   255
         Left            =   600
         TabIndex        =   36
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2040
      TabIndex        =   21
      Top             =   5475
      Width           =   1965
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "FrmDetalleHilTel.frx":0096
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Anterior"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "FrmDetalleHilTel.frx":0208
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Siguiente"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   0
         Picture         =   "FrmDetalleHilTel.frx":037A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Primero"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "FrmDetalleHilTel.frx":04EC
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Ultimo"
         Top             =   0
         Width           =   495
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
      Height          =   1890
      Left            =   30
      TabIndex        =   14
      Tag             =   "Detail"
      Top             =   3390
      Width           =   9510
      Begin VB.CommandButton cmdPesosBal 
         Caption         =   "..."
         Height          =   285
         Left            =   7800
         TabIndex        =   31
         Top             =   195
         Width           =   345
      End
      Begin VB.TextBox txtConos 
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
         Top             =   870
         Width           =   945
      End
      Begin VB.TextBox txtCod_OrdTra 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2745
         TabIndex        =   4
         Top             =   1575
         Width           =   825
      End
      Begin VB.TextBox TxtColor 
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
         TabIndex        =   3
         Top             =   1170
         Width           =   1935
      End
      Begin VB.TextBox TxtDesitem 
         Height          =   315
         Left            =   2130
         TabIndex        =   2
         Top             =   840
         Width           =   1665
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
         Left            =   1170
         MaxLength       =   8
         TabIndex        =   1
         Top             =   840
         Width           =   945
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
         TabIndex        =   5
         Text            =   "0"
         Top             =   180
         Width           =   945
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
         Width           =   2505
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
         TabIndex        =   10
         Top             =   510
         Width           =   3810
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
         TabIndex        =   6
         Text            =   "0"
         Top             =   510
         Width           =   945
      End
      Begin VB.TextBox TxtObs 
         Height          =   585
         Left            =   6780
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Text            =   "FrmDetalleHilTel.frx":065E
         Top             =   1230
         Width           =   2385
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Conos:"
         Height          =   195
         Index           =   3
         Left            =   5520
         TabIndex        =   30
         Top             =   915
         Width           =   495
      End
      Begin VB.Label lblCod_OrdTra 
         Caption         =   "O/T:"
         Height          =   195
         Left            =   2250
         TabIndex        =   29
         Top             =   1590
         Width           =   405
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
         TabIndex        =   28
         Top             =   1590
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Calidad:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   27
         Top             =   1590
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Color:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   1275
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad Mov:"
         Height          =   195
         Index           =   7
         Left            =   5520
         TabIndex        =   20
         Top             =   255
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hilo:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   945
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lote:"
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   18
         Top             =   255
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   17
         Top             =   585
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad Bultos:"
         Height          =   195
         Index           =   8
         Left            =   5520
         TabIndex        =   16
         Top             =   585
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   9
         Left            =   5520
         TabIndex        =   15
         Top             =   1275
         Width           =   1065
      End
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
      TabIndex        =   12
      Tag             =   "List"
      Top             =   60
      Width           =   9495
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2775
         Left            =   180
         TabIndex        =   13
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
   Begin VB.CommandButton lote 
      Caption         =   "&Ingresar Lote"
      Height          =   525
      Left            =   8280
      TabIndex        =   11
      Top             =   5400
      Width           =   1245
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   4020
      TabIndex        =   9
      Top             =   5430
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmDetalleHilTel.frx":0664
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   495
      Left            =   225
      TabIndex        =   38
      Top             =   5445
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      Custom          =   "0~0~VALORIZAR~Verdadero~Verdadero~&Valorizar Transferencias~0~0~1~~0~Falso~Falso~&Valorizar Transferencias~"
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmDetalleHilTel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Codigo
Public Descripcion
Public Cod_color As String
Public NewLote As String

'Public lote As String
Public Paso As Boolean
Dim Tip_item As String
Dim Tip_presentacion As String
Dim Cod_Calidad As String
Dim Cant_Anterior As Double

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
Public Flg_Partida_Generada As String
Public varCod_TipOrdTra As String
Public varCod_OrdTra As String

Dim Reg As New ADODB.Recordset
Dim Estado As String
Dim Num_Secuencia As String
Public Cod_ClaMov As String

Public varValida_Factura As Boolean

Sub ValidaHilo()
    Dim Temp
    
If Cod_ClaMov = "E" Then 'Entrada
    'Tiene O/C
    If Cod_ClaOrdComp <> "" Then
        If Tip_PtMp = "PT" Then
            If Flg_Partida_Generada = "S" Then
                BuscaHilo
            Else
                Set frmBusqGeneral.oParent = Me
                frmBusqGeneral.sQuery = "UP_AyudaHilten '1','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
                
                frmBusqGeneral.Cargar_Datos
                frmBusqGeneral.Show 1
                If Paso = True Then
                    TxtItem = Codigo
                    Sec_OrdComp = Descripcion
                    TxtDesitem = DevuelveCampo("select des_hiltel  from it_hilado where cod_hiltel='" & Codigo & "'", cConnect)
                End If
                TxtColor.Enabled = False
                TxtColor.Text = DevuelveCampo("select des_color from lb_color where cod_color='" & Cod_color & "'", cConnect)
            End If
        Else
            Set frmBusqGeneral.oParent = Me
            frmBusqGeneral.sQuery = "UP_AyudaHilten '3','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
            frmBusqGeneral.Cargar_Datos
            frmBusqGeneral.Show 1
            If Paso = True Then
                TxtItem = Codigo
                TxtDesitem = Descripcion
                Sec_OrdComp = DevuelveCampo("select sec_ordcomp from lg_ordcompitem where Ser_OrdComp='" & Ser_OrdComp & "' and Cod_OrdComp='" & Cod_OrdComp & "' and Cod_Item='" & Me.TxtItem & "' and cod_color='" & Me.Cod_color & "'", cConnect)
            End If
        End If
        TxtColor.Enabled = False
        TxtColor.Text = DevuelveCampo("select des_color from lb_color where cod_color='" & Cod_color & "'", cConnect)
    Else
        Set frmBusqGeneral.oParent = Me
        frmBusqGeneral.sQuery = "UP_AyudaHilten '4','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
        frmBusqGeneral.Cargar_Datos
        frmBusqGeneral.Show 1
        If Paso = True Then
            TxtItem = Codigo
            TxtDesitem = Descripcion
            Sec_OrdComp = DevuelveCampo("select sec_ordcomp from lg_ordcompitem where Ser_OrdComp='" & Ser_OrdComp & "' and Cod_OrdComp='" & Cod_OrdComp & "' and Cod_Item='" & Me.TxtItem & "'", cConnect)
            TxtColor.Enabled = True
        End If
        
    End If
Else
'salida
    'Tiene O/C
    If Cod_ClaOrdComp <> "" Then
        If Tip_PtMp = "PT" Then
            If Flg_Partida_Generada = "S" Then
                BuscaHilo
            Else
                Set frmBusqGeneral.oParent = Me
                frmBusqGeneral.sQuery = "UP_AyudaHilten '1','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
    
                frmBusqGeneral.Cargar_Datos
                frmBusqGeneral.Show 1
                If Paso = True Then
                    TxtItem = Codigo
                    Sec_OrdComp = Descripcion
                    TxtDesitem = DevuelveCampo("select des_hiltel  from it_hilado where cod_hiltel='" & Codigo & "'", cConnect)
                End If
                TxtColor.Enabled = False
            End If
        Else
            Set frmBusqGeneral.oParent = Me
            frmBusqGeneral.sQuery = "UP_AyudaHilten '2','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
            frmBusqGeneral.Cargar_Datos
            frmBusqGeneral.Show 1
            If Paso = True Then
                TxtItem = Codigo
                TxtDesitem = Descripcion
            End If
            TxtColor.Enabled = False
        End If
    Else
        Set frmBusqGeneral.oParent = Me
        frmBusqGeneral.sQuery = "UP_AyudaHilten '2','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
        frmBusqGeneral.Cargar_Datos
        frmBusqGeneral.Show 1
        If Paso = True Then
            TxtItem = Codigo
            TxtDesitem = Descripcion
        End If
        TxtColor.Enabled = False
    End If
    TxtColor.Text = DevuelveCampo("select des_color from lb_color where cod_color='" & Cod_color & "'", cConnect)
End If
    
' TxtColor.Text = DevuelveCampo("select des_color from lb_color where cod_color='" & Cod_color & "'", cCONNECT)
    
'    If Cod_ClaMov = "E" And Cod_ClaOrdComp <> "" Then
'
'            Set frmBusqGeneral.oParent = Me
'            frmBusqGeneral.sQuery = "UP_AyudaHilten '1','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
'
'            frmBusqGeneral.CARGAR_DATOS
'            frmBusqGeneral.Show 1
'            If Paso = True Then
'                TxtItem = Codigo
'                Sec_OrdComp = Descripcion
'                TxtDesitem = DevuelveCampo("select des_hiltel  from it_hilado where cod_hiltel='" & Codigo & "'", cCONNECT)
'            End If
'
'    Else
'        If Cod_ClaMov = "S" Then
'            Set frmBusqGeneral.oParent = Me
'            frmBusqGeneral.sQuery = "UP_AyudaHilten '2','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
'            frmBusqGeneral.CARGAR_DATOS
'            frmBusqGeneral.Show 1
'            If Paso = True Then
'                TxtItem = Codigo
'                TxtDesitem = Descripcion
''                TxtColor = Des_color
'            End If
'
'        ElseIf Cod_ClaMov = "E" Then
'            Set frmBusqGeneral.oParent = Me
'            frmBusqGeneral.sQuery = "UP_AyudaHilten '3','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
'            frmBusqGeneral.CARGAR_DATOS
'            frmBusqGeneral.Show 1
'            If Paso = True Then
'                TxtItem = Codigo
'                TxtDesitem = Descripcion
'                Sec_OrdComp = DevuelveCampo("select sec_ordcomp from lg_ordcompitem where Ser_OrdComp='" & Ser_OrdComp & "' and Cod_OrdComp='" & Cod_OrdComp & "' and Cod_Item='" & Me.TxtItem & "'", cCONNECT)
'                '**
'            End If
'
'        ElseIf Cod_ClaOrdComp <> "" Then
'            Set frmBusqGeneral.oParent = Me
'            frmBusqGeneral.sQuery = "UP_AyudaHilten '4','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
'            frmBusqGeneral.CARGAR_DATOS
'            frmBusqGeneral.Show 1
'            If Paso = True Then
'                TxtItem = Codigo
'                Sec_OrdComp = Descripcion
'                TxtDesitem = DevuelveCampo("select des_hiltel from it_hilado where cod_hiltel='" & Codigo & "'", cCONNECT)
'                TxtColor = DevuelveCampo("select a.des_color from lb_color a,lg_ordcompitem b where a.cod_color=b.cod_color and b.Cod_OrdComp='" & Cod_OrdComp & "' and Sec_OrdComp='" & Sec_OrdComp & "' and Cod_Item ='" & Codigo & "'", cCONNECT)
'                Cod_color = DevuelveCampo("select a.cod_color from lb_color a,lg_ordcompitem b where a.cod_color=b.cod_color and b.Cod_OrdComp='" & Cod_OrdComp & "' and Sec_OrdComp='" & Sec_OrdComp & "' and Cod_Item ='" & Codigo & "'", cCONNECT)
'            End If
'        End If
'
'End If
'
''TxtColor = DevuelveCampo("select a.des_color from lb_color a,lg_ordcompitem b where a.cod_color=b.cod_color and b.Cod_OrdComp='" & Cod_OrdComp & "' and Sec_OrdComp='" & Sec_OrdComp & "' and Cod_Item ='" & Codigo & "'", cCONNECT)
''Cod_color = DevuelveCampo("select a.cod_color from lb_color a,lg_ordcompitem b where a.cod_color=b.cod_color and b.Cod_OrdComp='" & Cod_OrdComp & "' and Sec_OrdComp='" & Sec_OrdComp & "' and Cod_Item ='" & Codigo & "'", cCONNECT)
End Sub


Public Sub Datos(Accion As String, EsAccion As Boolean)
On Error GoTo hand
Set Reg = Nothing
Reg.CursorLocation = adUseClient
If txtConos = "" Then
txtConos = 0
End If


If Accion = "V" Then
    Reg.Open "UP_Lg_MoviStkHilTen '" & Accion & "','" & Cod_Almacen & "','" & Num_MovStk & "'", cConnect
Else
    Reg.Open "UP_ACT_STOCKSHILTEN '" & Cod_Almacen & "','" & Num_MovStk & "','" & Accion & "','" & Num_Secuencia & "','" & _
            Trim(TxtLote) & "','" & Cod_Proveedor & "','" & TxtItem & "','" & Cod_color & "'," & Me.TxtCantidad & "," & Me.TxtBultos & ",'" & TxtObs.Text & "'," & _
            Cant_Anterior & ",'" & Sec_OrdComp & "','" & vusu & "', '" & _
            txtCod_OrdTra & "', " & txtConos, cConnect
End If
If EsAccion = False Then
    Set Me.DGridLista.DataSource = Reg
    DGridLista_RowColChange 0, 0
    Me.DGridLista.Columns("Cod_OrdTra").Visible = False
    Me.DGridLista.Columns("cod_proveedor").Visible = False
    Me.DGridLista.Columns("cod_color").Visible = False
Me.DGridLista.Columns("calidad").Visible = False
End If
Exit Sub
hand:
ErrorHandler err, "Datos"
End Sub

Sub Habilita()
TxtItem.Enabled = True
TxtDesitem.Enabled = True
Me.TxtCantidad.Enabled = True


Me.TxtBultos.Enabled = True
txtConos.Enabled = True
Me.TxtLote.Enabled = True
Me.TxtObs.Enabled = True
'Me.txtCod_Ordtra.Enabled = True
cmdPesosBal.Enabled = True
End Sub
Sub Deshabilita()
TxtItem.Enabled = False
TxtDesitem.Enabled = False
Me.TxtCantidad.Enabled = False

Me.TxtBultos.Enabled = False
txtConos.Enabled = False
Me.TxtLote.Enabled = False
Me.TxtObs.Enabled = False
'Me.txtCod_Ordtra.Enabled = False
cmdPesosBal.Enabled = False
End Sub

Sub Limpia()
TxtItem = ""
TxtDesitem = ""
Me.TxtCantidad = "0"
Me.TxtColor = ""
Me.TxtBultos = "0"
txtConos = "0"
Me.TxtLote = ""
Me.TxtObs = ""
'Me.txtCod_Ordtra = ""
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

Private Sub cmdPesosBal_Click()
    frmPesosBal.Show vbModal
    TxtCantidad = frmPesosBal.lblTotal
    Unload frmPesosBal
End Sub

Private Sub cmdPrevious_Click()
If Not Reg.BOF Then Reg.MovePrevious
End Sub

Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

If Not Reg.EOF And Not Reg.BOF Then
    Me.TxtBultos = Reg("Bultos")
    Me.txtConos = Reg("Conos")
    Me.TxtCantidad = Reg("Cant Movimiento")
    Me.TxtDesitem = Reg("Hilo")
    Me.TxtItem = Reg("Cod Hilo")
    Me.TxtLote = Reg("lote")
    Me.TxtObs = Reg("Observaciones")
    Me.TxtProveedor = Reg("Proveedor")
    Me.TxtColor = Reg("Color")
    Cod_color = Reg("Cod_Color")
    Cod_OrdTra = Reg("Cod_OrdTra")
    Cant_Anterior = Reg("Cant Movimiento")
    Num_Secuencia = Reg("secuencia")
    Cod_Proveedor = Reg("cod_proveedor")
    Label2.Caption = Reg("calidad")
    txtCod_OrdTra = Reg("OT")
End If
End Sub


Private Sub Form_Load()

Tip_item = DevuelveCampo("select tip_item from lg_almacen where cod_almacen='" & Me.Cod_Almacen & "'", cConnect)
Tip_presentacion = DevuelveCampo("select Tip_presentacion from lg_almacen where cod_almacen='" & Me.Cod_Almacen & "'", cConnect)
Cod_Calidad = DevuelveCampo("select isnull(Cod_Calidad,'') from lg_tiposmov where cod_tipmov='" & Me.Cod_TipMovi & "'", cConnect)
Cod_TipOrdTra = DevuelveCampo("select cod_tipordtra from Tx_TiposOrdTra where tip_item='" & Tip_item & "' and tip_presentacion='" & Tip_presentacion & "'", cConnect)
Me.TxtProveedor = DevuelveCampo("Select des_proveedor from lg_proveedor where cod_proveedor='" & Cod_Proveedor & "'", cConnect)
Limpia
Deshabilita
FormateaGrid Me.DGridLista
Datos "V", False
End Sub


Private Sub lote_Click()
Dim StrSql As String
Dim Rs As ADODB.Recordset

Set Rs = New ADODB.Recordset
StrSql = "select cod_clamov,Tip_Accion,TIP_PTMP,cod_tipanx,Flg_Partidas_Tinto from lg_tiposmov where cod_tipmov='" & Cod_TipMovi & "'"

Rs.Open StrSql, cConnect, adOpenStatic

If Rs.RecordCount Then
    If Rs("cod_clamov").Value = "E" And Rs("Tip_ACcion").Value = "E" And Trim(Rs("TIP_PTMP").Value) = "PT" And Trim(Rs("cod_tipanx")) = "P" And Trim(Rs("Flg_Partidas_Tinto")) <> "S" Then
        Set FRmLote.Padre = Me
        FRmLote.Cod_TipOrdTra = Me.Cod_TipOrdTra
        FRmLote.Cod_Proveedor = Me.Cod_Proveedor
        FRmLote.sSer_OrdComp = Me.Ser_OrdComp
        FRmLote.sCod_OrdComp = Me.Cod_OrdComp
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
'If Cod_ClaMov = "E" And Cod_ClaOrdComp <> "" Then
'    Set FRmLote.Padre = Me
'    FRmLote.Cod_TipOrdTra = Me.Cod_TipOrdTra
'    FRmLote.Show 1
'End If
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo hand
Select Case ActionName
    Case "ADICIONAR"
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            Limpia
            Habilita
            Estado = "NUEVO"
            Me.TxtLote.SetFocus
    Case "MODIFICAR"
        
        If Me.varValida_Factura = False Then
            MsgBox "No se puede modificar el registro por que el Mov. Almacen posee una factura relacionada.", vbInformation, "Mensaje"
            Exit Sub
        End If
        
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Estado = "MODIFICAR"
        Deshabilita
        TxtCantidad.Enabled = True
        Me.TxtBultos.Enabled = True
        txtConos.Enabled = True
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
        Datos "V", False
        Deshabilita
    Case "GRABAR"
        If Trim(TxtCantidad) = "" Or TxtCantidad = "0" Then MsgBox "Llene la cantidad", vbInformation: Exit Sub
        If Trim(TxtBultos) = "" Or TxtBultos <= "0" Then MsgBox "Llene la cantidad", vbInformation: Exit Sub
    '    If Trim(txtConos) = "" Or txtConos <= "0" Then MsgBox "se debe especficar conos", vbInformation: Exit Sub
        
        If Trim(TxtItem) = "" Then MsgBox "Debe seleccionar un item", vbInformation: Exit Sub
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
    TxtBultos.SelStart = 0
    TxtBultos.SelLength = Len(TxtBultos.Text)
End Sub

Private Sub txtBultos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
    SoloNumeros TxtBultos, KeyAscii, False, 0, 4
End Sub


Private Sub TxtCantidad_GotFocus()
    TxtCantidad.SelStart = 0
    TxtCantidad.SelLength = Len(TxtCantidad.Text)
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
    SoloNumeros TxtCantidad, KeyAscii, True, 3, 6

End Sub

Private Sub txtCod_OrdTra_KeyPress(KeyAscii As Integer)
Dim StrSql As String
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    Codigo = ""
    Descripcion = ""
    StrSql = "EXEC TJO_MUESTRA_OTS_TEJEDURIA_GRUPO_PROVEEDOR '" & _
             Ser_OrdComp & "', '" & Cod_OrdComp & "'"
    Set frmBusqGeneral.oParent = Me
    frmBusqGeneral.sQuery = StrSql
    frmBusqGeneral.Cargar_Datos
    frmBusqGeneral.gexList.Columns("Cod_TipOrdTra").Width = 500
    frmBusqGeneral.gexList.Columns("Cod_Ordtra").Width = 4995
    frmBusqGeneral.gexList.Columns("Cod_TipOrdTra").Caption = "Tipo"
    frmBusqGeneral.gexList.Columns("Cod_Ordtra").Caption = "Orden de Trabajo O/T"
    frmBusqGeneral.Show 1
    txtCod_OrdTra = Descripcion
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtColor_KeyPress(KeyAscii As Integer)
On Error GoTo hand
Dim cadena As String
If KeyAscii = 13 Then
    If Trim(TxtColor) = "" Then
        cadena = "select cod_color as Codigo, des_color as Descripcion from lb_color ORDER BY des_color"
    Else
        cadena = "select cod_color as Codigo, des_color as Descripcion from lb_color where des_color like '" & Trim(TxtColor.Text) & "%' ORDER BY des_color"
    End If
        Set frmBusqGeneral.oParent = Me
        frmBusqGeneral.sQuery = cadena
        frmBusqGeneral.Cargar_Datos
        frmBusqGeneral.Show 1
        TxtColor = Descripcion
        Cod_color = Codigo
        Codigo = ""
        Descripcion = ""
End If

Exit Sub

hand:
ErrorHandler err, "TxtColor_KeyPress"
End Sub

Private Sub txtConos_GotFocus()
    SelectionText txtConos
End Sub

Private Sub txtConos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    SoloNumeros TxtBultos, KeyAscii, False, 0, 4
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
If KeyAscii = 13 Then
    If Cod_ClaMov = "E" And Cod_ClaOrdComp <> "" And Tip_PtMp = "PT" Then
        If Flg_Partida_Generada = "S" Then
            BusqPartidasLote
        Else
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
                                        " a.cod_proveedor='" & Cod_Proveedor & "' and " & _
                                        " a.Cod_OrdProv like '%" & Me.TxtLote & "%'"
                frmBusqGeneral.Cargar_Datos
                frmBusqGeneral.Show 1
                TxtLote = Codigo
                Cod_Proveedor = Descripcion
            End If
        End If
    Else
        If Flg_Partida_Generada = "S" Then
            BusqPartidasLote
        ElseIf sFlg_Ot_Tejeduria_Generada = "S" Then
            txtCod_OrdTra = Trim(txtCod_OrdTra)
            If txtCod_OrdTra = "" Then
                MsgBox "Se debe especificar Una O/T", vbExclamation + vbOKOnly, "Busca Lote (Salida sFlg_Ot_Tejeduria_Generada = 'S')"
                If txtCod_OrdTra.Visible Then txtCod_OrdTra.SetFocus
                Exit Sub
            End If
            Set frmBusqGeneral.oParent = Me
            frmBusqGeneral.sQuery = "SELECT a.Cod_OrdProv AS Orden, " & _
            "b.Cod_Proveedor AS [Cod Prov], a.Cod_Color AS [Cod Color], c.Des_Proveedor AS Proveedor " & _
            "FROM  tj_ordtra_hilos_lotes a, LG_STOCKSHILTEN b, LG_PROVEEDOR c " & _
            "WHERE a.Cod_TipOrdTra = 'TJ' " & _
            "AND   a.Cod_OrdTra = '" & txtCod_OrdTra & "' " & _
            "AND   a.Cod_Almacen = '" & Cod_Almacen & "' " & _
            "AND   a.Cod_Almacen = b.Cod_Almacen " & _
            "AND   a.Cod_OrdProv = b.Cod_OrdProv " & _
            "AND   a.Cod_Hiltel = b.Cod_Hiltel " & _
            "AND   a.Cod_Color = b.Cod_Color " & _
            "AND   b.Cod_Proveedor = c.Cod_Proveedor " & _
            "GROUP BY a.Cod_OrdProv, b.Cod_Proveedor, a.Cod_Color, c.Des_Proveedor"
            frmBusqGeneral.Cargar_Datos
            frmBusqGeneral.Show 1
            TxtLote = Codigo
            Cod_Proveedor = Descripcion
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
                TxtLote = Codigo
                Cod_Proveedor = Descripcion
            End If
        End If
    End If
End If
Cod_OrdTra = DevuelveCampo("select cod_ordtra from tx_ordtra where Cod_TipOrdTra='" & Cod_TipOrdTra & _
        "' and Cod_Proveedor='" & Cod_Proveedor & "' and Cod_OrdProv='" & TxtLote & "'", cConnect)
Me.TxtProveedor = DevuelveCampo("Select des_proveedor from lg_proveedor where cod_proveedor='" & Cod_Proveedor & "'", cConnect)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub BusqPartidasLote()
    Load frmBusqPartidasLote
    frmBusqPartidasLote.varCod_OrdComp = Cod_OrdComp
    frmBusqPartidasLote.varSer_OrdComp = Ser_OrdComp
    frmBusqPartidasLote.CARGA_GRID
    Set frmBusqPartidasLote.oParent = Me
    frmBusqPartidasLote.Show 1
    Set frmBusqPartidasLote = Nothing
    Cod_TipOrdTra = varCod_TipOrdTra
    Cod_OrdTra = varCod_OrdTra
End Sub

Private Sub BuscaHilo()
With frmBusqGeneral3
    .Caption = "Seleccionar Hilo"
    .sQuery = "EXEC TH_UP_SEL_HILOS_TH_ORDTRA_ITEMS '" & _
    Cod_TipOrdTra & "', '" & Cod_OrdTra & "'"
    .Cargar_Datos
    
    .gexLista.Columns("Cod_TipOrdTra").Visible = False
    .gexLista.Columns("Cod_OrdTra").Visible = False
    .gexLista.Columns("Cod_HilTel").Visible = False
    .gexLista.Columns("Des_HilTel").Visible = False
    .gexLista.Columns("Cod_Color").Visible = False
    .gexLista.Columns("Des_Color").Visible = False
    .gexLista.Columns("Des_Color").Visible = False
    
    .gexLista.Columns("Num_Secuencia").Width = 360
    .gexLista.Columns("HILADO").Width = 2430
    .gexLista.Columns("COLOR").Width = 1965
    .gexLista.Columns("Bultos_Env.").Width = 990
    .gexLista.Columns("Bultos_Rec").Width = 960
    .gexLista.Columns("kgs_crudo").Width = 885
    .gexLista.Columns("kgs_tenidos_1ras").Width = 1380
    
    .gexLista.Columns("Num_Secuencia").Caption = "Sec."
    .gexLista.Columns("HILADO").Caption = "Hilado"
    .gexLista.Columns("COLOR").Caption = "Color"
    .gexLista.Columns("Bultos_Env.").Caption = "Bultos Env."
    .gexLista.Columns("Bultos_Rec").Caption = "Bultos Recib."
    .gexLista.Columns("kgs_crudo").Caption = "Kgs.Crudo"
    .gexLista.Columns("kgs_tenidos_1ras").Caption = "Kgs.Tenidos 1ras."
    
    .Show vbModal
    
    If .gexLista.RowCount > 0 And Not .bCancel Then
        TxtItem = .gexLista.Value(.gexLista.Columns("Cod_HilTel").Index)
        TxtDesitem = .gexLista.Value(.gexLista.Columns("Des_HilTel").Index)
        Cod_color = .gexLista.Value(.gexLista.Columns("Cod_Color").Index)
        TxtColor.Enabled = False
        TxtColor = .gexLista.Value(.gexLista.Columns("Des_Color").Index)
    End If
End With
Unload frmBusqGeneral3
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "VALORIZAR"
        FraValorizar.Visible = True
        TxtSoles.SetFocus
        
        TxtSoles.Text = DevuelveCampo("select imp_factura from lg_movistkhilten where cod_almacen ='" & Me.Cod_Almacen & "' and num_movstk='" & Me.Num_MovStk & "' and num_secuencia ='" & Reg("Secuencia") & "'", cConnect)
        TxtDolares.Text = DevuelveCampo("select imp_factura_dolares from lg_movistkhilten where cod_almacen ='" & Me.Cod_Almacen & "' and num_movstk='" & Me.Num_MovStk & "' and num_secuencia ='" & Reg("Secuencia") & "'", cConnect)
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
Dim StrSql As String

If Trim(TxtSoles.Text) = "" Then
    TxtSoles.Text = "0"
End If

If Trim(TxtDolares.Text) = "" Then
    TxtDolares.Text = "0"
End If

StrSql = "EXEC LG_MovistkItem_Valoriza_Transferencia '" & Me.Cod_Almacen & "','" & Me.Num_MovStk & "','" & Reg("Secuencia") & "'," & _
        CDbl(TxtSoles.Text) & "," & CDbl(TxtDolares.Text)
ExecuteSQL cConnect, StrSql

TxtSoles.Text = ""
TxtDolares.Text = ""
Exit Sub
errGrabar:
    ErrorHandler err, "Grabar"
End Sub

