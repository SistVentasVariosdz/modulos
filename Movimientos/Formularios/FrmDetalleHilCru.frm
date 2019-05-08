VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form FrmDetalleHilCru 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Hilos Crudos"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9660
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
      Top             =   1125
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
         Custom          =   $"FrmDetalleHilCru.frx":0000
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
         Index           =   5
         Left            =   600
         TabIndex        =   38
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton lote 
      Caption         =   "&Ingresar Lote"
      Height          =   525
      Left            =   8295
      TabIndex        =   25
      Top             =   5115
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
      Left            =   45
      TabIndex        =   18
      Tag             =   "List"
      Top             =   90
      Width           =   9525
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2775
         Left            =   180
         TabIndex        =   19
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
      Height          =   1620
      Left            =   60
      TabIndex        =   15
      Tag             =   "Detail"
      Top             =   3315
      Width           =   9510
      Begin VB.CommandButton cmdPesosBal 
         Caption         =   "..."
         Height          =   285
         Left            =   7830
         TabIndex        =   32
         Top             =   210
         Width           =   345
      End
      Begin VB.TextBox txtLote_Destino 
         Height          =   285
         Left            =   3690
         TabIndex        =   5
         Top             =   1245
         Width           =   1605
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
         TabIndex        =   8
         Text            =   "0"
         Top             =   840
         Width           =   945
      End
      Begin VB.TextBox txtCod_OrdTra 
         Height          =   285
         Left            =   2355
         TabIndex        =   4
         Top             =   1245
         Width           =   825
      End
      Begin VB.CommandButton cmdGetInfo 
         Height          =   285
         Left            =   3495
         Picture         =   "FrmDetalleHilCru.frx":0096
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Seleccionar Datos por Tela"
         Top             =   195
         Width           =   375
      End
      Begin VB.TextBox TxtObs 
         Height          =   315
         Left            =   6780
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
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
         Top             =   510
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
         Width           =   4005
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
         Width           =   2235
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
         TabIndex        =   6
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
         Left            =   1170
         MaxLength       =   8
         TabIndex        =   2
         Top             =   840
         Width           =   945
      End
      Begin VB.TextBox TxtDesitem 
         Height          =   315
         Left            =   2100
         TabIndex        =   3
         Top             =   840
         Width           =   2985
      End
      Begin VB.Label Label3 
         Caption         =   "Lote:"
         Height          =   195
         Left            =   3270
         TabIndex        =   31
         Top             =   1290
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Conos:"
         Height          =   195
         Index           =   2
         Left            =   5520
         TabIndex        =   30
         Top             =   915
         Width           =   495
      End
      Begin VB.Label lblCod_OrdTra 
         Caption         =   "O/T:"
         Height          =   195
         Left            =   1935
         TabIndex        =   29
         Top             =   1290
         Width           =   375
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
         Height          =   225
         Left            =   1200
         TabIndex        =   27
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Calidad:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   1290
         Width           =   570
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
         Top             =   585
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
         Caption         =   "Hilo:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   945
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad Mov:"
         Height          =   195
         Index           =   7
         Left            =   5520
         TabIndex        =   16
         Top             =   255
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2040
      TabIndex        =   10
      Top             =   5160
      Width           =   1965
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "FrmDetalleHilCru.frx":03A0
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ultimo"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "FrmDetalleHilCru.frx":0512
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Siguiente"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "FrmDetalleHilCru.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Anterior"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   0
         Picture         =   "FrmDetalleHilCru.frx":07F6
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Primero"
         Top             =   0
         Width           =   495
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   4305
      TabIndex        =   20
      Top             =   5115
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmDetalleHilCru.frx":0968
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   495
      Left            =   180
      TabIndex        =   33
      Top             =   5130
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
Attribute VB_Name = "FrmDetalleHilCru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Codigo
Public Descripcion
Public NewLote As String

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
Public Cod_TipOrdTra1 As String
Public Cod_OrdTra1 As String, bCancelSec As Boolean

Dim Reg As New ADODB.Recordset
Dim Estado As String, StrSql As String
Dim Num_Secuencia As String
Dim Num_Secuencia_OrdTra_Tinto As String
Public Cod_ClaMov As String, sTit As String, sErr As String

Public varValida_Factura As Boolean

Sub ValidaHilo()
Dim Temp
    If Cod_ClaMov = "E" And Cod_ClaOrdComp <> "" And Tip_PtMp = "PT" Then
        Set frmBusqGeneral.oParent = Me
        frmBusqGeneral.sQuery = "UP_AyudaHilCru '1','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
        
        frmBusqGeneral.Cargar_Datos
        frmBusqGeneral.Show 1
        If Paso = True Then
            TxtItem = Codigo
            Sec_OrdComp = Descripcion
            TxtDesitem = DevuelveCampo("select des_hiltel  from it_hilado where cod_hiltel='" & Codigo & "'", cConnect)
        End If
    ElseIf Cod_ClaMov = "E" And Cod_ClaOrdComp = "" And Tip_PtMp = "PT" And Cod_TipOrdPro = "HI" Then
        Set frmBusqGeneral.oParent = Me
        frmBusqGeneral.sQuery = "UP_AyudaHilCru '5','" & Cod_Almacen & "','" & _
        Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & _
        TxtLote & "','" & Cod_Proveedor & "', '" & TxtItem & "', '" & TxtDesitem & "'"
        
        frmBusqGeneral.Cargar_Datos
        frmBusqGeneral.Show 1
        If Paso = True Then
            TxtItem = Codigo
            Sec_OrdComp = Descripcion
            TxtDesitem = DevuelveCampo("select des_hiltel  from it_hilado where cod_hiltel='" & Codigo & "'", cConnect)
        End If
    Else
        If Cod_ClaMov = "S" Then
            Set frmBusqGeneral.oParent = Me
            frmBusqGeneral.sQuery = "UP_AyudaHilCru '2','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
            frmBusqGeneral.Cargar_Datos
            frmBusqGeneral.Show 1
            If Paso = True Then
                TxtItem = Codigo
                TxtDesitem = Descripcion
            End If
        
           If Cod_ClaOrdComp <> "" And Tip_PtMp = "PT" Then
            Set frmBusqGeneral.oParent = Me
            frmBusqGeneral.sQuery = "UP_AyudaHilCru '4','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
            frmBusqGeneral.Cargar_Datos
            frmBusqGeneral.Show 1
                If Paso = True Then
    '                TxtItem = Codigo
                    Sec_OrdComp = Descripcion
     '               TxtDesitem = DevuelveCampo("select des_hiltel from it_hilado where cod_hiltel='" & Codigo & "'", cCONNECT)
                End If
            End If
        ElseIf Cod_ClaMov = "E" Then
            Set frmBusqGeneral.oParent = Me
            frmBusqGeneral.sQuery = "UP_AyudaHilCru '5','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
            frmBusqGeneral.Cargar_Datos
            frmBusqGeneral.Show 1
            If Paso = True Then
                TxtItem = Codigo
                TxtDesitem = DevuelveCampo("select des_hiltel from it_hilado where cod_hiltel='" & Codigo & "'", cConnect)
            End If
        
'        ElseIf Cod_ClaOrdComp <> "" Then
'            Set frmBusqGeneral.oParent = Me
'            frmBusqGeneral.sQuery = "UP_AyudaHilCru '4','" & Cod_Almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
'            frmBusqGeneral.CARGAR_DATOS
'            frmBusqGeneral.Show 1
'            If Paso = True Then
'                TxtItem = Codigo
'                Sec_OrdComp = Descripcion
'                TxtDesitem = DevuelveCampo("select des_hiltel from it_hilado where cod_hiltel='" & Codigo & "'", cCONNECT)
'            End If
        End If
End If
Sec_OrdComp = DevuelveCampo("select sec_ordcomp from lg_ordcompitem where Ser_OrdComp='" & Ser_OrdComp & "' and Cod_OrdComp='" & Cod_OrdComp & "' and Cod_Item='" & Me.TxtItem & "'", cConnect)
End Sub

Public Sub Datos(Accion As String, EsAccion As Boolean)
On Error GoTo hand
Set Reg = Nothing
Reg.CursorLocation = adUseClient

If UCase(Accion) = "V" Then
    sTit = "Cargar Datos"
    Reg.Open "UP_Lg_MoviStkHilCru '" & Accion & "','" & Cod_Almacen & "','" & Num_MovStk & "'", cConnect
Else
    sTit = "Guardar Cambios"
    Reg.Open "UP_ACT_STOCKSHILCRU '" & Cod_Almacen & "','" & Num_MovStk & "','" & Accion & "','" & Num_Secuencia & "','" & _
            TxtLote & "','" & Cod_Proveedor & "','" & TxtItem & "'," & Me.TxtCantidad & "," & Me.txtBultos & ",'" & Me.TxtObs & "'," & _
            Cant_Anterior & ",'" & Sec_OrdComp & "','" & vusu & "', " & _
            Num_Secuencia_OrdTra_Tinto & ", '" & txtCod_OrdTra & "', " & txtConos & _
            ", '" & txtLote_Destino.Tag & "'", cConnect
End If
If EsAccion = False Then
    Set Me.DGridLista.DataSource = Reg
    DGridLista_RowColChange 0, 0
    Me.DGridLista.Columns("Cod_OrdTra").Visible = False
    Me.DGridLista.Columns("calidad").Visible = False
    Me.DGridLista.Columns("cod_proveedor").Visible = False
End If
Exit Sub
hand:
sErr = err.Description
MsgBox sErr, vbCritical + vbOKOnly, sTit
End Sub

Sub Habilita()

'TxtItem.Enabled = True
'TxtDesitem.Enabled = True
Me.TxtCantidad.Enabled = True

Me.txtBultos.Enabled = True
Me.txtConos.Enabled = True
'Me.TxtLote.Enabled = True
Me.TxtObs.Enabled = True

'Me.txtCod_OrdTra.Enabled = True

Me.txtLote_Destino.Enabled = True
cmdPesosBal.Enabled = True

If Flg_Partida_Generada <> "S" Then
    TxtLote.Enabled = True
    TxtItem.Enabled = True
    TxtDesitem.Enabled = True
    TxtLote.SetFocus
Else
    cmdGetInfo.Visible = True
    TxtCantidad.SetFocus
End If

End Sub

Sub Deshabilita()
TxtItem.Enabled = False
TxtDesitem.Enabled = False
Me.TxtCantidad.Enabled = False

Me.txtBultos.Enabled = False
Me.txtConos.Enabled = False
Me.TxtLote.Enabled = False
Me.TxtObs.Enabled = False
Me.txtCod_OrdTra.Enabled = False

Me.txtLote_Destino.Enabled = False

cmdGetInfo.Visible = False
cmdPesosBal.Enabled = False
End Sub

Sub Limpia()
TxtItem = ""
TxtDesitem = ""
Me.TxtCantidad = "0"

Me.txtBultos = "0"
txtConos = "0"
Me.TxtLote = ""
Me.TxtObs = ""
'txtCod_Ordtra = ""
txtLote_Destino = ""
txtLote_Destino.Tag = ""

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
With frmDetHilCruInfo
    .vCod_OrdTra = Cod_OrdTra1
    .vCod_TipOrdTra = Cod_TipOrdTra1
    .vCod_Almacen = Cod_Almacen
    .SM_AYUDA_ITEMS_DE_PARTIDA
    If .gexLotes.RowCount > 1 Then .Show vbModal
    If .gexLotes.RowCount > 0 And Not .bCancel Then
        TxtItem = .gexLotes.Value(.gexLotes.Columns("COD_HILTEL").Index)
        TxtDesitem = Trim(.gexLotes.Value(.gexLotes.Columns("DES_HILTEL").Index))
        'Cod_Comb = .gexLotes.Value(.gexLotes.Columns("COD_COMB").Index)
        'Label3 = Trim(.gexLotes.Value(.gexLotes.Columns("DES_COMB").Index))
        'Cod_Talla = .gexLotes.Value(.gexLotes.Columns("COD_TALLA").Index)
        'Label5 = Cod_Talla
        Num_Secuencia_OrdTra_Tinto = .gexLotes.Value(.gexLotes.Columns("NUM_SECUENCIA").Index)
        If Not .bCancelSec Then
            TxtLote = .lblCod_OrdProv
            Cod_Proveedor = .lblCod_Proveedor
            Txtproveedor = .lblDes_Proveedor
            Label2 = .lblCod_Calidad
            '.vStock
            'Cod_OrdTra
        End If
        'Aqui cargaremos los nuevos valores para los labels de cantidades
'        strSQL = "EXEC SM_ENCUENTRA_TIPOFAMTELA '" & TxtItem.Text & "',''"
'        varCod_TipFamTela = DevuelveCampo(strSQL, cConnect)
'        If varCod_TipFamTela = "N" Then
'            strSQL = "SELECT Cod_UniMed FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
'            Me.lblCantidad1.Caption = DevuelveCampo(strSQL, cConnect)
'            strSQL = "SELECT Cod_UniMedCnf FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
'            Me.lblCantidad2.Caption = DevuelveCampo(strSQL, cConnect)
'        Else
'            strSQL = "SELECT Cod_UniMedCnf FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
'            Me.lblCantidad1.Caption = DevuelveCampo(strSQL, cConnect)
'            strSQL = "SELECT Cod_UniMed FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
'            Me.lblCantidad2.Caption = DevuelveCampo(strSQL, cConnect)
'        End If
    End If
End With
Unload frmDetTelCruInfo
End Sub

Private Sub BuscPartidaGenEnt()
With frmDetHilCruEnt
    .vCod_OrdTra = Cod_OrdTra1
    .vCod_TipOrdTra = Cod_TipOrdTra1
    .vCod_Almacen = Cod_Almacen
    .SM_AYUDA_DEVOLUCION_TELA_CRUDA_DE_PARTIDAS
    If .gexLotes.RowCount > 1 Then .Show vbModal
    If .gexLotes.RowCount > 0 And Not .bCancel Then
        TxtItem = .gexLotes.Value(.gexLotes.Columns("COD_HILTEL").Index)
        TxtDesitem = Trim(.gexLotes.Value(.gexLotes.Columns("DES_HILTEL").Index))
        Num_Secuencia_OrdTra_Tinto = .gexLotes.Value(.gexLotes.Columns("NUM_SECUENCIA").Index)
        TxtLote = .gexLotes.Value(.gexLotes.Columns("COD_ORDPROV").Index)
        Cod_Proveedor = .gexLotes.Value(.gexLotes.Columns("COD_PROVEEDOR").Index)
        Txtproveedor = .gexLotes.Value(.gexLotes.Columns("DES_PROVEEDOR").Index)
        TxtCantidad = .gexLotes.Value(.gexLotes.Columns("KGS_ENVIADOS").Index)
        'Aqui cargaremos los nuevos valores para los labels de cantidades
'        strSQL = "EXEC SM_ENCUENTRA_TIPOFAMTELA '" & TxtItem.Text & "',''"
'        varCod_TipFamTela = DevuelveCampo(strSQL, cConnect)
'        If varCod_TipFamTela = "N" Then
'            strSQL = "SELECT Cod_UniMed FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
'            Me.lblCantidad1.Caption = DevuelveCampo(strSQL, cConnect)
'            strSQL = "SELECT Cod_UniMedCnf FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
'            Me.lblCantidad2.Caption = DevuelveCampo(strSQL, cConnect)
'        Else
'            strSQL = "SELECT Cod_UniMedCnf FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
'            Me.lblCantidad1.Caption = DevuelveCampo(strSQL, cConnect)
'            strSQL = "SELECT Cod_UniMed FROM TX_TELA WHERE Cod_Tela = '" & Me.TxtItem.Text & "'"
'            Me.lblCantidad2.Caption = DevuelveCampo(strSQL, cConnect)
'        End If
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

Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

If Not Reg.EOF And Not Reg.BOF Then
    Me.txtBultos = Reg("Bultos")
    Me.txtConos = Reg("Conos")
    Me.TxtCantidad = Reg("Cant Movimiento")
    Me.TxtDesitem = Reg("Hilo")
    Me.TxtItem = Reg("Cod Hilo")
    Me.TxtLote = Reg("lote")
    Me.TxtObs = Reg("Observaciones")
    Me.Txtproveedor = Reg("Proveedor")
    Cod_OrdTra = Reg("Cod_OrdTra")
    Cant_Anterior = Reg("Cant Movimiento")
    Num_Secuencia = Reg("secuencia")
    Cod_Proveedor = Reg("cod_proveedor")
    Label2.Caption = Reg("Calidad")
    txtCod_OrdTra = Trim(Reg("OT"))
    txtLote_Destino = Trim(Reg("LoteDestino"))
    txtLote_Destino.Tag = Reg("OTLoteDestino")
End If
End Sub


Private Sub Form_Load()
Num_Secuencia_OrdTra_Tinto = 0
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
Dim StrSql As String
Dim Rs As ADODB.Recordset

Set Rs = New ADODB.Recordset
StrSql = "select cod_clamov,Tip_Accion,TIP_PTMP,cod_tipanx,Flg_Partidas_Tinto from lg_tiposmov where cod_tipmov='" & Cod_TipMovi & "'"

Rs.Open StrSql, cConnect, adOpenStatic

If Rs.RecordCount Then
    If (Rs!Cod_ClaMov = "E" And Rs!Tip_Accion = "E" And _
    Trim(Rs!Tip_PtMp) = "PT" And Trim(Rs!Cod_TipAnx) = "P" _
    And Trim(Rs!Flg_Partidas_Tinto) <> "S") Or Cod_TipMovi = "ETE" Then
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
    Case "MODIFICAR"
        If Me.varValida_Factura = False Then
            MsgBox "No se puede modificar el registro por que el Mov. Almacen posee una factura relacionada.", vbInformation, "Mensaje"
            Exit Sub
        End If
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Estado = "MODIFICAR"
        Deshabilita
        TxtCantidad.Enabled = True
        Me.txtBultos.Enabled = True
        Me.txtConos.Enabled = True
        Me.TxtObs.Enabled = True
        cmdPesosBal.Enabled = True
        TxtCantidad.SetFocus
    Case "ELIMINAR"
        
        If Me.varValida_Factura = False Then
            MsgBox "No se puede eliminar el registro por que el Mov. Almacen posee una factura relacionada.", vbInformation, "Mensaje"
            Exit Sub
        End If

    
        Datos "E", True
        Limpia
        Datos "v", False
        Deshabilita
    Case "GRABAR"
        If Trim(TxtCantidad) = "" Or TxtCantidad = "0" Then MsgBox "Llene la cantidad", vbInformation: Exit Sub
        If Trim(txtBultos) = "" Or txtBultos <= "0" Then MsgBox "Ingrese un valor valido para los bultos", vbInformation: Exit Sub
        If Trim(txtConos) = "" Or txtConos <= "0" Then MsgBox "Ingrese un valor valido para los conos", vbInformation: Exit Sub
        If Trim(TxtItem) = "" Then MsgBox "Debe seleccionar un item", vbInformation: Exit Sub
        txtCod_OrdTra = Trim(txtCod_OrdTra)
        If txtCod_OrdTra = "" And txtCod_OrdTra.Visible Then MsgBox "Se debe especificar una O/T Valida", vbInformation: Exit Sub
        If Estado = "NUEVO" Then
            Datos "i", True
        Else
            Datos "M", True
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


Private Sub TxtCantidad_GotFocus()
    TxtCantidad.SelStart = 0
    TxtCantidad.SelLength = Len(TxtCantidad.Text)
End Sub

Private Sub txtcantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    SoloNumeros TxtCantidad, KeyAscii, True, 3, 6
End Sub

Private Sub txtCod_OrdTra_KeyPress(KeyAscii As Integer)
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

Private Sub txtConos_GotFocus()
    SelectionText txtConos
End Sub

Private Sub txtConos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    SoloNumeros txtConos, KeyAscii, False, 0, 4
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

Private Sub txtLote_Destino_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaLoteDes
        SendKeys "{TAB}"
    End If
End Sub

Private Sub BuscaLoteDes()
On Error GoTo ErrBusqLote
    sTit = "Busqueda de Lote Destino"
    
    txtLote_Destino = Trim(txtLote_Destino)
    
    StrSql = "SELECT Cod_OrdTra, Cod_OrdProv FROM TX_ORDTRA " & _
             "WHERE Cod_TipOrdTra = 'HI' AND Nivel_Costeo > 0 " & _
             "AND Cod_OrdProv LIKE '%" & txtLote_Destino & "%'"
    
    With frmBusqGeneral3
        .sQuery = StrSql
        
        .Caption = sTit
        
        .Cargar_Datos
        'Dar Formato al Grid
        .gexLista.Columns("Cod_OrdTra").Caption = "O/T"
        .gexLista.Columns("Cod_OrdProv").Caption = "Lote"
        .gexLista.Columns("Cod_OrdTra").Width = 1200
        .gexLista.Columns("Cod_OrdProv").Width = 5895
        
        txtLote_Destino = ""
        txtLote_Destino.Tag = ""
        
        If .gexLista.RowCount > 1 Then .Show vbModal
        bCancelSec = .bCancel
        If .gexLista.RowCount > 0 And Not .bCancel Then
            txtLote_Destino.Tag = .gexLista.Value(.gexLista.Columns("Cod_OrdTra").Index)
            txtLote_Destino = .gexLista.Value(.gexLista.Columns("Cod_OrdProv").Index)
        End If
    End With
    Unload frmBusqGeneral3
Exit Sub
ErrBusqLote:
    sErr = err.Description
    MsgBox sErr, vbCritical + vbOKOnly, sTit
End Sub

Private Sub txtLote_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
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
                                    " a.cod_proveedor='" & Cod_Proveedor & "' and " & _
                                    " a.Cod_OrdProv like '%" & Me.TxtLote & "%'"
            frmBusqGeneral.Cargar_Datos
            frmBusqGeneral.Show 1
            TxtLote = Codigo
            Cod_Proveedor = Descripcion
        End If
    Else
        If sFlg_Ot_Tejeduria_Generada = "S" Then
            txtCod_OrdTra = Trim(txtCod_OrdTra)
            If txtCod_OrdTra = "" Then
                MsgBox "Se debe especificar Una O/T", vbExclamation + vbOKOnly, "Busca Lote (Salida sFlg_Ot_Tejeduria_Generada = 'S')"
                If txtCod_OrdTra.Visible Then txtCod_OrdTra.SetFocus
                Exit Sub
            End If
            Set frmBusqGeneral.oParent = Me
            frmBusqGeneral.sQuery = "EXEC LG_VERIFICA_LOTE_HILO_CRUDO '" & _
            Cod_Almacen & "', 'TJ', '" & txtCod_OrdTra & "'"
            frmBusqGeneral.Cargar_Datos
            frmBusqGeneral.Show 1
            TxtLote = Codigo
            Cod_Proveedor = Descripcion
        Else
        Cod_TipOrdTra = "HI"
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
Me.Txtproveedor = DevuelveCampo("Select des_proveedor from lg_proveedor where cod_proveedor='" & Cod_Proveedor & "'", cConnect)

If KeyAscii = 13 Then SendKeys "{tab}"

Codigo = ""
Descripcion = ""
End Sub

Private Sub TxtObs_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Txtproveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "VALORIZAR"
        FraValorizar.Visible = True
        TxtSoles.SetFocus
        
        TxtSoles.Text = DevuelveCampo("select imp_factura from lg_movistkhilcru where cod_almacen ='" & Me.Cod_Almacen & "' and num_movstk='" & Me.Num_MovStk & "' and num_secuencia ='" & Reg("Secuencia") & "'", cConnect)
        TxtDolares.Text = DevuelveCampo("select imp_factura_dolares from lg_movistkhilcru where cod_almacen ='" & Me.Cod_Almacen & "' and num_movstk='" & Me.Num_MovStk & "' and num_secuencia ='" & Reg("Secuencia") & "'", cConnect)
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
