VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Begin VB.Form FrmDetalleTelCru 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tela Cruda"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton lote 
      Caption         =   "&Ingresar Lote"
      Height          =   525
      Left            =   8280
      TabIndex        =   27
      Top             =   6195
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
      Height          =   2730
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
         Left            =   5535
         TabIndex        =   33
         Top             =   2250
         Width           =   2220
      End
      Begin VB.CommandButton cmdGetInfo 
         Height          =   285
         Left            =   3210
         Picture         =   "FrmDetalleTelCru.frx":0000
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
         Text            =   "FrmDetalleTelCru.frx":030A
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
         Top             =   840
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
         Height          =   285
         Left            =   1185
         TabIndex        =   29
         Top             =   1260
         Width           =   2565
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
      Left            =   1815
      TabIndex        =   10
      Top             =   6135
      Width           =   2115
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1530
         Picture         =   "FrmDetalleTelCru.frx":0310
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ultimo"
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   1050
         Picture         =   "FrmDetalleTelCru.frx":0482
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Siguiente"
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   570
         Picture         =   "FrmDetalleTelCru.frx":05F4
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Anterior"
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   90
         Picture         =   "FrmDetalleTelCru.frx":0766
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Primero"
         Top             =   60
         Width           =   495
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   4020
      TabIndex        =   9
      Top             =   6165
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmDetalleTelCru.frx":08D8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "FrmDetalleTelCru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public codigo
Public Descripcion

Public NewLote As String

'Public lote As String
Public paso As Boolean
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
Public cod_almacen As String
Public Num_MovStk As String
Public Cod_OrdPro As String
Public Cod_TipOrdTra As String
Public Cod_OrdTra As String
Public Cod_Proveedor As String
Public Cod_TipOrdPro As String
Public Tip_PtMp As String
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

Public LoteaTransf As String
Public Cod_ProveedoraTransf As String
Public Cod_TelaaTransf As String
Public Cod_CombaTransf As String
Public Cod_CalidadaTransf As String
Public Cod_MedidaaTransf As String
Public bElijeDatos As Boolean


Sub Etiquetas2y3()
'Cod_Comb = DevuelveCampo("select cod_comb from lg_stockstelten  where Cod_Almacen='" & Cod_Almacen & "' and Cod_TipOrdTra='" & Cod_TipOrdTra & "' and Cod_OrdTra='" & Cod_OrdTra & "' and Cod_Tela='" & Me.TxtItem & "'", cCONNECT)
'Label3.Caption = DevuelveCampo("select Des_Comb from tx_telacomb where Cod_Comb='" & Cod_Comb & "' and Cod_tela='" & Me.TxtItem & "'", cCONNECT)
Label3.Caption = DevuelveCampo(" select a.des_comb " & _
                                " from tx_telacomb a  " & _
                                " Where a.cod_comb='" & Cod_Comb & "' and a.cod_tela='" & codigo & "'", cConnect)
'Cod_Talla = DevuelveCampo("select Cod_Talla  from lg_stockstelten  where Cod_Almacen='" & Cod_Almacen & "' and Cod_TipOrdTra='" & Cod_TipOrdTra & "' and Cod_OrdTra='" & Cod_OrdTra & "' and Cod_Tela='" & Me.TxtItem & "'", cCONNECT)
Label5.Caption = Cod_Talla
Label2.Caption = Cod_Calidad
End Sub


Sub ValidaHilo()
Dim Temp

codigo = ""
Descripcion = ""

strSQL = "select flg_partidas_tinto from Lg_TiposMov where cod_tipmov='" & Cod_TipMovi & "'"
If Cod_ClaMov = "E" And Cod_ClaOrdComp <> "" And DevuelveCampo(strSQL, cConnect) = "N" Then
    Set frmBusqGeneral.oParent = Me
    frmBusqGeneral.sQuery = "UP_AyudaTellCru '1','" & cod_almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
    
    frmBusqGeneral.CARGAR_DATOS
    frmBusqGeneral.Show 1
    If paso = True Then
        TxtItem = codigo
        Sec_OrdComp = Descripcion
        TxtDesitem = DevuelveCampo("select des_tela  from tx_tela where cod_tela='" & codigo & "'", cConnect)
    End If
Else
    If Cod_ClaMov = "S" Then
        Set frmBusqGeneral.oParent = Me
        frmBusqGeneral.sQuery = "UP_AyudaTellCru '2','" & cod_almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
        frmBusqGeneral.CARGAR_DATOS
        frmBusqGeneral.Show 1
        If paso = True Then
            TxtItem = codigo
            TxtDesitem = Descripcion
        End If
        Etiquetas2y3
    ElseIf Cod_ClaMov = "E" Then
        Set frmBusqGeneral.oParent = Me
        frmBusqGeneral.sQuery = "UP_AyudaTellCru '3','" & cod_almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
        frmBusqGeneral.CARGAR_DATOS
        frmBusqGeneral.Show 1
        If paso = True Then
            TxtItem = codigo
            TxtDesitem = Descripcion
            Sec_OrdComp = DevuelveCampo("select sec_ordcomp from lg_ordcompitem where Ser_OrdComp='" & Ser_OrdComp & "' and Cod_OrdComp='" & Cod_OrdComp & "' and Cod_Item='" & Me.TxtItem & "' AND cod_Comb='" & Me.Cod_Comb & "' and cod_talla='" & Me.Cod_Talla & "'", cConnect)
        End If
        Etiquetas2y3
    ElseIf Cod_ClaOrdComp <> "" Then
        Set frmBusqGeneral.oParent = Me
        frmBusqGeneral.sQuery = "UP_AyudaTellCru '4','" & cod_almacen & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_Calidad & "','" & TxtLote & "','" & Cod_Proveedor & "'"
        frmBusqGeneral.CARGAR_DATOS
        frmBusqGeneral.Show 1
        If paso = True Then
            TxtItem = codigo
            Sec_OrdComp = Descripcion
            TxtDesitem = DevuelveCampo("select des_tela from tx_tela where cod_tela='" & codigo & "'", cConnect)
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


Public Sub Datos(Accion As String, EsAccion As Boolean)
On Error GoTo hand
Set Reg = Nothing
Reg.CursorLocation = adUseClient

If UCase(Accion) = "V" Then
    Reg.Open "UP_Lg_MoviStkTelCru '" & Accion & "','" & cod_almacen & "','" & Num_MovStk & "'", cConnect
Else
    Reg.Open "UP_ACT_STOCKSTELCRU '" & cod_almacen & "','" & Num_MovStk & "','" & Accion & "','" & Num_Secuencia & "','" & _
            TxtLote & "','" & Cod_Proveedor & "','" & TxtItem & "','" & Cod_Comb & "','" & Cod_Talla & "'," & Me.TxtCantidad & "," & Me.TxtBultos & ",'" & TxtObs.Text & "'," & _
            Cant_Anterior & ",'" & Sec_OrdComp & "'," & Me.txtCan_Movimiento_2daunimed.Text & ",'" & vusu & "', " & Num_Secuencia_OrdTra_Tinto & ",'S' ,'" & Me.LoteaTransf & "','" & Me.Cod_ProveedoraTransf & "','" & Me.Cod_TelaaTransf & "','" & Me.Cod_CombaTransf & "','" & Me.Cod_MedidaaTransf & "'", cConnect
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
    TxtBultos.Enabled = True
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
End Sub

Sub Deshabilita()
TxtItem.Enabled = False
TxtDesitem.Enabled = False
Me.TxtCantidad.Enabled = False
txtCan_Movimiento_2daunimed.Enabled = False

Me.TxtBultos.Enabled = False
Me.TxtLote.Enabled = False
Me.TxtObs.Enabled = False
cmdGetInfo.Visible = False
End Sub

Sub Limpia()
TxtItem = ""
TxtDesitem = ""
Me.TxtCantidad = "0"
Me.txtCan_Movimiento_2daunimed = "0"
Me.TxtBultos = "0.00"
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
        .vCod_OrdTra = Cod_OrdTra1
        .vCod_TipOrdTra = Cod_TipOrdTra1
        .vCod_Almacen = cod_almacen
        .SM_AYUDA_ITEMS_DE_PARTIDA
        If .gexLotes.RowCount > 1 Then .Show vbModal
        If .gexLotes.RowCount > 0 And Not .bCancel Then
            TxtItem = .gexLotes.Value(.gexLotes.Columns("COD_TELA").Index)
            TxtDesitem = Trim(.gexLotes.Value(.gexLotes.Columns("DES_TELA").Index))
            Cod_Comb = .gexLotes.Value(.gexLotes.Columns("COD_COMB").Index)
            Label3 = Trim(.gexLotes.Value(.gexLotes.Columns("DES_COMB").Index))
            Cod_Talla = .gexLotes.Value(.gexLotes.Columns("COD_TALLA").Index)
            Label5 = Cod_Talla
            Num_Secuencia_OrdTra_Tinto = .gexLotes.Value(.gexLotes.Columns("NUM_SECUENCIA").Index)
            If Not .bCancelSec Then
                TxtLote = .lblCod_OrdProv
                Cod_Proveedor = .lblCod_Proveedor
                TxtProveedor = .lblDes_Proveedor
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
        .vCod_Almacen = cod_almacen
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
            TxtLote = .gexLotes.Value(.gexLotes.Columns("COD_ORDPROV").Index)
            Cod_Proveedor = .gexLotes.Value(.gexLotes.Columns("COD_PROVEEDOR").Index)
            TxtProveedor = .gexLotes.Value(.gexLotes.Columns("DES_PROVEEDOR").Index)
            TxtCantidad = .gexLotes.Value(.gexLotes.Columns("KGS_ENVIADOS").Index)
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

Private Sub cmdPrevious_Click()
If Not Reg.BOF Then Reg.MovePrevious
End Sub





Private Sub CmdTransferir_Click()
    Load frmDetalleTelCruTransf
    
    frmDetalleTelCruTransf.xlote = Me.TxtLote
    frmDetalleTelCruTransf.xCod_Proveedor = Mid(TxtProveedor.Text, 1, 12)
    frmDetalleTelCruTransf.xCod_Tela = TxtItem.Text
    frmDetalleTelCruTransf.xCod_Comb = Mid(Label3, 1, 3)
    frmDetalleTelCruTransf.xCod_Calidad = Mid(Label2, 1, 1)
    If Mid(Label5, 1, 3) <> "-" Then
       frmDetalleTelCruTransf.xCod_Medida = Mid(Label5, 1, 10)
    Else
       frmDetalleTelCruTransf.xCod_Medida = ""
    End If
    
    frmDetalleTelCruTransf.TxtLote = Me.TxtLote
    frmDetalleTelCruTransf.TxtProveedor = Me.TxtProveedor
    frmDetalleTelCruTransf.TxtItem = Me.TxtItem
    frmDetalleTelCruTransf.LlenaDatos
    frmDetalleTelCruTransf.Label2 = Label2
    frmDetalleTelCruTransf.Label5 = Label5
    

    frmDetalleTelCruTransf.xcod_almacen = Me.cod_almacen
    frmDetalleTelCruTransf.xNum_MovStk = Me.Num_MovStk
    frmDetalleTelCruTransf.Show vbModal
    
    If frmDetalleTelCruTransf.bOk Then
        Me.LoteaTransf = frmDetalleTelCruTransf.xlote
        Me.Cod_ProveedoraTransf = frmDetalleTelCruTransf.xCod_Proveedor
        Me.Cod_TelaaTransf = frmDetalleTelCruTransf.xCod_Tela
        Me.Cod_CombaTransf = frmDetalleTelCruTransf.xCod_Comb
        Me.Cod_CalidadaTransf = frmDetalleTelCruTransf.xCod_Calidad
        Me.Cod_MedidaaTransf = frmDetalleTelCruTransf.xCod_Medida
        bElijeDatos = True
    Else
        Me.LoteaTransf = ""
        Me.Cod_ProveedoraTransf = ""
        Me.Cod_TelaaTransf = ""
        Me.Cod_CombaTransf = ""
        Me.Cod_CalidadaTransf = ""
        Me.Cod_MedidaaTransf = ""
    End If
    
    Set frmDetalleTelCruTransf = Nothing

End Sub

Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo hand
If Not Reg.EOF And Not Reg.BOF Then
    Me.TxtBultos = Reg("Bultos")
    Me.TxtCantidad = Reg("Cant Movimiento")
    Me.TxtDesitem = Reg("Tela")
    Me.TxtItem = Reg("cod_tela")
    Me.TxtLote = Reg("lote")
    Me.TxtObs = Reg("Observaciones")
    Me.TxtProveedor = Reg("Proveedor")
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
Tip_item = DevuelveCampo("select tip_item from lg_almacen where cod_almacen='" & Me.cod_almacen & "'", cConnect)
Tip_presentacion = DevuelveCampo("select Tip_presentacion from lg_almacen where cod_almacen='" & Me.cod_almacen & "'", cConnect)
Cod_Calidad = DevuelveCampo("select isnull(Cod_Calidad,'') from lg_tiposmov where cod_tipmov='" & Me.Cod_TipMovi & "'", cConnect)
Cod_TipOrdTra = DevuelveCampo("select cod_tipordtra from Tx_TiposOrdTra where tip_item='" & Tip_item & "' and tip_presentacion='" & Tip_presentacion & "'", cConnect)
Me.TxtProveedor = DevuelveCampo("Select des_proveedor from lg_proveedor where cod_proveedor='" & Cod_Proveedor & "'", cConnect)
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
        Me.TxtBultos.Enabled = True
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
        Deshabilita
    Case "GRABAR"
        If Trim(TxtCantidad) = "" Or TxtCantidad = "0" Then MsgBox "Llene la cantidad", vbInformation: Exit Sub
        'If Trim(TxtBultos) = "" Or TxtBultos  "0" Then MsgBox "Llene la cantidad de bultos", vbInformation: Exit Sub
        If Trim(TxtBultos) = "" Then TxtBultos = "0"
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
            If Not bElijeDatos Then
                MsgBox "Debe seleccionar Destino de Transferencia ", vbExclamation, "Transferencia de Stocks"
                Exit Sub
            End If
            
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

Private Sub TxtBultos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
    SoloNumeros TxtBultos, KeyAscii, False, 0, 4
End Sub


Private Sub txtCan_Movimiento_2daunimed_GotFocus()
    Me.txtCan_Movimiento_2daunimed.SelStart = 0
    Me.txtCan_Movimiento_2daunimed.SelLength = Len(Me.txtCan_Movimiento_2daunimed.Text)
End Sub

Private Sub txtCan_Movimiento_2daunimed_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtBultos.SetFocus
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

Private Sub TxtLote_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Cod_ClaMov = "E" And Cod_ClaOrdComp <> "" And Tip_PtMp = "PT" Then
        If DevuelveCampo("select  count(*) from   TX_ORDTRA  Where  COD_TIPORDTRA='" & Cod_TipOrdTra & "' and " & _
                        "   Cod_OrdProv like '%" & Me.TxtLote & "%' and cod_proveedor='" & Cod_Proveedor & "'", cConnect) <= 0 Then
            MsgBox "Este Lote no existe", vbInformation
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
            frmBusqGeneral.CARGAR_DATOS
            frmBusqGeneral.Show 1
            TxtLote = codigo
            Cod_Proveedor = Descripcion
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
            frmBusqGeneral.CARGAR_DATOS
            frmBusqGeneral.Show 1
            TxtLote = codigo
            Cod_Proveedor = Descripcion
        End If
    End If
End If
Cod_OrdTra = DevuelveCampo("select cod_ordtra from tx_ordtra where Cod_TipOrdTra='" & Cod_TipOrdTra & _
             "' and Cod_Proveedor='" & Cod_Proveedor & "' and Cod_OrdProv='" & TxtLote & "'", cConnect)
Me.TxtProveedor = Descripcion & "-" & DevuelveCampo("Select des_proveedor from lg_proveedor where cod_proveedor='" & Cod_Proveedor & "'", cConnect)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub






