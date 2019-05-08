VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmShowFactVentasPrendasExportacion 
   Caption         =   "Documentos Ventas Exportación"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   12150
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFecCobRepro 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fecha Cobranza Reprogramada"
      Height          =   1650
      Left            =   4020
      TabIndex        =   46
      Top             =   6960
      Visible         =   0   'False
      Width           =   3555
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   510
         Left            =   510
         TabIndex        =   47
         Top             =   870
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   900
         Custom          =   $"frmShowFactVentasPrendasExportacion.frx":0000
         Orientacion     =   0
         Style           =   1
         Language        =   1
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   0
      End
      Begin NumBoxProject.NumBox txtFecCobRepro 
         Height          =   330
         Left            =   1080
         TabIndex        =   48
         Top             =   360
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         TypeVal         =   3
         Mask            =   "99/99/9999"
         Formato         =   "dd/MM/yyyy"
         AllowedMask     =   0
         MaskLen         =   20
         Aling           =   2
         Text            =   ""
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   0
      End
   End
   Begin VB.Frame fraPenalidad 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Penalidades del Embarque"
      Height          =   2130
      Left            =   4305
      TabIndex        =   41
      Top             =   4170
      Visible         =   0   'False
      Width           =   3555
      Begin VB.TextBox txtImp_Dscto_Penalidad 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1680
         TabIndex        =   42
         Tag             =   "SET"
         Top             =   690
         Width           =   1260
      End
      Begin FunctionsButtons.FunctButt FunctButt5 
         Height          =   510
         Left            =   555
         TabIndex        =   43
         Top             =   1230
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   900
         Custom          =   $"frmShowFactVentasPrendasExportacion.frx":0097
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   0
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Imp. Descuento por Penalidad"
         Height          =   360
         Left            =   225
         TabIndex        =   44
         Tag             =   "NUM_DUA"
         Top             =   675
         Width           =   1500
      End
   End
   Begin VB.Frame FraBuscar 
      Caption         =   "Argumentos de Registro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11970
      Begin VB.OptionButton optClienteComercial 
         Caption         =   "Cliente Comercial"
         Height          =   375
         Left            =   5520
         TabIndex        =   36
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   495
         Left            =   10200
         TabIndex        =   33
         Top             =   240
         Width           =   1545
      End
      Begin VB.OptionButton optFecha 
         Caption         =   "Fecha de Emision"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optCliente 
         Caption         =   "Consignatario"
         Height          =   375
         Left            =   1920
         TabIndex        =   31
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optAnoMes 
         Caption         =   "Año/ Mes"
         Height          =   375
         Left            =   3600
         TabIndex        =   30
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton optCorrelativo 
         Caption         =   "Correlativo Generico"
         Height          =   375
         Left            =   5520
         TabIndex        =   29
         Top             =   120
         Width           =   1815
      End
      Begin VB.OptionButton optNroDoc 
         Caption         =   "Nro de Documento"
         Height          =   375
         Left            =   3600
         TabIndex        =   28
         Top             =   480
         Width           =   1815
      End
      Begin VB.Frame frpo 
         Height          =   800
         Left            =   8400
         TabIndex        =   49
         Top             =   840
         Width           =   3375
         Begin VB.TextBox txt_po 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   480
            MaxLength       =   20
            TabIndex        =   50
            Top             =   240
            Width           =   2700
         End
         Begin VB.Label Label11 
            Caption         =   "P.O.:"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame frFecha 
         Height          =   800
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   7455
         Begin MSComCtl2.DTPicker dtpFecEmiIni 
            Height          =   315
            Left            =   1590
            TabIndex        =   2
            Top             =   360
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Format          =   94109697
            CurrentDate     =   37543
         End
         Begin MSComCtl2.DTPicker dtpFecEmiFin 
            Height          =   315
            Left            =   4470
            TabIndex        =   3
            Top             =   360
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Format          =   94109697
            CurrentDate     =   37543
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta :"
            Height          =   255
            Left            =   3720
            TabIndex        =   5
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Desde :"
            Height          =   255
            Left            =   960
            TabIndex        =   4
            Top             =   390
            Width           =   615
         End
      End
      Begin VB.Frame fraClienteCom 
         Height          =   800
         Left            =   120
         TabIndex        =   37
         Top             =   840
         Width           =   8295
         Begin VB.CheckBox chk_po 
            Caption         =   "P/O"
            Height          =   255
            Left            =   7440
            TabIndex        =   52
            Top             =   250
            Width           =   615
         End
         Begin VB.CommandButton cmdBusCliente 
            Caption         =   "..."
            Height          =   285
            Left            =   1440
            TabIndex        =   45
            Tag             =   "..."
            Top             =   260
            Width           =   300
         End
         Begin VB.TextBox txtAbr_Cliente 
            Height          =   315
            Left            =   840
            TabIndex        =   39
            Top             =   240
            Width           =   645
         End
         Begin VB.TextBox txtNom_Cliente 
            Height          =   315
            Left            =   1680
            TabIndex        =   38
            Top             =   240
            Width           =   5490
         End
         Begin VB.Label Label15 
            Caption         =   "Cliente"
            Height          =   225
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame frCorrelativo 
         Height          =   800
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   7455
         Begin VB.TextBox txtNum_Corre 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   2880
            MaxLength       =   15
            TabIndex        =   26
            Top             =   360
            Width           =   1425
         End
         Begin VB.Label Label7 
            Caption         =   "Correlativo Genérico"
            Height          =   255
            Left            =   1200
            TabIndex        =   27
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame frNroDoc 
         Height          =   800
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   7455
         Begin VB.TextBox txtNum_Docum 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   6120
            MaxLength       =   8
            TabIndex        =   21
            Top             =   375
            Width           =   1080
         End
         Begin VB.TextBox txtSer_Docum 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   4680
            MaxLength       =   3
            TabIndex        =   20
            Top             =   375
            Width           =   540
         End
         Begin VB.TextBox txtCod_TipDoc 
            Height          =   285
            Left            =   1080
            MaxLength       =   4
            TabIndex        =   19
            Top             =   375
            Width           =   480
         End
         Begin VB.TextBox txtDes_TipDoc 
            Height          =   285
            Left            =   1680
            TabIndex        =   18
            Top             =   375
            Width           =   1905
         End
         Begin VB.Label Label5 
            Caption         =   "Tipo Doc :"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   390
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Número :"
            Height          =   195
            Left            =   5400
            TabIndex        =   23
            Tag             =   "Number"
            Top             =   420
            Width           =   645
         End
         Begin VB.Label Label12 
            Caption         =   "Serie :"
            Height          =   255
            Left            =   4080
            TabIndex        =   22
            Top             =   390
            Width           =   495
         End
      End
      Begin VB.Frame frAnoMes 
         Height          =   800
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   7455
         Begin VB.TextBox txtMes 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   4440
            MaxLength       =   2
            TabIndex        =   14
            Top             =   345
            Width           =   480
         End
         Begin VB.TextBox txtAno 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   3000
            MaxLength       =   4
            TabIndex        =   13
            Top             =   345
            Width           =   660
         End
         Begin VB.Label Label6 
            Caption         =   "Año"
            Height          =   255
            Left            =   2520
            TabIndex        =   16
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label8 
            Caption         =   "Mes"
            Height          =   255
            Left            =   3960
            TabIndex        =   15
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame frCliente 
         Height          =   800
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   7455
         Begin VB.TextBox txtNum_Ruc 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   960
            MaxLength       =   11
            TabIndex        =   9
            Top             =   360
            Width           =   1200
         End
         Begin VB.TextBox txtDes_Anexo 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   2280
            MaxLength       =   30
            TabIndex        =   8
            Top             =   360
            Width           =   4050
         End
         Begin VB.TextBox txtCod_TipAnxo 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   6960
            MaxLength       =   1
            TabIndex        =   7
            Text            =   "C"
            Top             =   360
            Width           =   360
         End
         Begin VB.Label Label4 
            Caption         =   "Nro Ruc:"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   390
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "Tipo :"
            Height          =   255
            Left            =   6480
            TabIndex        =   10
            Top             =   360
            Width           =   495
         End
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   9420
      Left            =   10680
      TabIndex        =   34
      Top             =   1680
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   16616
      Custom          =   $"frmShowFactVentasPrendasExportacion.frx":012E
      Orientacion     =   1
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1300
      ControlHeigth   =   470
      ControlSeparator=   0
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   8610
      Left            =   120
      TabIndex        =   35
      Top             =   1800
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   15187
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      RowHeaders      =   -1  'True
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmShowFactVentasPrendasExportacion.frx":0898
      Column(2)       =   "frmShowFactVentasPrendasExportacion.frx":0960
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmShowFactVentasPrendasExportacion.frx":0A04
      FormatStyle(2)  =   "frmShowFactVentasPrendasExportacion.frx":0B3C
      FormatStyle(3)  =   "frmShowFactVentasPrendasExportacion.frx":0BEC
      FormatStyle(4)  =   "frmShowFactVentasPrendasExportacion.frx":0CA0
      FormatStyle(5)  =   "frmShowFactVentasPrendasExportacion.frx":0D78
      FormatStyle(6)  =   "frmShowFactVentasPrendasExportacion.frx":0E30
      FormatStyle(7)  =   "frmShowFactVentasPrendasExportacion.frx":0F10
      FormatStyle(8)  =   "frmShowFactVentasPrendasExportacion.frx":0FBC
      ImageCount      =   0
      PrinterProperties=   "frmShowFactVentasPrendasExportacion.frx":106C
   End
   Begin FunctionsButtons.FunctButt FunctButt3 
      Height          =   525
      Left            =   120
      TabIndex        =   53
      Top             =   10440
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   926
      Custom          =   $"frmShowFactVentasPrendasExportacion.frx":1244
      Orientacion     =   0
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1400
      ControlHeigth   =   500
      ControlSeparator=   0
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   10995
      Top             =   7980
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmShowFactVentasPrendasExportacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iRowAnterior As Long
Dim iColAnterior As Long
Dim bClickColSelec As Boolean
Dim bCargaGRid As Boolean
Dim bPuedeAutorizar  As Boolean
Dim sTipoDocAutorizar As String
Dim strOpcion As String
Public codigo As String, Descripcion As String, TipoAdd As String
Public strCod_Anxo As String, lvSW As Boolean
Public oParent As Object
Public var As String
Public indicegrilla As Long
Dim strSQL As String
Dim cod1 As String
Dim cod2 As String
Dim cod3 As String
Dim SNum_Corre As String
Dim Cod_Cliente As String

Private Sub chk_po_Click()
If Me.chk_po.Value Then
  frpo.Visible = True
  strOpcion = "P"
  txt_po.SetFocus
Else
  txt_po.Text = ""
  frpo.Visible = False
  strOpcion = "M"
End If

End Sub

Private Sub dtpFecEmiIni_Change()
  gridex1.ClearFields
  dtpFecEmiFin.Value = dtpFecEmiIni.Value
End Sub

Private Sub Form_Load()
  lvSW = True
  dtpFecEmiIni.Value = Date
  dtpFecEmiFin.Value = Date

  indicegrilla = 1
  FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name) & "/SALIR"
  FunctButt3.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
  
  strOpcion = "F"
  
  iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))
  frpo.Visible = False
  Me.fraClienteCom.Visible = False
End Sub

Private Sub cmdBuscar_Click()
  Buscar
End Sub
Public Sub Buscar()
On Error GoTo dprDepurar

Dim sSQL As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle

If Me.optClienteComercial.Value Then
    strSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente LIKE '" & Trim(txtAbr_Cliente.Text) & "%'"
    Cod_Cliente = DevuelveCampo(strSQL, cCONNECT)
End If

sSQL = "Ventas_Muestra_Doc_Ventas_Export  '" & strOpcion & "','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','" & "C" & "','" & strCod_Anxo & "','" & txtAno & "','" & txtMes & "','" & txtNum_Corre & "','" & txtCod_TipDoc & "','" & txtSer_Docum & "','" & txtNum_Docum & "','" & vusu & "','" & Cod_Cliente & "','" & Trim(Me.txt_po.Text) & "'"

Set gridex1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)

gridex1.Columns("Cod_TipDoc").Width = 375
gridex1.Columns("Cod_TipDoc").Caption = "Tip"
gridex1.Columns("Serie").Width = 525
gridex1.Columns("Serie").Caption = "Serie"
gridex1.Columns("Nro_Doc").Width = 810
gridex1.Columns("Nro_Doc").Caption = "Nro_Doc"
gridex1.Columns("Anexo").Width = 2865
gridex1.Columns("Anexo").Caption = "Anexo"
gridex1.Columns("Ruc").Width = 1410
gridex1.Columns("Ruc").Caption = "Ruc"
gridex1.Columns("Moneda").Width = 705
gridex1.Columns("Moneda").Caption = "Moneda"
gridex1.Columns("Imp_Neto").Width = 825
gridex1.Columns("Imp_Neto").Caption = "Imp Neto"
gridex1.Columns("Imp_Igv").Width = 705
gridex1.Columns("Imp_Igv").Caption = "Imp Igv"
gridex1.Columns("Imp_Gastos_Financieros").Caption = "Gastos Financieros"
gridex1.Columns("Imp_Gastos_Financieros").Width = 990
gridex1.Columns("Imp_Total").Width = 840
gridex1.Columns("Imp_Total").Caption = "Imp Total"
gridex1.Columns("Imp_Otros").Width = 870
gridex1.Columns("Imp_Otros").Caption = "Imp Otros"
gridex1.Columns("Emision").Width = 945
gridex1.Columns("Emision").Caption = "Emision"
gridex1.Columns("Registro").Width = 945
gridex1.Columns("Registro").Caption = "Registro"
gridex1.Columns("Vencimiento").Width = 945
gridex1.Columns("Vencimiento").Caption = "Vencimiento"
gridex1.Columns("Cancelado").Width = 1500
gridex1.Columns("Cancelado").Caption = "Cancelado"
gridex1.Columns("Ano_Registro").Width = 1095
gridex1.Columns("Ano_Registro").Caption = "Ano_Registro"
gridex1.Columns("Mes_Registro").Width = 1110
gridex1.Columns("Mes_Registro").Caption = "Mes_Registro"
gridex1.Columns("Num_Registro").Width = 1140
gridex1.Columns("Num_Registro").Caption = "Num_Registro"

'GridEX1.Columns("Emision").Format = "dd/mm/yyy"
'GridEX1.Columns("Registro").Format = "dd/mm/yyy"
'GridEX1.Columns("Vencimiento").Format = "dd/mm/yyy"
'GridEX1.Columns("Cancelado").Format = "dd/mm/yyy"

gridex1.ContinuousScroll = True


    
gridex1.RowSelected(indicegrilla) = True
'GridEX1.SetFocus
    

Exit Sub

dprDepurar:

errores err.Number
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo HandlerError

    Dim Msg As Variant
    Select Case ActionName
    Case "MODIFICA"
      If gridex1.RowCount = 0 Then Exit Sub
      
      If Not ifValidaDoc Then Exit Sub
      
      With frmAdicionaDocumVentasExport
        .strOption = "U"
        .Caption = "Modifica Docum Ventas"
        Carga_Data
        .strNum_Corre = gridex1.Value(gridex1.Columns("Num_Corre").Index)
        .Show 1
        'If .strNum_Corre <> "" Then
        '  optCorrelativo = True
        '  txtNum_Corre = .strNum_Corre
        '  Buscar
        'End If
      End With
    Case "VERDETALLE"
      If gridex1.RowCount = 0 Then Exit Sub
      Load frmMuestraDetalleDocumVentasExport
      With frmMuestraDetalleDocumVentasExport
        .Caption = gridex1.Value(gridex1.Columns("Cod_TipDoc").Index) & " Nro " & gridex1.Value(gridex1.Columns("Serie").Index) & "-" & gridex1.Value(gridex1.Columns("Nro_Doc").Index)
        .strSQL = "Ventas_Muestra_Detalle_Factura_Prendas '" & gridex1.Value(gridex1.Columns("Num_Corre").Index) & "'"
        .Num_Corre = gridex1.Value(gridex1.Columns("Num_Corre").Index)
        .Buscar
        .Show 1
        Buscar
      End With
      Set frmMuestraDetalleDocumVentasExport = Nothing
    Case "EXPORTADBF"

    Case "REIMPRESION"
        If gridex1.RowCount = 0 Then Exit Sub
        Imprimir gridex1.Value(gridex1.Columns("Num_Corre").Index), gridex1.Value(gridex1.Columns("Cod_TipDoc").Index)
    Case "ANULAR"
        If gridex1.RowCount = 0 Then Exit Sub
        If MsgBox("Esta Seguro de Anular este Documento", vbYesNo, "IMPORTANTE") = vbYes Then
           ExecuteCommandSQL cCONNECT, "Ventas_Man_Anula_Docum '" & gridex1.Value(gridex1.Columns("Num_Corre").Index) & "','" & vusu & "'"
           Buscar
        End If
    Case "REVIERTEDOCUM"
        If gridex1.RowCount = 0 Then Exit Sub
        If MsgBox("Esta Seguro de Revertir este Documento", vbYesNo, "IMPORTANTE") = vbYes Then
           ExecuteCommandSQL cCONNECT, "Ventas_Revierte_Docum '" & gridex1.Value(gridex1.Columns("Num_Corre").Index) & "','" & vusu & "'"
           Buscar
        End If
    Case "GENERAINFOCONT"
        If gridex1.RowCount = 0 Then Exit Sub
        GeneraInfoContable
    Case "VERVOUCHER"
        If gridex1.RowCount = 0 Then Exit Sub
        MuestraVoucher2
        
    Case "IMPRIMIRINVOICE"
    If gridex1.RowCount = 0 Then Exit Sub
   
          Call Reporte
        
    Case "IMPCOLOR"
    If gridex1.RowCount = 0 Then Exit Sub
   
          Call Reporte_COLOR
         
         
    Case "DESPEXT"
    If gridex1.RowCount = 0 Then Exit Sub
     
     
    'Load frmConfirmacionDespacho
    
    frmConfirmacionDespacho.Cod_TipDoc = gridex1.Value(gridex1.Columns("Cod_TipDoc").Index)
    frmConfirmacionDespacho.Serie = gridex1.Value(gridex1.Columns("Serie").Index)
    frmConfirmacionDespacho.Nro_doc = gridex1.Value(gridex1.Columns("Nro_Doc").Index)
    frmConfirmacionDespacho.Valor = gridex1.Value(gridex1.Columns("Despacho").Index)
     
     cod1 = gridex1.Value(gridex1.Columns("Cod_TipDoc").Index)
     cod2 = gridex1.Value(gridex1.Columns("Serie").Index)
     cod3 = gridex1.Value(gridex1.Columns("Nro_Doc").Index)
     
    Set frmConfirmacionDespacho.oParent = Me
    frmConfirmacionDespacho.Show vbModal
    Set frmConfirmacionDespacho = Nothing
       
    Case "PENALIDADES"
        LoadPenalidades
    
    Case "LDP/DDP"
       If gridex1.RowCount = 0 Then Exit Sub
        frmCompletarImportesLDPDDP.strNum_Corre = gridex1.Value(gridex1.Columns("Num_Corre").Index)
        frmCompletarImportesLDPDDP.txtFlete = gridex1.Value(gridex1.Columns("Imp_flete").Index)
        frmCompletarImportesLDPDDP.txtDesaduanaje = gridex1.Value(gridex1.Columns("imp_desaduanaje").Index)
        frmCompletarImportesLDPDDP.txtTransporte = gridex1.Value(gridex1.Columns("imp_transporte_pais_destino").Index)
        frmCompletarImportesLDPDDP.txtFob = gridex1.Value(gridex1.Columns("Imp_FOB").Index)
        frmCompletarImportesLDPDDP.txtCif = gridex1.Value(gridex1.Columns("Imp_CIF").Index)
        frmCompletarImportesLDPDDP.txtLdp = gridex1.Value(gridex1.Columns("Imp_LDP").Index)
        frmCompletarImportesLDPDDP.txtDdp = gridex1.Value(gridex1.Columns("Imp_DDP").Index)
       
        frmCompletarImportesLDPDDP.Show vbModal
        Set frmCompletarImportesLDPDDP = Nothing
        Buscar
    Case "IMPRESIONES"
        If gridex1.RowCount = 0 Then Exit Sub
        Imprimir1 gridex1.Value(gridex1.Columns("Num_Corre").Index), gridex1.Value(gridex1.Columns("Cod_TipDoc").Index)
   
    Case "RINFCONTABLE"
        If gridex1.RowCount = 0 Then Exit Sub
        If MsgBox("Esta Seguro de Revertir Ifx Contable de este Documento", vbYesNo, "IMPORTANTE") = vbYes Then
           ExecuteCommandSQL cCONNECT, "CN_REVIERTE_ASIENTO_VENTAS'" & gridex1.Value(gridex1.Columns("Num_Corre").Index) & "'"
           Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
           Buscar
        End If
        
    Case "FECDOC"
           
         If gridex1.RowCount = 0 Then Exit Sub
         
         frm_FecEnvDoc.Cod_TipDoc = gridex1.Value(gridex1.Columns("Cod_TipDoc").Index)
         frm_FecEnvDoc.Serie = gridex1.Value(gridex1.Columns("Serie").Index)
         frm_FecEnvDoc.Nro_doc = gridex1.Value(gridex1.Columns("Nro_Doc").Index)
         frm_FecEnvDoc.DTPFecha.Value = gridex1.Value(gridex1.Columns("Fec_Envio_Documentos_Cobranza").Index)
         Set frm_FecEnvDoc.oParent = Me
         frm_FecEnvDoc.Show vbModal
         Set frm_FecEnvDoc = Nothing
        Buscar
    
    Case "FECCOBREPRO"
        If gridex1.RowCount = 0 Then Exit Sub
        SNum_Corre = gridex1.Value(gridex1.Columns("NUM_CORRE").Index)
        txtFecCobRepro.Text = FixNulos(gridex1.Value(gridex1.Columns("Fec_Cobranza_Reprogramada").Index), vbString)
       
        fraFecCobRepro.Visible = True
    Case "SALIR"
       Unload Me
       
    Case "VEPACK"
        With FrmListaPacking
        .numCorre = gridex1.Value(gridex1.Columns("NUM_CORRE").Index)
        .cargarGrid
        End With
        FrmListaPacking.Show 1
    End Select
Exit Sub
Resume
HandlerError:
  errores err.Number
End Sub


Sub Reporte()
Dim oo As Object
On Error GoTo AceptarErr

Screen.MousePointer = 11

Set oo = CreateObject("excel.application")
oo.Workbooks.Open vRuta & "\Commercial_Invoice2.xlt"
oo.Visible = True
'oo.Run "Reporte", cCONNECT, "cf_extrae_datos_commercial_invoice_vans '" & Trim(GridEX1.Value(GridEX1.Columns("Serie").Index)) & "','" & Trim(GridEX1.Value(GridEX1.Columns("Nro_Doc").Index)) & "'", Trim(GridEX1.Value(GridEX1.Columns("Serie").Index)) & Trim(GridEX1.Value(GridEX1.Columns("Nro_Doc").Index))

oo.Run "Reporte", cCONNECT, "cf_extrae_datos_commercial_invoice_vans '" & Trim(gridex1.Value(gridex1.Columns("Serie").Index)) & "','" & Trim(gridex1.Value(gridex1.Columns("Nro_Doc").Index)) & "'", Trim(gridex1.Value(gridex1.Columns("Serie").Index)), Trim(gridex1.Value(gridex1.Columns("Nro_Doc").Index)), Trim(gridex1.Value(gridex1.Columns("Serie").Index)) & "-" & Trim(gridex1.Value(gridex1.Columns("Nro_Doc").Index))

Screen.MousePointer = 0

oo.Visible = True
Set oo = Nothing

Exit Sub
AceptarErr:
    MsgBox err.Description, vbCritical
    Screen.MousePointer = 0
End Sub

Sub Reporte_COLOR()
Dim oo As Object
On Error GoTo AceptarErr

Screen.MousePointer = 11

Set oo = CreateObject("excel.application")
oo.Workbooks.Open vRuta & "\CajasProgramadas_Gstar2.XLT"
oo.Visible = True

oo.Run "Reporte", cCONNECT, Trim(gridex1.Value(gridex1.Columns("Serie").Index)), Trim(gridex1.Value(gridex1.Columns("Nro_Doc").Index)), Trim(gridex1.Value(gridex1.Columns("Serie").Index)) & "-" & Trim(gridex1.Value(gridex1.Columns("Nro_Doc").Index))

Screen.MousePointer = 0

oo.Visible = True
Set oo = Nothing

Exit Sub
AceptarErr:
    MsgBox err.Description, vbCritical
    Screen.MousePointer = 0
End Sub



Private Function ifValidaDoc() As Boolean

Dim strMsg As String

'strMsg = DevuelveCampo("Select dbo.ventas_Valida_Documento_Manuales('" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "')", cCONNECT)
'If strMsg <> "" Then
'  MsgBox strMsg, vbInformation, "AVISO"
'  ifValidaDoc = False
'  Exit Function
'End If

ifValidaDoc = True

End Function


Sub Carga_Data()

Dim RS As Object
Set RS = CreateObject("ADODB.Recordset")

Set RS = CargarRecordSetDesconectado("Ventas_Up_Man_Exportacion 'V','" & gridex1.Value(gridex1.Columns("Num_Corre").Index) & "'", cCONNECT)

With RS
  If Not (.BOF Or .EOF) Then
    With frmAdicionaDocumVentasExport
    
      .txtCod_TipDoc = RS!Cod_TipDoc
      .txtDes_TipDoc = RS!Des_TipDoc
      .txtCod_TipVenta = RS!Cod_Tipo_Venta
      .txtDes_TipVenta = RS!Des_Tipo_Vent
      
      .txtSer_Docum = RS!Ser_Docum
      .txtNum_Docum = RS!Num_Docum_Ventas
      .strCod_Anxo = RS!Cod_Anxo
      .txtDes_TipAne = RS!DES_ANEXO
      .txtNum_Ruc = RS!Num_Ruc
      .inpFec_EmiDoc.Text = RS!Fec_EmiDoc
      .InpFec_RegDoc.Text = RS!Fec_RegDoc
      .txtTipo_Cambio.Text = RS!Tipo_Cambio
      .txtCod_Moneda = RS!Cod_Moneda
      .txtDes_Moneda = RS!Nom_Moneda
      .txtCod_ConPag = RS!Cod_CondVent
      .txtDes_ConPag = RS!Des_CondVent
      .txtNro_Guias = RS!Guias
      .txtNro_Ordener = RS!Pedidos
      
      .Imp_Gastos_Finacieros.Text = RS!Imp_Gastos_Financieros
      .Imp_Otros.Text = RS!Imp_Otros
      .Imp_Descuento.Text = RS!Imp_Descuento
      .txtImp_Desaduanaje.Text = RS!imp_desaduanaje
      .txtImp_Transporte_Pais_Destino.Text = RS!imp_transporte_pais_destino
      
      .txtGlosa = RS!Glosa
      
      .txtAbr_Cliente = RS!Abr_Cliente
      .txtAbr_Cliente.Tag = RS!Cod_Cliente
      .txtNom_Cliente = RS!Nom_Cliente
      .txtCod_LugEnt = RS!Cod_LugEnt
      .txtNum_CartaCredito = RS!Num_CartaCredito
      .imp_Seguro.Text = RS!imp_Seguro
      .txtCod_Termino_Venta = RS!Cod_Termino_Venta
      .txtDes_Termino_Venta = RS!Des_Termino_Venta
      .txtCod_TipoFact = RS!Cod_TipoFact
      .txtDes_TipoFact = RS!Des_TipoFact
      .txtCod_Embarque = RS!Cod_Embarque
      .txtDes_Embarque = RS!Des_Embarque
      .txtNom_Embarque = RS!Nom_Embarque
      .txtPie_Pagina1 = RS!Pie_Factura1
      .txtPie_Pagina2 = RS!Pie_Factura2
      .txtCod_Vendor = RS!Cod_Vendor
      .txtCod_Class = RS!Cod_Class
      
      .porc_comision.Text = RS!por_comision
      .imp_comision.Text = RS!imp_comision

      .txtCod_TipDoc.Enabled = False
      .txtDes_TipDoc.Enabled = False
      .txtSer_Docum.Enabled = False
      .txtNum_Docum.Enabled = False
      
      
      
      
      If gridex1.Value(gridex1.Columns("Transmision").Index) <> "P" Or gridex1.Value(gridex1.Columns("Impresion").Index) <> "N" Then .frMain.Enabled = False
      
      
      .Imp_Flete.Text = RS!Imp_Flete
      
      
      
      
      
      
      
      If RS!Cod_Mot_Nota <> "" Then
        .txtCod_TipVenta = RS!Cod_Mot_Nota
        .txtDes_TipVenta = RS!Des_MotAbono
      End If
      
    End With
  End If
End With

End Sub


Private Sub Genera_Voucher()
On Error GoTo Fin
Dim sTit As String
Dim sAccion As String, strSQL As String

sAccion = "D"
   sTit = "Generar Voucher De Ventas"
    
   If MsgBox("Genera Voucher De Ventas...?", vbQuestion + vbYesNo, sTit) = vbNo Then Exit Sub
    strSQL = "EXEC CN_GENERA_VOUCHER_VENTAS '" & gridex1.Value(gridex1.Columns("Num_Corre").Index) & "','" & vusu & "'"
    
    
    ExecuteCommandSQL cCONNECT, strSQL
  Buscar
    
Exit Sub
Fin:
    MsgBox err.Description, vbCritical + vbOKOnly, sTit
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
        Call ActualizaFechaCobranzaReprogramada
Case "CANCELAR"
        fraFecCobRepro.Visible = False
End Select
End Sub

Private Sub FunctButt3_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo HandlerError

    Dim Msg As Variant
    
    Select Case ActionName
    Case "VERCOBROS"
        If gridex1.RowCount = 0 Then Exit Sub
        Load FrmVer_Cobros
        FrmVer_Cobros.SNum_Corre = Trim(FixNulos(gridex1.Value(gridex1.Columns("NUM_CORRE").Index), vbString))
        FrmVer_Cobros.strSQL = "Ventas_Muestra_Cobranzas_del_Documento '" & Trim(gridex1.Value(gridex1.Columns("Num_Corre").Index)) & "'"
        FrmVer_Cobros.Buscar
        FrmVer_Cobros.Show vbModal
        Set FrmVer_Cobros = Nothing
    Case "ACTNFOB"
     If gridex1.RowCount > 0 Then
            actualizarNoFob
            Buscar
        Else
            MsgBox "Seleccione un Registro", vbExclamation, "Mensaje"
            Exit Sub
        End If
    Case "LIBERAR"
        If gridex1.RowCount > 0 Then
            If MsgBox("Desea Continuar con la operación ?", vbYesNo, "Mensaje") = vbYes Then
                strSQL = "CN_VENTAS_LIBERAR_ACTUALIZACION_PRECIOS_LDP_DDP '" & Trim(gridex1.Value(gridex1.Columns("Num_Corre").Index)) & "'"
                Call ExecuteCommandSQL(cCONNECT, strSQL)
                MsgBox ("Se Liberó con éxito")
            End If
        End If
        
        
        
    End Select
Exit Sub
Resume
HandlerError:
  errores err.Number
End Sub

Private Sub GridEX1_Click()
indicegrilla = gridex1.Row
End Sub

Sub actualizarNoFob()
On Error GoTo errGrabar
Dim numCorre As String
Dim vMessage As Variant
Dim strSQL As String

numCorre = gridex1.Value(gridex1.Columns("NUM_CORRE").Index)
        
vMessage = MsgBox("Desea Actualizar al estado NO FOB", 48 + 4, "Actualizar Factura")
    If vMessage = vbYes Then
        strSQL = "CN_VENTAS_ACTUALIZA_IMPORTE_NO_FOB '" & numCorre & "'"
        Call ExecuteCommandSQL(cCONNECT, strSQL)
        MsgBox "Transaccion Realizada con Exito", vbInformation, "Mensaje"
        Exit Sub
    End If
Exit Sub
errGrabar:
    MsgBox err.Description, vbCritical, "cerrarCarta"
End Sub


Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
Dim fmtCon  As JSFmtCondition
 
If gridex1.Columns.Count > 2 Then
    If RTrim(RowBuffer.Value(gridex1.Columns("NUM_REGISTRO").Index)) <> "" Then
        RowBuffer.RowStyle = "DOLARES"
    End If
End If
End Sub

Private Sub optAnoMes_Click()
  Limpiar
  
  frAnoMes.Visible = True
  frCliente.Visible = False
  fraClienteCom.Visible = False
  frCorrelativo.Visible = False
  frFecha.Visible = False
  frpo.Visible = False
  Me.chk_po.Value = 0
  strOpcion = "A"
  txtAno.SetFocus
End Sub

  
Private Sub optCliente_Click()
  Limpiar
  
  frAnoMes.Visible = False
  frCliente.Visible = True
  fraClienteCom.Visible = False
  frCorrelativo.Visible = False
  frFecha.Visible = False
  frpo.Visible = False
  Me.chk_po.Value = 0
  strOpcion = "C"
  txtNum_Ruc.SetFocus
End Sub

Private Sub OptClienteComercial_Click()
  Limpiar
  
  frAnoMes.Visible = False
  frCliente.Visible = False
  fraClienteCom.Visible = True
  frCorrelativo.Visible = False
  frFecha.Visible = False
  frpo.Visible = False
  Me.chk_po.Value = 0
  strOpcion = "M"
  txtAbr_Cliente.SetFocus
End Sub

Private Sub optCorrelativo_Click()
  Limpiar
  
  frAnoMes.Visible = False
  frCliente.Visible = False
  fraClienteCom.Visible = False
  frCorrelativo.Visible = True
  frFecha.Visible = False
  frpo.Visible = False
  Me.chk_po.Value = 0
  
  strOpcion = "O"
  txtNum_Corre.SetFocus
End Sub

Private Sub optFecha_Click()
  Limpiar
  
  frAnoMes.Visible = False
  frCliente.Visible = False
  fraClienteCom.Visible = False
  frCorrelativo.Visible = False
  frFecha.Visible = True
  frpo.Visible = False
  Me.chk_po.Value = 0
  strOpcion = "F"
  dtpFecEmiIni.SetFocus
End Sub
Sub Limpiar()

  frFecha.Visible = False
  frCliente.Visible = False
  frAnoMes.Visible = False
  frNroDoc.Visible = False
  frCorrelativo.Visible = False
  frpo.Visible = False
  txtNum_Corre.Text = ""
  txtNum_Ruc.Text = ""
  txtDes_Anexo.Text = ""
  txtAno.Text = ""
  txtMes.Text = ""
  txtCod_TipDoc.Text = ""
  txtDes_TipDoc.Text = ""
  txtSer_Docum.Text = ""
  txtNum_Docum.Text = ""
  txt_po.Text = ""

End Sub

Private Sub optNroDoc_Click()
  Limpiar
  frNroDoc.Visible = True
  frCliente.Visible = False
  fraClienteCom.Visible = False
  frCorrelativo.Visible = False
  frFecha.Visible = False
  frpo.Visible = False
  Me.chk_po.Value = 0
  strOpcion = "N"
  txtCod_TipDoc.SetFocus
End Sub

Private Sub txt_po_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar.SetFocus
End If
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
  If KeyAscii >= 48 And KeyAscii <= 57 _
    Or KeyAscii = 8 Or KeyAscii = 13 Then Else KeyAscii = 0
End Sub

Private Sub txtCod_TipAnxo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAnxo, txtDes_TipDoc, 1, Me)
End Sub

Private Sub txtCod_TipDoc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
   Call Busca_Opcion2("Cod_TipDoc", "Des_TipDoc", "CN_TiposDocum where Flg_Doc_Ventas = '*' and ", txtCod_TipDoc, txtDes_TipDoc, 1, Me)
   txtSer_Docum.SetFocus
    End If
End Sub

Private Sub txtDes_Anexo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion_Anexo("Num_Ruc", "Des_Anexo", " Cn_AnexosContables Where cod_tipanex ='" & Trim(txtCod_TipAnxo.Text) & "' and ", txtNum_Ruc, txtDes_Anexo, 2, Me)
End Sub

Private Sub txtDes_TipDoc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  Call Busca_Opcion2("Cod_TipDoc", "Des_TipDoc", "CN_TiposDocum where Flg_Doc_Ventas = '*' and ", txtCod_TipDoc, txtDes_TipDoc, 2, Me)

  End If
End Sub

Private Sub txtFecCobRepro_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    FunctButt2.SetFocus
End If
End Sub

Private Sub txtImp_Dscto_Penalidad_GotFocus()
    SelectionText txtImp_Dscto_Penalidad
End Sub

Private Sub txtMes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
  If KeyAscii >= 48 And KeyAscii <= 57 _
    Or KeyAscii = 8 Or KeyAscii = 13 Then Else KeyAscii = 0
End Sub

Private Sub TxtNom_Cliente_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    KeyAscii = 0
'    BuscaCliente 2
'    cmdBuscar.SetFocus
'End If

    If KeyAscii = 13 Then
        If Len(txtNom_Cliente) > 4 Then
            strSQL = "SELECT Abr_Cliente FROM TG_CLIENTE WHERE Nom_Cliente LIKE '" & Trim(txtNom_Cliente.Text) & "%'"
            txtAbr_Cliente.Text = DevuelveCampo(strSQL, cCONNECT)
            strSQL = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
            txtNom_Cliente.Text = DevuelveCampo(strSQL, cCONNECT)
            SendKeys "{TAB}"
       Else
            MsgBox ("El Texto Ingresado debe contar con un mínimo de 5 caracteres")
            txtNom_Cliente.SetFocus
        End If
    End If
End Sub

Private Sub txtNum_Corre_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNum_Corre_LostFocus()
  txtNum_Corre = Format(txtNum_Corre, "000000000000")
End Sub

Private Sub txtNum_Docum_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then cmdBuscar.SetFocus
End Sub

Private Sub txtNum_Docum_LostFocus()
  txtNum_Docum = Format(txtNum_Docum, "00000000")
End Sub

Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then Call Busca_Opcion_Anexo("Num_Ruc", "Des_Anexo", " Cn_AnexosContables Where cod_tipanex ='" & Trim(txtCod_TipAnxo.Text) & "' and ", txtNum_Ruc, txtDes_Anexo, 1, Me)
End Sub

Private Sub txtSer_Docum_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then txtNum_Docum.SetFocus
End Sub

Private Sub txtSer_Docum_LostFocus()
  txtSer_Docum = Format(txtSer_Docum, "000")
End Sub


Private Sub txtAbr_Cliente_Change()
        txtAbr_Cliente.Tag = ""
End Sub

Private Sub TxtAbr_Cliente_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        BuscaCliente 1
'        SendKeys "{TAB}"
'    End If
    If KeyAscii = 13 Then
        If Trim(txtAbr_Cliente.Text) = "" Then
            cmdBusCliente_Click
        Else
            strSQL = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE Abr_Cliente LIKE '" & Trim(txtAbr_Cliente.Text) & "%'"
            txtNom_Cliente.Text = DevuelveCampo(strSQL, cCONNECT)
            'SendKeys "{TAB}"
            Me.cmdBuscar.SetFocus
        End If
    End If
End Sub

Private Sub cmdBusCliente_Click()
    Dim oTipo As New frmBusqGeneral
    Dim RS As Object
    Set RS = CreateObject("ADODB.Recordset")
    Set oTipo.oParent = Me
    oTipo.SQuery = "SELECT Abr_Cliente as Código, nom_cliente as Descripción FROM TG_Cliente ORDER BY Abr_Cliente"
    oTipo.CARGAR_DATOS
    oTipo.Show 1
    If codigo <> "" Then
        txtAbr_Cliente.Text = codigo
        txtNom_Cliente.Text = Descripcion
        SendKeys "{TAB}"
        codigo = ""
    End If
    Set oTipo = Nothing
    Set RS = Nothing
End Sub

Public Sub BuscaCliente(opcion As String)
Dim rstAux As Object
Dim strSQL As String

Set rstAux = CreateObject("ADODB.Recordset")

    strSQL = "SELECT Cod_Cliente, Abr_Cliente, Nom_Cliente FROM TG_CLIENTE WHERE "
    
    txtAbr_Cliente = Trim(txtAbr_Cliente)
    txtNom_Cliente = Trim(txtNom_Cliente)
    
    Select Case opcion
    Case 1: strSQL = strSQL & "Abr_Cliente LIKE '%" & txtAbr_Cliente & "%'"
    Case 2: strSQL = strSQL & "Nom_Cliente LIKE '%" & txtNom_Cliente & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    
    frmBusqGeneral3.gexLista.Columns("Cod_Cliente").Visible = False
    frmBusqGeneral3.gexLista.Columns("Abr_Cliente").Width = 570
    frmBusqGeneral3.gexLista.Columns("Nom_Cliente").Width = 2370
    
    frmBusqGeneral3.gexLista.Columns("Abr_Cliente").Caption = "Abrev."
    frmBusqGeneral3.gexLista.Columns("Nom_Cliente").Caption = "Cliente"
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtAbr_Cliente.Tag = ""
    txtAbr_Cliente = ""
    txtNom_Cliente = ""
    If codigo <> "" Then
        txtAbr_Cliente = Descripcion
        txtNom_Cliente = TipoAdd
        txtAbr_Cliente.Tag = codigo
    End If
    codigo = ""
    Descripcion = ""
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
End Sub


Private Sub Imprimir(ByVal SNum_Corre As String, ByVal SCod_TipDoc As String)
On Error GoTo ErrorImpresion
Dim oo As Object
Dim sFormato_Invoice As String
Dim strSQL As String
    Dim sRutaLogo As String
    strSQL = "SELECT Ruta_Logo = ISNULL(Ruta_Logo, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
    sRutaLogo = DevuelveCampo(strSQL, cCONNECT)
    
    Dim sEmpresa As String
    strSQL = "SELECT Des_Empresa = ISNULL(Des_Empresa, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
    sEmpresa = DevuelveCampo(strSQL, cCONNECT)
        

    sFormato_Invoice = DevuelveCampo("SELECT FORMATO_INVOICE FROM TG_CLIENTE WHERE COD_CLIENTE = '" & gridex1.Value(gridex1.Columns("COD_CLIENTE").Index) & "'", cCONNECT)
    Set oo = CreateObject("excel.application")
    Select Case SCod_TipDoc
        Case "FA"
            oo.Workbooks.Open vRuta & "\Invoice" & sFormato_Invoice & ".XLT"
    End Select
    oo.Visible = True
    oo.displayalerts = False
    
    If sFormato_Invoice = "01" Then
        oo.Run "reporte", cCONNECT, SNum_Corre, sEmpresa, sRutaLogo
    Else
        oo.Run "reporte", cCONNECT, SNum_Corre
    End If
    
    Set oo = Nothing
       
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub


Private Sub Imprimir1(ByVal SNum_Corre As String, ByVal SCod_TipDoc As String)
On Error GoTo ErrorImpresion
Dim oo As Object
Dim sFormato_Invoice As String

    sFormato_Invoice = DevuelveCampo("SELECT FORMATO_INVOICE FROM TG_CLIENTE WHERE COD_CLIENTE = '" & gridex1.Value(gridex1.Columns("COD_CLIENTE").Index) & "'", cCONNECT)
    Set oo = CreateObject("excel.application")
    Select Case SCod_TipDoc
        Case "FA"
        If vemp = "01" Then
            oo.Workbooks.Open vRuta & "\Invoice05.XLT"
        ElseIf vemp = "03" Then
            oo.Workbooks.Open vRuta & "\Invoice05_inka.XLT"
        End If
    End Select
    oo.Visible = True
    oo.displayalerts = False
    oo.Run "reporte", cCONNECT, SNum_Corre
    Set oo = Nothing
       
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub

Private Sub GeneraInfoContable()
On Error GoTo errx
Dim vResp As Variant
Dim sSQL As String

vResp = MsgBox("Confirma Generación Contable de Documento ? ", vbYesNo, "CONFIRMACION")
If vResp = vbNo Then Exit Sub

sSQL = "CN_GENERA_ASIENTO_VENTAS '" & gridex1.Value(gridex1.Columns("Num_Corre").Index) & "'"

ExecuteCommandSQL cCONNECT, sSQL
Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
Buscar
Exit Sub

errx:
    errores err.Number
End Sub
Private Sub MuestraVoucher2()

On Error GoTo errx
Dim sSQL As String
Dim rsAsientos As Object


If gridex1.RowCount = 0 Then Exit Sub
  
  If RTrim(gridex1.Value(gridex1.Columns("num_registro").Index)) = "" Then
    MsgBox "No se le ha Generado Voucher", vbInformation, "AVISO"
    Exit Sub
  End If

  Load frmShowVoucher
  frmShowVoucher.sCod_TipoDiario = RTrim(DevuelveCampo("select Cod_TipodiarioVentas  from cn_control ", cCONNECT))
  frmShowVoucher.sano = RTrim(gridex1.Value(gridex1.Columns("Ano_Registro").Index))
  frmShowVoucher.smes = RTrim(gridex1.Value(gridex1.Columns("Mes_registro").Index))
  frmShowVoucher.lNum_Registro = RTrim(gridex1.Value(gridex1.Columns("Num_Registro").Index))
  frmShowVoucher.Num_Corre = gridex1.Value(gridex1.Columns("Num_Corre").Index)
  'frmShowVoucher.dImporte = GridEX1.Value(GridEX1.Columns("Imp_Total").Index)
  'frmShowVoucher.sFlg_Status = GridEX1.Value(GridEX1.Columns("Estatus_Letra").Index)
  frmShowVoucher.Buscar
  frmShowVoucher.FunctButt1.ChangeProperty "ENABLED", 1, False
  frmShowVoucher.Show vbModal
  Set frmShowVoucher = Nothing
  
Exit Sub

Resume
errx:
    errores err.Number

End Sub




Private Sub LoadPenalidades()
SNum_Corre = gridex1.Value(gridex1.Columns("NUM_CORRE").Index)
fraPenalidad.Visible = True
txtImp_Dscto_Penalidad.Text = gridex1.Value(gridex1.Columns("Imp_Descuento_Penalidad").Index)
txtImp_Dscto_Penalidad.SetFocus
End Sub

Private Sub FunctButt5_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            GrabarPenalidad
        Case "CANCELAR"
            Me.fraPenalidad.Visible = False
    End Select
End Sub

Private Sub ActualizaFechaCobranzaReprogramada()
On Error GoTo errores
Dim sSQL As String
Dim sFlg_Pendalidad As String

If txtFecCobRepro.Text = "" Then
    sSQL = "CN_VENTAS_ACTUALIZA_FEC_COBRANZA_REPROGRAMADA '" & SNum_Corre & "',null"
Else
    sSQL = "CN_VENTAS_ACTUALIZA_FEC_COBRANZA_REPROGRAMADA '" & SNum_Corre & "','" & txtFecCobRepro.Text & "'"
End If


  
ExecuteCommandSQL cCONNECT, sSQL

Me.fraFecCobRepro.Visible = False
Buscar

Exit Sub

errores:
    errores err.Number
End Sub

Private Sub GrabarPenalidad()
On Error GoTo errores
Dim sSQL As String
Dim sFlg_Pendalidad As String

sSQL = "TG_EMBARQUE_DATOS_PENALIDAD_DESCUENTO '$','$'"
sSQL = VBsprintf(sSQL, SNum_Corre, txtImp_Dscto_Penalidad.Text)
  
ExecuteCommandSQL cCONNECT, sSQL

Me.fraPenalidad.Visible = False
Buscar

Exit Sub

errores:
    errores err.Number
End Sub

