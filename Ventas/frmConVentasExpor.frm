VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmConVentasExpor 
   Caption         =   "Reporte de Ventas de Exportacion"
   ClientHeight    =   7845
   ClientLeft      =   345
   ClientTop       =   780
   ClientWidth     =   11640
   Icon            =   "frmConVentasExpor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   11640
   WindowState     =   2  'Maximized
   Begin VB.Frame FraBuscar 
      Caption         =   "Argumentos de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11520
      Begin VB.CheckBox ChkMuestraAnulada 
         Alignment       =   1  'Right Justify
         Caption         =   "Muestras Anuladas"
         Height          =   195
         Left            =   9555
         TabIndex        =   28
         Top             =   1245
         Width           =   1740
      End
      Begin VB.TextBox TxtNom_Cliente 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   3000
         MaxLength       =   30
         TabIndex        =   27
         Top             =   1200
         Visible         =   0   'False
         Width           =   5610
      End
      Begin VB.TextBox TxtAbr_cliente 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   2280
         MaxLength       =   11
         TabIndex        =   26
         Top             =   1200
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.OptionButton OptAnexo 
         Caption         =   "&Por Anexo"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptClienteComercial 
         Caption         =   "&Por Cliente Comercial"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Frame fraOP 
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   45
         TabIndex        =   20
         Top             =   1440
         Width           =   9270
         Begin VB.TextBox txtCod_Fabrica 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   795
            TabIndex        =   6
            Top             =   180
            Width           =   570
         End
         Begin VB.TextBox txtNom_Fabrica 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1395
            TabIndex        =   7
            Top             =   180
            Width           =   1500
         End
         Begin VB.TextBox txtCod_OrdPro 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3930
            TabIndex        =   8
            Top             =   210
            Width           =   750
         End
         Begin VB.Label Label5 
            Caption         =   "Fabrica"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   90
            TabIndex        =   23
            Top             =   210
            Width           =   705
         End
         Begin VB.Label lblTit_OP 
            Caption         =   "N/P"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3150
            TabIndex        =   22
            Top             =   240
            Width           =   750
         End
         Begin VB.Label lblDes_EstPro 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4725
            TabIndex        =   21
            Top             =   210
            Width           =   4440
         End
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   675
         Left            =   8280
         TabIndex        =   9
         Top             =   120
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   1191
         Custom          =   $"frmConVentasExpor.frx":030A
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1000
         ControlHeigth   =   650
         ControlSeparator=   40
      End
      Begin VB.CheckBox chkMuestras 
         Alignment       =   1  'Right Justify
         Caption         =   "Solo &Muestras"
         Height          =   255
         Left            =   9840
         TabIndex        =   19
         Top             =   855
         Width           =   1455
      End
      Begin VB.TextBox txtDes_TipAxo 
         Height          =   285
         Left            =   8400
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtCod_TipAnxo 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   9240
         MaxLength       =   1
         TabIndex        =   16
         Text            =   "C"
         Top             =   840
         Width           =   360
      End
      Begin VB.TextBox txtNum_Ruc 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   2280
         MaxLength       =   11
         TabIndex        =   4
         Top             =   840
         Width           =   1200
      End
      Begin VB.TextBox txtDes_Anexo 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   3480
         MaxLength       =   30
         TabIndex        =   5
         Top             =   840
         Width           =   5130
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   5520
         TabIndex        =   13
         Top             =   120
         Width           =   2655
         Begin VB.OptionButton optResumido 
            Caption         =   "&Resumido"
            Height          =   195
            Left            =   1320
            TabIndex        =   3
            Top             =   270
            Width           =   1095
         End
         Begin VB.OptionButton optDetallado 
            Caption         =   "&Detallado"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   270
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   330
         Left            =   1200
         TabIndex        =   0
         Top             =   337
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   582
         _Version        =   393216
         Format          =   62652417
         CurrentDate     =   37987
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   330
         Left            =   3840
         TabIndex        =   1
         Top             =   337
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   582
         _Version        =   393216
         Format          =   62652417
         CurrentDate     =   37987
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo :"
         Height          =   255
         Left            =   8760
         TabIndex        =   17
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Nro Ruc:"
         Height          =   255
         Left            =   1560
         TabIndex        =   15
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final :"
         Height          =   195
         Left            =   2880
         TabIndex        =   14
         Top             =   405
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial :"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   405
         Width           =   990
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5580
      Left            =   0
      TabIndex        =   10
      Top             =   2280
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   9843
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmConVentasExpor.frx":03BF
      Column(2)       =   "frmConVentasExpor.frx":0487
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmConVentasExpor.frx":052B
      FormatStyle(2)  =   "frmConVentasExpor.frx":0663
      FormatStyle(3)  =   "frmConVentasExpor.frx":0713
      FormatStyle(4)  =   "frmConVentasExpor.frx":07C7
      FormatStyle(5)  =   "frmConVentasExpor.frx":089F
      FormatStyle(6)  =   "frmConVentasExpor.frx":0957
      FormatStyle(7)  =   "frmConVentasExpor.frx":0A37
      FormatStyle(8)  =   "frmConVentasExpor.frx":0AE3
      ImageCount      =   0
      PrinterProperties=   "frmConVentasExpor.frx":0B93
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   10875
      Top             =   5985
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmConVentasExpor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public codigo As String, Descripcion As String, TipoAdd As String
Dim strCod_Anxo As String
Dim strSql As String

Private Sub dtpFecIni_Change()
  dtpFecFin = dtpFecIni
End Sub

Private Sub Form_Load()
  ChkMuestraAnulada.Value = 1
  dtpFecIni = Date
  dtpFecFin = Date
End Sub

Private Sub Buscar()

On Error GoTo drDepurar

Dim ssql As String
Dim fmtCon As JSFmtCondition
Dim VerAnulados  As String
  
  If ChkMuestraAnulada.Value = 1 Then
   VerAnulados = "S"
  Else
   VerAnulados = "N"
  End If
  
  ssql = "Ventas_Muestra_Documento_Exportacion '" & dtpFecIni & "','" & dtpFecFin & "','" & IIf(optDetallado, "D", "R") & "','" & IIf(txtNum_Ruc = "", "", txtCod_TipAnxo) & "','" & IIf(txtNum_Ruc = "", "", strCod_Anxo) & "','" & chkMuestras.Value & "','" & txtCod_Fabrica.Text & "','" & txtCod_OrdPro & "','" & IIf(OptAnexo, "1", "2") & "','" & TxtAbr_cliente.Tag & "','" & VerAnulados & "'"
  
  Set GridEX1.ADORecordset = CargarRecordSetDesconectado(ssql, cCONNECT)
  
  GridEX1.Columns("Fecha").Width = 1020
  GridEX1.Columns("Factura").Width = 795
  GridEX1.Columns("Cliente").Width = 2715
  GridEX1.Columns("NP").Width = 645
  GridEX1.Columns("guias").Width = 1110
  GridEX1.Columns("Destino").Width = 1515
  GridEX1.Columns("Prendas").Width = 720
  GridEX1.Columns("Total_FOB").Width = 900
  GridEX1.Columns("Comision").Width = 765
  GridEX1.Columns("Fletes").Width = 555
  GridEX1.Columns("Peso_Neto").Width = 930
  GridEX1.Columns("Peso_Bruto").Width = 960
  GridEX1.Columns("Ship_Date").Width = 945
  GridEX1.Columns("Observacion").Width = 4680
  GridEX1.Columns("Total_FOB").Format = "###,###.00"
  
Exit Sub
Resume
drDepurar:
  errores Err.Number
End Sub

Public Sub Reporte()
  
On Error GoTo ErrorImpresion

    VB.Screen.MousePointer = vbHourglass
    
    Dim oo As Object
    Set oo = CreateObject("excel.application")
    
    If optDetallado Then
      oo.Workbooks.Open vRuta & "\ReporteDocumentosExportaciones.xlt"
      oo.Visible = True
      oo.Run "REPORTE", GridEX1.ADORecordset, "DOCUMENTOS DE EXPORTACION  DETALLADO DESDE EL " & dtpFecIni & " HASTA EL " & dtpFecFin & IIf(txtNum_Ruc <> "", " DEL CLIENTE " & txtDes_Anexo, "") & IIf(chkMuestras, " MUESTRAS ", "")
    Else
      oo.Workbooks.Open vRuta & "\ReporteDocumentosExportaciones.xlt"
      oo.Visible = True
      oo.Run "REPORTE", GridEX1.ADORecordset, "DOCUMENTOS DE EXPORTACION  RESUMIDO DESDE EL " & dtpFecIni & " HASTA EL " & dtpFecFin & IIf(txtNum_Ruc <> "", " DEL CLIENTE " & txtDes_Anexo, "") & IIf(chkMuestras, " MUESTRAS ", "")
    End If
    
    
    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
    
    Exit Sub
    Resume
ErrorImpresion:
    Screen.MousePointer = vbNormal
    Set oo = Nothing
    Error Err.Number
End Sub


Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Msg As Variant
    Select Case ActionName
    Case "BUSCAR"
      Buscar
    Case "IMPRIMIR"
        If GridEX1.RowCount = 0 Then Exit Sub
        Reporte
    Case "SALIR"
       Unload Me
    End Select
End Sub

Private Sub OptAnexo_Click()
    txtNum_Ruc.Visible = True
    txtDes_Anexo.Visible = True
    TxtAbr_cliente.Visible = False
    TxtNom_Cliente.Visible = False
    TxtAbr_cliente.Text = ""
    TxtNom_Cliente.Text = ""
End Sub

Private Sub optClienteComercial_Click()
    Label4.Visible = False
    txtNum_Ruc.Visible = False
    txtDes_Anexo.Visible = False
    TxtAbr_cliente.Visible = True
    TxtNom_Cliente.Visible = True
    txtNum_Ruc.Text = ""
    txtDes_Anexo.Text = ""
End Sub



Private Sub TxtAbr_cliente_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    BuscaCliente 1
    FunctButt1.SetFocus
End If
End Sub

Private Sub txtCod_TipAnxo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAnxo, txtDes_TipAxo, 1, Me)
End Sub

Private Sub txtDes_Anexo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then Call Busca_Opcion_Anexo("Num_Ruc", "Des_Anexo", " Cn_AnexosContables where cod_tipanex = '" & txtCod_TipAnxo & "' and ", txtNum_Ruc, txtDes_Anexo, 2)
End Sub

Private Sub TxtNom_Cliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    BuscaCliente 2
    FunctButt1.SetFocus
End If
End Sub

Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then Call Busca_Opcion_Anexo("Num_Ruc", "Des_Anexo", " Cn_AnexosContables where cod_tipanex = '" & txtCod_TipAnxo & "' and ", txtNum_Ruc, txtDes_Anexo, 1)
End Sub


Sub Busca_Opcion_Anexo(strCampo1 As String, strCampo2 As String, StrTabla As String, txtCod As TextBox, txtDes As TextBox, Opcion As Integer)

On Error GoTo Fin

Dim rstAux As ADODB.Recordset, strSql As String
    strSql = "select Cod_Anxo as Cod,Des_Anexo as Nombre,Num_Ruc as Ruc from cn_anexoscontables where cod_tipanex = 'C' and "

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    Select Case Opcion
    Case 1: strSql = strSql & strCampo1 & " like '%" & txtCod & "%'"
    Case 2: strSql = strSql & strCampo2 & " like '%" & txtDes & "%'"
    End Select
    txtCod = ""
    txtDes = ""
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = strSql
        .Cargar_Datos
        
        codigo = ".."
        .DGridLista.Columns("Cod").Visible = False
        .DGridLista.Columns("Nombre").Width = 4575
        .DGridLista.Columns("RUC").Width = 1695
        Set rstAux = .DGridLista.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If codigo <> "" And rstAux.RecordCount > 0 Then
            strCod_Anxo = Trim(rstAux!Cod)
            txtDes = Trim(rstAux!Nombre)
            txtCod = Trim(rstAux!Ruc)
            Select Case Opcion
            Case 1: SendKeys "{TAB}": SendKeys "{TAB}"
            Case 2: SendKeys "{TAB}"
            End Select
        Else
            SendKeys "{TAB}"
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & Opcion & ")"
End Sub



Private Sub txtCod_Fabrica_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If txtCod_Fabrica.Text = "" Then
            BuscaFabrica 2
        Else
            BuscaFabrica 1
        End If
    End If
End Sub

Private Sub txtCod_OrdPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        txtCod_OrdPro = Format(txtCod_OrdPro, "00000")
        strSql = "SELECT b.Des_EstPro FROM ES_ORDPRO a, ES_ESTPRO b " & _
                 "WHERE a.Cod_OrdPro = '" & txtCod_OrdPro & "' " & _
                 "AND   a.Cod_EstPro = b.Cod_EstPro"
        lblDes_EstPro = DevuelveCampo(strSql, cCONNECT)
        SendKeys "{TAB}"
    End If
End Sub




Private Sub BuscaFabrica(Tipo As Integer)


    Select Case Tipo
        Case 1:
                              
                    strSql = "Select  Nom_Fabrica From TG_FABRICA WHERE Cod_Fabrica = '" & Trim(txtCod_Fabrica) & "'"
                    txtNom_Fabrica.Text = Trim(DevuelveCampo(strSql, cCONNECT))
                    SendKeys "{TAB}"
                    
        Case 2:
                    Dim oTipo As New frmBusqGeneral
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    oTipo.SQuery = "Select Cod_Fabrica, Nom_Fabrica From TG_FABRICA"
                    oTipo.Cargar_Datos
                    oTipo.Show 1
                    If codigo <> "" Then
'                         Set GridEX1.Recordset = Nothing
                         txtCod_Fabrica.Text = Trim(codigo)
                         txtNom_Fabrica.Text = Trim(Descripcion)
                         codigo = "": Descripcion = ""
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
                    SendKeys "{TAB}"
    End Select
    
End Sub


Private Sub txtNom_Fabrica_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If txtCod_Fabrica = "" Then
            BuscaFabrica 2
        End If
        SendKeys "{TAB}"
    End If
End Sub

Public Sub BuscaCliente(Opcion As String)
Dim rstAux As ADODB.Recordset
Dim strSql As String

    strSql = "SELECT Cod_Cliente, Abr_Cliente, Nom_Cliente FROM TG_CLIENTE WHERE "
    
    TxtAbr_cliente = Trim(TxtAbr_cliente)
    TxtNom_Cliente = Trim(TxtNom_Cliente)
    
    Select Case Opcion
    Case 1: strSql = strSql & "Abr_Cliente LIKE '%" & TxtAbr_cliente & "%'"
    Case 2: strSql = strSql & "Nom_Cliente LIKE '%" & TxtNom_Cliente & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSql
    frmBusqGeneral3.Cargar_Datos
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
    
    TxtAbr_cliente.Tag = ""
    TxtAbr_cliente = ""
    TxtNom_Cliente = ""
    If codigo <> "" Then
        TxtAbr_cliente = Descripcion
        TxtNom_Cliente = TipoAdd
        TxtAbr_cliente.Tag = codigo
    End If
    codigo = ""
    Descripcion = ""
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
End Sub

