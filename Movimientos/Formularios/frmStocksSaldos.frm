VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmStocksSaldos 
   Caption         =   "Ver Saldo de Stocks"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   4620
      TabIndex        =   4
      Top             =   6150
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "0~0~IMPRIMIR~True~True~&Imprimir~0~0~1~~0~False~False~&Imprimir~~1~0~SALIR~True~True~&Salir~0~0~2~~0~False~False~&Salir~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   150
      TabIndex        =   5
      Top             =   60
      Width           =   11655
      Begin VB.TextBox txtDes_Fabrica 
         Height          =   285
         Left            =   1800
         TabIndex        =   18
         Top             =   705
         Width           =   2175
      End
      Begin VB.TextBox txtCod_Fabrica 
         Height          =   285
         Left            =   975
         TabIndex        =   16
         Top             =   705
         Width           =   780
      End
      Begin VB.TextBox txtCod_OrdPro 
         Height          =   285
         Left            =   5070
         TabIndex        =   1
         Top             =   735
         Width           =   780
      End
      Begin VB.ComboBox cboAlmacen 
         Height          =   315
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   255
         Width           =   3315
      End
      Begin FunctionsButtons.FunctButt fnbBuscar 
         Height          =   495
         Left            =   10215
         TabIndex        =   2
         Top             =   345
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~True~True~&Buscar~0~0~1~~0~False~False~&Buscar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label1 
         Caption         =   "Fabrica"
         Height          =   240
         Left            =   135
         TabIndex        =   17
         Top             =   750
         Width           =   810
      End
      Begin VB.Label lblCod_OrdTra 
         Caption         =   "OrdTra"
         Height          =   240
         Left            =   4215
         TabIndex        =   15
         Top             =   795
         Width           =   810
      End
      Begin VB.Label lblCod_PurOrd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   6225
         TabIndex        =   14
         Top             =   765
         Width           =   1530
      End
      Begin VB.Label lblCod_EstCli 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4830
         TabIndex        =   13
         Top             =   270
         Width           =   1650
      End
      Begin VB.Label Label7 
         Caption         =   "PO"
         Height          =   210
         Left            =   5910
         TabIndex        =   12
         Top             =   810
         Width           =   300
      End
      Begin VB.Label Label6 
         Caption         =   "EST"
         Height          =   195
         Left            =   4455
         TabIndex        =   11
         Top             =   300
         Width           =   300
      End
      Begin VB.Label Label4 
         Caption         =   "Desp."
         Height          =   195
         Left            =   7845
         TabIndex        =   10
         Top             =   780
         Width           =   435
      End
      Begin VB.Label lblFecDesp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   8340
         TabIndex        =   9
         Top             =   765
         Width           =   1665
      End
      Begin VB.Label Label5 
         Caption         =   "CLIENTE"
         Height          =   195
         Left            =   6705
         TabIndex        =   8
         Top             =   300
         Width           =   675
      End
      Begin VB.Label lblNom_Cli 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   7440
         TabIndex        =   7
         Top             =   270
         Width           =   2520
      End
      Begin VB.Label Label3 
         Caption         =   "Alamcén:"
         Height          =   255
         Left            =   135
         TabIndex        =   6
         Top             =   300
         Width           =   750
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4680
      Left            =   90
      TabIndex        =   3
      Top             =   1275
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   8255
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ContScroll      =   -1  'True
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmStocksSaldos.frx":0000
      Column(2)       =   "frmStocksSaldos.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmStocksSaldos.frx":016C
      FormatStyle(2)  =   "frmStocksSaldos.frx":02A4
      FormatStyle(3)  =   "frmStocksSaldos.frx":0354
      FormatStyle(4)  =   "frmStocksSaldos.frx":0408
      FormatStyle(5)  =   "frmStocksSaldos.frx":04E0
      FormatStyle(6)  =   "frmStocksSaldos.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmStocksSaldos.frx":0678
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   1770
      Top             =   6090
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmStocksSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Codigo As String, Descripcion As String
Dim Strsql As String, sTit_OP As String, rstAux As ADODB.Recordset

Private Sub cboAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub fnbBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    If cboAlmacen.ListIndex = -1 Then
        MsgBox "Se debe elegir un Almacen", vbOKOnly + vbExclamation, "Ver Saldo de Stocks"
    End If
    Strsql = "EXEC SM_MUESTRA_CF_STOCKS_SALDOS '" & Left(cboAlmacen, 2) & _
             "', '" & txtCod_Fabrica & "', '" & txtCod_OrdPro & "'"
    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(Strsql, cConnect)
    
    GridEX1.Columns("Cli.").Width = 450
    GridEX1.Columns("OP").Width = 570
    GridEX1.Columns("PO").Width = 1290
    GridEX1.Columns("Estilo_Propio").Width = 1050
    GridEX1.Columns("Estilo_Cliente").Width = 1365
    GridEX1.Columns("Color").Width = 1245
    GridEX1.Columns("Talla").Width = 795
    GridEX1.Columns("Calidad").Width = 660
    GridEX1.Columns("Desc.Calidad").Width = 1080
    GridEX1.Columns("Stock").Width = 800
    GridEX1.Columns("Fecha_Entrada").Width = 1500
    
    GridEX1.Columns("OP").Caption = sTit_OP
    
    If GridEX1.RowCount > 0 Then
        GridEX1.SetFocus
    Else
        cboAlmacen.SetFocus
    End If
    
End Sub

Private Sub FillAlmacen()
Dim rstAlm As ADODB.Recordset

    Strsql = "SELECT Cod_Almacen, Nom_Almacen FROM CF_Almacen " & _
             "WHERE Flg_Saldos = '*' OR Flg_Saldos_Por_Ingresar = '*'"
    Set rstAlm = CargarRecordSetDesconectado(Strsql, cConnect)
    rstAlm.MoveFirst
    cboAlmacen.Clear
    Do Until rstAlm.EOF
        cboAlmacen.AddItem rstAlm!Cod_Almacen & " " & rstAlm!Nom_Almacen
        rstAlm.MoveNext
    Loop
    rstAlm.Close
    Set rstAlm = Nothing
End Sub

Private Sub Form_Load()
    Strsql = "Select Top 1 Titulo_Abr_Orden From TG_Control"
    sTit_OP = DevuelveCampo(Strsql, cConnect)
    lblCod_OrdTra = sTit_OP
    FillAlmacen
    Strsql = "Select count(Cod_Fabrica) From TG_FABRICA"
    If DevuelveCampo(Strsql, cConnect) = 1 Then BuscaFabrica 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "IMPRIMIR"
        Reporte
    Case "SALIR"
        Unload Me
    End Select
End Sub

Private Sub Reporte()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String

    Ruta = vRuta & "\StocksSaldos.xlt"
    Screen.MousePointer = 11
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    Strsql = "SELECT Ruta_Logo FROM SEGURIDAD..SEG_EMPRESAS WHERE cod_EMPRESA ='" & _
             Trim(vemp1) & "'"
    oo.Run "reporte", sTit_OP, IIf(txtCod_OrdPro = "", "", txtCod_Fabrica), txtCod_OrdPro, Left(cboAlmacen, 2), Mid(cboAlmacen, 4), DevuelveCampo(Strsql, cConnect), cConnect
    Set oo = Nothing
    Screen.MousePointer = 0
Exit Sub
hand:
    Screen.MousePointer = 0
    ErrorHandler Err, "Reporte"
    Set oo = Nothing
End Sub

Private Sub txtCod_Fabrica_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        BuscaFabrica 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtcod_ordpro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtCod_OrdPro.Text = Format(txtCod_OrdPro, "00000")
        
        Strsql = "SELECT a.Cod_EstCli, a.Cod_PurOrd, a.Fec_DespachoAct, b.Nom_Cliente " & _
                 "FROM ES_ORDPRO a, TG_CLIENTE b " & _
                 "WHERE a.Cod_OrdPro = '" & txtCod_OrdPro & "' " & _
                 "AND   b.Cod_Cliente = a.Cod_Cliente " & _
                 "AND   a.Cod_Fabrica = '" & txtCod_Fabrica & "'"
        
        Set rstAux = CargarRecordSetDesconectado(Strsql, cConnect)
        If rstAux.RecordCount = 0 Then
            MsgBox "La Orden de Trabajo no Existe", vbCritical, "Orden de Trabajo"
            lblCod_PurOrd = ""
            lblCod_EstCli = ""
            lblFecDesp = ""
            lblNom_Cli = ""
            txtCod_OrdPro = ""
        Else
            rstAux.MoveFirst
            lblCod_PurOrd = rstAux!Cod_PurOrd
            lblCod_EstCli = rstAux!Cod_EstCli
            lblNom_Cli = rstAux!Nom_Cliente
            lblFecDesp = Format(rstAux!Fec_DespachoAct, "dd/mm/yyyy")
        End If
        SendKeys "{TAB}"
    End If
End Sub

Private Sub BuscaFabrica(Opcion As Integer)
On Error GoTo Fin
Dim sTit As String, sCod_Fabrica As String, iRegs As Long
    sTit = "Buscar Fabrica"
    
    txtCod_Fabrica = Trim(txtCod_Fabrica)
    txtDes_Fabrica = Trim(txtDes_Fabrica)
    
    Strsql = "Select Cod_Fabrica, Nom_Fabrica From TG_FABRICA WHERE "
    Select Case Opcion
    Case 1: Strsql = Strsql & "Cod_Fabrica like '%" & txtCod_Fabrica & "%'"
    Case 2: Strsql = Strsql & "Des_Fabrica like '%" & txtDes_Fabrica & "%'"
    End Select
    
    sCod_Fabrica = txtCod_Fabrica
    
    txtCod_Fabrica = ""
    txtDes_Fabrica = ""
    With frmBusqGeneral2
        Set .oParent = Me
        .sQuery = Strsql
        .Cargar_Datos
        Set rstAux = .DGridLista.DataSource
        iRegs = rstAux.RecordCount
        If iRegs > 1 Then .Show vbModal
        If iRegs = 1 Then .cmdAceptar_Click
        
        If Codigo <> "" Then
            If sCod_Fabrica = "" And Opcion = 1 And iRegs = 1 Then
                txtCod_Fabrica.Enabled = False
                txtDes_Fabrica.Enabled = False
            End If
            txtCod_Fabrica = Codigo
            txtDes_Fabrica = Descripcion
        End If
    End With
    Unload frmBusqGeneral2
Exit Sub
Fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, sTit
End Sub

Private Sub txtDes_Fabrica_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        BuscaFabrica 2
        SendKeys "{TAB}"
    End If
End Sub
