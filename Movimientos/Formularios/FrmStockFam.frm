VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmStockFam 
   Caption         =   "Stocks por Familia"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   90
      TabIndex        =   4
      Top             =   1020
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Stocks"
      TabPicture(0)   =   "FrmStockFam.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Grilla1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FunctButt1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FraObsoletos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FraUbicacion"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame FraUbicacion 
         Caption         =   "Ubicacion Fisica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   3240
         TabIndex        =   11
         Top             =   2520
         Visible         =   0   'False
         Width           =   4695
         Begin FunctionsButtons.FunctButt FunctButt3 
            Height          =   510
            Left            =   1080
            TabIndex        =   14
            Top             =   960
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   900
            Custom          =   $"FrmStockFam.frx":001C
            Orientacion     =   0
            Style           =   0
            Language        =   0
            TypeImageList   =   0
            ControlWidth    =   1155
            ControlHeigth   =   490
            ControlSeparator=   110
         End
         Begin VB.TextBox TxtUbicacion 
            Height          =   495
            Left            =   1080
            TabIndex        =   13
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label Label3 
            Caption         =   "Ubicacion"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame FraObsoletos 
         Height          =   1335
         Left            =   3480
         TabIndex        =   7
         Top             =   2520
         Visible         =   0   'False
         Width           =   3975
         Begin FunctionsButtons.FunctButt FunctButt2 
            Height          =   510
            Left            =   720
            TabIndex        =   10
            Top             =   720
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   900
            Custom          =   $"FrmStockFam.frx":00B2
            Orientacion     =   0
            Style           =   0
            Language        =   0
            TypeImageList   =   0
            ControlWidth    =   1155
            ControlHeigth   =   490
            ControlSeparator=   110
         End
         Begin MSComCtl2.DTPicker DTPFecha 
            Height          =   255
            Left            =   1800
            TabIndex        =   9
            Top             =   300
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            _Version        =   393216
            Format          =   73531393
            CurrentDate     =   38663
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Máxima:"
            Height          =   195
            Left            =   360
            TabIndex        =   8
            Top             =   380
            Width           =   1080
         End
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   585
         Left            =   960
         TabIndex        =   6
         Top             =   4320
         Width           =   8790
         _ExtentX        =   15505
         _ExtentY        =   1032
         Custom          =   $"FrmStockFam.frx":014B
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   560
         ControlSeparator=   110
      End
      Begin GridEX20.GridEX Grilla1 
         Height          =   3840
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   10380
         _ExtentX        =   18309
         _ExtentY        =   6773
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
         Column(1)       =   "FrmStockFam.frx":03C0
         Column(2)       =   "FrmStockFam.frx":0488
         FormatStylesCount=   6
         FormatStyle(1)  =   "FrmStockFam.frx":052C
         FormatStyle(2)  =   "FrmStockFam.frx":0664
         FormatStyle(3)  =   "FrmStockFam.frx":0714
         FormatStyle(4)  =   "FrmStockFam.frx":07C8
         FormatStyle(5)  =   "FrmStockFam.frx":08A0
         FormatStyle(6)  =   "FrmStockFam.frx":0958
         ImageCount      =   0
         PrinterProperties=   "FrmStockFam.frx":0A38
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Buscar"
      Height          =   525
      Left            =   9300
      TabIndex        =   3
      Top             =   300
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   10575
      Begin VB.CheckBox ChkStock 
         Caption         =   "Solo Con Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5250
         TabIndex        =   16
         Top             =   480
         Value           =   1  'Checked
         Width           =   2850
      End
      Begin VB.CheckBox ChkStockComprometido 
         Caption         =   "Incluye Stock Comprometido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5250
         TabIndex        =   5
         Top             =   195
         Width           =   2850
      End
      Begin VB.ComboBox CmbAlmacen 
         Height          =   315
         Left            =   1095
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   2355
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Almacen:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   315
         Width           =   660
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   0
      Top             =   720
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmStockFam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Reg As New ADODB.Recordset, sFam As String
Public tipoimp As Integer
Dim strSQL   As String
Sub Buscar()
On Error GoTo Fin
    Screen.MousePointer = 11
    Set Reg = Nothing
    
    Reg.CursorLocation = adUseClient
    
    If frmSelecFamilias.ExisteHilo = "S" Then
        Reg.Open "UP_RepStockFamHilo '" & Right(Me.CmbAlmacen, 2) & "','" & sFam & "'," & frmSelecFamilias.orderby & ",'" & frmSelecFamilias.sCod_Prov & "'", cConnect
        strSQL = "UP_RepStockFamHilo '" & Right(Me.CmbAlmacen, 2) & "','" & sFam & "'," & frmSelecFamilias.orderby & ",'" & frmSelecFamilias.sCod_Prov & "'"
    Else
        If ChkStockComprometido.Value = Checked Then
            tipoimp = 1
        Else
            tipoimp = 0
        End If
        Reg.Open "UP_RepStockFam '" & Right(Me.CmbAlmacen, 2) & "','" & sFam & "'," & tipoimp & "," & IIf(ChkStock.Value, 1, 0), cConnect
        strSQL = "UP_RepStockFam '" & Right(Me.CmbAlmacen, 2) & "','" & sFam & "'," & tipoimp & "," & IIf(ChkStock.Value, 1, 0)
    End If
    
    'Set Grilla.DataSource = Reg
    Set Grilla1.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    
    Grilla1.Columns("pre_ultcomp").Visible = False
    Grilla1.Columns("importe").Visible = False
    
    Screen.MousePointer = 0
Exit Sub
Fin:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical + vbOKOnly, "Mostrar Resultados de Busqueda"
End Sub

Private Sub CmbAlmacen_Click()
    Load frmSelecFamilias
    frmSelecFamilias.vCod_Almacen = Right(Me.CmbAlmacen, 2)
    frmSelecFamilias.carga_lista
    frmSelecFamilias.Show vbModal
    sFam = frmSelecFamilias.sFam
    Unload frmSelecFamilias
End Sub

Private Sub Command1_Click()
If sFam = "" Then
    MsgBox "Seleccione una familia", vbInformation
    Exit Sub
End If

Buscar
End Sub

Private Sub Form_Load()
FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp1, Me.Name)
LlenaCombo CmbAlmacen, "Select Nom_Almacen+space(100)+ Cod_Almacen from lg_almacen where tip_item='I' order by 1", cConnect
'FormateaGrid Grilla1
Buscar
End Sub

Sub GeneraRepoStkCritico()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String
Dim RutaLogo As String
Dim strSQL As String
    
    Ruta = vRuta & "\stocksCriticos.xlt"

    strSQL = "SELECT Ruta_Logo FROM SEGURIDAD..SEG_EMPRESAS WHERE cod_EMPRESA ='" & Trim(vemp1) & "'"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "Repo", DevuelveCampo(strSQL, cConnect), cConnect
    'oo.Run "Repo"
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub

Sub GeneraRepo()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String
    
    
    If frmSelecFamilias.ExisteHilo = "S" Then
        If frmSelecFamilias.orderby = 0 Then
         Ruta = vRuta & "\stockfamHilo.xlt"
         Else
         Ruta = vRuta & "\stockfamHiloProv.xlt"
         End If
    Else
        Ruta = vRuta & "\stockfam.xlt"
    End If

    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    'oo.Run "Repo", CStr(Left(Me.CmbAlmacen, 20)), CStr(Left(Me.CmbFamilia, 20)), vemp1, Reg, cConnect
    oo.Run "Repo", CStr(Left(Me.CmbAlmacen, 20)), sFam, vemp1, tipoimp, Reg, cConnect
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub


Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
    GeneraRepo
Case "VALORIZADO"
    GeneraRepoVal
Case "CRITICO"
    GeneraRepoStkCritico
Case "RESUMEN"
    Resumen
Case "OBSOLETOS"
    DTPFecha.Value = Date
    FraObsoletos.Visible = True
    DTPFecha.SetFocus
Case "UBICACION"
    If Grilla1.RowCount = 0 Then Exit Sub
    If frmSelecFamilias.ExisteHilo <> "S" Then
        TxtUbicacion.Text = Trim(Grilla1.Value(Grilla1.Columns("ubicacion_fisica").Index))
        FraUbicacion.Visible = True
        TxtUbicacion.SetFocus
    End If
Case "SALIR"
    Unload Me
End Select
End Sub

Sub GeneraRepoVal()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String
    
    If frmSelecFamilias.ExisteHilo = "S" Then
        Ruta = vRuta & "\stockfamHiloVal.xlt"
    Else
        Ruta = vRuta & "\stockfamVal.xlt"
    End If

    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    'oo.Run "Repo", CStr(Left(Me.CmbAlmacen, 20)), CStr(Left(Me.CmbFamilia, 20)), vemp1, Reg, cConnect
    oo.Run "Repo", CStr(Left(Me.CmbAlmacen, 20)), sFam, vemp1, tipoimp, Reg, cConnect
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub


Sub Resumen()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String
    
    Ruta = vRuta & "\RptStockFam_Resumido.XLT"
    
    strSQL = "UP_RepStockFam_Resumido '" & Right(Me.CmbAlmacen, 2) & "','" & sFam & "'"
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "Reporte", CStr(Left(Me.CmbAlmacen, 20)), sFam, strSQL, cConnect
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub

Sub Obsoletos()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String
    
    Ruta = vRuta & "\RptItemsCandidatosaObsoletos.XLT"
    
    strSQL = "UP_RepItemCandidatos_Obsoletos '" & Right(Me.CmbAlmacen, 2) & "','" & sFam & "','" & DTPFecha.Value & "'"
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "Reporte", CStr(Left(Me.CmbAlmacen, 20)), sFam, Format(DTPFecha.Value, "DD/MM/YYYY"), strSQL, cConnect
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
    Call Obsoletos
    FraObsoletos.Visible = False
Case "CANCELAR"
    FraObsoletos.Visible = False
End Select
End Sub

Private Sub FunctButt3_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    Call Graba_Ubicacion
Case "CANCELAR"
    FraUbicacion.Visible = False
End Select
End Sub

Sub Graba_Ubicacion()
On Error GoTo errUbicacion

strSQL = "LG_ACTUALIZA_UBICACION_FISICA '" & Right(Me.CmbAlmacen, 2) & "','" & Grilla1.Value(Grilla1.Columns("Cod_Item").Index) & "','" & Grilla1.Value(Grilla1.Columns("Cod_comb").Index) & "','" & _
            Grilla1.Value(Grilla1.Columns("Cod_color").Index) & "','" & Grilla1.Value(Grilla1.Columns("Cod_talla").Index) & "','" & Grilla1.Value(Grilla1.Columns("Cod_destino").Index) & "','" & _
            Grilla1.Value(Grilla1.Columns("cod_estcli").Index) & "','" & Trim(TxtUbicacion.Text) & "'"
            
ExecuteSQL cConnect, strSQL
FraUbicacion.Visible = False
Buscar

Exit Sub
errUbicacion:
    MsgBox err.Description, vbCritical, "Ubicacion"
End Sub

Private Sub TxtUbicacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub
