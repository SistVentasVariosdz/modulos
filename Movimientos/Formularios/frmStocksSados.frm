VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmStocksSados 
   Caption         =   "Ver Saldo de Stocks"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3555
      TabIndex        =   3
      Top             =   6270
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
      Height          =   1110
      Left            =   90
      TabIndex        =   4
      Top             =   60
      Width           =   9300
      Begin VB.ComboBox cboAlmacen 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   435
         Width           =   1605
      End
      Begin FunctionsButtons.FunctButt fnbBuscar 
         Height          =   495
         Left            =   7635
         TabIndex        =   1
         Top             =   315
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
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   5490
         TabIndex        =   6
         Top             =   405
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   61931523
         CurrentDate     =   37809
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   3420
         TabIndex        =   8
         Top             =   420
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   61931523
         CurrentDate     =   37809
      End
      Begin VB.Label Label2 
         Caption         =   "Desde:"
         Height          =   240
         Left            =   2790
         TabIndex        =   9
         Top             =   465
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta:"
         Height          =   240
         Left            =   4890
         TabIndex        =   7
         Top             =   450
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "Alamcén:"
         Height          =   255
         Left            =   210
         TabIndex        =   5
         Top             =   480
         Width           =   750
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4680
      Left            =   90
      TabIndex        =   2
      Top             =   1275
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   8255
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmStocksSados.frx":0000
      Column(2)       =   "frmStocksSados.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmStocksSados.frx":016C
      FormatStyle(2)  =   "frmStocksSados.frx":02A4
      FormatStyle(3)  =   "frmStocksSados.frx":0354
      FormatStyle(4)  =   "frmStocksSados.frx":0408
      FormatStyle(5)  =   "frmStocksSados.frx":04E0
      FormatStyle(6)  =   "frmStocksSados.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmStocksSados.frx":0678
   End
End
Attribute VB_Name = "frmStocksSados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String

Private Sub cboAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub fnbBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    If cboAlmacen.ListIndex = -1 Then
        MsgBox "Se debe elegir un Almacen", vbOKOnly + vbExclamation, "Ver Saldo de Stocks"
    End If
    strSQL = "EXEC SM_MUESTRA_CF_STOCKS_SALDOS '" & Left(cboAlmacen, 2) & "'"
    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    
    If GridEX1.RowCount > 0 Then: GridEX1.SetFocus
    Else: cboAlmacen.SetFocus
    End If
    
End Sub

Private Sub FillAlmacen()
Dim rstAlm As ADODB.Recordset

    strSQL = "SELECT Cod_Almacen, Nom_Almacen FROM CF_Almacen"
    Set rstAlm = CargarRecordSetDesconectado(strSQL, cConnect)
    rstAlm.MoveFirst
    cboAlmacen.Clear
    Do Until rstAlm.EOF
        cboAlmacen.AddItem rstAlm!Cod_Almacen & " " & rstAlm!Nom_Almacen
        rstAlm.MoveNext
    Loop
    rstAlm.Close
    Set rstAlm = Nothing
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "SALIR"
        Unload Me
    End Select
End Sub
