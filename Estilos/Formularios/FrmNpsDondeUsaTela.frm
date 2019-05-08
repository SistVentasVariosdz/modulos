VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form FrmNpsDondeUsaTela 
   Caption         =   "Nps donde se usa"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2400
      TabIndex        =   6
      Top             =   6720
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmNpsDondeUsaTela.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7215
      Begin VB.TextBox TxtDes_Tela 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   200
         Width           =   4335
      End
      Begin VB.TextBox TxtCod_Tela 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   200
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tela"
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
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   390
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5985
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   7185
      Begin GridEX20.GridEX GridEX1 
         Height          =   5640
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   9948
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
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "FrmNpsDondeUsaTela.frx":0090
         Column(2)       =   "FrmNpsDondeUsaTela.frx":0158
         FormatStylesCount=   8
         FormatStyle(1)  =   "FrmNpsDondeUsaTela.frx":01FC
         FormatStyle(2)  =   "FrmNpsDondeUsaTela.frx":0334
         FormatStyle(3)  =   "FrmNpsDondeUsaTela.frx":03E4
         FormatStyle(4)  =   "FrmNpsDondeUsaTela.frx":0498
         FormatStyle(5)  =   "FrmNpsDondeUsaTela.frx":0570
         FormatStyle(6)  =   "FrmNpsDondeUsaTela.frx":0628
         FormatStyle(7)  =   "FrmNpsDondeUsaTela.frx":0708
         FormatStyle(8)  =   "FrmNpsDondeUsaTela.frx":07B4
         ImageCount      =   0
         PrinterProperties=   "FrmNpsDondeUsaTela.frx":0864
      End
   End
End
Attribute VB_Name = "FrmNpsDondeUsaTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
    Call Reporte
Case "SALIR"
    Unload Me
 End Select
End Sub

Sub CARGA_GRID()
strSQL = "es_muestra_Nps_Asociadas_Tela '" & Trim(TxtCod_Tela.Text) & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

GridEX1.Columns("Fabrica").Width = 700
GridEX1.Columns("cod_OrdPro").Width = 800
GridEX1.Columns("Descripcion").Width = 2800
GridEX1.Columns("Cliente").Width = 2300

GridEX1.Columns("cod_OrdPro").Caption = "NP"

End Sub

Sub Reporte()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String

    strSQL = "es_muestra_Nps_Asociadas_Tela '" & Trim(TxtCod_Tela.Text) & "'"
    Ruta = vRuta & "\RptNpsAsociadasTela.xlt"
    
    Set oo = CreateObject("excel.application")
    oo.workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.run "Reporte", Trim(TxtCod_Tela.Text) & "-" & Trim(TxtDes_Tela.Text), strSQL, cCONNECT
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler Err, "Reporte"
    Set oo = Nothing
End Sub
