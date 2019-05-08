VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form FrmEnviosPostTenido 
   Caption         =   "Envios a Post Teñido"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   9240
      TabIndex        =   4
      Top             =   5400
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   900
      Custom          =   "0~0~SALIR~Verdadero~Verdadero~&Salir~0~0~1~~0~Falso~Falso~&Salir~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Orden Compra"
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label LblCod_Ordcomp 
         BackColor       =   &H8000000E&
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
         Height          =   270
         Left            =   2760
         TabIndex        =   2
         Top             =   285
         Width           =   1665
      End
      Begin VB.Label LblSer_OrdComp 
         BackColor       =   &H8000000E&
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
         Height          =   270
         Left            =   1920
         TabIndex        =   1
         Top             =   285
         Width           =   705
      End
   End
   Begin GridEX20.GridEX gexLista 
      Height          =   4455
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   7858
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "FrmEnviosPostTenido.frx":0000
      Column(2)       =   "FrmEnviosPostTenido.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "FrmEnviosPostTenido.frx":016C
      FormatStyle(2)  =   "FrmEnviosPostTenido.frx":02A4
      FormatStyle(3)  =   "FrmEnviosPostTenido.frx":0354
      FormatStyle(4)  =   "FrmEnviosPostTenido.frx":0408
      FormatStyle(5)  =   "FrmEnviosPostTenido.frx":04E0
      FormatStyle(6)  =   "FrmEnviosPostTenido.frx":0598
      ImageCount      =   0
      PrinterProperties=   "FrmEnviosPostTenido.frx":0678
   End
End
Attribute VB_Name = "FrmEnviosPostTenido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Strsql As String
Public sSer_OrdComp As String, sCod_OrdComp As String

Public Sub CARGA_GRID()

    Strsql = "EXEC lg_muestra_envios_post_tenido_por_oc '" & sSer_OrdComp & "','" & sCod_OrdComp & "'"
    Set gexLista.ADORecordset = CargarRecordSetDesconectado(Strsql, cConnect)
    'SetGeneralGridEX gexLista, 0, 1
    Call CONFIGURAR_GRID
End Sub

Public Sub CONFIGURAR_GRID()
    gexLista.Columns("Partida").Width = 650
    gexLista.Columns("Color").Width = 1800
    gexLista.Columns("Alm_Mov").Width = 900
    gexLista.Columns("Fec_MovStk").Width = 1000
    gexLista.Columns("Numero_Guia").Width = 900
    gexLista.Columns("Kgs_Enviados").Width = 1000
    gexLista.Columns("Kgs_Devueltos").Width = 1000
    gexLista.Columns("Tela").Width = 2200
    gexLista.Columns("Comb").Width = 1400
    gexLista.Columns("Talla").Width = 850
    
    gexLista.Columns("Fec_MovStk").Caption = "Fec. Mov."
    gexLista.Columns("Numero_Guia").Caption = "Num. Guia"
    gexLista.Columns("Kgs_Enviados").Caption = "Kgs.Enviados"
    gexLista.Columns("Kgs_Devueltos").Caption = "Kgs.Devueltos"
    
    gexLista.FrozenColumns = 5
    
End Sub

Public Function SetGeneralGridEX(ByRef GridEx As GridEX20.GridEx, ByVal iFixsCols As Integer, ByVal iTipoColorBack As Integer)

    If iFixsCols > 0 Then
        GridEx.FrozenColumns = iFixsCols
    End If
    
    If iTipoColorBack = 1 Then
        GridEx.BackColor = &H80000018
        GridEx.BackColorBkg = &H80000018
        GridEx.GridLines = jgexGLVertical
        GridEx.GridLineStyle = jgexGLSSmallDots
    Else
        GridEx.BackColor = &H80000005
        GridEx.BackColorBkg = &H80000005
        GridEx.GridLines = jgexGLBoth
        GridEx.GridLineStyle = jgexGLSSmallDots
    End If
End Function

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "SALIR"
    Unload Me
End Select
End Sub
