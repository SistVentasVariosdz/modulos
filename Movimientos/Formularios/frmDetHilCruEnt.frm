VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmDetHilCruEnt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   690
      Left            =   75
      TabIndex        =   4
      Top             =   0
      Width           =   8145
      Begin VB.Label lblOT 
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
         Height          =   285
         Left            =   3375
         TabIndex        =   6
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label6 
         Caption         =   "OT"
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
         Left            =   2925
         TabIndex        =   5
         Top             =   285
         Width           =   285
      End
   End
   Begin VB.Frame Frame2 
      Height          =   810
      Left            =   75
      TabIndex        =   2
      Top             =   4680
      Width           =   8145
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   2820
         TabIndex        =   3
         Top             =   195
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   "0~0~ACEPTAR~True~True~&Aceptar~0~0~1~~0~False~False~&Aceptar~~1~0~CANCELAR~True~True~&Cancelar~0~0~2~~0~False~False~&Cancelar~"
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3990
      Left            =   75
      TabIndex        =   1
      Top             =   675
      Width           =   8145
      Begin GridEX20.GridEX gexLotes 
         Height          =   3585
         Left            =   60
         TabIndex        =   0
         Top             =   210
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   6324
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         DataMode        =   1
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "frmDetHilCruEnt.frx":0000
         Column(2)       =   "frmDetHilCruEnt.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmDetHilCruEnt.frx":016C
         FormatStyle(2)  =   "frmDetHilCruEnt.frx":02A4
         FormatStyle(3)  =   "frmDetHilCruEnt.frx":0354
         FormatStyle(4)  =   "frmDetHilCruEnt.frx":0408
         FormatStyle(5)  =   "frmDetHilCruEnt.frx":04E0
         FormatStyle(6)  =   "frmDetHilCruEnt.frx":0598
         ImageCount      =   0
         PrinterProperties=   "frmDetHilCruEnt.frx":0678
      End
   End
End
Attribute VB_Name = "frmDetHilCruEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vCod_TipOrdTra As String, vCod_OrdTra As String, bCancel As Boolean, vCod_Almacen As String

Public Sub SM_AYUDA_DEVOLUCION_TELA_CRUDA_DE_PARTIDAS()
Dim Strsql As String
    bCancel = False
    lblOT = vCod_TipOrdTra & " - " & vCod_OrdTra
    Strsql = "EXEC TH_SM_AYUDA_DEVOLUCION_TELA_CRUDA_DE_PARTIDAS '" & vCod_OrdTra & "'"
    Set gexLotes.ADORecordset = CargarRecordSetDesconectado(Strsql, cConnect)
    With gexLotes
        .Columns("COD_ORDPROV").Caption = "Lote"
        .Columns("COD_PROVEEDOR").Caption = "Cod.Proveedor"
        .Columns("DES_PROVEEDOR").Caption = "Proveedor"
        .Columns("COD_HILTEL").Caption = "Cod.Hilado"
        .Columns("DES_HILTEL").Caption = "Hilado"
        .Columns("KGS_ENVIADOS").Caption = "Kgs.Enviados"
        '.Columns("UNI_ENVIADOS").Caption = "Uni.Enviadas"
        .Columns("NUM_SECUENCIA").Caption = "Sec."
        
        .Columns("COD_ORDPROV").Width = 1000
        .Columns("COD_PROVEEDOR").Width = 2000
        .Columns("DES_PROVEEDOR").Width = 2000
        .Columns("COD_HILTEL").Width = 1000
        .Columns("DES_HILTEL").Width = 2000
        .Columns("KGS_ENVIADOS").Width = 1000
        '.Columns("UNI_ENVIADOS").Width = 1000
        .Columns("NUM_SECUENCIA").Width = 500
    End With
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim Strsql As String
    Select Case ActionName
    Case "ACEPTAR"
        Me.Hide
    Case "CANCELAR"
        Me.Hide
        bCancel = True
    End Select
End Sub

Private Sub gexLotes_DblClick()
    FunctButt1_ActionClick 0, 0, "ACEPTAR"
End Sub

Private Sub gexLotes_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        FunctButt1_ActionClick 0, 0, "ACEPTAR"
    Case vbKeyEscape
        FunctButt1_ActionClick 0, 0, "CANCELAR"
    End Select
End Sub
