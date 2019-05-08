VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBusqGeneral3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3900
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8145
      Begin GridEX20.GridEX gexLista 
         Height          =   3480
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   7950
         _ExtentX        =   14023
         _ExtentY        =   6138
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
         Column(1)       =   "frmBusqGeneral3.frx":0000
         Column(2)       =   "frmBusqGeneral3.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmBusqGeneral3.frx":016C
         FormatStyle(2)  =   "frmBusqGeneral3.frx":02A4
         FormatStyle(3)  =   "frmBusqGeneral3.frx":0354
         FormatStyle(4)  =   "frmBusqGeneral3.frx":0408
         FormatStyle(5)  =   "frmBusqGeneral3.frx":04E0
         FormatStyle(6)  =   "frmBusqGeneral3.frx":0598
         ImageCount      =   0
         PrinterProperties=   "frmBusqGeneral3.frx":0678
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2835
      TabIndex        =   1
      Top             =   3990
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
Attribute VB_Name = "frmBusqGeneral3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sQuery As String, bCancel As Boolean
Public oParent As Object

Public Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    If ActionName = "CANCELAR" Then bCancel = True
    Me.Hide
End Sub

Public Sub CARGAR_DATOS()
    bCancel = False
    Set gexLista.ADORecordset = CargarRecordSetDesconectado(sQuery, cConnect)
    If gexLista.RowCount > 0 Then
        gexLista.Row = 1
        gexLista.ADORecordset.AbsolutePosition = gexLista.Row
    End If
End Sub

Private Sub gexLista_DblClick()
    FunctButt1_ActionClick 0, 0, "ACEPTAR"
End Sub

Private Sub gexLista_KeyDown(KeyCode As Integer, Shift As Integer)
'Al presionar Enter la fila baja y no se selecciona el regsitro deseado
    Select Case KeyCode
    Case vbKeyReturn
        FunctButt1_ActionClick 0, 0, "ACEPTAR"
    Case vbKeyEscape
        FunctButt1_ActionClick 0, 0, "CANCELAR"
    End Select
    
End Sub
