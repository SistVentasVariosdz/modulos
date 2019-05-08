VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBuscaMovGuia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Movimiento por Nro de Guia"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin GridEX20.GridEX gexMovs 
      Height          =   3180
      Left            =   90
      TabIndex        =   3
      Top             =   1095
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   5609
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmBuscaMovGuia.frx":0000
      Column(2)       =   "frmBuscaMovGuia.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmBuscaMovGuia.frx":016C
      FormatStyle(2)  =   "frmBuscaMovGuia.frx":02A4
      FormatStyle(3)  =   "frmBuscaMovGuia.frx":0354
      FormatStyle(4)  =   "frmBuscaMovGuia.frx":0408
      FormatStyle(5)  =   "frmBuscaMovGuia.frx":04E0
      FormatStyle(6)  =   "frmBuscaMovGuia.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmBuscaMovGuia.frx":0678
   End
   Begin VB.Frame Frame1 
      Height          =   945
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   7920
      Begin VB.TextBox txtNro_Guia 
         Height          =   285
         Left            =   2040
         TabIndex        =   0
         Top             =   345
         Width           =   2730
      End
      Begin FunctionsButtons.FunctButt fnbBuscar 
         Height          =   495
         Left            =   6495
         TabIndex        =   2
         Top             =   225
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
         Caption         =   "Nro de Guia Contiene"
         Height          =   225
         Left            =   270
         TabIndex        =   5
         Top             =   375
         Width           =   1620
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   555
      Left            =   6750
      TabIndex        =   4
      Top             =   4470
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      Custom          =   "0~0~SALIR~True~True~&Salir~0~0~1~~0~False~False~&Salir~"
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmBuscaMovGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Almacen As String
Dim strSQL As String

Public Sub fnbBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo Fin
Dim sTit As String
    sTit = "Buscar Movimiento por Guía"
    
    txtNro_Guia = Trim(txtNro_Guia)
    If txtNro_Guia = "" Then
        MsgBox "Se debe especificar una parte de la Guía", vbExclamation + vbOKOnly, sTit
        txtNro_Guia.SetFocus
        Exit Sub
    End If
    
    strSQL = "EXEC SM_BUSCA_GUIA_EN_ALMACEN '" & sCod_Almacen & "', '" & txtNro_Guia & "'"
    
    Set gexMovs.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    
    gexMovs.Columns("Num_MovStk").Width = 1110
    gexMovs.Columns("Fec_MovStk").Width = 1050
    gexMovs.Columns("Num_Guia").Width = 1500
    gexMovs.Columns("Nro_Guia_Propia").Width = 1350
    gexMovs.Columns("Des_Proveedor").Width = 2385
    
    gexMovs.Columns("Num_MovStk").Caption = "Nro.Mov."
    gexMovs.Columns("Fec_MovStk").Caption = "Fecha"
    gexMovs.Columns("Num_Guia").Caption = "Nro.Guia"
    gexMovs.Columns("Nro_Guia_Propia").Caption = "Nro.Guia Propia"
    gexMovs.Columns("Des_Proveedor").Caption = "Proveedor"
    
Exit Sub
Fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, sTit
End Sub

Private Sub fnbBuscar_GotFocus()
    fnbBuscar_ActionClick 0, 0, ""
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Unload Me
End Sub

Private Sub txtNro_Guia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub
