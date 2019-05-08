VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form FrmBitacoraItems 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bitacora Item"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   9240
      TabIndex        =   2
      Top             =   4080
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   900
      Custom          =   "0~0~SALIR~True~True~&Salir~0~0~1~~0~False~False~&Salir~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin GridEX20.GridEX GridEX1 
         Height          =   3570
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10140
         _ExtentX        =   17886
         _ExtentY        =   6297
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
         Column(1)       =   "FrmBitacoraItems.frx":0000
         Column(2)       =   "FrmBitacoraItems.frx":00C8
         FormatStylesCount=   8
         FormatStyle(1)  =   "FrmBitacoraItems.frx":016C
         FormatStyle(2)  =   "FrmBitacoraItems.frx":02A4
         FormatStyle(3)  =   "FrmBitacoraItems.frx":0354
         FormatStyle(4)  =   "FrmBitacoraItems.frx":0408
         FormatStyle(5)  =   "FrmBitacoraItems.frx":04E0
         FormatStyle(6)  =   "FrmBitacoraItems.frx":0598
         FormatStyle(7)  =   "FrmBitacoraItems.frx":0678
         FormatStyle(8)  =   "FrmBitacoraItems.frx":0724
         ImageCount      =   0
         PrinterProperties=   "FrmBitacoraItems.frx":07D4
      End
   End
End
Attribute VB_Name = "FrmBitacoraItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public Cod_Item As String

Sub CARGA_GRID()
strSQL = "exec SM_MUESTRA_BITACORA_ITEMS '" & Me.Cod_Item & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

GridEX1.Columns("Cod_HilTel").Visible = False
GridEX1.Columns("Fac_Conversion").Visible = False
GridEX1.Columns("Por_Mertin").Visible = False
GridEX1.Columns("Cod_TipCar").Visible = False
GridEX1.Columns("Dir_Icono").Visible = False
GridEX1.Columns("Sec.").Visible = False

GridEX1.Columns("Flg_Status").Width = 600
GridEX1.Columns("Fec. Modificacion").Width = 1800
GridEX1.Columns("Modificado por").Width = 1100
GridEX1.Columns("Des_Item").Width = 3000
GridEX1.Columns("Cod_GruItem").Width = 600
GridEX1.Columns("Cod_UniMed").Width = 600
GridEX1.Columns("Cod_ClaItem").Width = 600
GridEX1.Columns("Can_PtoReor").Width = 900
GridEX1.Columns("Can_LotPed").Width = 900
GridEX1.Columns("Rep_PreDol").Width = 900
GridEX1.Columns("Cod_Origen").Width = 550
GridEX1.Columns("Ide_Talla").Width = 500
GridEX1.Columns("Ide_Color").Width = 500
GridEX1.Columns("Ide_EsCli").Width = 580
GridEX1.Columns("Ide_Destino").Width = 600
GridEX1.Columns("Ide_Po").Width = 450
GridEX1.Columns("Cod_MotPrePro").Width = 800
GridEX1.Columns("Comentario").Width = 900

GridEX1.Columns("Cod_GruItem").Caption = "GruItem"
GridEX1.Columns("Cod_ClaItem").Caption = "ClaItem"
GridEX1.Columns("Cod_MotPrePro").Caption = "Mot.PreProd"
GridEX1.Columns("Fec. Modificacion").Caption = "Fec. Cambio"
GridEX1.Columns("Ide_Talla").Caption = "Talla"
GridEX1.Columns("Ide_Color").Caption = "Color"
GridEX1.Columns("Ide_EsCli").Caption = "Est.Cli."
GridEX1.Columns("Ide_Destino").Caption = "Destino"
GridEX1.Columns("Dir_Icono").Caption = "Icono"
GridEX1.Columns("Ide_Po").Caption = "PO"
GridEX1.Columns("Cod_UniMed").Caption = "UniMed."
GridEX1.Columns("Cod_Origen").Caption = "Origen"
GridEX1.Columns("Flg_Status").Caption = "Status"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Unload Me
End Sub
