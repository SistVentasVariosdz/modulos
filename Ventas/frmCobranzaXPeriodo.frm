VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmCobranzaXPeriodo 
   Caption         =   "Cobranzas por Periodo"
   ClientHeight    =   6075
   ClientLeft      =   180
   ClientTop       =   495
   ClientWidth     =   7365
   Icon            =   "frmCobranzaXPeriodo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraBuscar 
      Caption         =   "Argumentos de Registro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   7305
      Begin VB.TextBox txtAno 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   2
         Top             =   360
         Width           =   660
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   615
         Left            =   5835
         TabIndex        =   4
         Top             =   220
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~~~0~Verdadero~Falso~&Buscar~"
         Orientacion     =   1
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1000
         ControlHeigth   =   500
         ControlSeparator=   40
      End
      Begin VB.Label Label6 
         Caption         =   "Año"
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4155
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   7329
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
      RowHeaders      =   -1  'True
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmCobranzaXPeriodo.frx":030A
      Column(2)       =   "frmCobranzaXPeriodo.frx":03D2
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmCobranzaXPeriodo.frx":0476
      FormatStyle(2)  =   "frmCobranzaXPeriodo.frx":05AE
      FormatStyle(3)  =   "frmCobranzaXPeriodo.frx":065E
      FormatStyle(4)  =   "frmCobranzaXPeriodo.frx":0712
      FormatStyle(5)  =   "frmCobranzaXPeriodo.frx":07EA
      FormatStyle(6)  =   "frmCobranzaXPeriodo.frx":08A2
      FormatStyle(7)  =   "frmCobranzaXPeriodo.frx":0982
      FormatStyle(8)  =   "frmCobranzaXPeriodo.frx":0A2E
      ImageCount      =   0
      PrinterProperties=   "frmCobranzaXPeriodo.frx":0ADE
   End
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   525
      Left            =   2122
      TabIndex        =   5
      Top             =   5520
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   926
      Custom          =   $"frmCobranzaXPeriodo.frx":0CB6
      Orientacion     =   0
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1000
      ControlHeigth   =   500
      ControlSeparator=   40
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   5760
      Top             =   6000
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmCobranzaXPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iRowAnterior As Long
Dim iColAnterior As Long
Dim bClickColSelec As Boolean
Dim bCargaGRid As Boolean
Dim bPuedeAutorizar  As Boolean
Dim sTipoDocAutorizar As String
Dim strOpcion As String
Public CODIGO As String, Descripcion As String, strCod_Anxo As String, TipoAdd As String
Dim lvSw As Boolean

Private Sub cmdBuscar_Click()
If txtAno.Text = "" Then
    MsgBox "Debe ingresar un año"
    Exit Sub
 Else
  Buscar
 End If
End Sub

Sub Buscar()

On Error GoTo dprDepurar

Dim ssql As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle

ssql = "CF_VENTAS_COBRANZAS_PERIODO '" & Trim(txtAno.Text) & "'"

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(ssql, cCONNECT)

GridEX1.Columns("ano").Width = 0
GridEX1.Columns("ano").Caption = "ano"

GridEX1.Columns("mes").Width = 1000
GridEX1.Columns("mes").Caption = "Mes"

GridEX1.Columns("Imp_Dol_Parte_Cobranzas").Width = 2000
GridEX1.Columns("Imp_Dol_Parte_Cobranzas").Caption = "Imp. Dol. Parte Cobranzas"

GridEX1.Columns("Imp_Dol_Letras_Abonadas").Width = 2000
GridEX1.Columns("Imp_Dol_Letras_Abonadas").Caption = "Imp. Dol. Letras Abonadas"

GridEX1.Columns("Imp_Dol_Cobranza_Exterior").Width = 2000
GridEX1.Columns("Imp_Dol_Cobranza_Exterior").Caption = "Imp. Dol. Cobranza Exterior"

GridEX1.ContinuousScroll = True

Exit Sub

dprDepurar:

errores Err.Number
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub



Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo HandlerError

    Dim Msg As Variant
    
    Select Case ActionName
    
    Case "BUSCAR"
      If txtAno.Text = "" Then
    MsgBox "Debe ingresar un año"
    Exit Sub
 Else
  Buscar
 End If
   
    End Select
Exit Sub
Resume
HandlerError:
  errores Err.Number
End Sub



Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "AGREGAR"
    If txtAno.Text = "" Then
    MsgBox "Debe ingresar un año"
    Exit Sub
 Else
    FrmDetalleCobrabzaXPerido.sAno = txtAno.Text
    FrmDetalleCobrabzaXPerido.sAccion = "I"
    FrmDetalleCobrabzaXPerido.Show 1
    Set FrmDetalleCobrabzaXPerido = Nothing
    Buscar
    End If
Case "MODIFICAR"
If txtAno.Text = "" Then
    MsgBox "Debe ingresar un año"
    Exit Sub
 Else
    If GridEX1.RowCount = 0 Then Exit Sub
    FrmDetalleCobrabzaXPerido.sAccion = "U"
    FrmDetalleCobrabzaXPerido.txtMes = GridEX1.Value(GridEX1.Columns("mes").Index)
    FrmDetalleCobrabzaXPerido.txtcobranza = GridEX1.Value(GridEX1.Columns("Imp_Dol_Parte_Cobranzas").Index)
    FrmDetalleCobrabzaXPerido.txtabonadas = GridEX1.Value(GridEX1.Columns("Imp_Dol_Letras_Abonadas").Index)
    FrmDetalleCobrabzaXPerido.txtexterior = GridEX1.Value(GridEX1.Columns("Imp_Dol_Cobranza_Exterior").Index)
    FrmDetalleCobrabzaXPerido.sAno = GridEX1.Value(GridEX1.Columns("ano").Index)
    FrmDetalleCobrabzaXPerido.txtMes.Enabled = False
    FrmDetalleCobrabzaXPerido.Show 1
    Buscar
    Set FrmDetalleCobrabzaXPerido = Nothing
   End If
Case "SALIR"
    Unload Me
   End Select
End Sub

Private Sub txtano_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
  If KeyAscii >= 48 And KeyAscii <= 57 _
    Or KeyAscii = 8 Or KeyAscii = 13 Then Else KeyAscii = 0
End Sub



