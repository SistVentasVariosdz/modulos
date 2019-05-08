VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmShowAlternativasPesoAncho 
   Caption         =   "Alternativas Peso / Ancho "
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   2310
      Left            =   9240
      TabIndex        =   1
      Top             =   0
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   4075
      Custom          =   $"FrmShowAlternativasPesoAncho.frx":0000
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4890
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   8625
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
      Column(1)       =   "FrmShowAlternativasPesoAncho.frx":012C
      Column(2)       =   "FrmShowAlternativasPesoAncho.frx":01F4
      FormatStylesCount=   8
      FormatStyle(1)  =   "FrmShowAlternativasPesoAncho.frx":0298
      FormatStyle(2)  =   "FrmShowAlternativasPesoAncho.frx":03D0
      FormatStyle(3)  =   "FrmShowAlternativasPesoAncho.frx":0480
      FormatStyle(4)  =   "FrmShowAlternativasPesoAncho.frx":0534
      FormatStyle(5)  =   "FrmShowAlternativasPesoAncho.frx":060C
      FormatStyle(6)  =   "FrmShowAlternativasPesoAncho.frx":06C4
      FormatStyle(7)  =   "FrmShowAlternativasPesoAncho.frx":07A4
      FormatStyle(8)  =   "FrmShowAlternativasPesoAncho.frx":0850
      ImageCount      =   0
      PrinterProperties=   "FrmShowAlternativasPesoAncho.frx":0900
   End
End
Attribute VB_Name = "FrmShowAlternativasPesoAncho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vCod_Tela As String
Dim strSQL As String

Sub CARGA_GRID()
On Error GoTo errCarga_Grid

strSQL = "EXEC tx_muestra_tx_tela_alternativas '" & vCod_Tela & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

GridEX1.Columns("Alternativa").Width = 1000
GridEX1.Columns("Descripcion").Width = 3200
GridEX1.Columns("Gramaje_Acab").Width = 1000
GridEX1.Columns("Ancho_Acab").Width = 1000

Exit Sub
errCarga_Grid:
    ErrorHandler Err, "Carga Grid"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ADICIONAR"
    Load FrmAlternativasPesoAncho
    Set FrmAlternativasPesoAncho.oParent = Me
    FrmAlternativasPesoAncho.vAccion = "I"
    FrmAlternativasPesoAncho.vCod_Tela = Me.vCod_Tela
    FrmAlternativasPesoAncho.TxtCod_Tela.Text = Me.vCod_Tela
    FrmAlternativasPesoAncho.TxtDes_Tela.Text = DevuelveCampo("select des_tela from tx_tela where cod_tela ='" & Me.vCod_Tela & "'", cCONNECT)
    FrmAlternativasPesoAncho.Show vbModal
    Set FrmAlternativasPesoAncho = Nothing
Case "MODIFICAR"
    Load FrmAlternativasPesoAncho
    Set FrmAlternativasPesoAncho.oParent = Me
    FrmAlternativasPesoAncho.vAccion = "U"
    FrmAlternativasPesoAncho.vCod_Tela = Me.vCod_Tela
    FrmAlternativasPesoAncho.TxtCod_Tela.Text = Me.vCod_Tela
    FrmAlternativasPesoAncho.TxtDes_Tela.Text = DevuelveCampo("select des_tela from tx_tela where cod_tela ='" & Me.vCod_Tela & "'", cCONNECT)
    FrmAlternativasPesoAncho.TxtAlternativa.Text = GridEX1.Value(GridEX1.Columns("alternativa").Index)
    FrmAlternativasPesoAncho.TxtDescripcion.Text = GridEX1.Value(GridEX1.Columns("Descripcion").Index)
    FrmAlternativasPesoAncho.TxtGramaje.Text = GridEX1.Value(GridEX1.Columns("gramaje_acab").Index)
    FrmAlternativasPesoAncho.TxtAncho.Text = GridEX1.Value(GridEX1.Columns("ancho_acab").Index)
    FrmAlternativasPesoAncho.Show vbModal
    Set FrmAlternativasPesoAncho = Nothing
Case "ELIMINAR"
    Load FrmAlternativasPesoAncho
    Set FrmAlternativasPesoAncho.oParent = Me
    FrmAlternativasPesoAncho.vAccion = "D"
    FrmAlternativasPesoAncho.vCod_Tela = Me.vCod_Tela
    FrmAlternativasPesoAncho.TxtCod_Tela.Text = Me.vCod_Tela
    FrmAlternativasPesoAncho.TxtDes_Tela.Text = DevuelveCampo("select des_tela from tx_tela where cod_tela ='" & Me.vCod_Tela & "'", cCONNECT)
    FrmAlternativasPesoAncho.TxtAlternativa.Text = GridEX1.Value(GridEX1.Columns("alternativa").Index)
    FrmAlternativasPesoAncho.TxtDescripcion.Text = GridEX1.Value(GridEX1.Columns("Descripcion").Index)
    FrmAlternativasPesoAncho.TxtGramaje.Text = GridEX1.Value(GridEX1.Columns("gramaje_acab").Index)
    FrmAlternativasPesoAncho.TxtAncho.Text = GridEX1.Value(GridEX1.Columns("ancho_acab").Index)
    FrmAlternativasPesoAncho.FraDatos.Enabled = False
    FrmAlternativasPesoAncho.Show vbModal
    Set FrmAlternativasPesoAncho = Nothing
Case "SALIR"
    Unload Me
End Select
End Sub

