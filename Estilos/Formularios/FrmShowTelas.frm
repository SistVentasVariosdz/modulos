VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmShowTelas 
   Caption         =   "Rutas"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   10650
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   4110
      Left            =   9240
      TabIndex        =   1
      Top             =   15
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   7250
      Custom          =   $"FrmShowTelas.frx":0000
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1300
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
      Column(1)       =   "FrmShowTelas.frx":0241
      Column(2)       =   "FrmShowTelas.frx":0309
      FormatStylesCount=   8
      FormatStyle(1)  =   "FrmShowTelas.frx":03AD
      FormatStyle(2)  =   "FrmShowTelas.frx":04E5
      FormatStyle(3)  =   "FrmShowTelas.frx":0595
      FormatStyle(4)  =   "FrmShowTelas.frx":0649
      FormatStyle(5)  =   "FrmShowTelas.frx":0721
      FormatStyle(6)  =   "FrmShowTelas.frx":07D9
      FormatStyle(7)  =   "FrmShowTelas.frx":08B9
      FormatStyle(8)  =   "FrmShowTelas.frx":0965
      ImageCount      =   0
      PrinterProperties=   "FrmShowTelas.frx":0A15
   End
End
Attribute VB_Name = "FrmShowTelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vCod_Tela As String
Public vDes_Tela As String
Public vFamite As String
Public vCOD_ORDTRA As String
Dim strSQL As String

Sub CARGA_GRID()
On Error GoTo errCarga_Grid

strSQL = "EXEC  tx_sm_muestra_Tela_DatTecnicos_cabecera '" & vCod_Tela & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

GridEX1.Columns("Cod_Ruta").Width = 1500
GridEX1.Columns("Descripcion").Width = 3200
GridEX1.Columns("Fec_Creacion").Width = 1300
GridEX1.Columns("Status").Width = 1000

GridEX1.Columns("Cod_Ruta").Caption = "Ruta"
GridEX1.Columns("Descripcion").Caption = "Descripción"
GridEX1.Columns("Fec_Creacion").Caption = "Fecha Creación"
GridEX1.Columns("Status").Caption = "Status"

Exit Sub
errCarga_Grid:
    ErrorHandler Err, "Carga Grid"
End Sub

Private Sub Form_Load()
Dim sSeguridad  As String
sSeguridad = get_botones1(Me, vper, vemp, Me.Name)

FunctButt1.FunctionsUser = sSeguridad
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ADICIONAR"
    Load FrmDetalleTelas
    Set FrmDetalleTelas.oParent = Me
    FrmDetalleTelas.vAccion = "I"
    FrmDetalleTelas.vCod_Tela = Me.vCod_Tela
    FrmDetalleTelas.vRuta = GridEX1.Value(GridEX1.Columns("Cod_Ruta").Index)
   FrmDetalleTelas.Show vbModal
    Set FrmAlternativasPesoAncho = Nothing

Case "MODIFICAR"
    Load FrmDetalleTelas
    Set FrmDetalleTelas.oParent = Me
    FrmDetalleTelas.vAccion = "U"
    FrmDetalleTelas.vCod_Tela = Me.vCod_Tela
    FrmDetalleTelas.TxtDescripcion.Text = GridEX1.Value(GridEX1.Columns("Descripcion").Index)
    FrmDetalleTelas.vRuta = GridEX1.Value(GridEX1.Columns("Cod_Ruta").Index)
    FrmDetalleTelas.Show vbModal
    Set FrmDetalleTelas = Nothing

Case "ELIMINAR"
    If GridEX1.RowCount = 0 Then Exit Sub

    If MsgBox("Esta seguro de eliminar esta Descripción", vbYesNo, "IMPORTANTE") = vbYes Then
      strSQL = "tx_up_man_Tela_DatTecnicos_cabecera 'D" & "','" & Me.vCod_Tela & "','" & GridEX1.Value(GridEX1.Columns("Cod_Ruta").Index) & "','" _
          & GridEX1.Value(GridEX1.Columns("Descripcion").Index) & "'"

      Call ExecuteCommandSQL(cCONNECT, strSQL)
      Call CARGA_GRID
    End If

Case "DETALLE"
            If GridEX1.RowCount = 0 Then Exit Sub
            If vFamite = "DE" Then
                MsgBox "Debe ingresar los datos por Combinacion", vbCritical
                Exit Sub
            End If
            Load FrmManTelasDet
            FrmManTelasDet.sCod_Tela = vCod_Tela
            FrmManTelasDet.sDes_tela = vDes_Tela
            FrmManTelasDet.sFamite = vFamite
            FrmManTelasDet.vRuta = GridEX1.Value(GridEX1.Columns("Cod_Ruta").Index)
            FrmManTelasDet.Carga_Datos
            FrmManTelasDet.Show 1
            Set FrmManTelasDet = Nothing
             
Case "HOJARUTA"
            If GridEX1.RowCount = 0 Then Exit Sub
            Call Hoja_Ruta

Case "PRUEBAS"
            If GridEX1.RowCount = 0 Then Exit Sub
            'vFamite
            If vFamite = "DE" Then
                MsgBox "Debe ingresar los datos por Combinacion", vbCritical
                Exit Sub
            End If
            Load FrmManTelasDatTecAdd
            FrmManTelasDatTecAdd.sCod_Tela = vCod_Tela
            FrmManTelasDatTecAdd.sCod_Ruta = GridEX1.Value(GridEX1.Columns("cod_ruta").Index)
            FrmManTelasDatTecAdd.sFamite = vFamite
            FrmManTelasDatTecAdd.Carga_Datos
            FrmManTelasDatTecAdd.Show 1
            Set FrmManTelasDatTecAdd = Nothing

Case "SALIR"
    Unload Me
End Select
End Sub

Sub Hoja_Ruta()
Dim sRuta As String
On Error GoTo hand
    Dim oo As Object
    Dim strSQL As String
    Screen.MousePointer = 11
    
    sRuta = vRuta
    
    Set oo = CreateObject("excel.application")
    oo.workbooks.Open sRuta & "\Hoja_Tecnica_Rutas.xlt"
    oo.Visible = True
    oo.run "Reporte", vCod_Tela, cCONNECT, sRuta, vCOD_ORDTRA, GridEX1.Value(GridEX1.Columns("Cod_Ruta").Index), GridEX1.Value(GridEX1.Columns("Descripcion").Index)
    Screen.MousePointer = vbNormal
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler Err, "Telas No Operativas"
    Screen.MousePointer = vbNormal
    Set oo = Nothing
End Sub
