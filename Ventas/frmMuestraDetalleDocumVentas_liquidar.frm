VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMuestraDetalleDocumVentas_liquidar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3675
   ClientLeft      =   240
   ClientTop       =   1500
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   9975
   Begin VB.Frame fraAjuste 
      Caption         =   "Ajuste Importe"
      Height          =   1335
      Left            =   3120
      TabIndex        =   3
      Top             =   1260
      Visible         =   0   'False
      Width           =   3615
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   480
         Left            =   960
         TabIndex        =   6
         Top             =   750
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   847
         Custom          =   $"frmMuestraDetalleDocumVentas_liquidar.frx":0000
         Orientacion     =   0
         Style           =   1
         Language        =   1
         TypeImageList   =   0
         ControlWidth    =   1000
         ControlHeigth   =   450
         ControlSeparator=   10
      End
      Begin VB.TextBox txtImporteTotal 
         Height          =   285
         Left            =   1410
         TabIndex        =   5
         Top             =   315
         Width           =   1905
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Importe Total : "
         Height          =   195
         Left            =   270
         TabIndex        =   4
         Top             =   360
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   480
      Left            =   7320
      TabIndex        =   2
      Top             =   3165
      Width           =   1450
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   615
      Left            =   8760
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
      _ExtentX        =   2117
      _ExtentY        =   900
      Custom          =   "0~0~AJUSIMP~Verdadero~Verdadero~&Ajuste Importe~0~0~1~~0~Falso~Falso~&Ajuste Importe~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   5530
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmMuestraDetalleDocumVentas_liquidar.frx":0097
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmMuestraDetalleDocumVentas_liquidar.frx":03E9
      Column(2)       =   "frmMuestraDetalleDocumVentas_liquidar.frx":04B1
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmMuestraDetalleDocumVentas_liquidar.frx":0555
      FormatStyle(2)  =   "frmMuestraDetalleDocumVentas_liquidar.frx":068D
      FormatStyle(3)  =   "frmMuestraDetalleDocumVentas_liquidar.frx":073D
      FormatStyle(4)  =   "frmMuestraDetalleDocumVentas_liquidar.frx":07F1
      FormatStyle(5)  =   "frmMuestraDetalleDocumVentas_liquidar.frx":08C9
      FormatStyle(6)  =   "frmMuestraDetalleDocumVentas_liquidar.frx":0981
      FormatStyle(7)  =   "frmMuestraDetalleDocumVentas_liquidar.frx":0A61
      FormatStyle(8)  =   "frmMuestraDetalleDocumVentas_liquidar.frx":0F19
      ImageCount      =   1
      ImagePicture(1) =   "frmMuestraDetalleDocumVentas_liquidar.frx":1365
      PrinterProperties=   "frmMuestraDetalleDocumVentas_liquidar.frx":16B7
   End
End
Attribute VB_Name = "frmMuestraDetalleDocumVentas_liquidar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strSQL As String, Num_Corre As String, strCod_TipDoc As String

Public sDOCUMENTO As String

Public Function BUSCAR() As Boolean

On Error GoTo errores

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
GridEX1.Columns("T").Width = 240
GridEX1.Columns("Codigo").Width = 1455
GridEX1.Columns("Codigo").Caption = "Codigo"
GridEX1.Columns("Articulo").Width = 4455
GridEX1.Columns("Articulo").Caption = "Articulo"
GridEX1.Columns("Cantidad").Width = 765
GridEX1.Columns("Cantidad").Caption = "Cantidad"
GridEX1.Columns("Uni_Med").Width = 780
GridEX1.Columns("Uni_Med").Caption = "Uni Med"
GridEX1.Columns("Valor_Unitario").Width = 1125
GridEX1.Columns("Valor_Unitario").Caption = "Valor Unitario"
GridEX1.Columns("Valor_Venta").Width = 1005
GridEX1.Columns("Valor_Venta").Caption = "Valor Venta"
GridEX1.Columns("Num_Corre").Visible = False
GridEX1.Columns("Secuencia").Visible = False
GridEX1.Columns("Origen").Visible = False

Exit Function
Resume
errores:
    errores err.Number
End Function
Private Function ifValidaDoc() As Boolean

Dim strMsg As String

strMsg = DevuelveCampo("Select dbo.ventas_Valida_Documento_Manuales_Det('" & Num_Corre & "')", cCONNECT)
If strMsg <> "" Then
  MsgBox strMsg, vbInformation, "AVISO"
  ifValidaDoc = False
  Exit Function
End If

ifValidaDoc = True

End Function

Private Sub cmdImprimir_Click()
On Error GoTo ERROR
Dim oo As Object, vRutaLogo As Variant, _
    sRutaLogo As String, sTitulo As String, Ruta As String
    
    If GridEX1.ADORecordset.RecordCount > 0 Then
        sTitulo = Trim(sDOCUMENTO) & " : " & Trim(Num_Corre)
        
        If MsgBox("Desea imprimir usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir") = vbYes Then
            Set oo = CreateObject("excel.application")
            oo.Workbooks.Open vRuta & "\RptDetalleDeDocumento.XLT"
            oo.DisplayAlerts = False
            oo.Visible = True
            
            strSQL = "SELECT Ruta_Logo From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
            sRutaLogo = DevuelveCampo(strSQL, cCONNECT)
            sRutaLogo = CStr(IIf(IsNull(sRutaLogo), "", sRutaLogo))
            
            oo.Run "REPORTE", CStr(sRutaLogo), GridEX1.ADORecordset, sTitulo
       
        End If
        Set oo = Nothing
    Else
        MsgBox "No se han encontrado datos para imprirmir....", vbInformation
    End If
Exit Sub
ERROR:
    ErrorHandler err, "[VENTAS] : Ranking de Ventas por Pais-Destino"
End Sub

Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

Dim lvSql As String

On Error GoTo DrpDepurar

Select Case ActionName
Case Is = "ADICIONAR"
  
  If Not ifValidaDoc Then Exit Sub
  
  If strCod_TipDoc = "NC" Or strCod_TipDoc = "ND" Then
    Load frmAdicionaDetalleDocumAsigNotas
    With frmAdicionaDetalleDocumAsigNotas
      .Caption = "Adicion " + Me.Caption
      .strNum_Corre_Ori = Num_Corre
      .Show 1
    End With
  Else
    Load frmAdicionaDetalleDocum
    With frmAdicionaDetalleDocum
      .Caption = "Adicion " + Me.Caption
      .strNum_Corre_Detalle = Num_Corre
      .IntSencuencia = 0
      .strOption = "I"
      .Show 1
    End With
  End If
  
  BUSCAR
  Call GridEX1.Find(GridEX1.Columns("Secuencia").Index, jgexEqual, frmAdicionaDetalleDocum.IntSencuencia)
      
Case Is = "MODIFICAR"

  If GridEX1.RowCount = 0 Then Exit Sub
  
  If Not ifValidaDoc Then Exit Sub
  
  Load frmAdicionaDetalleDocum
  With frmAdicionaDetalleDocum
    .Caption = "Modificar " + Me.Caption
    .strNum_Corre_Detalle = GridEX1.Value(GridEX1.Columns("Num_Corre").Index)
    .IntSencuencia = GridEX1.Value(GridEX1.Columns("Secuencia").Index)
    .strOption = "U"
    .strNum_Corre_Detalle = GridEX1.Value(GridEX1.Columns("Num_Corre").Index)
    .txtTip_Item = GridEX1.Value(GridEX1.Columns("T").Index)
    .txtCod_Producto = GridEX1.Value(GridEX1.Columns("Codigo").Index)
    .txtDescripcion = GridEX1.Value(GridEX1.Columns("Articulo").Index)
    .txtCantidad.Text = GridEX1.Value(GridEX1.Columns("Cantidad").Index)
    .txtUnida_Medida.Text = GridEX1.Value(GridEX1.Columns("Uni_Med").Index)
    .txtImp_Unitario.Text = GridEX1.Value(GridEX1.Columns("Valor_Unitario").Index)
    .txtImp_Total.Text = GridEX1.Value(GridEX1.Columns("Valor_Venta").Index)
    .txtPorc_Commision.Text = GridEX1.Value(GridEX1.Columns("Porcentaje_Commision").Index)
    .txtCantUniAlter.Text = GridEX1.Value(GridEX1.Columns("Cantidad_Item_NC_ND").Index)
    If Trim(GridEX1.Value(GridEX1.Columns("T").Index)) = "P" Then
        .cmdBuscar.Visible = True
    Else
        .cmdBuscar.Visible = False
    End If
    .Show 1
    BUSCAR
    Call GridEX1.Find(GridEX1.Columns("Secuencia").Index, jgexEqual, .IntSencuencia)
  End With
Case Is = "ELIMINAR"

  If Not ifValidaDoc Then Exit Sub
  
  If GridEX1.RowCount = 0 Then Exit Sub
  If MsgBox("Esta Seguro de Eliminar este Registro", vbYesNo, "ADVERTENCIA") = vbYes Then
    lvSql = "Ventas_Up_Man_Detalle '" & "D" & "','" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "'," & GridEX1.Value(GridEX1.Columns("Secuencia").Index)
    ExecuteCommandSQL cCONNECT, lvSql
    BUSCAR
  End If
Case Is = "AJUSIMP"
    fraAjuste.Visible = True
    txtImporteTotal.Text = GridEX1.Value(GridEX1.Columns("Valor_Venta").Index)
    GridEX1.Enabled = True
    txtImporteTotal.SetFocus
Case Is = "SALIR"
  Unload Me
End Select

Exit Sub

DrpDepurar:

errores err.Number

End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo lblError
Dim lvSql  As String
Select Case ActionName
Case "ACEPTAR"
  If GridEX1.RowCount = 0 Then Exit Sub
  If MsgBox("Esta Seguro de ajustar este Registro", vbYesNo, "ADVERTENCIA") = vbYes Then
    lvSql = "EXEC CN_VENTAS_AJUSTA_IMPORTE_VENTA '" & GridEX1.Value(GridEX1.Columns("Num_Corre").Index) & "'," & GridEX1.Value(GridEX1.Columns("Secuencia").Index) & "," & txtImporteTotal.Text & ""
    ExecuteCommandSQL cCONNECT, lvSql
    BUSCAR
    fraAjuste.Visible = False
    txtImporteTotal.Text = ""
  End If
Case "CANCELAR"
    fraAjuste.Visible = False
    txtImporteTotal.Text = ""
End Select
Exit Sub
lblError:
    MsgBox err.Description, vbCritical, "Mensaje del Sistema"
    Exit Sub
End Sub


Private Sub txtImporteTotal_KeyPress(KeyAscii As Integer)
Call SoloNumeros(txtImporteTotal, KeyAscii, True, 2, 16)
End Sub


