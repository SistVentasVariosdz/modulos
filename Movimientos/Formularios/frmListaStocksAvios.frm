VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmListaStocksAvios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stocks del "
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   510
      Left            =   2760
      TabIndex        =   1
      Top             =   3105
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   510
      Left            =   4680
      TabIndex        =   2
      Top             =   3105
      Width           =   1395
   End
   Begin GridEX20.GridEX gexLista 
      Height          =   3000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5292
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmListaStocksAvios.frx":0000
      Column(2)       =   "frmListaStocksAvios.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmListaStocksAvios.frx":016C
      FormatStyle(2)  =   "frmListaStocksAvios.frx":02A4
      FormatStyle(3)  =   "frmListaStocksAvios.frx":0354
      FormatStyle(4)  =   "frmListaStocksAvios.frx":0408
      FormatStyle(5)  =   "frmListaStocksAvios.frx":04E0
      FormatStyle(6)  =   "frmListaStocksAvios.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmListaStocksAvios.frx":0678
   End
End
Attribute VB_Name = "frmListaStocksAvios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String

Public varCod_Almacen As String
Public varCod_Item As String

Public oParent As Object
Public varBusqueda As String
Public varOpcionBusq As String

Sub CARGA_GRID()
    
    'Esta cadena es para devolver el Codigo de Cliente
    strSQL = "EXEC SM_AYUDA_ITEMS_CON_STOCKS '" & Me.varCod_Almacen & "','" & Me.varCod_Item & "'"
    
    Set gexLista.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    
    SetGeneralGridEX gexLista, 0, 1
    
    Call Me.gexLista.Find(4, jgexEqual, varBusqueda)
    'If gexLista.RowCount > 0 Then
    '    gexLista.MoveFirst
    'End If
    
    Call Configurar_Grid
    
End Sub

Private Sub cmdAceptar_Click()
    If gexLista.RowCount > 0 Then
        With oParent
            If RTrim(varOpcionBusq) = "" Then
                '.CmbCombinacion_DropDown
                'BuscaCombo gexLista.Value(gexLista.Columns("cod_comb").Index), 2, .CmbCombinacion
                .TxtCod_Comb = Mid(gexLista.Value(gexLista.Columns("cod_comb").Index), 1, 3)
                .TxtDes_comb = Mid(gexLista.Value(gexLista.Columns("cod_comb").Index), 5)
                
                .TxtDetalle = gexLista.Value(gexLista.Columns("Des_Color").Index)
                .CmbColor = gexLista.Value(gexLista.Columns("Cod_Color").Index)
                
                'BuscaCombo gexLista.Value(gexLista.Columns("Dest.").Index), 2, .CmbDestino
                .Txtcod_Destino = gexLista.Value(gexLista.Columns("Dest.").Index)
                .TxtDes_Destino = DevuelveCampo("select des_destino from tg_destino where cod_destino='" & gexLista.Value(gexLista.Columns("Dest.").Index) & "'", cConnect)
                
                'BuscaCombo RTrim(gexLista.Value(gexLista.Columns("Est.Cli").Index)), 1, .CmbEstilo
                .TxtCod_EstCli = RTrim(gexLista.Value(gexLista.Columns("Est.Cli").Index))
                
                .varTallaProv = gexLista.Value(gexLista.Columns("cod_talla").Index)
                '.CmbTalla_DropDown
                'BuscaCombo gexLista.Value(gexLista.Columns("cod_talla").Index), 2, .CmbTalla
                .TxtCod_Medida = gexLista.Value(gexLista.Columns("cod_talla").Index)
                .TxtDes_Medida = Mid(gexLista.Value(gexLista.Columns("medida").Index), 12)
                .TxtCodProv = gexLista.Value(gexLista.Columns("cod_PROV").Index)
            Else
                '.CmbCombinacion_DropDown
                BuscaCombo gexLista.Value(gexLista.Columns("cod_comb").Index), 2, .CmbCombinacion
                
                .TxtDetalle = gexLista.Value(gexLista.Columns("Des_Color").Index)
                .CmbColor = gexLista.Value(gexLista.Columns("Cod_Color").Index)
                
                BuscaCombo gexLista.Value(gexLista.Columns("Dest.").Index), 2, .CmbDestino
                
                BuscaCombo RTrim(gexLista.Value(gexLista.Columns("Est.Cli").Index)), 1, .CmbEstilo
                
                .varTallaProv = gexLista.Value(gexLista.Columns("cod_talla").Index)
                .CmbTalla_DropDown
                BuscaCombo gexLista.Value(gexLista.Columns("cod_talla").Index), 2, .CmbTalla
                .TxtCodProv = gexLista.Value(gexLista.Columns("cod_PROV").Index)
            
            End If
        End With
    End If
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub gexLista_DblClick()
    Call cmdAceptar_Click
End Sub

Private Sub gexLista_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call cmdAceptar_Click
    End If
End Sub

Private Sub gexLista_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then
    '    Call cmdAceptar_Click
    'End If
End Sub

Public Sub Configurar_Grid()
    Me.gexLista.Columns("cod_comb").Visible = False
    Me.gexLista.Columns("cod_color").Visible = False
    Me.gexLista.Columns("des_color").Visible = False
    Me.gexLista.Columns("cod_talla").Visible = False
    Me.gexLista.Columns("UM").Visible = False

    Me.gexLista.Columns("Comb.").Width = 1800
    Me.gexLista.Columns("Color").Width = 1800
    Me.gexLista.Columns("Medida").Width = 1000
    Me.gexLista.Columns("Dest.").Width = 1500
    Me.gexLista.Columns("Est.Cli").Width = 1500
    Me.gexLista.Columns("Stock").Width = 1000
    
End Sub


