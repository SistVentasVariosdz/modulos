VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form Frm_DetalleTejeduria 
   Caption         =   "Detalle Rollos"
   ClientHeight    =   4365
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   11490
   StartUpPosition =   3  'Windows Default
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   4560
      TabIndex        =   1
      Top             =   3840
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"Frm_DetalleTejeduria.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin GridEX20.GridEX GridEX1 
         Height          =   3495
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   6165
         Version         =   "2.0"
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "Frm_DetalleTejeduria.frx":0096
         Column(2)       =   "Frm_DetalleTejeduria.frx":015E
         FormatStylesCount=   6
         FormatStyle(1)  =   "Frm_DetalleTejeduria.frx":0202
         FormatStyle(2)  =   "Frm_DetalleTejeduria.frx":033A
         FormatStyle(3)  =   "Frm_DetalleTejeduria.frx":03EA
         FormatStyle(4)  =   "Frm_DetalleTejeduria.frx":049E
         FormatStyle(5)  =   "Frm_DetalleTejeduria.frx":0576
         FormatStyle(6)  =   "Frm_DetalleTejeduria.frx":062E
         ImageCount      =   0
         PrinterProperties=   "Frm_DetalleTejeduria.frx":070E
      End
   End
End
Attribute VB_Name = "Frm_DetalleTejeduria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Scod_ordtra As String
Public sCod_Almacen As String
Public sCod_TipMov As String
Public sNum_MovStk As String

Private Sub Form_Load()
BUSCAR
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "ACEPTAR"
        GridEX1.MoveFirst
        For i = 1 To GridEX1.RowCount
        
        If GridEX1.Value(GridEX1.Columns("Sel").Index) <> 0 Then
            StrSql = "EXEC lg_UP_MAN_TX_MOVISTK_DETALLE_ROLLOS 'I', '" & _
            sCod_Almacen & "', '" & sNum_MovStk & "', '" & sNum_Secuencia & "', '" & _
            GridEX1.Value(GridEX1.Columns("Prefijo_Maquina").Index) & "', '" & GridEX1.Value(GridEX1.Columns("Codigo_Rollo").Index) & "', '" & GridEX1.Value(GridEX1.Columns("Peso_Rollo_Actual").Index) & "', 0, '', '', '0','" & vusu & "'"
    
            ExecuteSQL cConnect, StrSql
        End If
         GridEX1.MoveNext
        Next

Unload Me
    Case "CANCELAR"
        Unload Me
End Select
End Sub

Sub BUSCAR()
On Error GoTo Errores
Dim sSQL As String
Dim vBookmark As Variant

sSQL = "Tj_Muestra_Rollos_Tejeduria  '$'"
sSQL = VBsprintf(sSQL, Scod_ordtra)

vBookmark = GridEX1.Row
GridEX1.ClearFields

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cConnect)

GridEX1.Row = vBookmark

GridEX1.ContinuousScroll = True


GridEX1.Columns("SEL").ColumnType = jgexCheckBox
GridEX1.Columns("SEL").Visible = True
GridEX1.Columns("SEL").EditType = jgexEditCheckBox
GridEX1.Columns("SEL").Width = 500



 GridEX1.Columns("Num_Secuencia").Visible = False
 GridEX1.Columns("Flg_Status").Visible = False
 GridEX1.Columns("Unidades_Rollo_Actual").Visible = False
 
  GridEX1.Columns("Cod_Calidad").Visible = False
 
GridEX1.Columns("Cod_TipoOC_Tejeduria").Visible = False
GridEX1.Columns("Des_TipoOC_Tejeduria").Visible = False
GridEX1.Columns("Kilos_Requeridos").Visible = False
GridEX1.Columns("Cliente").Visible = False

 
 GridEX1.Columns("Cod_OrdTra").Width = 700
 GridEX1.Columns("Cod_OrdTra").Caption = "Ot"
 
 
 GridEX1.Columns("Prefijo_Maquina").Width = 500
 GridEX1.Columns("Prefijo_Maquina").Caption = "Maq"
 
 
 GridEX1.Columns("Codigo_Rollo").Width = 850
 GridEX1.Columns("Codigo_Rollo").Caption = "Rollo"
 
 
 GridEX1.Columns("Des_Tela").Width = 4000
 GridEX1.Columns("Des_Tela").Caption = "Tela"
 
 
  GridEX1.Columns("Peso_Rollo_Actual").Width = 800
 GridEX1.Columns("Peso_Rollo_Actual").Caption = "Kg"



GridEX1.FrozenColumns = 2

Exit Sub

Errores:
    err.Raise err.Number, err.Source, err.Description
End Sub

Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
 If ColIndex <> GridEX1.Columns("SEL").Index Then
        Cancel = True
    End If
End Sub

