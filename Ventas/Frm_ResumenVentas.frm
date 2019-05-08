VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form Frm_Resumen_Ventas 
   Caption         =   "Resumen Deuda"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   14970
   StartUpPosition =   3  'Windows Default
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   12360
      TabIndex        =   1
      Top             =   7080
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"Frm_ResumenVentas.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      Begin GridEX20.GridEX GridEX1 
         Height          =   6735
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   11880
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "Frm_ResumenVentas.frx":0090
         Column(2)       =   "Frm_ResumenVentas.frx":0158
         FormatStylesCount=   6
         FormatStyle(1)  =   "Frm_ResumenVentas.frx":01FC
         FormatStyle(2)  =   "Frm_ResumenVentas.frx":0334
         FormatStyle(3)  =   "Frm_ResumenVentas.frx":03E4
         FormatStyle(4)  =   "Frm_ResumenVentas.frx":0498
         FormatStyle(5)  =   "Frm_ResumenVentas.frx":0570
         FormatStyle(6)  =   "Frm_ResumenVentas.frx":0628
         ImageCount      =   0
         PrinterProperties=   "Frm_ResumenVentas.frx":0708
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   4080
      Top             =   7320
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "Frm_Resumen_Ventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SBLOQUEO As Integer

Private Sub Form_Load()
Buscar
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
    Imprimir
Case "SALIR"
    Unload Me
    
End Select

End Sub

Private Sub Imprimir()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim strSQL As String
Dim sEmpresa As String

    Ruta = vRuta & "\Rpt_DeudaClientes.XLT"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "reporte", cCONNECT
    
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub



Sub Buscar()

On Error GoTo dprDepurar

Dim sSql As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle

sSql = "Cn_Resumen_DeudaPorCobrar "

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSql, cCONNECT)

GridEX1.Columns("Cod_Tipanex").Width = 0
GridEX1.Columns("Cod_Anxo").Width = 0


GridEX1.FrozenColumns = 5

GridEX1.Columns("Des_Anexo").Width = 3000
GridEX1.Columns("Des_Anexo").Caption = "Cliente"

GridEX1.Columns("Facturasol").Width = 1000
GridEX1.Columns("Facturasol").Caption = "Fact Sol"

GridEX1.Columns("facturaDol").Width = 1000
GridEX1.Columns("facturaDol").Caption = "Fact Dol"

GridEX1.Columns("PorAceptarSol").Width = 1000
GridEX1.Columns("PorAceptarSol").Caption = "Aceptar Sol"

GridEX1.Columns("PorAceptarDol").Width = 1000
GridEX1.Columns("PorAceptarDol").Caption = "Aceptar Dol"


GridEX1.Columns("AceptadaSol").Width = 1000
GridEX1.Columns("AceptadaSol").Caption = "Aceptar Sol"

GridEX1.Columns("AceptadaDol").Width = 1000
GridEX1.Columns("AceptadaDol").Caption = "Aceptar Dol"

GridEX1.Columns("DecuentoSol").Width = 1000
GridEX1.Columns("DecuentoSol").Caption = "Descuento Sol"

GridEX1.Columns("DecuentoDol").Width = 1000
GridEX1.Columns("DecuentoDol").Caption = "Descuento Dol"

GridEX1.Columns("AbonarSol").Width = 1000
GridEX1.Columns("AbonarSol").Caption = "Abonar Sol"

GridEX1.Columns("AbonarDol").Width = 1000
GridEX1.Columns("AbonarDol").Caption = "Abonar Dol"

GridEX1.Columns("ImporteTotal").Width = 1000
GridEX1.Columns("ImporteTotal").Caption = "Importe Total"

GridEX1.Columns("Limite_Dolares").Format = "#,##0.00"

GridEX1.Columns("AceptadaSol").Format = "#,##0.00"
GridEX1.Columns("AceptadaDol").Format = "#,##0.00"

GridEX1.Columns("ImporteTotal").Format = "#,##0.00"

GridEX1.Columns("Facturasol").Format = "#,##0.00"

GridEX1.Columns("facturaDol").Format = "#,##0.00"





GridEX1.Columns("SEL").ColumnType = jgexCheckBox
GridEX1.Columns("SEL").Visible = True
GridEX1.Columns("SEL").EditType = jgexEditCheckBox
GridEX1.Columns("SEL").Width = 500

Exit Sub

dprDepurar:

errores err.Number
  
End Sub

Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
 If ColIndex <> GridEX1.Columns("SEL").Index Then
            Cancel = True
 End If
  SendKeys "{ENTER}"
End Sub

Private Sub GridEX1_BeforeColUpdate(ByVal Row As Long, ByVal ColIndex As Integer, ByVal OldValue As String, ByVal Cancel As GridEX20.JSRetBoolean)
If GridEX1.Value(ColIndex) = -1 Then
    SBLOQUEO = 1
Else
    SBLOQUEO = 0
End If
    

    strSQL = "Ti_Bloquea_Despacho '" & GridEX1.Value(GridEX1.Columns("Cod_TipAnEX").Index) & "','" & GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index) & "','" & SBLOQUEO & "','" & vusu & "','" & ComputerName & "'"
    ExecuteSQL cCONNECT, strSQL

                        
End Sub
