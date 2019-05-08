VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmShowTiposCambio 
   Caption         =   "Tipos de Cambio"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   11145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione Rango de Fechas :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   10932
      Begin VB.Frame Frame2 
         Caption         =   "Tipo Cambio Promedio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   5250
         TabIndex        =   8
         Top             =   120
         Width           =   4335
         Begin VB.Label lblVentaPromedio 
            BackColor       =   &H8000000D&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   225
            Left            =   1140
            TabIndex        =   12
            Top             =   345
            Width           =   1005
         End
         Begin VB.Label Label5 
            Caption         =   "Tipo Venta :"
            Height          =   240
            Left            =   90
            TabIndex        =   11
            Top             =   330
            Width           =   1035
         End
         Begin VB.Label lblCompraPromedio 
            BackColor       =   &H8000000D&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   225
            Left            =   3270
            TabIndex        =   10
            Top             =   345
            Width           =   1005
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Compra :"
            Height          =   240
            Left            =   2220
            TabIndex        =   9
            Top             =   330
            Width           =   1035
         End
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   612
         Left            =   9600
         TabIndex        =   1
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin MSComCtl2.DTPicker DTFecha_Ini 
         Height          =   315
         Left            =   1260
         TabIndex        =   2
         Top             =   390
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   94109697
         CurrentDate     =   37504
      End
      Begin MSComCtl2.DTPicker DTFecha_Fin 
         Height          =   315
         Left            =   3735
         TabIndex        =   3
         Top             =   390
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   94109697
         CurrentDate     =   37504
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final"
         Height          =   240
         Left            =   2835
         TabIndex        =   5
         Top             =   450
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial"
         Height          =   240
         Left            =   180
         TabIndex        =   4
         Top             =   450
         Width           =   1305
      End
   End
   Begin GridEX20.GridEX gridex1 
      Height          =   5856
      Left            =   60
      TabIndex        =   6
      Top             =   1092
      Width           =   9672
      _ExtentX        =   17066
      _ExtentY        =   10319
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorBkg    =   -2147483624
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmShowTiposCambio.frx":0000
      Column(2)       =   "frmShowTiposCambio.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmShowTiposCambio.frx":016C
      FormatStyle(2)  =   "frmShowTiposCambio.frx":02A4
      FormatStyle(3)  =   "frmShowTiposCambio.frx":0354
      FormatStyle(4)  =   "frmShowTiposCambio.frx":0408
      FormatStyle(5)  =   "frmShowTiposCambio.frx":04E0
      FormatStyle(6)  =   "frmShowTiposCambio.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmShowTiposCambio.frx":0678
   End
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   4116
      Left            =   9840
      TabIndex        =   7
      Top             =   1092
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   7250
      Custom          =   $"frmShowTiposCambio.frx":0850
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   10080
      Top             =   5880
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmShowTiposCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim sSeguridad As String
sSeguridad = get_botones1(Me, vper, vemp, Me.Name)
    
Me.FunctButt2.FunctionsUser = sSeguridad
Me.DTFecha_Ini = Date
Me.DTFecha_Fin = Date
Buscar
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Buscar
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ADICIONAR"
            Load frmManTipoCambio
            frmManTipoCambio.DTFecha.Value = DTFecha_Ini.Value
            frmManTipoCambio.sAccion = "I"
            Set frmManTipoCambio.oParent = Me
            frmManTipoCambio.Show vbModal
            Set frmManTipoCambio = Nothing
        Case "MODIFICAR"
            Load frmManTipoCambio
            
            frmManTipoCambio.sAccion = "U"
            CargarDatos
            Set frmManTipoCambio.oParent = Me
            frmManTipoCambio.Show vbModal
            Set frmManTipoCambio = Nothing
        Case "CONSULTAR"
            Load frmManTipoCambio
            CargarDatos
            frmManTipoCambio.FunctButt2.ChangeProperty "ENABLED", 0, False
            Set frmManTipoCambio.oParent = Me
            frmManTipoCambio.Show vbModal
            Set frmManTipoCambio = Nothing
        Case "IMPRIMIR"
            ReporteTC
        Case "SALIR"
            Unload Me
    End Select
End Sub



Private Sub CargarDatos()
'
frmManTipoCambio.DTFecha.Value = FixNulos(gridex1.Value(gridex1.Columns("Fecha").Index), vbDouble)
frmManTipoCambio.TxtCambio = FixNulos(gridex1.Value(gridex1.Columns("Tipo_Cambio").Index), vbDouble)
frmManTipoCambio.TxtCompra = FixNulos(gridex1.Value(gridex1.Columns("Tipo_Compra").Index), vbDouble)
frmManTipoCambio.TxtVenta = FixNulos(gridex1.Value(gridex1.Columns("Tipo_Venta").Index), vbDouble)
frmManTipoCambio.txtEuros = FixNulos(gridex1.Value(gridex1.Columns("Tipo_Cambio_Euros").Index), vbDouble)
frmManTipoCambio.txtFrancos = FixNulos(gridex1.Value(gridex1.Columns("Tipo_Cambio_Francos").Index), vbDouble)
frmManTipoCambio.txtMarcos = FixNulos(gridex1.Value(gridex1.Columns("Tipo_Cambio_Marcos").Index), vbDouble)
frmManTipoCambio.txtYen = FixNulos(gridex1.Value(gridex1.Columns("Tipo_Cambio_Yen").Index), vbDouble)
frmManTipoCambio.txtEurosCompra = FixNulos(gridex1.Value(gridex1.Columns("Tipo_Compra_Euros").Index), vbDouble)

End Sub

Private Sub Buscar()
On Error GoTo errx

Dim strSQL As String
strSQL = "CN_MUESTRA_TIPO_CAMBIO_PROMEDIO '" & DTFecha_Ini.Value & "','" & DTFecha_Fin.Value & "'"
Dim adoRs As Object
Set adoRs = CreateObject("ADODB.Recordset")
Set adoRs = CargarRecordSetDesconectado(strSQL, cCONNECT)
If adoRs.State <> adStateClosed Then
    If adoRs.RecordCount > 0 Then
        lblCompraPromedio.Caption = adoRs("TIPO_COMPRA").Value
        lblVentaPromedio.Caption = adoRs("TIPO_VENTA").Value
    End If
End If

Dim sSQL As String
Dim vBookmark  As Variant

sSQL = "SM_CN_TIPOCAMBIO '$','$'"
sSQL = VBsprintf(sSQL, DTFecha_Ini.Value, DTFecha_Fin.Value)

vBookmark = gridex1.Row

gridex1.ColumnHeaderHeight = 500
gridex1.ClearFields

Set gridex1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)

gridex1.Columns("Tipo_Cambio").Caption = "Tipo Cambio"
gridex1.Columns("Tipo_Cambio").Caption = "Tipo Cambio"
gridex1.Columns("Tipo_Cambio").Width = 1000
gridex1.Columns("Tipo_Venta").Caption = "Tipo Venta"
gridex1.Columns("Tipo_Venta").Width = 1100
gridex1.Columns("Tipo_Compra").Caption = "Tipo Compra"
gridex1.Columns("Tipo_Compra").Width = 1100
gridex1.Columns("Tipo_Cambio_Euros").Caption = "Tipo Venta Euros"
gridex1.Columns("Tipo_Cambio_Euros").Width = 1000
gridex1.Columns("Tipo_Cambio_Marcos").Caption = "Tipo Cambio Marcos"
gridex1.Columns("Tipo_Cambio_Marcos").Width = 1000
gridex1.Columns("Tipo_Cambio_Francos").Caption = "Tipo Cambio Francos"
gridex1.Columns("Tipo_Cambio_Francos").Width = 1000
gridex1.Columns("Tipo_Cambio_Yen").Caption = "Tipo Cambio Yen"
gridex1.Columns("Tipo_Cambio_Yen").Width = 1000
gridex1.Columns("Tipo_Compra_Euros").Caption = "Tipo Compra Euros"
gridex1.Columns("Tipo_Compra_Euros").Width = 1000



gridex1.Row = vBookmark


Exit Sub
errx:
    ErrorHandler err, "SALVAR_DATOS"
End Sub

Sub ReporteTC()

Dim RS As Object
Set RS = CreateObject("ADODB.Recordset")
Dim oo As Object
Dim empresa As String

On Error GoTo ErrorImpresion


    empresa = DevuelveCampo("select des_empresa from seguridad..seg_empresas where cod_empresa='" & vemp & "'", cCONNECT)
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\RptTipoCambio.XLT"
    
    oo.Visible = True
    oo.Run "REPORTE", gridex1.ADORecordset, Me.DTFecha_Ini, Me.DTFecha_Fin, empresa
    oo.Visible = True
    Set oo = Nothing
    Screen.MousePointer = vbNormal

    
Exit Sub
Resume
ErrorImpresion:
    Set oo = Nothing
    Set RS = Nothing
    Screen.MousePointer = vbNormal
    MsgBox "Hubo error en la impresion del Reporte " & err.Description, vbCritical, "Impresion"
End Sub
