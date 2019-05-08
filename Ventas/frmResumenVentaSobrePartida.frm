VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmResumenVentaSobrePartida 
   Caption         =   "Reporte Por SubPartida Arancelaria"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   14760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Detallado"
      Height          =   615
      Left            =   5400
      TabIndex        =   7
      Top             =   1380
      Width           =   9252
      Begin VB.OptionButton optAnexoEstiloD 
         Caption         =   "Por Anexo/Estilo"
         Height          =   255
         Left            =   7200
         TabIndex        =   15
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optAnexoD 
         Caption         =   "Por Anexo"
         Height          =   255
         Left            =   5400
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optClienteEstiloD 
         Caption         =   "Por Cliente/Estilo"
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optClienteD 
         Caption         =   "Por Cliente"
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optGeneralD 
         Caption         =   "General"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Resumido"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   1380
      Width           =   5055
      Begin VB.OptionButton optAnexoR 
         Caption         =   "Por Anexo"
         Height          =   255
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optClienteR 
         Caption         =   "Por Cliente"
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optGeneralR 
         Caption         =   "General"
         Height          =   255
         Left            =   480
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   12255
      Begin VB.OptionButton optPendiente 
         Caption         =   "Pendiente de Envío Draw Back"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   2835
      End
      Begin VB.OptionButton optFec_EnvioDrawBack 
         Caption         =   "Fecha de Envío Draw Back"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   660
         Width           =   2295
      End
      Begin VB.OptionButton optFec_EmiDoc 
         Caption         =   "Fecha de Emisión"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   270
         Left            =   3120
         TabIndex        =   2
         Top             =   300
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         _Version        =   393216
         Format          =   94109697
         CurrentDate     =   39402
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   270
         Left            =   5760
         TabIndex        =   3
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
         _Version        =   393216
         Format          =   94109697
         CurrentDate     =   39402
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   492
         Left            =   10920
         TabIndex        =   17
         Top             =   180
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   714
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   400
         ControlSeparator=   110
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta  :"
         Height          =   255
         Left            =   4920
         TabIndex        =   5
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   300
         Width           =   735
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4215
      Left            =   120
      TabIndex        =   8
      Top             =   2220
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   7435
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmResumenVentaSobrePartida.frx":0000
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmResumenVentaSobrePartida.frx":0352
      Column(2)       =   "frmResumenVentaSobrePartida.frx":041A
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmResumenVentaSobrePartida.frx":04BE
      FormatStyle(2)  =   "frmResumenVentaSobrePartida.frx":05F6
      FormatStyle(3)  =   "frmResumenVentaSobrePartida.frx":06A6
      FormatStyle(4)  =   "frmResumenVentaSobrePartida.frx":075A
      FormatStyle(5)  =   "frmResumenVentaSobrePartida.frx":0832
      FormatStyle(6)  =   "frmResumenVentaSobrePartida.frx":08EA
      FormatStyle(7)  =   "frmResumenVentaSobrePartida.frx":09CA
      FormatStyle(8)  =   "frmResumenVentaSobrePartida.frx":0E82
      ImageCount      =   1
      ImagePicture(1) =   "frmResumenVentaSobrePartida.frx":12CE
      PrinterProperties=   "frmResumenVentaSobrePartida.frx":1620
   End
   Begin FunctionsButtons.FunctButt FunctButt3 
      Height          =   510
      Left            =   6120
      TabIndex        =   16
      Top             =   6600
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmResumenVentaSobrePartida.frx":17F8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   480
      Top             =   6000
      _cx             =   677
      _cy             =   677
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmResumenVentaSobrePartida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim opcion As String
Dim strSQL As String
Dim sModo  As String


Private Sub DesactivaDetallado()
    Me.optGeneralD.Value = False
    Me.optClienteD.Value = False
    Me.optClienteEstiloD.Value = False
    Me.optAnexoD.Value = False
    Me.optAnexoEstiloD = False
End Sub

Private Sub DesactivaResumido()
    Me.optGeneralR.Value = False
    Me.optClienteR.Value = False
    Me.optAnexoR.Value = False
End Sub





Private Sub Form_Load()
    Me.dtpDesde = Date
    Me.dtpHasta = Date
    sModo = "1"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
  Case "BUSCAR"
    Buscar
End Select
End Sub

Public Sub Buscar()
Dim fmtCon As JSFmtCondition
On Error GoTo Err_Buscar
    
strSQL = "EXEC cn_Resumen_Ventas_Por_SobrePartidaArancelaria '" & opcion & "','" & Me.dtpDesde & "','" & Me.dtpHasta & "','" & sModo & "'"

                                            
Set gridex1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
gridex1.ColumnHeaderHeight = 500
gridex1.Columns("Des. Partida").Width = 7500
gridex1.Columns("Tipo").Width = 0


If opcion = "3" Or opcion = "7" Or opcion = "8" Then
    gridex1.Columns("Cod. TipAnex").Width = 400
    gridex1.Columns("Cod. Anexo").Width = 550
End If
If opcion = "6" Or opcion = "8" Then
    gridex1.Columns("Cod. Estilo Cliente").Width = 1000
End If
gridex1.Columns("Num. Partida Arancelaria").Width = 1100
gridex1.Columns("Sec.Partida Arancelaria").Width = 430
If opcion = "4" Or opcion = "5" Or opcion = "6" Or opcion = "7" Or opcion = "8" Then gridex1.Columns("Factura").Width = 1100
gridex1.Columns("Num. Prendas").Width = 900
gridex1.Columns("Imp. Total").Width = 1500
If opcion = "2" Or opcion = "5" Or opcion = "6" Then gridex1.Columns("Cliente").Width = 1300

If opcion = "6" Or opcion = "7" Or opcion = "8" Then gridex1.Columns("Des. Partida").Width = 6000

Set fmtCon = gridex1.FmtConditions.Add(gridex1.Columns("tipo").Index, jgexEqual, "2")
fmtCon.FormatStyle.BackColor = &HFFFFC0
  
gridex1.Columns("Num. Prendas").Format = "###,###"
gridex1.Columns("Imp. Total").Format = "###,###.00"

Exit Sub
Err_Buscar:
    MsgBox err.Description, vbCritical + vbOKOnly, "Stock Pre Saldos"
End Sub

Private Sub FunctButt3_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "IMPRIMIR"
        Reporte
    Case "SALIR"
        Unload Me
End Select
End Sub

Private Sub optAnexoD_Click()
    opcion = "7"
    DesactivaResumido
End Sub

Private Sub optAnexoEstiloD_Click()
    opcion = "8"
    DesactivaResumido
End Sub

Private Sub optAnexoR_Click()
    opcion = "3"
    DesactivaDetallado
End Sub

Private Sub optClienteD_Click()
    opcion = "5"
    DesactivaResumido
End Sub

Private Sub optClienteEstiloD_Click()
    opcion = "6"
    DesactivaResumido
End Sub

Private Sub optClienteR_Click()
    opcion = "2"
    DesactivaDetallado
End Sub

Private Sub optFec_EmiDoc_Click()
    sModo = "1"
End Sub

Private Sub optFec_EnvioDrawBack_Click()
    sModo = "2"
End Sub

Private Sub optGeneralD_Click()
    opcion = "4"
    DesactivaResumido
End Sub

Private Sub optGeneralR_Click()
    opcion = "1"
    DesactivaDetallado
End Sub
Sub Reporte()

Dim RS As Object
Set RS = CreateObject("ADODB.Recordset")
Dim oo As Object
Dim strSQL As String
Dim sEmpresa As String
On Error GoTo ErrorImpresion

    strSQL = "SELECT DES_EMPRESA FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA='" & vemp & "'"
    sEmpresa = DevuelveCampo(strSQL, cCONNECT)

    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\RptVPartidaArancelaria.XLT"
    
    oo.Visible = True
    oo.Run "REPORTE", Me.gridex1.ADORecordset, opcion, Me.dtpDesde, Me.dtpHasta, sEmpresa
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

Private Sub Option1_Click()

End Sub

Private Sub optPendiente_Click()
    sModo = "3"
End Sub
