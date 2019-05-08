VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmTransFactVentas 
   Caption         =   "Transmision Facturas Venta a INKA"
   ClientHeight    =   8670
   ClientLeft      =   180
   ClientTop       =   495
   ClientWidth     =   11715
   Icon            =   "frmTransFactVentas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   11715
   Begin VB.Frame frAnoMes 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   11535
      Begin VB.CheckBox ChkPorTransmitir 
         Caption         =   "Solo Incluye por Transmitir"
         Height          =   255
         Left            =   4680
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtMes 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   3480
         MaxLength       =   2
         TabIndex        =   5
         Top             =   345
         Width           =   480
      End
      Begin VB.TextBox txtAno 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   0
         Top             =   345
         Width           =   660
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   570
         Left            =   9480
         TabIndex        =   9
         Top             =   195
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   1005
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~~~0~Verdadero~Falso~&Buscar~"
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1450
         ControlHeigth   =   550
         ControlSeparator=   35
      End
      Begin VB.Label Label6 
         Caption         =   "Año Registro :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Mes Registro : "
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7095
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   11535
      Begin GridEX20.GridEX GridEX1 
         Height          =   6720
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   11853
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
         Column(1)       =   "frmTransFactVentas.frx":030A
         Column(2)       =   "frmTransFactVentas.frx":03D2
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmTransFactVentas.frx":0476
         FormatStyle(2)  =   "frmTransFactVentas.frx":05AE
         FormatStyle(3)  =   "frmTransFactVentas.frx":065E
         FormatStyle(4)  =   "frmTransFactVentas.frx":0712
         FormatStyle(5)  =   "frmTransFactVentas.frx":07EA
         FormatStyle(6)  =   "frmTransFactVentas.frx":08A2
         FormatStyle(7)  =   "frmTransFactVentas.frx":0982
         FormatStyle(8)  =   "frmTransFactVentas.frx":0A2E
         ImageCount      =   0
         PrinterProperties=   "frmTransFactVentas.frx":0ADE
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   570
      Left            =   3600
      TabIndex        =   1
      Top             =   8040
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   1005
      Custom          =   $"frmTransFactVentas.frx":0CB6
      Orientacion     =   0
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1450
      ControlHeigth   =   550
      ControlSeparator=   35
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   5520
      Top             =   9480
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmTransFactVentas"
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
Public codigo As String, Descripcion As String, strCod_Anxo As String, TipoAdd As String
Dim lvSW As Boolean

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "TRANSMITIR"
          If GridEX1.RowCount = 0 Then Exit Sub
          Load frmAddTransaccion
          frmAddTransaccion.dtpFecha.Value = GridEX1.Value(GridEX1.Columns("Fec_Registro").Index)
          frmAddTransaccion.scorrelativo = GridEX1.Value(GridEX1.Columns("Num_Corre").Index)
                    
          'modificado 07.04.2010 - afernandez
          frmAddTransaccion.txtCod_TipIte = GridEX1.Value(GridEX1.Columns("Cod_Tipite_Asociado").Index)
          frmAddTransaccion.txtDes_TipIte = GridEX1.Value(GridEX1.Columns("Des_TipIte").Index)
          frmAddTransaccion.TxtTipo_Retencion = GridEX1.Value(GridEX1.Columns("Tipo_Detraccion_Asociado").Index)
          frmAddTransaccion.TxtDes_Retencion = GridEX1.Value(GridEX1.Columns("Des_Retencion").Index)
          frmAddTransaccion.TxtPorcentaje_Retencion = GridEX1.Value(GridEX1.Columns("Porcentaje").Index)
          
          Dim C As String
          C = GridEX1.Value(GridEX1.Columns("Flg_Pagar_Detraccion").Index)
          
          If C = "N" Then
          frmAddTransaccion.chkDetraccion.Value = 0
          Else
          frmAddTransaccion.chkDetraccion.Value = 1
          End If
          
          frmAddTransaccion.Show vbModal
          
          Set frmAddTransaccion = Nothing
          Call Buscar
          
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
Dim SDATO As String
Dim strSQL As String
Dim sEmpresa As String

    strSQL = "SELECT DES_EMPRESA FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA='" & vemp & "'"
    sEmpresa = DevuelveCampo(strSQL, cCONNECT)

If ChkPorTransmitir.Value = 1 Then
    SDATO = "S"
Else
    SDATO = "N"
End If

    Ruta = vRuta & "\RPTTransFactVentas.XLT"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "reporte", txtAno.Text, txtMes, SDATO, cCONNECT, sEmpresa
    
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub


Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Msg As Variant
    Select Case ActionName
    Case "BUSCAR"
      Buscar
    End Select
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtMes.SetFocus
    End If
End Sub

Private Sub txtMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FunctButt2.SetFocus
    End If

End Sub

Sub Buscar()

On Error GoTo dprDepurar

Dim sSQL As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle
Dim SDATO As String

If ChkPorTransmitir.Value = 1 Then
    SDATO = "S"
Else
    SDATO = "N"
End If

sSQL = "CN_MUESTRA_FACTURAS_VENTAS_INKAD  '" & Trim(txtAno.Text) & "','" & Trim(txtMes.Text) & "','" & Trim(SDATO) & "'"

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)

GridEX1.Columns("Factura").Width = 1500
GridEX1.Columns("Factura").Caption = "Factura"
GridEX1.Columns("Fec_Registro").Width = 1000
GridEX1.Columns("Fec_Registro").Caption = "Fec.Registro"
GridEX1.Columns("Fec_Emision").Width = 1000
GridEX1.Columns("Fec_Emision").Caption = "Fec.Emision"
GridEX1.Columns("Imp_Neto").Width = 1000
GridEX1.Columns("Imp_Neto").Caption = "Imp.Neto"
GridEX1.Columns("Imp_Total").Width = 1000
GridEX1.Columns("Imp_Total").Caption = "Imp.Total"
GridEX1.Columns("Guias").Width = 1000
GridEX1.Columns("Guias").Caption = "Guias"
GridEX1.Columns("Pedidos").Width = 1000
GridEX1.Columns("Pedidos").Caption = "Pedidos"
GridEX1.Columns("Num_Corre").Width = 1500
GridEX1.Columns("Num_Corre").Caption = "Num.Corre"
GridEX1.Columns("Num_Corre_Docum_Relacionado").Caption = "Num.Corre.Docum.Relacionado"
GridEX1.Columns("Num_Corre_Docum_Relacionado").Width = 1500


'GridEX1.ContinuousScroll = True

Exit Sub

dprDepurar:

errores err.Number
  
End Sub
