VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmTiposCobranza 
   Caption         =   "Registro de Transacciones de Cobranzas"
   ClientHeight    =   6930
   ClientLeft      =   285
   ClientTop       =   720
   ClientWidth     =   13680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   13680
   StartUpPosition =   2  'CenterScreen
   Begin GridEX20.GridEX GridEX1 
      Height          =   6060
      Left            =   0
      TabIndex        =   0
      Top             =   75
      Width           =   13680
      _ExtentX        =   24130
      _ExtentY        =   10689
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
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
      Column(1)       =   "frmTiposCobranza.frx":0000
      Column(2)       =   "frmTiposCobranza.frx":00C8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmTiposCobranza.frx":016C
      FormatStyle(2)  =   "frmTiposCobranza.frx":02A4
      FormatStyle(3)  =   "frmTiposCobranza.frx":0354
      FormatStyle(4)  =   "frmTiposCobranza.frx":0408
      FormatStyle(5)  =   "frmTiposCobranza.frx":04E0
      FormatStyle(6)  =   "frmTiposCobranza.frx":0598
      FormatStyle(7)  =   "frmTiposCobranza.frx":0678
      FormatStyle(8)  =   "frmTiposCobranza.frx":0724
      ImageCount      =   0
      PrinterProperties=   "frmTiposCobranza.frx":07D4
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   675
      Left            =   4920
      TabIndex        =   1
      Top             =   6240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1191
      Custom          =   $"frmTiposCobranza.frx":09AC
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1075
      ControlHeigth   =   650
      ControlSeparator=   75
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   3840
      Top             =   240
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmTiposCobranza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public codigo As String, Descripcion As String
Public sTipoBusq As String


Private Sub Form_Load()
Buscar
End Sub

Sub Buscar()

Dim strSQL
On Error GoTo errores

strSQL = "Ventas_Mantenimiento_TipoCobranza ' ','','B','',''"
Set gridex1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

Dim colTemp As JSColumn

gridex1.ColumnHeaderHeight = 500

gridex1.Columns("Cod_Tipcobranza").Width = 750
gridex1.Columns("Cod_Tipcobranza").Caption = "Cod. Cobranza"
gridex1.Columns("Descripcion").Width = 3500
gridex1.Columns("Descripcion").Caption = "Descripcion "
gridex1.Columns("Flg_Requiere_Banco").Width = 750
gridex1.Columns("Flg_Requiere_Banco").Caption = "Flg. Req. Banco"
gridex1.Columns("Flg_Cobranza_Simple").Width = 750
gridex1.Columns("Flg_Cobranza_Simple").Caption = "Flg. Cobranza Simple"
gridex1.Columns("Cod_Concepto_Cobranza_Debe").Width = 750
gridex1.Columns("Cod_Concepto_Cobranza_Debe").Caption = "Cod. Concepto Cobranza Debe"
gridex1.Columns("DES1").Width = 3500
gridex1.Columns("DES1").Caption = "Desc. Concepto Cobranza Debe"
gridex1.Columns("Cod_Concepto_Cobranza_Haber").Width = 750
gridex1.Columns("Cod_Concepto_Cobranza_Haber").Caption = "Cod. Concepto Cobranza Haber"
gridex1.Columns("DES2").Width = 2000
gridex1.Columns("DES2").Caption = "Desc. Concepto Cobranza Haber"
gridex1.Columns("Flg_Ventas_Tienda").Width = 750
gridex1.Columns("Flg_Ventas_Tienda").Caption = "Flg. Ventas Tienda"
gridex1.Columns("Flg_Cobranza_Multiple").Width = 750
gridex1.Columns("Flg_Cobranza_Multiple").Caption = "Flg. Cobranza Multiple"
gridex1.Columns("Flg_Efectivo").Width = 750
gridex1.Columns("Flg_Efectivo").Caption = "Flg. Efectivo"
gridex1.Columns("Flg_Seleccionable").Width = 750
gridex1.Columns("Flg_Seleccionable").Caption = "Flg. Seleccionable"
gridex1.Columns("Flg_NotaAbono").Width = 750
gridex1.Columns("Flg_NotaAbono").Caption = "Flg. NotaAbono"
gridex1.Columns("Flg_Canje").Width = 750
gridex1.Columns("Flg_Canje").Caption = "Flg. Canje"
gridex1.Columns("Flg_Parte_Cobranza").Width = 750
gridex1.Columns("Flg_Parte_Cobranza").Caption = "Flg. Parte Cobranza"
gridex1.Columns("Flg_Anticipos").Width = 750
gridex1.Columns("Flg_Anticipos").Caption = "Flg. Anticipos"
gridex1.Columns("Tip_StoredProc").Width = 750
gridex1.Columns("Tip_StoredProc").Caption = "Tip. StoredProc"
gridex1.Columns("Flg_Descuento_Letras_Banco").Width = 750
gridex1.Columns("Flg_Descuento_Letras_Banco").Caption = "Flg. Desc. Letras Banco"
gridex1.Columns("Cod_Transaccion").Width = 750
gridex1.Columns("Cod_Transaccion").Caption = "Cod. Transaccion"
gridex1.Columns("Flg_CanjeTesoreria").Width = 750
gridex1.Columns("Flg_CanjeTesoreria").Caption = "Flg. CanjeTesoreria"
gridex1.Columns("Flg_Genera_Concepto_Automatico").Width = 750
gridex1.Columns("Flg_Genera_Concepto_Automatico").Caption = "Flg. Genera Concepto Automatico"
gridex1.Columns("Flg_Cobranza_Dudosa").Width = 750
gridex1.Columns("Flg_Cobranza_Dudosa").Caption = "Flg. Cobranza Dudosa"

Exit Sub
Resume
errores:
    errores err.Number
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo hand

Select Case ActionName
        Case Is = "AGREGAR"
               With frmTiposCobranzaAdd
                  .strOption = "I"
                  .sCod_Tipcobranza = LPad(DevuelveCampo("Select max(Cod_Tipcobranza)+1 from cn_ventas_tipos_cobranza ", cCONNECT), 3, "0")
                  .txtcodigo.Text = LPad(DevuelveCampo("Select max(Cod_Tipcobranza)+1 from cn_ventas_tipos_cobranza ", cCONNECT), 3, "0")
                  .txtcodigo.Enabled = False
                  .Show 1
                    Buscar
                End With
    
        Case Is = "MODIFICAR"
         If gridex1.RowCount = 0 Then Exit Sub
         With frmTiposCobranzaAdd
           .txtcodigo.Text = gridex1.Value(gridex1.Columns("Cod_Tipcobranza").Index)
           .txtcodigo.Enabled = False
           .txtDescripcion.Text = gridex1.Value(gridex1.Columns("Descripcion").Index)
           .txtCod_Cobranza.Text = gridex1.Value(gridex1.Columns("Cod_Concepto_Cobranza_Debe").Index)
           .txtDes_Cobranza.Text = DevuelveCampo("Select Descripcion from cn_ventas_conceptos_cobranza where Cod_Concepto_Cobranza ='" & gridex1.Value(gridex1.Columns("Cod_Concepto_Cobranza_Debe").Index) & "'", cCONNECT)
           .txtCod_Cobranza1.Text = gridex1.Value(gridex1.Columns("Cod_Concepto_Cobranza_Haber").Index)
           .txtDes_Cobranza1.Text = DevuelveCampo("Select Descripcion from cn_ventas_conceptos_cobranza where Cod_Concepto_Cobranza ='" & gridex1.Value(gridex1.Columns("Cod_Concepto_Cobranza_Haber").Index) & "'", cCONNECT)
           .strOption = "U"
           varSecuencia = gridex1.Value(gridex1.Columns("Cod_Tipcobranza").Index)
           frmTiposCobranzaAdd.Show 1
           Buscar
           Call gridex1.Find(gridex1.Columns("Cod_Tipcobranza").Index, jgexEqual, varSecuencia)
         End With
        Case Is = "ELIMINAR"
             If gridex1.RowCount = 0 Then Exit Sub
    
                
                If MsgBox("Esta seguro de Eliminar este Tipo de Cobranza", vbYesNo, "IMPORTANTE") = vbYes Then
                  lvSql = "Ventas_Mantenimiento_TipoCobranza '" & gridex1.Value(gridex1.Columns("Cod_Tipcobranza").Index) & "' ,'','D','',''"
                  Call ExecuteCommandSQL(cCONNECT, lvSql)
                  Buscar
                End If
                     
        Case Is = "SALIR"
           Unload Me
End Select

Exit Sub
Resume
hand:

errores err.Number

End Sub

