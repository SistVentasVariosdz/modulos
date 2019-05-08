VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmConceptoCobranza 
   Caption         =   "Registro de Tipo Cobranza"
   ClientHeight    =   6345
   ClientLeft      =   285
   ClientTop       =   720
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin GridEX20.GridEX GridEX1 
      Height          =   5340
      Left            =   0
      TabIndex        =   0
      Top             =   75
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   9419
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
      Column(1)       =   "frmConceptoCobranza.frx":0000
      Column(2)       =   "frmConceptoCobranza.frx":00C8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmConceptoCobranza.frx":016C
      FormatStyle(2)  =   "frmConceptoCobranza.frx":02A4
      FormatStyle(3)  =   "frmConceptoCobranza.frx":0354
      FormatStyle(4)  =   "frmConceptoCobranza.frx":0408
      FormatStyle(5)  =   "frmConceptoCobranza.frx":04E0
      FormatStyle(6)  =   "frmConceptoCobranza.frx":0598
      FormatStyle(7)  =   "frmConceptoCobranza.frx":0678
      FormatStyle(8)  =   "frmConceptoCobranza.frx":0724
      ImageCount      =   0
      PrinterProperties=   "frmConceptoCobranza.frx":07D4
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   675
      Left            =   2475
      TabIndex        =   1
      Top             =   5535
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1191
      Custom          =   $"frmConceptoCobranza.frx":09AC
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
Attribute VB_Name = "frmConceptoCobranza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public codigo As String, Descripcion As String
Public sTipoBusq As String


Private Sub Form_Load()
BUSCAR
End Sub

Sub BUSCAR()

Dim strSQL
On Error GoTo errores
         Dim sano4 As String
             sano4 = DevuelveCampo("select Ultimo_Ano_Cerrado from Cn_Control_Ventas", cCONNECT)
             
strSQL = "Ventas_Mantenimiento_ConceptoCobranza ' ','','B','','','','" & sano4 & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

Dim colTemp As JSColumn

GridEX1.ColumnHeaderHeight = 500

GridEX1.Columns("Cod_Concepto_Cobranza").Width = 1000
GridEX1.Columns("Cod_Concepto_Cobranza").Caption = "Cod. Concepto Cobranza"
GridEX1.Columns("Descripcion").Width = 2700
GridEX1.Columns("Descripcion").Caption = "Descripcion"
GridEX1.Columns("Flg_Debe_Haber").Width = 750
GridEX1.Columns("Flg_Debe_Haber").Caption = "Flg. Debe Haber"
GridEX1.Columns("Cod_CtaCont").Width = 1200
GridEX1.Columns("Cod_CtaCont").Caption = "Cod. CtaCont"
GridEX1.Columns("descrip").Width = 2800
GridEX1.Columns("descrip").Caption = "Desc. CtaCont"
GridEX1.Columns("Flg_Asociado_Factura").Width = 750
GridEX1.Columns("Flg_Asociado_Factura").Caption = "Flg. Asociado Factura"
GridEX1.Columns("Flg_Seleccionable").Width = 750
GridEX1.Columns("Flg_Seleccionable").Caption = "Flg. Seleccionable"
GridEX1.Columns("Flg_Asociado_Docum_Pago").Width = 750
GridEX1.Columns("Flg_Asociado_Docum_Pago").Caption = "Flg. Asociado Docum. Pago"
GridEX1.Columns("Flg_Anticipo").Width = 750
GridEX1.Columns("Flg_Anticipo").Caption = "Flg. Anticipo"
GridEX1.Columns("Flg_automatico").Width = 750
GridEX1.Columns("Flg_automatico").Caption = "Flg. Automatico"
GridEX1.Columns("Flg_Resumen").Width = 750
GridEX1.Columns("Flg_Resumen").Caption = "Flg. Resumen"
GridEX1.Columns("Flg_CanjeServicios").Width = 750
GridEX1.Columns("Flg_CanjeServicios").Caption = "Flg. CanjeServicios"
GridEX1.Columns("Flg_NotaAbono").Width = 750
GridEX1.Columns("Flg_NotaAbono").Caption = "Flg. Nota Abono"
GridEX1.Columns("Cod_Concepto_Finanzas").Width = 750
GridEX1.Columns("Cod_Concepto_Finanzas").Caption = "Flg. Concepto Finanzas"
GridEX1.Columns("Flg_Canje_Servicios_Tesoreria").Width = 750
GridEX1.Columns("Flg_Canje_Servicios_Tesoreria").Caption = "Flg. Canje Servicios Tesoreria"
GridEX1.Columns("Flg_Financiamientos").Width = 750
GridEX1.Columns("Flg_Financiamientos").Caption = "Tip. Financiamientos"
GridEX1.Columns("Tip_Docum_Otros").Width = 750
GridEX1.Columns("Tip_Docum_Otros").Caption = "Flg. Docum. Otros"
GridEX1.Columns("Flg_Diversos").Width = 750
GridEX1.Columns("Flg_Diversos").Caption = "Flg. Diversos"

Exit Sub
Resume
errores:
    errores err.Number
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo hand

Select Case ActionName
        Case Is = "AGREGAR"
               With frmConceptosCobranzaADD
                  .StrOption = "I"
                  .sCod_Tipcobranza = LPad(DevuelveCampo("Select max(Cod_Concepto_Cobranza)+1 from cn_ventas_conceptos_cobranza ", cCONNECT), 3, "0")
                  .txtcodigo.Text = LPad(DevuelveCampo("SELECT max(Cod_Concepto_Cobranza)  +1 FROM   cn_ventas_conceptos_cobranza   where (Cod_Concepto_Cobranza <> '125' and Cod_Concepto_Cobranza <> '800' and Cod_Concepto_Cobranza <> '801' and Cod_Concepto_Cobranza <> '998' and Cod_Concepto_Cobranza <> '999')", cCONNECT), 3, "0")
                  .txtcodigo.Enabled = False
                  .Show 1
                    BUSCAR
                End With
    
        Case Is = "MODIFICAR"
         If GridEX1.RowCount = 0 Then Exit Sub
         Dim sano1 As String
             sano1 = DevuelveCampo("select Ultimo_Ano_Cerrado from Cn_Control_Ventas", cCONNECT)

         With frmConceptosCobranzaADD
           .txtcodigo.Text = GridEX1.Value(GridEX1.Columns("Cod_Concepto_Cobranza").Index)
           .txtcodigo.Enabled = False
           .txtDescripcion.Text = GridEX1.Value(GridEX1.Columns("Descripcion").Index)
           .txtCod_Cobranza.Text = GridEX1.Value(GridEX1.Columns("Cod_CtaCont").Index)
           .txtDes_Cobranza.Text = DevuelveCampo("Select Des_CtaCont from CN_PLANCONTABLE where Cod_CtaCont ='" & GridEX1.Value(GridEX1.Columns("Cod_CtaCont").Index) & "'  and  ano = '" & sano1 & "'", cCONNECT)
           .txtCod_Cobranza1.Text = GridEX1.Value(GridEX1.Columns("Cod_Concepto_Finanzas").Index)
           .txtDes_Cobranza1.Text = DevuelveCampo("Select Des_Concepto_Finanzas from FI_CONCEPTOS where Cod_Concepto_Finanzas ='" & GridEX1.Value(GridEX1.Columns("Cod_Concepto_Finanzas").Index) & "'", cCONNECT)
           If GridEX1.Value(GridEX1.Columns("Flg_Debe_Haber").Index) = "D" Then
                .opt1.Value = True
                .opt2.Value = False
           Else
                .opt2.Value = True
                .opt1.Value = False
           End If
           
           .StrOption = "U"
           varSecuencia = GridEX1.Value(GridEX1.Columns("Cod_Concepto_Cobranza").Index)
           frmConceptosCobranzaADD.Show 1
           BUSCAR
           Call GridEX1.Find(GridEX1.Columns("Cod_Concepto_Cobranza").Index, jgexEqual, varSecuencia)
         End With
        Case Is = "ELIMINAR"
             If GridEX1.RowCount = 0 Then Exit Sub
    
                
                If MsgBox("Esta seguro de Eliminar este Concepto de Cobranza", vbYesNo, "IMPORTANTE") = vbYes Then
                  lvSql = "Ventas_Mantenimiento_ConceptoCobranza '" & GridEX1.Value(GridEX1.Columns("Cod_Concepto_Cobranza").Index) & "' ,'','D','','','',''"
                  Call ExecuteCommandSQL(cCONNECT, lvSql)
                  BUSCAR
                End If
                     
        Case Is = "SALIR"
           Unload Me
End Select

Exit Sub
Resume
hand:

errores err.Number

End Sub

