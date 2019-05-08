VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmTelasEnAlm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   14055
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   5820
      TabIndex        =   1
      Top             =   5430
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "0~0~IMPRIMIR~True~True~&Imprimir~0~0~1~~0~False~False~&Imprimir~~1~0~SALIR~True~True~&Salir~1~0~3~~0~False~False~&Salir~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX gexTel 
      Height          =   4455
      Left            =   90
      TabIndex        =   0
      Top             =   870
      Width           =   13890
      _ExtentX        =   24500
      _ExtentY        =   7858
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ContScroll      =   -1  'True
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmTelasEnAlm.frx":0000
      Column(2)       =   "frmTelasEnAlm.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmTelasEnAlm.frx":016C
      FormatStyle(2)  =   "frmTelasEnAlm.frx":02A4
      FormatStyle(3)  =   "frmTelasEnAlm.frx":0354
      FormatStyle(4)  =   "frmTelasEnAlm.frx":0408
      FormatStyle(5)  =   "frmTelasEnAlm.frx":04E0
      FormatStyle(6)  =   "frmTelasEnAlm.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmTelasEnAlm.frx":0678
   End
   Begin VB.Label lblNom_Cliente 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1200
      TabIndex        =   5
      Top             =   90
      Width           =   6075
   End
   Begin VB.Label Label2 
      Caption         =   "Cliente"
      Height          =   195
      Left            =   195
      TabIndex        =   4
      Top             =   150
      Width           =   945
   End
   Begin VB.Label lblDes_OrdPro 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   6075
   End
   Begin VB.Label Label1 
      Caption         =   "Descripcion"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   510
      Width           =   945
   End
End
Attribute VB_Name = "frmTelasEnAlm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Fabrica As String, sCod_OrdPro As String
Dim strSQL As String, sTit As String, sErr As String, rstAux As ADODB.Recordset

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "IMPRIMIR"
        Reporte
    Case "SALIR"
        Unload Me
    End Select
End Sub

Public Sub MostrarTelasEnAlm()
On Error GoTo ErrTel
Dim fmtAux As JSFmtCondition
    
    strSQL = "EXEC ES_OBTIENE_DATOS_OP '" & sCod_Fabrica & "', '" & sCod_OrdPro & "'"
    
    Set rstAux = CargarRecordSetDesconectado(strSQL, cConnect)
    lblNom_Cliente = ""
    lblDes_OrdPro = ""
    If rstAux.RecordCount > 0 Then
        lblNom_Cliente = rstAux!Nom_Cliente
        lblDes_OrdPro = rstAux!Des_EstPro
    End If
    rstAux.Close
    Set rstAux = Nothing
    
    strSQL = "EXEC TX_TELAS_EN_ALMACEN_POR_ORDEN '" & sCod_Fabrica & _
             "', '" & sCod_OrdPro & "'"
    Me.Caption = "Telas en Almacen O/P : " & sCod_OrdPro
    Screen.MousePointer = 11
    Set gexTel.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    
    Set fmtAux = gexTel.FmtConditions.Add(gexTel.Columns("SubTipo").Index, jgexEqual, 2)
    fmtAux.FormatStyle.BackColor = &HC0FFFF
    
    Set fmtAux = gexTel.FmtConditions.Add(gexTel.Columns("Tipo").Index, jgexEqual, 3)
    fmtAux.FormatStyle.BackColor = &HE0E0E0
    
    gexTel.Columns("Tela").Width = 4380
    gexTel.Columns("Comb").Width = 900
    gexTel.Columns("Color").Width = 1290
    gexTel.Columns("Proveedor").Width = 1680
    gexTel.Columns("Partida").Width = 975
    gexTel.Columns("Stock").Width = 930
    gexTel.Columns("Total_Partida").Width = 1080
    gexTel.Columns("Total_Requerimiento").Width = 1605
    gexTel.Columns("Porcentaje").Width = 900
    gexTel.Columns("Tipo").Width = 240
    gexTel.Columns("SubTipo").Width = 285
    
    gexTel.Columns("Tela").Caption = "Tela"
    gexTel.Columns("Comb").Caption = "Combinación"
    gexTel.Columns("Color").Caption = "Color"
    gexTel.Columns("Proveedor").Caption = "Proveedor"
    gexTel.Columns("Partida").Caption = "Partida"
    gexTel.Columns("Stock").Caption = "Stock"
    gexTel.Columns("Total_Partida").Caption = "Tot.Partida"
    gexTel.Columns("Total_Requerimiento").Caption = "Tot.Requer."
    gexTel.Columns("Porcentaje").Caption = "Porc."
    gexTel.Columns("Tipo").Visible = False
    gexTel.Columns("SubTipo").Visible = False
    
    gexTel.Columns("Stock").Format = "#,##0.00"
    gexTel.Columns("Total_Partida").Format = "#,##0.00"
    gexTel.Columns("Total_Requerimiento").Format = "#,##0.00"
    gexTel.Columns("Porcentaje").Format = "#,##0.00"
    
    Screen.MousePointer = 0
Exit Sub
ErrTel:
    sErr = Err.Description
    Screen.MousePointer = 0
    MsgBox sErr, vbCritical + vbOKOnly, sTit
End Sub

Public Sub Reporte()
On Error GoTo ErrorImpresion
    Dim oo As Object
    strSQL = "select ruta_logo from seguridad..seg_empresas where cod_Empresa='" & vemp1 & "'"
    
    Screen.MousePointer = 11
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\TelasEnAlm.xlt"
    'oo.diplaywarnings
    oo.Visible = True
    
    oo.Run "REPORTE", gexTel.ADORecordset, lblNom_Cliente, sCod_Fabrica, _
           sCod_OrdPro, lblDes_OrdPro, DevuelveCampo(strSQL, cConnect), cConnect
    Screen.MousePointer = vbNormal
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    Screen.MousePointer = vbNormal
    MsgBox "Hubo error en la impresion del Reporte  " & Err.Description, vbCritical, "Impresion"
End Sub

