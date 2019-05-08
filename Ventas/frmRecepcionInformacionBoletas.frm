VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmRecepcionInformacionBoletas 
   Caption         =   "Recepción de Información - Boletas"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin GridEX20.GridEX GridEX1 
      Height          =   3885
      Left            =   210
      TabIndex        =   0
      Top             =   270
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   6853
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmRecepcionInformacionBoletas.frx":0000
      Column(2)       =   "frmRecepcionInformacionBoletas.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmRecepcionInformacionBoletas.frx":016C
      FormatStyle(2)  =   "frmRecepcionInformacionBoletas.frx":02A4
      FormatStyle(3)  =   "frmRecepcionInformacionBoletas.frx":0354
      FormatStyle(4)  =   "frmRecepcionInformacionBoletas.frx":0408
      FormatStyle(5)  =   "frmRecepcionInformacionBoletas.frx":04E0
      FormatStyle(6)  =   "frmRecepcionInformacionBoletas.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmRecepcionInformacionBoletas.frx":0678
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2160
      TabIndex        =   1
      Top             =   4350
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   900
      Custom          =   $"frmRecepcionInformacionBoletas.frx":0850
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmRecepcionInformacionBoletas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Dim StrSql As String

 

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName

Case "RECEPCIONARINFO"
    If MsgBox("Esta seguro de realizar la recepción de datos", vbInformation + vbYesNo, "AVISO") = vbYes Then
         Call RecepcionarInformacion
    End If
Case "PROCESARTRANSMICION"
    If GridEX1.RowCount = 0 Then Exit Sub
    If MsgBox("Esta seguro de procesar los datos recibidos", vbInformation + vbYesNo, "AVISO") = vbYes Then
         Call ProcesarInformacion
    End If
Case "SALIR"
  Unload Me
End Select

End Sub


Sub ProcesarInformacion()
On Error GoTo ProcesarInformacion

     StrSql = " EXEC Ventas_PROCESA_RECEPCION_Boletas_Saldos_TIENDAS"
     
     If ExecuteSQL(cCONNECT, StrSql) = -1 Then
        MsgBox "Los datos fueron procesados con éxito", vbInformation, "Información"
     
     End If
     Exit Sub
     
ProcesarInformacion:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
 

Sub RecepcionarInformacion()
On Error GoTo RecepcionarInformacion

     StrSql = " EXEC VT_CARGA_INFORMACION_TIENDA_TRANSMISION_VENTAS"
     
     If ExecuteSQL(cCONNECT, StrSql) = -1 Then
        MsgBox "La información fue recepcionada con éxito", vbInformation, "Información"
        Call mostrar
     End If
     Exit Sub
     
RecepcionarInformacion:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
 



Sub mostrar()

On Error GoTo Fin

StrSql = " exec VT_MUESTRA_VENTAS_RECEPCION_TIENDA "
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(StrSql, cCONNECT)

GridEX1.Columns("Tipo Anexo").Width = 900
GridEX1.Columns("Anexo").Width = 1000
GridEX1.Columns("Fecha Emisión").Width = 1200
    
   


Exit Sub
Fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub




