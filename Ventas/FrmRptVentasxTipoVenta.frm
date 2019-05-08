VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmRptVentasxTipoVenta 
   Caption         =   "Emision de Venta por Tipo Venta"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7455
      Begin VB.TextBox txtDesTipoVenta 
         Height          =   285
         Left            =   2250
         TabIndex        =   9
         Top             =   195
         Width           =   4455
      End
      Begin VB.TextBox txtCodTipoVenta 
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   195
         Width           =   1035
      End
      Begin VB.CheckBox chkresumido 
         Caption         =   "Resumido por Item"
         Height          =   255
         Left            =   5400
         TabIndex        =   2
         Top             =   675
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   675
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   94109697
         CurrentDate     =   38590
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   285
         Left            =   3705
         TabIndex        =   4
         Top             =   675
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   94109697
         CurrentDate     =   38590
      End
      Begin VB.CheckBox chkResumidoCliente 
         Caption         =   "Resumido por Cliente"
         Height          =   255
         Left            =   5400
         TabIndex        =   10
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Venta :"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         Height          =   195
         Left            =   3120
         TabIndex        =   5
         Top             =   720
         Width           =   510
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   570
      Left            =   2520
      TabIndex        =   0
      Top             =   1680
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   1005
      Custom          =   $"FrmRptVentasxTipoVenta.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1200
      ControlHeigth   =   550
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   0
      Top             =   1440
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmRptVentasxTipoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public codigo, Descripcion As String, strCod_Anxo As String
Dim strSQL As String





Private Sub Form_Load()
  dtpFecIni = Date
  dtpFecFin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
If chkresumido.Value = Checked Or chkResumidoCliente.Value = Checked Then

If chkresumido.Value = Checked Then
Call Reporte2
End If

If chkResumidoCliente.Value = Checked Then
Call Reporte3

End If

Else

Call Reporte

End If

 

Case "SALIR"
    Unload Me
End Select
End Sub

Sub Reporte()

On Error GoTo ERROR
Dim sSQL As String, RS As Object
Set RS = CreateObject("ADODB.Recordset")
Dim oo As Object
Dim Ruta As String


sSQL = "Gerencia_Muestra_Detalle_Ventas_por_Tipo_Venta '" & txtCodTipoVenta.Text & "','" & dtpFecIni.Value & "','" & dtpFecFin.Value & "' ,'D'"
Set RS = CargarRecordSetDesconectado(sSQL, cCONNECT)

If RS.BOF Then
  MsgBox "No hay Registro ha Imprinmir", vbCritical + vbInformation, "AVISO"
  Exit Sub
End If

If MsgBox("Imprimir usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir") = vbYes Then
    Ruta = vRuta & "\RptDetalleVentasXTipoVenta.XLT"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.displayalerts = False
            
    oo.Run "Reporte", RS, dtpFecIni & " - " & dtpFecFin & "          Tipo Venta: " & txtCodTipoVenta.Text & "-" & txtDesTipoVenta.Text, "D"
Else
    Ruta = vRuta & "\RptDetalleVentasXTipoVenta.OTS"
    Set oo = CreateObject("ooBusiness.Calc")
    oo.OfficeTemplateSheet = Ruta
    oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
    oo.MacroLibraryName = "Library1"
    oo.MacroModuleName = "Module1"
    oo.MacroName = "Reporte"
    
    oo.Run RS.Source, dtpFecIni & " - " & dtpFecFin & "          Tipo Venta: " & _
            txtCodTipoVenta.Text & "-" & txtDesTipoVenta.Text, "D", cCONNECT
End If
Set oo = Nothing

Exit Sub
ERROR:
    errores err.Number

End Sub


Sub Reporte2()

On Error GoTo ERROR
Dim sSQL As String, RS As Object
Set RS = CreateObject("ADODB.Recordset")
Dim oo As Object
Dim Ruta As String

sSQL = "Gerencia_Muestra_Detalle_Ventas_por_Tipo_Venta '" & txtCodTipoVenta.Text & "','" & dtpFecIni.Value & "','" & dtpFecFin.Value & "' ,'R'"
Set RS = CargarRecordSetDesconectado(sSQL, cCONNECT)

If RS.BOF Then
    MsgBox "No hay Registro ha Imprinmir", vbCritical + vbInformation, "AVISO"
    Exit Sub
End If

If MsgBox("Imprimir usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir") = vbYes Then
    Ruta = vRuta & "\RptDetalleVentasXTipoVentaCliente.XLT"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.displayalerts = False
            
    oo.Run "Reporte", RS, dtpFecIni & " - " & dtpFecFin & "          Tipo Venta: " & txtCodTipoVenta.Text & "-" & txtDesTipoVenta.Text, "R"
Else
    Ruta = vRuta & "\RptDetalleVentasXTipoVenta.OTS"
    Set oo = CreateObject("ooBusiness.Calc")
    oo.OfficeTemplateSheet = Ruta
    oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
    oo.MacroLibraryName = "Library1"
    oo.MacroModuleName = "Module1"
    oo.MacroName = "Reporte"
    
    oo.Run RS.Source, dtpFecIni & " - " & dtpFecFin & "          Tipo Venta: " & _
           txtCodTipoVenta.Text & "-" & txtDesTipoVenta.Text, "R", cCONNECT
End If
Set oo = Nothing

Exit Sub
ERROR:
    errores err.Number

End Sub


Sub Reporte3()

On Error GoTo ERROR
Dim sSQL As String, RS As Object
Set RS = CreateObject("ADODB.Recordset")
Dim oo As Object
Dim Ruta As String

sSQL = "Gerencia_Muestra_Detalle_Ventas_por_Tipo_Venta '" & txtCodTipoVenta.Text & "','" & dtpFecIni.Value & "','" & dtpFecFin.Value & "' ,'C'"
Set RS = CargarRecordSetDesconectado(sSQL, cCONNECT)

If RS.BOF Then
    MsgBox "No hay Registro ha Imprinmir", vbCritical + vbInformation, "AVISO"
    Exit Sub
End If

If MsgBox("Imprimir usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir") = vbYes Then
    Ruta = vRuta & "\RptDetalleVentasXTipoVenta.XLT"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.displayalerts = False
            
    oo.Run "Reporte", RS, dtpFecIni & " - " & dtpFecFin & "          Tipo Venta: " & txtCodTipoVenta.Text & "-" & txtDesTipoVenta.Text, "R"
Else
    Ruta = vRuta & "\RptDetalleVentasXTipoVenta.OTS"
    Set oo = CreateObject("ooBusiness.Calc")
    oo.OfficeTemplateSheet = Ruta
    oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
    oo.MacroLibraryName = "Library1"
    oo.MacroModuleName = "Module1"
    oo.MacroName = "Reporte"
    
    oo.Run RS.Source, dtpFecIni & " - " & dtpFecFin & "          Tipo Venta: " & _
           txtCodTipoVenta.Text & "-" & txtDesTipoVenta.Text, "R", cCONNECT
End If
Set oo = Nothing

Exit Sub
ERROR:
    errores err.Number

End Sub




Private Sub txtCodTipoVenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Busca_Opcion("Cod_Tipo_Venta", "Descripcion", "Cn_Tipos_Venta where ", txtCodTipoVenta, txtDesTipoVenta, 1, Me)
End Sub

Private Sub txtDesTipoVenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Busca_Opcion("Cod_Tipo_Venta", "Descripcion", "Cn_Tipos_Venta where ", txtCodTipoVenta, txtDesTipoVenta, 2, Me)
End Sub

 

