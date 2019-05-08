VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmRptVentasxCliente 
   Caption         =   "Emision de Venta por Cliente"
   ClientHeight    =   2640
   ClientLeft      =   660
   ClientTop       =   1545
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7455
      Begin VB.TextBox txtCod_TipDoc 
         Height          =   285
         Left            =   960
         MaxLength       =   4
         TabIndex        =   5
         Top             =   1200
         Width           =   480
      End
      Begin VB.TextBox txtDes_TipDoc 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   1200
         Width           =   1905
      End
      Begin VB.CheckBox chkresumido 
         Caption         =   "Resumido"
         Height          =   255
         Left            =   5400
         TabIndex        =   4
         Top             =   675
         Width           =   1935
      End
      Begin VB.TextBox txtDes_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         MaxLength       =   30
         TabIndex        =   0
         Top             =   240
         Width           =   3945
      End
      Begin VB.TextBox txtDes_TipAnex 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4200
         MaxLength       =   4
         TabIndex        =   10
         Text            =   "C"
         Top             =   240
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.TextBox txtNum_Ruc 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5835
         MaxLength       =   11
         TabIndex        =   1
         Top             =   240
         Width           =   1545
      End
      Begin VB.TextBox txtCod_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4800
         MaxLength       =   4
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "C"
         Top             =   240
         Width           =   360
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   285
         Left            =   840
         TabIndex        =   2
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
         TabIndex        =   3
         Top             =   675
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   94109697
         CurrentDate     =   38590
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Doc :"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1215
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         Height          =   195
         Left            =   3120
         TabIndex        =   14
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   285
         Width           =   570
      End
      Begin VB.Label Label28 
         Caption         =   "R.U.C."
         Height          =   255
         Left            =   5280
         TabIndex        =   11
         Top             =   255
         Width           =   495
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   570
      Left            =   1800
      TabIndex        =   7
      Top             =   1920
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   1005
      Custom          =   $"FrmRptVentasxCliente.frx":0000
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
      Top             =   1920
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmRptVentasxCliente"
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
If chkresumido.Value = 1 Then
    Call Reporte
Else
    Call Reporte2
End If
Case "IMPRIDNI"
    Call Reporte_DNI
Case "SALIR"
    Unload Me
End Select
End Sub

Sub Reporte_DNI()
On Error GoTo Fail
Dim strSQL As String
Dim adoRs As Object
Set adoRs = CreateObject("ADODB.Recordset")
Dim sEmpresa As String, Ruta As String
Dim oo As Object
    
    strSQL = "SELECT DES_EMPRESA FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA='" & vemp & "'"
    sEmpresa = DevuelveCampo(strSQL, cCONNECT)
    strSQL = "Gerencia_Muestra_Ventas_por_Cliente_Fecha_Agrupado '" & txtCod_TipAne.Text & "','" & strCod_Anxo & "','" & Trim(txtNum_Ruc.Text) & "','" & dtpFecIni & "','" & dtpFecFin & "','" & Trim(txtCod_TipDoc.Text) & "'"
    Set adoRs = CargarRecordSetDesconectado(strSQL, cCONNECT)
    
    If MsgBox("Imprimir usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir x DNI") = vbYes Then
        Set oo = CreateObject("Excel.Application")
        oo.Workbooks.Open vRuta & "\Rpt_MuesVentClieAgr.xlt"
        oo.Visible = True
        oo.displayalerts = False
        
        oo.Run "Reporte", sEmpresa, adoRs, dtpFecIni & " - " & dtpFecFin, txtDes_TipAne, txtDes_TipDoc
    Else
        Ruta = vRuta & "\Rpt_MuesVentClieAgr.OTS"
        Set oo = CreateObject("ooBusiness.Calc")
        oo.OfficeTemplateSheet = Ruta
        oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
        oo.MacroLibraryName = "Library1"
        oo.MacroModuleName = "Module1"
        oo.MacroName = "Reporte"
        
        oo.Run sEmpresa, adoRs.Source, dtpFecIni & " - " & dtpFecFin, txtDes_TipAne, txtDes_TipDoc, cCONNECT
    End If
    Set oo = Nothing
Exit Sub
Fail:
    MsgBox err.Description
End Sub

Sub Reporte()
On Error GoTo ERROR
Dim sSQL As String, RS As Object
Set RS = CreateObject("ADODB.Recordset")
Dim oo As Object, Rs2 As Object
Set Rs2 = CreateObject("ADODB.Recordset")
Dim Ruta As String

sSQL = "Gerencia_Muestra_Detalle_Ventas_por_Cliente_Fecha_Agrupado '" & txtCod_TipAne & "','" & strCod_Anxo & "','Trim(txtNum_Ruc.Text)','" & dtpFecIni & "','" & dtpFecFin & "','D','" & Trim(txtCod_TipDoc.Text) & "'"
Set RS = CargarRecordSetDesconectado(sSQL, cCONNECT)

sSQL = "Gerencia_Muestra_Detalle_Ventas_por_Cliente_Fecha_Agrupado '" & txtCod_TipAne & "','" & strCod_Anxo & "','Trim(txtNum_Ruc.Text)','" & dtpFecIni & "','" & dtpFecFin & "','R','" & Trim(txtCod_TipDoc.Text) & "'"
Set Rs2 = CargarRecordSetDesconectado(sSQL, cCONNECT)

If RS.BOF Or RS.EOF Then
    MsgBox "No hay Registro ha Imprinmir", vbCritical + vbInformation, "AVISO"
    Exit Sub
End If

If MsgBox("Imprimir usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir") = vbYes Then
    Ruta = vRuta & "\RptDetalleVentas2.XLT"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.displayalerts = False
    
    oo.Run "Reporte", RS, Rs2, dtpFecIni & " - " & dtpFecFin & "          Cliente: " & txtDes_TipAne.Text
Else
    Ruta = vRuta & "\RptDetalleVentas2.OTS"
    Set oo = CreateObject("ooBusiness.Calc")
    oo.OfficeTemplateSheet = Ruta
    oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
    oo.MacroLibraryName = "Library1"
    oo.MacroModuleName = "Module1"
    oo.MacroName = "Reporte"
    
    oo.Run RS.Source, Rs2.Source, dtpFecIni & " - " & dtpFecFin & "          Cliente: " & txtDes_TipAne.Text, cCONNECT
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

If strCod_Anxo = "" Or txtCod_TipAne = "" Then
  MsgBox "Seleccione Un Cliente", vbCritical + vbInformation, "AVISO"
  Exit Sub
End If

sSQL = "Gerencia_Muestra_Detalle_Ventas_por_Cliente_Fecha '" & txtCod_TipAne & "','" & strCod_Anxo & "','','" & dtpFecIni & "','" & dtpFecFin & "'"
Set RS = CargarRecordSetDesconectado(sSQL, cCONNECT)

If RS.BOF Or RS.EOF Then
  MsgBox "No hay Registro ha Imprinmir", vbCritical + vbInformation, "AVISO"
  Exit Sub
End If

If MsgBox("Imprimir usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir") = vbYes Then
    Ruta = vRuta & "\RptDetalleVentas.XLT"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.displayalerts = False
    
    oo.Run "Reporte", RS, txtDes_TipAne, txtNum_Ruc, dtpFecIni & " - " & dtpFecFin
Else
    Ruta = vRuta & "\RptDetalleVentas.OTS"
    Set oo = CreateObject("ooBusiness.Calc")
    oo.OfficeTemplateSheet = Ruta
    oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
    oo.MacroLibraryName = "Library1"
    oo.MacroModuleName = "Module1"
    oo.MacroName = "Reporte"
    
    oo.Run RS.Source, txtDes_TipAne, txtNum_Ruc, dtpFecIni & " - " & dtpFecFin, cCONNECT
End If
Set oo = Nothing

Exit Sub
ERROR:
    errores err.Number

End Sub

Private Sub txtCod_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAne, txtDes_TipAnex, 1, Me)
End Sub

Private Sub txtCod_TipDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call Busca_Opcion("Cod_TipDoc", "Des_TipDoc", "CN_TiposDocum where Flg_Doc_Ventas = '*' and ", txtCod_TipDoc, txtDes_TipDoc, 1, Me)
  Me.FunctButt1.SetFocus
End If
End Sub

Private Sub txtDes_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 2, Me)
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtDes_TipDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_TipDoc", "Des_TipDoc", "CN_TiposDocum where Flg_Doc_Ventas = '*' and ", txtCod_TipDoc, txtDes_TipDoc, 2, Me)
    Me.FunctButt1.SetFocus
End If
End Sub

Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 1, Me)
End Sub

