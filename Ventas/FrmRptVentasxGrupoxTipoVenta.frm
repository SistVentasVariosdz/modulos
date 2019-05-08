VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRptVentasxGrupoxTipoVenta 
   Caption         =   "Emision Reporte Ventas por Grupo "
   ClientHeight    =   3465
   ClientLeft      =   2280
   ClientTop       =   2385
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   6240
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   516
      Left            =   1560
      TabIndex        =   14
      Top             =   2640
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmRptVentasxGrupoxTipoVenta.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   2532
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6132
      Begin VB.TextBox txtNum_Ruc 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         MaxLength       =   11
         TabIndex        =   10
         Top             =   2040
         Width           =   1068
      End
      Begin VB.TextBox txtCod_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "C"
         Top             =   2040
         Width           =   360
      End
      Begin VB.TextBox txtDes_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2640
         TabIndex        =   13
         Top             =   2040
         Width           =   3228
      End
      Begin VB.TextBox txtCod_TipVenta 
         Height          =   285
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "0"
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtDes_TipVenta 
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Top             =   1080
         Width           =   2505
      End
      Begin VB.ComboBox CboOrigen 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   1500
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   56754177
         CurrentDate     =   38590
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   285
         Left            =   1185
         TabIndex        =   1
         Top             =   645
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   56754177
         CurrentDate     =   38590
      End
      Begin VB.Label Label4 
         Caption         =   "Ruc :"
         Height          =   180
         Left            =   240
         TabIndex        =   11
         Tag             =   "Anexo Type"
         Top             =   2088
         Width           =   432
      End
      Begin VB.Label Label13 
         Caption         =   "Tipo Venta :"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1095
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   690
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   285
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Origen : "
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   600
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   120
      Top             =   1080
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmRptVentasxGrupoxTipoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public codigo As String, Descripcion As String, strCod_Anxo As String

Private Sub dtpFecFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtCod_TipVenta.SetFocus
    End If
End Sub

Private Sub dtpFecIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.dtpFecFin.SetFocus
    End If
End Sub

Private Sub Form_Load()
Call LLENA_COMBO
dtpFecIni = Date
dtpFecFin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
    Call Reporte
Case "SALIR"
    Unload Me
End Select
End Sub

Public Sub Reporte()
Dim oo As Object
Dim strSQL As String, sEmpresa As String, Ruta As String
On Error GoTo ErrorImpresion
    
    strSQL = "SELECT DES_EMPRESA FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA='" & vemp & "'"
    sEmpresa = DevuelveCampo(strSQL, cCONNECT)
    
    VB.Screen.MousePointer = vbHourglass
    
    If txtNum_Ruc.Text = "" Then
        strCod_Anxo = ""
    End If
    
    strSQL = "Ventas_Emision_Articulos_por_Grupo '','','" & Mid(CboOrigen.Text, 1, 1) & "','" & dtpFecIni & "','" & dtpFecFin & "','" & txtCod_TipVenta & "','" & strCod_Anxo & "'"
    
    If MsgBox("Imprimir usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir") = vbYes Then
        Set oo = CreateObject("excel.application")
        oo.Workbooks.Open vRuta & "\RptVentasxGrupo.XLT"
        oo.Visible = True
        
        oo.Run "REPORTE", strSQL, "DESDE EL " & dtpFecIni & " HASTA EL " & dtpFecFin, cCONNECT, CboOrigen.Text, sEmpresa
    Else
        Ruta = vRuta & "\RptVentasxGrupo.OTS"
        Set oo = CreateObject("ooBusiness.Calc")
        oo.OfficeTemplateSheet = Ruta
        oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
        oo.MacroLibraryName = "Library1"
        oo.MacroModuleName = "Module1"
        oo.MacroName = "Reporte"
        
        oo.Run strSQL, "DESDE EL " & dtpFecIni & " HASTA EL " & dtpFecFin, cCONNECT, CboOrigen.Text, sEmpresa
    End If
    
    Screen.MousePointer = vbNormal
    Set oo = Nothing
    
Exit Sub
ErrorImpresion:
    Screen.MousePointer = vbNormal
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte" & err.Description, vbCritical, "Impresion"
End Sub

Sub LLENA_COMBO()
CboOrigen.AddItem "N - Nacional"
CboOrigen.AddItem "E - Extranjero"
CboOrigen.AddItem "T - Todos"
CboOrigen.AddItem "G - Transferencia Gratuita"

CboOrigen.ListIndex = -1
End Sub

Private Sub txtCod_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAne, txtDes_TipAne, 1, Me)
End Sub

Private Sub txtCod_TipVenta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Grupo_Ventas", "Descripcion", "Cn_Grupos_Ventas where ", txtCod_TipVenta, txtDes_TipVenta, 1, Me)
End Sub

Private Sub txtCod_TipVenta_LostFocus()
  If txtCod_TipVenta = "" Then txtCod_TipVenta = 0
End Sub

Private Sub txtDes_TipVenta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Grupo_Ventas", "Descripcion", "Cn_Grupos_Ventas where ", txtCod_TipVenta, txtDes_TipVenta, 2, Me)
End Sub

Private Sub txtDes_TipVenta_LostFocus()
  If txtCod_TipVenta = "" Then txtCod_TipVenta = 0
End Sub

Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 1, Me)
    SendKeys "{TAB}"
  End If
End Sub

