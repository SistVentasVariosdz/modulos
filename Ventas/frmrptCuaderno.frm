VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmrptCuaderno 
   Caption         =   "Impresion  Documentos Ventas "
   ClientHeight    =   7635
   ClientLeft      =   345
   ClientTop       =   735
   ClientWidth     =   11610
   Icon            =   "frmrptCuaderno.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   11610
   Begin VB.Frame FraBuscar 
      Caption         =   "Opciones de Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11520
      Begin VB.CheckBox chkBordes 
         Alignment       =   1  'Right Justify
         Caption         =   "Con Bordes"
         Height          =   255
         Left            =   4920
         TabIndex        =   14
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtInterlineado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   6240
         MaxLength       =   1
         TabIndex        =   3
         Text            =   "1"
         Top             =   840
         Width           =   540
      End
      Begin VB.TextBox txtSer_Docum 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   4200
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "001"
         Top             =   840
         Width           =   540
      End
      Begin VB.TextBox txtCod_TipDoc 
         Height          =   285
         Left            =   960
         MaxLength       =   4
         TabIndex        =   0
         Text            =   "FA"
         Top             =   840
         Width           =   480
      End
      Begin VB.TextBox txtDes_TipDoc 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Text            =   "FACTURAS"
         Top             =   840
         Width           =   1905
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   675
         Left            =   7200
         TabIndex        =   4
         Top             =   270
         Width           =   4170
         _ExtentX        =   7355
         _ExtentY        =   1191
         Custom          =   $"frmrptCuaderno.frx":030A
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1000
         ControlHeigth   =   650
         ControlSeparator=   40
      End
      Begin MSComCtl2.DTPicker dtpFecEmiIni 
         Height          =   315
         Left            =   960
         TabIndex        =   5
         Top             =   330
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61997057
         CurrentDate     =   37543
      End
      Begin MSComCtl2.DTPicker dtpFecEmiFin 
         Height          =   315
         Left            =   3480
         TabIndex        =   6
         Top             =   330
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61997057
         CurrentDate     =   37543
      End
      Begin VB.Label Label3 
         Caption         =   "Nro Interlineado :"
         Height          =   255
         Left            =   4920
         TabIndex        =   13
         Top             =   855
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Doc :"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   855
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Serie :"
         Height          =   255
         Left            =   3600
         TabIndex        =   11
         Top             =   855
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         Height          =   195
         Left            =   2760
         TabIndex        =   10
         Top             =   390
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   390
         Width           =   555
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   6060
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1320
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   10689
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
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmrptCuaderno.frx":0446
      Column(2)       =   "frmrptCuaderno.frx":050E
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmrptCuaderno.frx":05B2
      FormatStyle(2)  =   "frmrptCuaderno.frx":06EA
      FormatStyle(3)  =   "frmrptCuaderno.frx":079A
      FormatStyle(4)  =   "frmrptCuaderno.frx":084E
      FormatStyle(5)  =   "frmrptCuaderno.frx":0926
      FormatStyle(6)  =   "frmrptCuaderno.frx":09DE
      FormatStyle(7)  =   "frmrptCuaderno.frx":0ABE
      FormatStyle(8)  =   "frmrptCuaderno.frx":0B6A
      ImageCount      =   0
      PrinterProperties=   "frmrptCuaderno.frx":0C1A
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   10680
      Top             =   5760
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmrptCuaderno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public codigo As String, Descripcion As String

Private Sub dtpFecEmiIni_Change()
  dtpFecEmiFin.Value = dtpFecEmiIni.Value
End Sub

Private Sub Form_Load()
    
  iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))
  
  
  dtpFecEmiIni = Date - 1
  dtpFecEmiFin = Date - 1
    
End Sub

Private Sub Buscar()

Dim sSQL As String

sSQL = "Ventas_Muestra_Relacion_Facturas_Cuaderno '" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','" & txtCod_TipDoc & "','" & txtSer_Docum & "'"

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)

GridEX1.Columns("Nro_Documento").Width = 1305
GridEX1.Columns("Nro_Documento").Caption = "Nro_Factura"
GridEX1.Columns("Cliente").Width = 4800
GridEX1.Columns("Cliente").Caption = "Cliente"
GridEX1.Columns("Fecha").Width = 945
GridEX1.Columns("Fecha").Caption = "Fecha"

End Sub

Private Sub Doc_Finan()

On Error GoTo ErrorImpresion

Dim sSQL As String, Rs As New ADODB.Recordset
Dim oo As Object

sSQL = "Ventas_Muestra_Doc_Gastos_Financieros '" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','" & txtCod_TipDoc & "','" & txtSer_Docum & "'"

Set Rs = CargarRecordSetDesconectado(sSQL, cCONNECT)

If Not (Rs.BOF And Rs.EOF) Then

  Set oo = CreateObject("excel.application")
  oo.Workbooks.Open vRuta & "\ReporteDocumentosReq.xlt"
  oo.Visible = True
  
  oo.Run "REPORTE", Rs, " DESDE EL " & dtpFecEmiIni & "  HASTA EL " & dtpFecEmiFin

  Screen.MousePointer = vbNormal
  oo.Visible = True
  Set oo = Nothing

Else
  MsgBox "No hay Registros q Imprimir", vbInformation, "AVISO"
End If

Exit Sub

Resume
ErrorImpresion:
    Screen.MousePointer = vbNormal
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte" & err.Description, vbCritical, "Impresion"
End Sub


Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Msg As Variant
    Select Case ActionName
    Case "BUSCAR"
      Buscar
    Case "IMPRIMIR"
      If GridEX1.RowCount = 0 Then Exit Sub
      Reporte
    Case "FINAN"
       Doc_Finan
    Case "SALIR"
       Unload Me
    End Select
End Sub

Public Sub Reporte()
  
On Error GoTo ErrorImpresion

    VB.Screen.MousePointer = vbHourglass
    
    Dim oo As Object
    Set oo = CreateObject("excel.application")
    
    If txtCod_TipDoc = "81" Then
      oo.Workbooks.Open vRuta & "\ReporteDocVentasLetras.xlt"
    Else
      oo.Workbooks.Open vRuta & "\ReporteDocVentas.xlt"
    End If
    
    oo.Visible = True

    oo.Run "REPORTE", GridEX1.ADORecordset, " DESDE EL " & dtpFecEmiIni & "  HASTA EL " & dtpFecEmiFin, txtInterlineado, chkBordes

    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
    
    Exit Sub
    Resume
ErrorImpresion:
    Screen.MousePointer = vbNormal
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte" & err.Description, vbCritical, "Impresion"
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtCod_TipDoc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipDoc", "Des_TipDoc", "CN_TiposDocum where Flg_Doc_Ventas = '*' and ", txtCod_TipDoc, txtDes_TipDoc, 1)
End Sub

Private Sub txtDes_TipDoc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipDoc", "Des_TipDoc", "CN_TiposDocum where Flg_Doc_Ventas = '*' and ", txtCod_TipDoc, txtDes_TipDoc, 1)
End Sub

Sub Busca_Opcion(strCampo1 As String, strCampo2 As String, StrTabla As String, txtCod As TextBox, txtDes As TextBox, Opcion As Integer)

On Error GoTo Fin

Dim rstAux As ADODB.Recordset, strSQL As String

    strSQL = "Select " & strCampo1 & " AS Cod," & strCampo2 & " as Descripcion from " & StrTabla

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    Select Case Opcion
    Case 1: strSQL = strSQL & strCampo1 & " like '%" & txtCod & "%'"
    Case 2: strSQL = strSQL & strCampo2 & " like '%" & txtDes & "%'"
    End Select
    txtCod = ""
    txtDes = ""
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = strSQL
        .CARGAR_DATOS
        
        codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If codigo <> "" And rstAux.RecordCount > 0 Then
            txtCod = Trim(rstAux!Cod)
            txtDes = Trim(rstAux!Descripcion)
            Select Case Opcion
            Case 1: SendKeys "{TAB}": SendKeys "{TAB}"
            Case 2: SendKeys "{TAB}"
            End Select
        Else
            SendKeys "{TAB}"
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & Opcion & ")"
End Sub

Private Sub txtInterlineado_Change()
If txtInterlineado = "" Or txtInterlineado = "0" Then txtInterlineado = 1
End Sub

Private Sub txtInterlineado_GotFocus()
SendKeys "{HOME}+{END}"
End Sub

Private Sub txtInterlineado_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtSer_Docum_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtSer_Docum_LostFocus()
  txtSer_Docum = Format(txtSer_Docum, "000")
End Sub
