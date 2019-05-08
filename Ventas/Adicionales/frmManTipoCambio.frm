VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmManTipoCambio 
   Caption         =   "Mantenimiento de Tipo de Cambio"
   ClientHeight    =   3720
   ClientLeft      =   2220
   ClientTop       =   3510
   ClientWidth     =   7890
   Icon            =   "frmManTipoCambio.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   7890
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Cambio Otras de Monedas a Soles"
      Height          =   1395
      Left            =   75
      TabIndex        =   16
      Top             =   1620
      Width           =   7680
      Begin VB.TextBox txtEurosCompra 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   825
         TabIndex        =   5
         Text            =   "0"
         Top             =   795
         Width           =   975
      End
      Begin VB.TextBox txtYen 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6525
         TabIndex        =   8
         Text            =   "0"
         Top             =   320
         Width           =   975
      End
      Begin VB.TextBox txtEuros 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   825
         TabIndex        =   4
         Text            =   "0"
         Top             =   320
         Width           =   975
      End
      Begin VB.TextBox txtFrancos 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2670
         TabIndex        =   6
         Text            =   "0"
         Top             =   320
         Width           =   975
      End
      Begin VB.TextBox txtMarcos 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4560
         TabIndex        =   7
         Text            =   "0"
         Top             =   320
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Euros - Compra:"
         Height          =   435
         Left            =   120
         TabIndex        =   22
         Top             =   765
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Yen:"
         Height          =   405
         Left            =   5805
         TabIndex        =   20
         Top             =   405
         Width           =   750
      End
      Begin VB.Label Label6 
         Caption         =   "Euros - Venta:"
         Height          =   450
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Francos Suizos:"
         Height          =   405
         Left            =   2010
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Marcos Alemanes:"
         Height          =   405
         Left            =   3840
         TabIndex        =   17
         Top             =   360
         Width           =   750
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   510
      Left            =   2715
      TabIndex        =   9
      Top             =   3150
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmManTipoCambio.frx":030A
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Cambio Dólares a Soles"
      Height          =   855
      Left            =   75
      TabIndex        =   11
      Top             =   690
      Width           =   7680
      Begin VB.TextBox TxtCompra 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4560
         TabIndex        =   3
         Text            =   "0"
         Top             =   320
         Width           =   975
      End
      Begin VB.TextBox TxtVenta 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2640
         TabIndex        =   2
         Text            =   "0"
         Top             =   320
         Width           =   975
      End
      Begin VB.TextBox TxtCambio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Text            =   "0"
         Top             =   320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Compra :"
         Height          =   255
         Left            =   3840
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Venta :"
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Cambio :"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   75
      TabIndex        =   10
      Top             =   -30
      Width           =   7680
      Begin MSComCtl2.DTPicker DTFecha 
         Height          =   300
         Left            =   810
         TabIndex        =   0
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   94109697
         CurrentDate     =   37504
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha :"
         Height          =   255
         Left            =   150
         TabIndex        =   21
         Top             =   255
         Width           =   1155
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt3 
      Height          =   510
      Left            =   330
      TabIndex        =   15
      Top             =   4575
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   900
      Custom          =   "0~0~ACTUALIZAR~Verdadero~Verdadero~&Actualizar~0~0~1~~0~Falso~Falso~&Actualizar~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmManTipoCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public sAccion As String

Private Sub DTFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    Dim sSeguridad As String
    
    DTFecha.Value = Date
    'sSeguridad = get_botones1(Me, vper, vemp, Me.Name)
    
    'Me.FunctButt3.FunctionsUser = sSeguridad
    'If InStr(sSeguridad, "ACTUALIZAR") <> 0 Then
    '    Frame2.Enabled = True
    '    Me.FunctButt2.Visible = True
    'Else
    '    Frame2.Enabled = False
    '    Me.FunctButt2.Visible = False
    'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim Reg As ADODB.Recordset
On Error GoTo hand

'Set Reg = CreateObject("ADODB.Recordset")
'Reg.ActiveConnection = cCONNECT
'Reg.CursorLocation = adUseClient

'Reg.Open "select * from CN_TipoCambio where fecha='" & DTFecha.Value & "'"
'If Reg.RecordCount Then
'    TxtCambio.Text = Reg("tipo_cambio")
'    TxtVenta.Text = Reg("Tipo_venta")
'    TxtCompra.Text = Reg("Tipo_compra")
'Else
'    MsgBox "No existen tipos de cambio con la fecha seleccionada", vbInformation, Me.Caption
'    TxtCambio.Text = ""
'    TxtVenta.Text = ""
'    TxtCompra.Text = ""
'End If
'Set Reg = Nothing

Exit Sub
hand:
    ErrorHandler err, "Buscar"
    Set Reg = Nothing
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            Salvar_Datos
        Case "CANCELAR"
            Unload Me
    End Select
End Sub

Private Sub TxtCambio_GotFocus()
    SelectionText TxtCambio
End Sub

Private Sub TxtCambio_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    Else
        Call SoloNumeros(TxtCambio, KeyAscii, True, 4, 5)
    End If
End Sub

Private Sub TxtCompra_GotFocus()
    SelectionText TxtCompra
End Sub

Private Sub TxtCompra_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    Else
        Call SoloNumeros(TxtCompra, KeyAscii, True, 4, 5)
    End If
End Sub

Private Sub txtEuros_GotFocus()
    SelectionText txtEuros
End Sub

Private Sub txtEuros_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    Else
        Call SoloNumeros(txtEuros, KeyAscii, True, 4, 5)
    End If
End Sub

Private Sub txtEurosCompra_GotFocus()
    SelectionText txtEurosCompra
End Sub

Private Sub txtEurosCompra_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    Else
        Call SoloNumeros(txtEurosCompra, KeyAscii, True, 4, 5)
    End If
End Sub

Private Sub txtFrancos_GotFocus()
    SelectionText txtFrancos
End Sub

Private Sub txtFrancos_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    Else
        Call SoloNumeros(txtFrancos, KeyAscii, True, 4, 5)
    End If
End Sub

Private Sub txtMarcos_GotFocus()
    SelectionText txtMarcos
End Sub

Private Sub txtMarcos_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    Else
        Call SoloNumeros(txtMarcos, KeyAscii, True, 4, 5)
    End If
End Sub

Private Sub TxtVenta_GotFocus()
    SelectionText TxtVenta
End Sub

Private Sub TxtVenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    Else
        Call SoloNumeros(TxtVenta, KeyAscii, True, 4, 5)
    End If
End Sub

Private Sub Salvar_Datos()
Dim sSQL As String

On Error GoTo hand
    sSQL = "UP_MAN_CN_TIPOCAMBIO_ALL '$','$',$,$,$,'$' ,'$',$,$,$,$ , $"
    sSQL = VBsprintf(sSQL, sAccion, DTFecha.Value, TxtCambio.Text, TxtVenta.Text, TxtCompra.Text, vusu, ComputerName, txtEuros.Text, txtFrancos.Text, txtMarcos.Text, txtYen.Text, txtEurosCompra.Text)
    ExecuteCommandSQL cCONNECT, sSQL
'    If vemp = "01" Then cmdTransmitir
    'If vemp = "01" Then cmdTransmitir_SLIN
    MsgBox "Datos actualizados", vbInformation, "Tipo Cambio"
    
Exit Sub
hand:
    ErrorHandler err, "SALVAR_DATOS"
    
End Sub








Function FixData(wtexto As Variant, ofield As ADODB.FIELD)
    If IsNull(wtexto) Or Len(Trim(wtexto)) = 0 Then
        Select Case ofield.Type
        Case adBigInt, adCurrency, adDecimal, adDouble, adInteger, adNumeric, adSingle
            wtexto = 0
        Case adBoolean
            wtexto = False
        Case adDate
            wtexto = Empty
        Case adChar, adVarChar
            wtexto = ""
        End Select
    End If
    FixData = wtexto
End Function

Private Sub txtYen_GotFocus()
    SelectionText txtYen
End Sub

Private Sub txtYen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    Else
        Call SoloNumeros(txtYen, KeyAscii, True, 4, 5)
    End If
End Sub

