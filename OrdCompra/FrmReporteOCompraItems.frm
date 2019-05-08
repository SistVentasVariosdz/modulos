VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmReporteOCompraItems 
   Caption         =   "Ordenes De Compra Por Items"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Ordenes De Compra Por Items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6885
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
         Height          =   540
         Left            =   2205
         TabIndex        =   3
         Top             =   1680
         Width           =   1380
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   540
         Left            =   3675
         TabIndex        =   6
         Top             =   1680
         Width           =   1380
      End
      Begin VB.Frame Frame2 
         Caption         =   "Rango Opcional"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   105
         TabIndex        =   9
         Top             =   840
         Width           =   6555
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   330
            Left            =   1410
            TabIndex        =   4
            Top             =   315
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   100990977
            CurrentDate     =   37924
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   330
            Left            =   4035
            TabIndex        =   5
            Top             =   315
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   100990977
            CurrentDate     =   37924
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inicial"
            Height          =   195
            Left            =   465
            TabIndex        =   11
            Top             =   420
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Final"
            Height          =   195
            Left            =   3090
            TabIndex        =   10
            Top             =   420
            Width           =   825
         End
      End
      Begin VB.CommandButton CmdBuscaItem 
         Caption         =   "..."
         Height          =   330
         Left            =   1995
         TabIndex        =   7
         Top             =   420
         Width           =   435
      End
      Begin VB.TextBox TxtDesItem 
         Height          =   330
         Left            =   2415
         TabIndex        =   2
         Top             =   420
         Width           =   4245
      End
      Begin VB.TextBox TxtCodItem 
         Height          =   330
         Left            =   840
         TabIndex        =   1
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Item"
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   525
         Width           =   300
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   5280
      Top             =   2400
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmReporteOCompraItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Fecha1 As String
Public Fecha2 As String
Public codigo, Descripcion As String
Dim strSql As String

Private Sub cmdAceptar_Click()
If IsNull(DTPicker1.Value) And IsNull(DTPicker2.Value) Then
    Fecha1 = ""
    Fecha2 = ""
    Call Reporte
Else
    If IsNull(DTPicker1.Value) Or IsNull(DTPicker2.Value) Then
        MsgBox "Debe ingresar ambas fechas o en su defecto ninguna", vbCritical
    Else
        Fecha1 = CStr(DTPicker1.Value)
        Fecha2 = CStr(DTPicker2.Value)
        Call Reporte
    End If
End If
    DTPicker1.Value = ""
    DTPicker2.Value = ""
End Sub

Private Sub CmdBuscaItem_Click()
    Call Me.BUSCA_ITEM(2)
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Sub Reporte()
Dim oo As Object
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\RptOrdenesCompraItems.xlt"
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "Reporte", TxtCodItem.Text, Fecha1, Fecha2, cConnect
    Set oo = Nothing
End Sub

Private Sub Form_Load()
    Fecha1 = ""
    Fecha2 = ""
End Sub

Private Sub TxtCodItem_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Me.TxtCodItem.Text) = "" Then
        Call Me.BUSCA_ITEM(2)
    Else
        Call Me.BUSCA_ITEM(1)
    End If
End If
CmdAceptar.SetFocus
End Sub

Public Sub BUSCA_ITEM(Tipo As Integer)
    
    Select Case Tipo
        Case 1:
                    strSql = "SELECT COD_ITEM, DES_ITEM FROM LG_ITEM WHERE cod_ITEM = '" & Trim(Me.TxtCodItem.Text) & "' ORDER BY cod_ITEM"
                    Me.TxtDesItem.Text = Trim(DevuelveCampo(strSql, cConnect))
        Case 2:
                    Dim oTipo As New frmBusqGeneral
                    Dim Rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    oTipo.sQuery = "SELECT cod_ITEM as 'Código', DES_ITEM as 'Descripción' FROM LG_ITEM ORDER BY cod_ITEM"
                    
                    oTipo.CARGAR_DATOS
                    oTipo.Show 1
                    If codigo <> "" Then
                        Me.TxtCodItem.Text = Trim(codigo)
                        Me.TxtDesItem.Text = Trim(Descripcion)
                        codigo = "": Descripcion = ""
                    End If
                    Set oTipo = Nothing
                    Set Rs = Nothing
    End Select
    Me.TxtCodItem.SetFocus
End Sub
