VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmShowGeneracionContable 
   Caption         =   "Generación de Información Contable"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtDes_TipoDiario 
      Height          =   285
      Left            =   1485
      TabIndex        =   1
      Top             =   75
      Width           =   3900
   End
   Begin VB.Frame Frame1 
      Caption         =   "Modo de Generación"
      Height          =   1110
      Left            =   105
      TabIndex        =   12
      Top             =   825
      Width           =   5295
      Begin VB.Frame fraRango 
         Caption         =   "Rango de Comprobantes"
         Height          =   750
         Left            =   2460
         TabIndex        =   13
         Top             =   255
         Visible         =   0   'False
         Width           =   2655
         Begin VB.TextBox txtHasta 
            Height          =   285
            Left            =   1905
            TabIndex        =   7
            Top             =   330
            Width           =   525
         End
         Begin VB.TextBox txtDesde 
            Height          =   285
            Left            =   645
            TabIndex        =   6
            Top             =   330
            Width           =   525
         End
         Begin VB.Label Label5 
            Caption         =   "Hasta"
            Height          =   285
            Left            =   1380
            TabIndex        =   15
            Top             =   345
            Width           =   435
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
            Height          =   285
            Left            =   105
            TabIndex        =   14
            Top             =   345
            Width           =   525
         End
      End
      Begin VB.OptionButton optRango 
         Caption         =   "Rango de Comprobantes"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   645
         Width           =   2205
      End
      Begin VB.OptionButton optPendientes 
         Caption         =   "Pendientes de Transmisión"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   315
         Value           =   -1  'True
         Width           =   2205
      End
   End
   Begin VB.TextBox TxtCod_TipoDiario 
      Height          =   285
      Left            =   1005
      TabIndex        =   0
      Top             =   75
      Width           =   405
   End
   Begin VB.TextBox TxtPeriodo 
      Height          =   285
      Left            =   2550
      TabIndex        =   3
      Top             =   450
      Width           =   405
   End
   Begin VB.TextBox TxtAno 
      Height          =   285
      Left            =   1005
      TabIndex        =   2
      Top             =   435
      Width           =   615
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1515
      TabIndex        =   8
      Top             =   2055
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "0~0~ACEPTAR~True~True~&Aceptar~0~0~1~~0~False~False~&Aceptar~~1~0~CANCELAR~True~True~&Cancelar~0~0~2~~0~False~False~&Cancelar~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Label Label3 
      Caption         =   "SubDiario"
      Height          =   300
      Left            =   150
      TabIndex        =   11
      Top             =   90
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Período"
      Height          =   300
      Left            =   1770
      TabIndex        =   10
      Top             =   465
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Año"
      Height          =   300
      Left            =   150
      TabIndex        =   9
      Top             =   465
      Width           =   600
   End
End
Attribute VB_Name = "frmShowGeneracionContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public Codigo As String
Public Descripcion As String
Public TipoAdd As String

Private Sub Form_Load()
    TxtAno = Year(Date)
    TxtPeriodo = Month(Date)
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            Generar
        Case "CANCELAR"
            Unload Me
    End Select
End Sub

Private Sub optPendientes_Click()
    fraRango.Visible = False
        
End Sub

Private Sub optRango_Click()
    fraRango.Visible = True
    txtDesde.SetFocus
End Sub

Private Sub TxtCod_TipoDiario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_SubDiario("1")
End If
End Sub

Private Sub TxtDes_TipoDiario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_SubDiario("2")
End If
End Sub

Sub Busca_SubDiario(Tipo As String)
Dim oTipo As New frmBusqGeneral3
Dim iCol As Long
Dim rstAux As adodb.Recordset

Set oTipo.oParent = Me

If Tipo = "1" Then
    strSQL = "SELECT cod_tipodiario as Codigo, Des_TipoDiario as Descripcion, flg_canjefacturasporpagarconletras as Flg from cn_tipodiario where cod_tipodiario like '" & Trim(TxtCod_TipoDiario.Text) & "%'"
Else
    strSQL = "SELECT cod_tipodiario as Codigo, Des_TipoDiario as Descripcion, flg_canjefacturasporpagarconletras as Flg from cn_tipodiario where des_tipodiario like '%" & Trim(TxtDes_TipoDiario.Text) & "%'"
End If
With oTipo
    Set .oParent = Me
    .SQuery = strSQL
    .Cargar_Datos
    .Caption = "Selccionar SubDiario"
    Codigo = ".."
    Set rstAux = .gexLista.ADORecordset
    
    .gexLista.Columns("Codigo").Width = 700
    .gexLista.Columns("Descripcion").Width = 5000

'    For iCol = 3 To .gexLista.Columns.count
'        .gexLista.Columns(iCol).Visible = False
'    Next iCol
    
    If rstAux.RecordCount = 1 Then
        Codigo = Trim(rstAux!Codigo)
        Descripcion = Trim(rstAux!Descripcion)
    End If
    
    If rstAux.RecordCount > 1 Then .Show vbModal
    
    If Codigo <> "" And rstAux.RecordCount > 0 Then
        TxtCod_TipoDiario = Codigo
        TxtDes_TipoDiario = Descripcion
        TxtAno.SetFocus
    End If
End With

Codigo = "": Descripcion = ""
Unload oTipo
Set oTipo = Nothing
rstAux.Close
Set rstAux = Nothing
End Sub

Private Sub TxtAno_GotFocus()
SelectionText TxtAno
End Sub

Private Sub TxtAno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If txtDesde.Text <> "" Then
        txtDesde = StrZero(txtDesde, 4)
    End If
    txtHasta.SetFocus
End If
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtHasta.Text <> "" Then
            txtHasta = StrZero(txtHasta, 4)
        End If
        FunctButt1.SetFocus
    End If
End Sub

Private Sub TxtPeriodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If TxtPeriodo.Text <> "" Then
            TxtPeriodo = StrZero(TxtPeriodo, 2)
        End If
        FunctButt1.SetFocus
    End If
End Sub


Private Sub Generar()
On Error GoTo errx
Dim sSql As String

If optPendientes Then
    txtDesde.Text = ""
    txtHasta.Text = ""
Else
    If Val(txtDesde) < 0 Then
        Aviso "Rango Inicial inválido.Revisar", 2
        Exit Sub
    End If
    If Val(txtHasta) > 9999 Then
        Aviso "Rango Final inválido.Revisar", 2
        Exit Sub
    End If
End If

sSql = "CN_GENERACION_CONTABLE '$','$','$','$','$'"
sSql = VBsprintf(sSql, TxtCod_TipoDiario.Text, TxtAno.Text, TxtPeriodo.Text, txtDesde.Text, txtHasta.Text)


Exit Sub
errx:
    errores Err.Number
End Sub
