VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form FRmLote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Lote"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   Icon            =   "FRmLote.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtGrupo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Top             =   550
      Width           =   1300
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1305
      TabIndex        =   7
      Top             =   2580
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FRmLote.frx":030A
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.TextBox TxtObs 
      Height          =   975
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "FRmLote.frx":03A6
      Top             =   1350
      Width           =   3375
   End
   Begin VB.TextBox TxtLote 
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   950
      Width           =   1305
   End
   Begin VB.TextBox Txt_Des_Proveedor 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2610
      TabIndex        =   2
      Top             =   180
      Width           =   2025
   End
   Begin VB.TextBox TxtCod_Proveedor 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   180
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Grupo O/C :"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   620
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones:"
      Height          =   165
      Index           =   2
      Left            =   90
      TabIndex        =   5
      Top             =   1380
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Lote / Partida:"
      Height          =   285
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   1000
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Proveedor:"
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   250
      Width           =   975
   End
End
Attribute VB_Name = "FRmLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Reg As New ADODB.Recordset
Public Cod_TipOrdTra As String
Public Grupo As String
Public Codigo As String
Public Descripcion As String, sSer_OrdComp As String, sCod_OrdComp As String
Public Padre As Object
Public Paso As Boolean
Public Cod_Proveedor

Sub Adicionar()
On Error GoTo hand
Set Reg = Nothing
Reg.CursorLocation = adUseClient
Reg.Open "UP_Tx_OrdTra 'I','" & Cod_TipOrdTra & "', '" & Me.TxtCod_Proveedor & _
         "','" & Me.TxtLote & "','O','" & Me.TxtObs & "','" & Grupo & "', '" & _
         sSer_OrdComp & "', '" & sCod_OrdComp & "'", cConnect
Exit Sub
hand:
TxtLote.Text = ""
ErrorHandler err, "Adicionar"
Set Reg = Nothing
End Sub

Sub Limpia()
Me.Txt_Des_Proveedor = ""
Me.TxtCod_Proveedor = ""
Me.TxtGrupo = ""
Me.TxtLote = ""
Me.TxtObs = ""
End Sub

Private Sub Form_Load()
Limpia
TxtGrupo.Text = Grupo
TxtCod_Proveedor = Cod_Proveedor
Txt_Des_Proveedor.Text = DevuelveCampo("select des_proveedor from lg_proveedor where cod_proveedor='" & Cod_Proveedor & "'", cConnect)
End Sub


Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo hand
Select Case ActionName
    Case "ADICIONAR"
        If Trim(TxtCod_Proveedor) = "" Then MsgBox "Ingrese un Proveedor", vbInformation: Exit Sub
        If Trim(TxtLote) = "" Then MsgBox "Ingrese un Lote", vbInformation: Exit Sub
        
        Adicionar
        
        Padre.Cod_OrdTra = DevuelveCampo("select cod_ordtra from tx_ordtra where Cod_TipOrdTra='" & Cod_TipOrdTra & _
        "' and Cod_Proveedor='" & Me.TxtCod_Proveedor & "' and Cod_OrdProv='" & TxtLote & "'", cConnect)
        
        Padre.NewLote = Me.TxtLote
        Unload Me
    Case "CANCELAR"
        Unload Me
End Select
Exit Sub
hand:
    ErrorHandler err, "FunctButt1_ActionClick"
End Sub


Private Sub Txt_Des_Proveedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If DevuelveCampo("select count(*) from lg_proveedor where des_proveedor like '" & Txt_Des_Proveedor & "'", cConnect) = 1 Then
        Me.TxtCod_Proveedor = DevuelveCampo("select cod_proveedor from lg_proveedor where des_proveedor like '" & Me.Txt_Des_Proveedor & "'", cConnect)
    Else
        Set frmBusqGeneral.oParent = Me
        frmBusqGeneral.sQuery = "select cod_proveedor as Codigo, des_proveedor as Descripcion from lg_proveedor where des_proveedor like  '%" & Txt_Des_Proveedor & "%'"
        frmBusqGeneral.Cargar_Datos
        frmBusqGeneral.Show 1
        
        Me.TxtCod_Proveedor = Codigo
        Me.Txt_Des_Proveedor = Descripcion
    End If
End If
End Sub


Private Sub TxtCod_Proveedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If DevuelveCampo("select count(*) from lg_proveedor where cod_proveedor='" & Me.TxtCod_Proveedor & "'", cConnect) <= 0 Then
        MsgBox "El Proveedor no existe", vbInformation
        Exit Sub
    Else
        Me.Txt_Des_Proveedor = DevuelveCampo("select des_proveedor from lg_proveedor where cod_proveedor='" & Me.TxtCod_Proveedor & "'", cConnect)
    End If
End If
End Sub


