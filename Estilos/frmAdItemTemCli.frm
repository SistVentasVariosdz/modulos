VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmAdItemTemCli 
   Caption         =   "Adicionar Bord.Estamp. Aplicac. a Temporada"
   ClientHeight    =   1305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   ScaleHeight     =   1305
   ScaleWidth      =   6900
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcod_item 
      Height          =   285
      Left            =   930
      MaxLength       =   8
      TabIndex        =   0
      Top             =   150
      Width           =   1005
   End
   Begin VB.TextBox txtdes_item 
      Height          =   285
      Left            =   2250
      TabIndex        =   1
      Top             =   150
      Width           =   4200
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2190
      TabIndex        =   2
      Top             =   660
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmAdItemTemCli.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Label Label1 
      Caption         =   "Item"
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   180
      Width           =   720
   End
End
Attribute VB_Name = "frmAdItemTemCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Codigo As String
Public Descripcion As String
Public sCod_Cliente As String
Public sCod_Temcli As String
Public oParent As Object

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            Grabar
        Case "CANCELAR"
            Unload Me
    End Select
End Sub

Private Sub txtcod_item_KeyPress(KeyAscii As Integer)
    Dim StrSQL As String
    If KeyAscii = 13 Then
        If Trim(txtcod_item.Text) = "" Then
            Call MUESTRA_ITEMS(1)
            'Call MsgBox("Sirvase ingresar un codigo de Item", vbInformation)
        Else
            txtcod_item.Text = CompletaCodigo(Trim(txtcod_item.Text), 8, 2)
            
            'Esta consulta es para obtener el Codigo de Cliente
            StrSQL = "SELECT Des_Item FROM LG_ITEM WHERE Cod_Item='" & txtcod_item.Text & "'"
            txtdes_item.Text = DevuelveCampo(StrSQL, cCONNECT)
            FunctButt1.SetFocus
        End If
        
    End If
End Sub

Sub MUESTRA_ITEMS(Tipo As Integer)
    Dim oTipo As New frmBusqGeneral
    Dim rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    If Tipo = 1 Then
        oTipo.sQuery = "SELECT cod_Item as Codigo, des_Item as Descripcion FROM LG_Item ORDER BY cod_Item"
    ElseIf Tipo = 2 Then
        oTipo.sQuery = "SELECT cod_Item as Codigo, des_Item as Descripcion FROM LG_Item where des_item like '%" & Trim(Me.txtdes_item.Text) & "%' ORDER BY des_Item"
    ElseIf Tipo = 3 Then
        oTipo.sQuery = "SELECT cod_Item as Codigo, des_Item as Descripcion FROM LG_Item ORDER BY Des_Item"
    End If
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtcod_item.Text = Codigo
        txtdes_item.Text = Descripcion
        
        FunctButt1.SetFocus
        Codigo = ""
    End If
    Set oTipo = Nothing
    Set rs = Nothing
End Sub




Public Function CompletaCodigo(CodOrigen As String, longcodfinal As Integer, PosfinalCod As Integer) As String
' CodOrigen     = Es el codigo que sera pasado por parametro
' LongCodFinal  = Es el tamaño del Codigo a devolver
' PosFinalCod   = Es la posicion de la 1era parte del codigo
    Dim Contador As Integer
    CompletaCodigo = Mid(CodOrigen, 1, PosfinalCod)
    For Contador = 1 To longcodfinal - Len(CodOrigen)
        CompletaCodigo = CompletaCodigo & "0"
    Next
    Contador = Len(CodOrigen) - PosfinalCod
    If Contador < 0 Then
        Contador = 0
    End If
    CompletaCodigo = CompletaCodigo & Right(CodOrigen, Contador)
End Function


Private Sub Grabar()
On Error GoTo errx:
Dim ssql As String

ssql = "LG_Adicionar_Item_a_Temporada_Cliente '$','$','$'"
ssql = VBsprintf(ssql, txtcod_item, sCod_Cliente, sCod_Temcli)

ExecuteCommandSQL cCONNECT, ssql

oParent.CargaLista
Unload Me

Exit Sub

errx:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
