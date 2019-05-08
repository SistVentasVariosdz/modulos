VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form FrmCopiarComb 
   Caption         =   "Alternativas Peso / Ancho"
   ClientHeight    =   1590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmCopiarComb.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame FraDatos 
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7335
      Begin VB.TextBox txtfamilia 
         Height          =   285
         Left            =   1515
         TabIndex        =   0
         Top             =   360
         Width           =   1245
      End
      Begin VB.CommandButton cmdBusFamItem 
         Caption         =   "..."
         Height          =   300
         Left            =   2805
         TabIndex        =   4
         Tag             =   "..."
         Top             =   360
         Width           =   360
      End
      Begin VB.TextBox txtdes_familia 
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Combinación :"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   405
         Width           =   1005
      End
   End
End
Attribute VB_Name = "FrmCopiarComb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vCodCombD As String
Public vAccion As String
Public vRuta As String
Public Codigo_tela As String
Public Codigo, Descripcion As String
Public oParent As Object
Dim StrSQL As String

Private Sub cmdBusFamItem_Click()
    Dim oTipo As New frmBusqGeneral4
    Dim Rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.sQuery = "SELECT Cod_Comb as Código, Des_Comb as Descripción FROM TX_TELACOMB WHERE Cod_Tela='" & Codigo_tela & "'"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        txtfamilia.Text = Codigo
        txtdes_familia.Text = Descripcion
        
    End If
    Set oTipo = Nothing
    Set Rs = Nothing
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    Call Grabar
Case "CANCELAR"
    Unload Me
End Select
End Sub


Sub Grabar()
On Error GoTo errGrabar

If Trim(txtfamilia.Text) = "" Then
    MsgBox "Debe ingresar una Combinación"
End If


StrSQL = "TX_TELA_COPIAR_COMB '" & Codigo_tela & "','" & txtfamilia.Text & "','" & vCodCombD & "'"
            
ExecuteCommandSQL cCONNECT, StrSQL

frmMantTelaComb.CARGA_GRID
Unload Me

Exit Sub
errGrabar:
    MsgBox Err.Description, vbCritical, "Grabar"
End Sub

Private Sub txtfamilia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtfamilia.Text) = "" Then
            cmdBusFamItem_Click
        Else
                StrSQL = "SELECT Des_Comb as Descripción FROM TX_TELACOMB WHERE Cod_Comb='" & txtfamilia.Text & "' AND Cod_Tela='" & Codigo_tela & "'"
                txtdes_familia.Text = DevuelveCampo(StrSQL, cCONNECT)

        End If
    
    FunctButt1.SetFocus
    End If
    
End Sub
