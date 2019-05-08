VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form FrmCambiosColor 
   Caption         =   "Cambios Color"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Datos Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6105
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   1785
         TabIndex        =   5
         Top             =   1155
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"FrmCambiosColor.frx":0000
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.TextBox TxtDes_Color 
         Height          =   330
         Left            =   1365
         TabIndex        =   2
         Top             =   735
         Width           =   4635
      End
      Begin VB.TextBox TxtCod_Color 
         Height          =   330
         Left            =   1365
         MaxLength       =   6
         TabIndex        =   1
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   4
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   3
         Top             =   420
         Width           =   600
      End
   End
End
Attribute VB_Name = "FrmCambiosColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Cod_Almacen As String
Public Num_MovStk As String
Dim strSQL As String
Public Codigo As String
Public Descripcion As String
Public Paso As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "ACEPTAR"
        Call CAMBIOS_COLOR
    Case "CANCELAR"
        Unload Me
End Select
End Sub

Sub CAMBIOS_COLOR()
On Error GoTo err

If TxtCod_Color = "" Then
    MsgBox "Ingrese Codigo Color"
    Exit Sub
End If
If txtDes_Color = "" Then
    MsgBox "Ingrese Descripcion Color"
    Exit Sub
End If

strSQL = "EXEC UP_MODI_COLOR_PARTIDA '" & Me.Cod_Almacen & _
    "', '" & Me.Num_MovStk & "', '" & Trim(Me.TxtCod_Color) & _
    "', '" & Trim(Me.txtDes_Color) & "'"
    
ExecuteSQL cConnect, strSQL
Unload Me
Exit Sub
err:
    ErrorHandler err, "CAMBIOS_COLOR"
End Sub

Private Sub TxtCod_Color_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TxtCod_Color <> "" Then
        busca_color (1)
    End If
    SendKeys "{TAB}"
End If
End Sub


Sub busca_color(Tipo As Integer)
Dim oTipo As New frmBusqGeneral
Dim Rs As New ADODB.Recordset
Set oTipo.oParent = Me

If Tipo = 1 Then
    oTipo.sQuery = "SELECT cod_color AS 'Código', des_color AS 'Descripción' FROM lb_color where cod_color like '" & Me.TxtCod_Color & "%'"
ElseIf Tipo = 2 Then
    oTipo.sQuery = "SELECT cod_color AS 'Código', des_color AS 'Descripción' FROM lb_color where des_color like '" & Me.txtDes_Color & "%'"
End If

oTipo.CARGAR_DATOS
oTipo.Show 1
If Codigo <> "" Then
    Me.TxtCod_Color.Text = Trim(Codigo)
    Me.txtDes_Color.Text = Trim(Descripcion)
    Codigo = "": Descripcion = ""
End If
Set oTipo = Nothing
Set Rs = Nothing
End Sub

Private Sub TxtDes_Color_GotFocus()
txtDes_Color.SelStart = 0
txtDes_Color.SelLength = Len(txtDes_Color)
End Sub

Private Sub TxtDes_Color_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtDes_Color <> "" Then
        busca_color (2)
    End If
End If
End Sub
