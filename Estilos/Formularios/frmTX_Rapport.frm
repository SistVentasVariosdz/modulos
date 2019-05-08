VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmTX_Rapport 
   Caption         =   "Mantenimiento Rapport"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1995
      TabIndex        =   4
      Top             =   1365
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "0~0~ACEPTAR~True~True~&Aceptar~0~0~1~~0~False~False~&Aceptar~~1~0~CANCELAR~True~True~&Cancelar~1~0~3~~0~False~False~&Cancelar~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6525
      Begin VB.TextBox TxtDes_Cliente 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3255
         TabIndex        =   11
         Top             =   735
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   330
         Left            =   2730
         TabIndex        =   10
         Top             =   735
         Width           =   435
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   330
         Left            =   2730
         TabIndex        =   9
         Top             =   360
         Width           =   435
      End
      Begin VB.TextBox TxtDes_Tela 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3255
         TabIndex        =   8
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox txtCod_Tela 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   2
         Top             =   360
         Width           =   990
      End
      Begin VB.TextBox txtDesRapport 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   210
         Visible         =   0   'False
         Width           =   3870
      End
      Begin VB.TextBox TxtCod_Cliente 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1695
         TabIndex        =   3
         Top             =   765
         Width           =   930
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Rapport:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   105
         TabIndex        =   7
         Tag             =   "Code"
         Top             =   285
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Tela"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   105
         TabIndex        =   6
         Tag             =   "Family :"
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   105
         TabIndex        =   5
         Tag             =   "Family :"
         Top             =   840
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmTX_Rapport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Opcion As String
Dim StrSql As String

Public codigo As String
Public descripcion As String

Dim Cliente As String
Public Rapport As Integer

Private Sub Command1_Click()
Dim oTipo As New frmBusqGeneral
    Dim Rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.sQuery = "SELECT Cod_Tela as Código, Des_Tela as Descripción FROM TX_TELA"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If codigo <> "" Then
        txtCod_Tela.Text = codigo
        TxtDes_Tela.Text = descripcion
    End If
    Set oTipo = Nothing
    Set Rs = Nothing
End Sub

Private Sub Command2_Click()
Set frmBusqGeneral.oParent = Me
frmBusqGeneral.sQuery = "Select abr_cliente as Codigo,nom_cliente as Descripcion from tg_cliente order by 1"
frmBusqGeneral.Cargar_Datos

frmBusqGeneral.Show 1
TxtDes_Cliente = descripcion
TxtCod_Cliente.Text = codigo
Cliente = DevuelveCampo("Select cod_cliente from tg_cliente  where abr_cliente='" & TxtCod_Cliente.Text & "'", cCONNECT)
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    If VALIDA = False Then
        Exit Sub
    End If
    Call GRABAR_RAPPORT
Case "CANCELAR"
    Unload Me
End Select
End Sub


Sub GRABAR_RAPPORT()
Dim con As New ADODB.Connection
On Error GoTo Salvar_DatosErr
Dim StrSql As String
Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    con.ConnectionString = cCONNECT
    con.Open
    
    con.BeginTrans
    
    If Opcion = "I" Then
        Rapport = CInt(DevuelveCampo("SELECT RAPPORT_NUMBER FROM TG_CONTROL", cCONNECT)) + 1
    End If

    StrSql = "EXEC UP_MAN_TX_RAPPORT '" & Opcion & "'," & Rapport & ",'" & Me.txtDesRapport & "','" & Me.txtCod_Tela & "','" & Cliente & " ','" & vusu & "','" & Format(Now, "DD/MM/YYYY") & "','" & ComputerName & "'"
                
    con.Execute StrSql
    con.CommitTrans
    
    Screen.MousePointer = vbDefault
    Unload Me
    If Opcion = "I" Then
        MsgBox "Generado el Rapport " & Rapport, vbInformation
    End If
    Exit Sub
Salvar_DatosErr:
    con.RollbackTrans
    Set con = Nothing
    Screen.MousePointer = vbDefault
    ErrorHandler Err, "GRABAR_RAPPORT"
End Sub


Private Sub TxtCod_Cliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxtCod_Cliente.Text) = "" Then
        Command2_Click
    Else
        Cliente = DevuelveCampo("Select cod_cliente from tg_cliente  where abr_cliente='" & TxtCod_Cliente.Text & "'", cCONNECT)
        StrSql = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE cod_cliente='" & Cliente & "'"
        TxtDes_Cliente.Text = DevuelveCampo(StrSql, cCONNECT)
    End If
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtCod_Tela_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_Tela.Text) = "" Then
            Command1_Click
        Else
            Me.txtCod_Tela.Text = CompletaCodigo(Trim(txtCod_Tela.Text), 8, 2)
            
            
            
            StrSql = "SELECT Des_Tela FROM TX_TELA WHERE Cod_Tela='" & txtCod_Tela.Text & "'"
            TxtDes_Tela.Text = DevuelveCampo(StrSql, cCONNECT)
        End If
        SendKeys "{TAB}"
    End If
End Sub


Private Sub TxtDes_Cliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ExisteCampo("abr_cliente", "tg_cliente", TxtDes_Cliente, cCONNECT, True) = False Then
        MsgBox "El cliente no existe", vbInformation
        Exit Sub
    Else
        Cliente = DevuelveCampo("Select cod_cliente from tg_cliente  where abr_cliente='" & txtCod_Tela.Text & "'", cCONNECT)
    End If
End If
End Sub


Function VALIDA() As Boolean
    VALIDA = False
    If Trim(txtCod_Tela.Text) = "" Then
        MsgBox "Ingrese Tela"
        VALIDA = False
        Exit Function
    End If
    If Trim(TxtCod_Cliente.Text) = "" Then
        MsgBox "Ingrese Cliente"
        VALIDA = False
        Exit Function
    End If
    VALIDA = True
    
End Function

Private Sub txtDesRapport_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If

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
