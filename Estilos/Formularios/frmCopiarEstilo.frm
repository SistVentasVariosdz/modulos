VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmCopiarEstilo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copiar Estilos"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBuscar 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7485
      Begin VB.ComboBox cboCod_TemCli 
         Height          =   315
         Left            =   1950
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   210
         Width           =   2955
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Temporada Origen"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   300
         Width           =   1320
      End
   End
   Begin VB.Frame fraOpciones 
      Height          =   795
      Left            =   0
      TabIndex        =   1
      Top             =   4980
      Width           =   7485
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   480
         Left            =   4050
         TabIndex        =   10
         Top             =   210
         Width           =   1230
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Grabar"
         Height          =   480
         Left            =   2130
         TabIndex        =   9
         Top             =   210
         Width           =   1245
      End
   End
   Begin VB.Frame fraEstilos 
      Caption         =   "Elegir Estilos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4275
      Left            =   0
      TabIndex        =   0
      Top             =   660
      Width           =   7485
      Begin SSDataWidgets_B.SSDBGrid lstEstilosSelec 
         Height          =   3885
         Left            =   4080
         TabIndex        =   3
         Top             =   270
         Width           =   3300
         _Version        =   196617
         DataMode        =   2
         BackColorOdd    =   12648447
         RowHeight       =   423
         ExtraHeight     =   53
         Columns.Count   =   2
         Columns(0).Width=   1746
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "Cod_EstCli"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   4
         Columns(1).Width=   3200
         Columns(1).Caption=   "Descripción"
         Columns(1).Name =   "Des_EstCli"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Style=   4
         _ExtentX        =   5821
         _ExtentY        =   6853
         _StockProps     =   79
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B.SSDBGrid lstEstilos 
         Height          =   3885
         Left            =   120
         TabIndex        =   2
         Top             =   270
         Width           =   3300
         _Version        =   196617
         DataMode        =   2
         BackColorOdd    =   12648447
         RowHeight       =   423
         ExtraHeight     =   53
         Columns.Count   =   2
         Columns(0).Width=   1720
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "Cod_EstCli"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   4
         Columns(1).Width=   3254
         Columns(1).Caption=   "Descripción"
         Columns(1).Name =   "Des_EstCli"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Style=   4
         _ExtentX        =   5821
         _ExtentY        =   6853
         _StockProps     =   79
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdIzqAll 
         Caption         =   "<<"
         Height          =   480
         Left            =   3510
         TabIndex        =   4
         Top             =   2475
         Width           =   480
      End
      Begin VB.CommandButton cmdIzq 
         Caption         =   "<"
         Height          =   480
         Left            =   3510
         TabIndex        =   5
         Top             =   2010
         Width           =   480
      End
      Begin VB.CommandButton cmdDer 
         Caption         =   ">"
         Height          =   480
         Left            =   3510
         TabIndex        =   6
         Top             =   1545
         Width           =   480
      End
      Begin VB.CommandButton cmdDerAll 
         Caption         =   ">>"
         Height          =   480
         Left            =   3510
         TabIndex        =   7
         Top             =   1080
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmCopiarEstilo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim CadConn As ADODB.Connection
Dim varString As String
Dim Rs_Lista As ADODB.Recordset

Public varCod_Cliente As String
Public varCod_TemCli_origen As String

Public Sub CARGA_ESTILOS()
    
    strSQL = "SELECT Cod_EstCli, Des_EstCli FROM tg_estclitem " & _
            " WHERE  Cod_cliente = '" & Me.varCod_Cliente & "' and " & _
                    "Cod_TemCli  = '" & Mid(Me.cboCod_TemCli.Text, 1, 3) & "'"
    
    Me.lstEstilos.RemoveAll
    Me.lstEstilosSelec.RemoveAll
    
    Set Rs_Lista = New ADODB.Recordset
    
    Rs_Lista.ActiveConnection = cCONNECT
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    Rs_Lista.Open strSQL
    
    If Rs_Lista.RecordCount > 0 Then
        Rs_Lista.MoveFirst
        Do Until Rs_Lista.EOF
            varString = Rs_Lista("Cod_EstCli").Value & vbTab & Rs_Lista("Des_EstCli").Value
            Me.lstEstilos.AddItem varString
            Rs_Lista.MoveNext
        Loop
    End If
    
    Rs_Lista.Close
    Set Rs_Lista = Nothing
End Sub

Private Sub cboCod_TemCli_Click()
    Call CARGA_ESTILOS
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo ErrorAceptar:

    Set CadConn = New ADODB.Connection
    CadConn.Open cCONNECT
    Dim j As Integer
    
    Dim varItem As String
    Dim LonItem As Integer
    
    If lstEstilosSelec.Rows > 0 Then
        'Esto es para adicionar las tallas
        For j = 0 To lstEstilosSelec.Rows - 1
            
            lstEstilosSelec.Bookmark = j
            
            strSQL = "EXEC up_copia_estcli_temp '" & _
            Me.varCod_Cliente & "','" & _
            varCod_TemCli_origen & "','" & _
            Mid(cboCod_TemCli.Text, 1, 3) & "','" & _
            Me.lstEstilosSelec.Columns("cod_estcli").Value & "'"
            
            CadConn.Execute strSQL
                
        Next
        Set CadConn = Nothing
        
        Call MsgBox("El Copiado de Estilos fue exitoso", vbInformation, "Copiar Esilos")
   
        Unload Me
    Else
        Call MsgBox("Debe de seleccionar un registro como minimo. Sirvase verificar", vbInformation, "Copiar Esilos")
    End If
    Exit Sub
ErrorAceptar:
    Set CadConn = Nothing
    ErrorHandler Err, "Error Aceptar"
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdDerAll_Click()
    Dim j As Integer
    Me.lstEstilos.Row = 0
    For j = 0 To Me.lstEstilos.Rows
        Me.lstEstilos.Bookmark = j
        If Me.lstEstilos.Rows > 0 And CStr(Me.lstEstilos.Bookmark) >= "0" And CStr(Me.lstEstilos.Columns(0).Value) <> "" Then
            Me.lstEstilosSelec.AddItem lstEstilos.Columns("Cod_EstCli").Value & vbTab & lstEstilos.Columns("Des_EstCli").Value
            If Me.lstEstilos.Rows > 1 Then
                Me.lstEstilos.RemoveItem Me.lstEstilos.AddItemRowIndex(Me.lstEstilos.Bookmark)
                j = -1
            Else
                Me.lstEstilos.RemoveAll
            End If
        Else
            Exit Sub
        End If
    Next
    'Me.lstEstilos.RemoveAll
End Sub

Private Sub cmdDer_Click()
    Dim varAdicion As String
    On Error GoTo ErrComp
    If Me.lstEstilos.Rows > 0 And CStr(Me.lstEstilos.Bookmark) >= "0" And CStr(Me.lstEstilos.Columns(0).Value) <> "" Then
        'Me.lstEstilosSelec.AddItem lstEstilos.Columns("Tipo").Value & vbTab & lstEstilos.Columns("Cod_Tela").Value & vbTab & lstEstilos.Columns("Des_Tela").Value
        varAdicion = lstEstilos.Columns("Cod_EstCli").Value & vbTab & lstEstilos.Columns("Des_EstCli").Value
        If Me.lstEstilos.Rows > 1 Then
            Me.lstEstilos.RemoveItem Me.lstEstilos.AddItemRowIndex(Me.lstEstilos.Bookmark)
        Else
            Me.lstEstilos.RemoveAll
        End If
        Me.lstEstilosSelec.AddItem varAdicion
        
        If Me.lstEstilos.Rows > 0 Then
            lstEstilos.Bookmark = 0
        End If
    Else
        Exit Sub
    End If
    Exit Sub
ErrComp:
    MsgBox "Sirvase volver a seleccionar el registro", vbInformation, "Information"
End Sub

Private Sub cmdIzq_Click()
Dim varAdicion As String
On Error GoTo ErrComp
    If Me.lstEstilosSelec.Rows > 0 And CStr(Me.lstEstilosSelec.Bookmark) >= "0" And CStr(Me.lstEstilosSelec.Columns(0).Value) <> "" Then
        varAdicion = lstEstilosSelec.Columns("Cod_EstCli").Value & vbTab & lstEstilosSelec.Columns("Des_EstCli").Value
        If Me.lstEstilosSelec.Rows > 1 Then
            Me.lstEstilosSelec.RemoveItem Me.lstEstilosSelec.AddItemRowIndex(Me.lstEstilosSelec.Bookmark)
        Else
            Me.lstEstilosSelec.RemoveAll
        End If
        Me.lstEstilos.AddItem varAdicion
        
        If Me.lstEstilosSelec.Rows > 0 Then
            lstEstilosSelec.Bookmark = 0
        End If
    Else
        Exit Sub
    End If
    Exit Sub
ErrComp:
    MsgBox "Sirvase volver a seleccionar el registro", vbInformation, "Information"
End Sub

Private Sub cmdIzqAll_Click()
    Dim j As Integer
    For j = 0 To Me.lstEstilosSelec.Rows
        Me.lstEstilosSelec.Bookmark = j
        If Me.lstEstilosSelec.Rows > 0 And CStr(Me.lstEstilosSelec.Bookmark) >= "0" And CStr(Me.lstEstilosSelec.Columns(0).Value) <> "" Then
            Me.lstEstilos.AddItem lstEstilosSelec.Columns("Cod_EstCli").Value & vbTab & lstEstilosSelec.Columns("Des_EstCli").Value
            If Me.lstEstilosSelec.Rows > 1 Then
                Me.lstEstilosSelec.RemoveItem Me.lstEstilosSelec.AddItemRowIndex(Me.lstEstilosSelec.Bookmark)
            Else
                Me.lstEstilosSelec.RemoveAll
            End If
        Else
            Exit Sub
        End If
    Next
End Sub

Public Sub CARGA_TEMPORADA()
    strSQL = "SELECT Cod_TemCli + ' - ' + Nom_TemCli FROM TG_TEMCLI WHERE Cod_Cliente = '" & Me.varCod_Cliente & "' AND Cod_TemCli <> '" & Me.varCod_TemCli_origen & "'"
    Call LlenaCombo(Me.cboCod_TemCli, strSQL, cCONNECT)
End Sub

