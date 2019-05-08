VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmMantEstCliColAll 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adicion de Colores"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Colores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Left            =   30
      TabIndex        =   2
      Top             =   360
      Width           =   6480
      Begin SSDataWidgets_B.SSDBGrid lstColoresSelec 
         Height          =   3405
         Left            =   3540
         TabIndex        =   9
         Top             =   240
         Width           =   2775
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   2
         BackColorOdd    =   12648447
         RowHeight       =   423
         ExtraHeight     =   79
         Columns.Count   =   2
         Columns(0).Width=   1720
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "Cod_ColCli"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   4
         Columns(1).Width=   3200
         Columns(1).Caption=   "Descripción"
         Columns(1).Name =   "Nom_ColCli"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Style=   4
         _ExtentX        =   4895
         _ExtentY        =   6006
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
      Begin SSDataWidgets_B.SSDBGrid LstColores 
         Height          =   3405
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2775
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   2
         BackColorOdd    =   12648447
         RowHeight       =   423
         ExtraHeight     =   79
         Columns.Count   =   2
         Columns(0).Width=   1667
         Columns(0).Caption=   "Código"
         Columns(0).Name =   "Cod_ColCli"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   4
         Columns(1).Width=   3200
         Columns(1).Caption=   "Descripción"
         Columns(1).Name =   "Nom_ColCli"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Style=   4
         _ExtentX        =   4895
         _ExtentY        =   6006
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
         Left            =   2970
         TabIndex        =   3
         Top             =   2430
         Width           =   480
      End
      Begin VB.CommandButton cmdIzq 
         Caption         =   "<"
         Height          =   480
         Left            =   2970
         TabIndex        =   4
         Top             =   1965
         Width           =   480
      End
      Begin VB.CommandButton cmdDer 
         Caption         =   ">"
         Height          =   480
         Left            =   2970
         TabIndex        =   5
         Top             =   1500
         Width           =   480
      End
      Begin VB.CommandButton cmdDerAll 
         Caption         =   ">>"
         Height          =   480
         Left            =   2970
         TabIndex        =   6
         Top             =   1035
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Grabar"
      Height          =   480
      Left            =   1665
      TabIndex        =   1
      Top             =   4305
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   480
      Left            =   3585
      TabIndex        =   0
      Top             =   4305
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Selección de Colores"
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
      Left            =   2205
      TabIndex        =   7
      Top             =   75
      Width           =   1815
   End
End
Attribute VB_Name = "frmMantEstCliColAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Dim strSQL As String
Dim Rs_Lista As ADODB.Recordset
Dim CadConn  As New ADODB.Connection
Dim varString As String
'Variables pasadas como parametro
Public varCod_Cliente As String
Public varCod_TemCli As String
Public varCod_EstCli As String

Public Sub CARGA_COLORES()

    'Aqui llenamos con los colores pendientes
    'strSQL = "SELECT Cod_ColCli, Nom_ColCli FROM TG_COLCLITEM WHERE Cod_Cliente = '" & varCod_Cliente & "' AND Cod_TemCli = '" & varCod_TemCli & "'"
    strSQL = "EXEC UP_SEL_TRAE_COLORES_ESTCLI '" & varCod_Cliente & "','" & varCod_TemCli & "','" & varCod_EstCli & "'"
    'Call LlenaCombo(Me.LstColores, strSQL, cCONNECT)
    
    Me.LstColores.RemoveAll
    Me.lstColoresSelec.RemoveAll
    
    Set Rs_Lista = New ADODB.Recordset
    
    Rs_Lista.ActiveConnection = cCONNECT
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    Rs_Lista.Open strSQL
    
    If Rs_Lista.RecordCount > 0 Then
        Rs_Lista.MoveFirst
        Do Until Rs_Lista.EOF
            varString = Rs_Lista("Cod_ColCli").Value & vbTab & Rs_Lista("Nom_ColCli").Value
            Me.LstColores.AddItem varString
            Rs_Lista.MoveNext
        Loop
    End If
    
    Rs_Lista.Close
    Set Rs_Lista = Nothing
    
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo ErrorAceptar:

    Set CadConn = Nothing
    CadConn.Open cCONNECT
    Dim j As Integer
    
    Dim varItem As String
    Dim LonItem As Integer
    
    'Esto es para adicionar las tallas
    For j = 0 To lstColoresSelec.Rows - 1
        lstColoresSelec.Bookmark = j
        'Strsql = "SELECT Nom_ColCli FROM TG_COLCLITEM WHERE Cod_Cliente = '" & varCod_Cliente & "' AND Cod_TemCli = '" & varCod_TemCli & "' AND Cod_ColCli='" & lstColoresSelec.List(j) & "'"
        'Aqui se pone el codigo para la eliminacion
        
        strSQL = "EXEC UP_MAN_ESTCLICOL " & _
        "I" & ",'" & _
        varCod_Cliente & "','" & _
        varCod_TemCli & "','" & _
        varCod_EstCli & "','" & _
        lstColoresSelec.Columns("Cod_ColCli").Value & "','" & _
        lstColoresSelec.Columns("Nom_ColCli").Value & "'"
        
        'Strsql = "EXEC UP_MAN_TALLASDET 'I','" & varCod_grutal & "','" & varItem & "',0"
        CadConn.Execute strSQL
            
    Next
    Set CadConn = Nothing
    
    Call MsgBox("La actualizacion de Colores fue exitosa", vbInformation, "Tallas")
    'Call Me.CARGA_COLORES
    
    Unload Me
    Exit Sub
ErrorAceptar:
    Set CadConn = Nothing
    ErrorHandler Err, "Error Aceptar"


End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub lstcolores_DblClick()
    cmdDer_Click
End Sub

Private Sub lstcoloresSelec_DblClick()
    cmdIzq_Click
End Sub

Private Sub cmdDerAll_Click()
    Dim j As Integer
    Me.LstColores.Row = 0
    For j = 0 To Me.LstColores.Rows
        Me.LstColores.Bookmark = j
        If Me.LstColores.Rows > 0 And CStr(Me.LstColores.Bookmark) >= "0" And CStr(Me.LstColores.Columns(0).Value) <> "" Then
            Me.lstColoresSelec.AddItem LstColores.Columns("Cod_ColCli").Value & vbTab & LstColores.Columns("Nom_ColCli").Value
            If Me.LstColores.Rows > 1 Then
                Me.LstColores.RemoveItem Me.LstColores.AddItemRowIndex(Me.LstColores.Bookmark)
                j = -1
            Else
                Me.LstColores.RemoveAll
            End If
        Else
            Exit Sub
        End If
    Next
    'Me.lstColores.RemoveAll
End Sub

Private Sub cmdDer_Click()
    Dim varAdicion As String
    On Error GoTo ErrComp
    If Me.LstColores.Rows > 0 And CStr(Me.LstColores.Bookmark) >= "0" And CStr(Me.LstColores.Columns(0).Value) <> "" Then
        'Me.lstColoresSelec.AddItem lstColores.Columns("Tipo").Value & vbTab & lstColores.Columns("Cod_Tela").Value & vbTab & lstColores.Columns("Des_Tela").Value
        varAdicion = LstColores.Columns("Cod_ColCli").Value & vbTab & LstColores.Columns("Nom_ColCli").Value
        If Me.LstColores.Rows > 1 Then
            Me.LstColores.RemoveItem Me.LstColores.AddItemRowIndex(Me.LstColores.Bookmark)
        Else
            Me.LstColores.RemoveAll
        End If
        Me.lstColoresSelec.AddItem varAdicion
        
        If Me.LstColores.Rows > 0 Then
            LstColores.Bookmark = 0
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
    If Me.lstColoresSelec.Rows > 0 And CStr(Me.lstColoresSelec.Bookmark) >= "0" And CStr(Me.lstColoresSelec.Columns(0).Value) <> "" Then
        varAdicion = lstColoresSelec.Columns("Cod_ColCli").Value & vbTab & lstColoresSelec.Columns("Nom_ColCli").Value
        If Me.lstColoresSelec.Rows > 1 Then
            Me.lstColoresSelec.RemoveItem Me.lstColoresSelec.AddItemRowIndex(Me.lstColoresSelec.Bookmark)
        Else
            Me.lstColoresSelec.RemoveAll
        End If
        Me.LstColores.AddItem varAdicion
        
        If Me.lstColoresSelec.Rows > 0 Then
            lstColoresSelec.Bookmark = 0
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
    For j = 0 To Me.lstColoresSelec.Rows
        Me.lstColoresSelec.Bookmark = j
        If Me.lstColoresSelec.Rows > 0 And CStr(Me.lstColoresSelec.Bookmark) >= "0" And CStr(Me.lstColoresSelec.Columns(0).Value) <> "" Then
            Me.LstColores.AddItem lstColoresSelec.Columns("Cod_ColCli").Value & vbTab & lstColoresSelec.Columns("Nom_ColCli").Value
            If Me.lstColoresSelec.Rows > 1 Then
                Me.lstColoresSelec.RemoveItem Me.lstColoresSelec.AddItemRowIndex(Me.lstColoresSelec.Bookmark)
            Else
                Me.lstColoresSelec.RemoveAll
            End If
        Else
            Exit Sub
        End If
    Next
End Sub


