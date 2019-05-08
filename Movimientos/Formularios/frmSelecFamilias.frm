VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmSelecFamilias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar Familias"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Orden"
      Height          =   1020
      Left            =   195
      TabIndex        =   3
      Top             =   3885
      Visible         =   0   'False
      Width           =   5190
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   2070
         TabIndex        =   7
         Top             =   195
         Width           =   1275
      End
      Begin VB.OptionButton OptProveedor 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   2820
         TabIndex        =   6
         Top             =   600
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.OptionButton OptItem 
         Caption         =   "Item"
         Height          =   195
         Left            =   690
         TabIndex        =   5
         Top             =   615
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Cod Proveedor"
         Height          =   270
         Left            =   810
         TabIndex        =   4
         Top             =   195
         Width           =   1140
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Familias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3705
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   5355
      Begin VB.ListBox LstFamilias 
         Height          =   3210
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   270
         Width           =   5085
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1635
      TabIndex        =   1
      Top             =   5025
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmSelecFamilias.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmSelecFamilias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sFam As String
Public ExisteHilo As String
Dim strSQL As String, rstLista As ADODB.Recordset
Public vCod_Almacen As String
Public orderby As Integer
Public sCod_Prov As String

Private Sub Form_Load()
'    LlenaCombo LstFamilias, "SELECT Des_FamItem + SPACE(100) + Cod_FamItem FROM LG_FAMITE ORDER BY 1", cConnect
End Sub

Public Function Cadena_Fam(ByRef pListBox As Object) As String
Dim Contador As Integer
Dim Cadena As String    'Este es el prov del resultado
    
    Cadena = ""
    If pListBox.ListCount < 1 Then
        Cadena_Fam = ""
        Exit Function
    End If
    For Contador = 0 To pListBox.ListCount - 1
        If pListBox.Selected(Contador) Then
            If Trim(Right(pListBox.List(Contador), 3)) = "HI" Then
                ExisteHilo = "S"
                Cadena = "." & Trim(Right(pListBox.List(Contador), 3)) & "."
                Cadena_Fam = Cadena
                sCod_Prov = Text1.Text
                If OptItem.Value Then
                    orderby = 0
                    Else
                    orderby = 1
                End If
                Exit Function
            End If
                Cadena = Cadena & "." & Trim(Right(pListBox.List(Contador), 3)) & ".,"
        End If
    Next
    Cadena = Trim(Cadena)
    If Cadena <> "" Then
        Cadena = Mid(Cadena, 1, Len(Cadena) - 1)
    End If
    ExisteHilo = "N"
    Cadena_Fam = Cadena

End Function

Private Sub Form_Unload(Cancel As Integer)
    If Me.Visible Then
        Me.Hide
        Cancel = 200
    End If
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "ACEPTAR"
        sFam = Cadena_Fam(LstFamilias)
        Unload Me
    Case "CANCELAR"
        sFam = ""
        Unload Me
    End Select
End Sub

Private Sub LstFamilias_Click()
    If Trim(Right(LstFamilias.List(LstFamilias.ListIndex), 3)) = "HI" Then
        Frame1.Visible = True
    Else
        Frame1.Visible = False
    End If

End Sub

Private Sub LstFamilias_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    
End Sub


Sub carga_lista()
    LlenaCombo LstFamilias, "LG_MUESTRA_FAMILIAS_ITEMS_CON_STOCK '" & vCod_Almacen & "'", cConnect
End Sub
