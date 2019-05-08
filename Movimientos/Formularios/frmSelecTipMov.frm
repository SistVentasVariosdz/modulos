VERSION 5.00
Begin VB.Form frmSelectipMov 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de Movimientos"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   Icon            =   "frmSelecTipMov.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   525
      Left            =   2880
      TabIndex        =   3
      Top             =   4560
      Width           =   1245
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   555
      Left            =   1200
      TabIndex        =   2
      Top             =   4560
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Caption         =   "Movimientos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4230
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   5010
      Begin VB.ListBox LstTipMov 
         Height          =   3660
         Left            =   300
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   330
         Width           =   4515
      End
   End
End
Attribute VB_Name = "frmSelectipMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Dim strSQL As String
Dim Rs_Lista As ADODB.Recordset
'Public varSer_OrdComp As String
Public varCod_Almacen As String
Public varTipoReq As String

Private Sub cmdAceptar_Click()
   oParent.varCadena_Movs = CADENA_MOV(Me.LstTipMov)
   Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
    oParent.varCadena_Movs = ""
End Sub

Public Sub CARGA_MOVIMIENTOS()
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.ActiveConnection = cConnect
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    strSQL = "select a.cod_tipmov, b.des_tipmov from lg_tipmovIalm a, lg_tiposmov b " & _
            "where a.cod_tipmov=b.cod_tipmov and a.cod_almacen='" & varCod_Almacen & "'"
        
            
    Rs_Lista.Open strSQL
    Me.LstTipMov.Clear
    If Rs_Lista.RecordCount > 0 Then
        Rs_Lista.MoveFirst
        While Not Rs_Lista.EOF And Not Rs_Lista.BOF
            Me.LstTipMov.AddItem Rs_Lista.Fields("cod_tipmov").Value & " - " & Rs_Lista.Fields("des_tipmov").Value
            Rs_Lista.MoveNext
        Wend
    Else
        Me.LstTipMov.Enabled = False
        Exit Sub
    End If
    
    Rs_Lista.Close
    Set Rs_Lista = Nothing
    
End Sub

Public Function CADENA_MOV(ByRef pListBox As Object) As String
    Dim Contador As Integer
    Dim Cadena As String    'Este es el prov del resultado
    
    Cadena = ""
    If pListBox.ListCount < 1 Then
        CADENA_MOV = ""
        Exit Function
    End If
    For Contador = 0 To pListBox.ListCount - 1
        If pListBox.Selected(Contador) Then
            Cadena = Cadena & "." & Trim(Left(pListBox.List(Contador), 3)) & ".,"
        End If
    Next
    Cadena = Trim(Cadena)
    If Cadena <> "" Then
        Cadena = Mid(Cadena, 1, Len(Cadena) - 1)
    End If
    
    CADENA_MOV = Cadena
    
End Function

