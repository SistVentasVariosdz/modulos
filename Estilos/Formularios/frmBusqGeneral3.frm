VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBusqGeneral3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3900
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8145
      Begin GridEX20.GridEX DGridLista 
         Height          =   3585
         Left            =   75
         TabIndex        =   0
         Top             =   210
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   6324
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         DataMode        =   1
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmBusqGeneral3.frx":0000
         FormatStyle(2)  =   "frmBusqGeneral3.frx":0138
         FormatStyle(3)  =   "frmBusqGeneral3.frx":01E8
         FormatStyle(4)  =   "frmBusqGeneral3.frx":029C
         FormatStyle(5)  =   "frmBusqGeneral3.frx":0374
         FormatStyle(6)  =   "frmBusqGeneral3.frx":042C
         FormatStyle(7)  =   "frmBusqGeneral3.frx":050C
         ImageCount      =   0
         PrinterProperties=   "frmBusqGeneral3.frx":052C
      End
   End
End
Attribute VB_Name = "frmBusqGeneral3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sQuery As String, bCancel As Boolean
Public oParent As Object
Dim Campo As Integer

Private Sub CmdAceptar_Click()
DGridLista_DblClick
'Me.Hide
End Sub

Public Sub Cargar_Datos()
    bCancel = False
    Set DGridLista.ADORecordset = CargarRecordSetDesconectado(sQuery, cCONNECT)
End Sub

Private Sub CmdCancelar_Click()
bCancel = True
Unload Me
End Sub

'Public Sub Form_Unload(Cancel As Integer)
'bCancel = True
'UnloadForm Me
'End Sub

Private Sub DGridLista_DblClick()
With oParent
        'If DGridLista.RowCount > 0 Then DGridLista.ADORecordset.AbsolutePosition = DGridLista.RowIndex(DGridLista.Row)
        .Codigo = DGridLista.Value(DGridLista.Columns(1).Index)
        
        If DGridLista.Columns.Count > 1 Then
            If IsNull(DGridLista.Value(DGridLista.Columns(2).Index)) Then
                .Descripcion = ""
            Else
                .Descripcion = DGridLista.Value(DGridLista.Columns(2).Index)
            End If
        End If
        
        If DGridLista.Columns.Count >= 3 Then
            .TipoAdd = DGridLista.Value(DGridLista.Columns(3).Index)
        End If
        
        If DGridLista.Columns.Count >= 4 Then
            .TipoAdd2 = DGridLista.Value(DGridLista.Columns(4).Index)
        End If
        
    End With
If Me.Visible Then
    Me.Hide
End If
'Unload Me
End Sub

Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        DGridLista_DblClick
    End If
    If KeyCode = vbKeyEscape Then CmdCancelar_Click
End Sub

Private Sub DGridLista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DGridLista_DblClick
    Else
        If DGridLista.RowCount > 0 Then
            Dim rs As New ADODB.Recordset
            Set rs = DGridLista.ADORecordset
            rs.MoveFirst
            Call BuscaCampo(rs, rs(Campo).Name, UCase(Chr(KeyAscii)))
            DGridLista.MoveToBookmark rs.Bookmark
    '        Call DGridLista.Find(1, jgexContains, UCase(Chr(KeyAscii)))
            Set rs = Nothing
        End If
    End If
End Sub

