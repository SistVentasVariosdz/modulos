VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBusqGeneral_Lis 
   Caption         =   "Búsqueda"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5595
      TabIndex        =   1
      Tag             =   "&Cancel"
      Top             =   3465
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4350
      TabIndex        =   0
      Tag             =   "&OK"
      Top             =   3465
      Width           =   1215
   End
   Begin GridEX20.GridEX DGridLista 
      Height          =   3300
      Left            =   135
      TabIndex        =   2
      Top             =   75
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   5821
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmBusqGeneral_Lis.frx":0000
      FormatStyle(2)  =   "frmBusqGeneral_Lis.frx":0138
      FormatStyle(3)  =   "frmBusqGeneral_Lis.frx":01E8
      FormatStyle(4)  =   "frmBusqGeneral_Lis.frx":029C
      FormatStyle(5)  =   "frmBusqGeneral_Lis.frx":0374
      FormatStyle(6)  =   "frmBusqGeneral_Lis.frx":042C
      FormatStyle(7)  =   "frmBusqGeneral_Lis.frx":050C
      ImageCount      =   0
      PrinterProperties=   "frmBusqGeneral_Lis.frx":052C
   End
End
Attribute VB_Name = "frmBusqGeneral_Lis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oParent As Object
Public SQuery As String, Campo As Integer
Dim Rs_Carga As Object

Sub CARGAR_DATOS()
On Error GoTo Cargar_DatosErr

Set DGridLista.ADORecordset = CargarRecordSetDesconectado(SQuery, cCONNECT)

If DGridLista.Columns.Count = 1 Then
    DGridLista.Columns(1).Width = 6000
End If
If DGridLista.Columns.Count = 2 Then
    DGridLista.Columns(2).Width = 4000
End If

With oParent
    .codigo = ""
    .Descripcion = ""
    
'    .Linea1 = ""
'
'    .Linea2 = ""
       
    
    If DGridLista.Columns.Count >= 3 Then
        .estado = ""
    End If
    
    If DGridLista.Columns.Count >= 4 Then
        .Tipoa = ""
    End If
    
    If DGridLista.Columns.Count >= 5 Then
        .Tipob = ""
    End If
    
End With
Exit Sub
Cargar_DatosErr:
    Set Rs_Carga = Nothing
    ErrorHandler err, "Cargar_Datos"
End Sub

Private Sub DGridlista_DblClick()
    With oParent
        .codigo = DGridLista.Value(DGridLista.Columns(1).Index)
        
        If DGridLista.Columns.Count > 1 Then
            If IsNull(DGridLista.Value(DGridLista.Columns(2).Index)) Then
                .Descripcion = ""
            Else
                .Descripcion = DGridLista.Value(DGridLista.Columns(2).Index)
            End If
        End If
        
        If DGridLista.Columns.Count >= 3 Then
            .estado = DGridLista.Value(DGridLista.Columns(3).Index)
        End If
        
        
        If DGridLista.Columns.Count >= 4 Then
            .Tipoa = DGridLista.Value(DGridLista.Columns(4).Index)
        End If
        
        If DGridLista.Columns.Count >= 5 Then
            .Tipob = DGridLista.Value(DGridLista.Columns(5).Index)
        End If
        
    End With
        bCancel = False
Unload Me
End Sub

Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        DGridlista_DblClick
    End If
End Sub

Private Sub DGridLista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DGridlista_DblClick
    Else
        If DGridLista.RowCount > 0 Then
            Dim RS As Object
            Set RS = DGridLista.ADORecordset
            RS.MoveFirst
            Call BuscaCampo(RS, RS(Campo).Name, UCase(Chr(KeyAscii)))
            DGridLista.MoveToBookmark RS.Bookmark
    '        Call DGridLista.Find(1, jgexContains, UCase(Chr(KeyAscii)))
            Set RS = Nothing
        End If
    End If
End Sub

Private Sub Form_Activate()
    DGridLista.SetFocus
End Sub

Private Sub Form_Load()
Set RS = CreateObject("ADODB.Recordset")

Call FormSet(Me)
   SetGeneralGridEX DGridLista, 0, 1
   Campo = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Rs_Carga = Nothing
End Sub
Public Sub cmdAceptar_Click()
    DGridlista_DblClick
End Sub
Private Sub cmdCancelar_Click()
oParent.codigo = ""
Unload Me
End Sub


