VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBusqGeneral_Cliente 
   Caption         =   "Busqueda"
   ClientHeight    =   4170
   ClientLeft      =   4140
   ClientTop       =   3795
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7005
   Visible         =   0   'False
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   840
      Top             =   3960
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      Left            =   4335
      TabIndex        =   1
      Tag             =   "&OK"
      Top             =   3600
      Width           =   1215
   End
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
      Left            =   5580
      TabIndex        =   0
      Tag             =   "&Cancel"
      Top             =   3600
      Width           =   1215
   End
   Begin GridEX20.GridEX DGridLista 
      Height          =   3300
      Left            =   120
      TabIndex        =   2
      Top             =   210
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
      FormatStyle(1)  =   "frmBusqGeneral_cliente1.frx":0000
      FormatStyle(2)  =   "frmBusqGeneral_cliente1.frx":0138
      FormatStyle(3)  =   "frmBusqGeneral_cliente1.frx":01E8
      FormatStyle(4)  =   "frmBusqGeneral_cliente1.frx":029C
      FormatStyle(5)  =   "frmBusqGeneral_cliente1.frx":0374
      FormatStyle(6)  =   "frmBusqGeneral_cliente1.frx":042C
      FormatStyle(7)  =   "frmBusqGeneral_cliente1.frx":050C
      ImageCount      =   0
      PrinterProperties=   "frmBusqGeneral_cliente1.frx":052C
   End
End
Attribute VB_Name = "frmBusqGeneral_Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oParent As Object
Public sQuery As String
Dim Rs_Carga As New ADODB.Recordset

Sub CARGAR_DATOS()
On Error GoTo Cargar_DatosErr

Set DGridLista.ADORecordset = CargarRecordSetDesconectado(sQuery, cConnect)

If DGridLista.Columns.Count = 2 Then
    DGridLista.Columns(2).Width = 4000
End If

With oParent
    .Codigo = ""
    .Descripcion = ""
    
    If DGridLista.Columns.Count = 3 Then
        .TipoAdd = ""
    End If
    
End With
Exit Sub
Cargar_DatosErr:
    Set Rs_Carga = Nothing
    ErrorHandler Err, "Cargar_Datos"
End Sub

Private Sub DGridlista_DblClick()
    With oParent
        If DGridLista.RowCount > 0 Then DGridLista.ADORecordset.AbsolutePosition = DGridLista.RowIndex(DGridLista.Row)
        'DGridLista.ADORecordset.AbsolutePosition = DGridLista.RowIndex(DGridLista.Row)
        
        .Codigo = DGridLista.Value(DGridLista.Columns(1).Index)
        
        If DGridLista.Columns.Count > 1 Then
            If IsNull(DGridLista.Value(DGridLista.Columns(2).Index)) Then
                .Descripcion = ""
            Else
                .Descripcion = DGridLista.Value(DGridLista.Columns(2).Index)
            End If
        End If
        
        If DGridLista.Columns.Count = 3 Then
            .TipoAdd = DGridLista.Value(DGridLista.Columns(3).Index)
        End If
        
    End With
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
            Dim Rs As New ADODB.Recordset
            Set Rs = DGridLista.ADORecordset
            Rs.MoveFirst
            Call BuscaCampo(Rs, Rs(0).Name, UCase(Chr(KeyAscii)))
            DGridLista.MoveToBookmark Rs.Bookmark
    '        Call DGridLista.Find(1, jgexContains, UCase(Chr(KeyAscii)))
            Set Rs = Nothing
        End If
    End If
End Sub

Private Sub Form_Activate()

    DGridLista.SetFocus
End Sub

Private Sub Form_Load()
Call FormSet(Me)
   SetGeneralGridEX DGridLista, 0, 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Rs_Carga = Nothing
End Sub
Private Sub cmdAceptar_Click()
    DGridlista_DblClick
End Sub
Private Sub cmdCancelar_Click()
oParent.Codigo = ""
Unload Me
End Sub

