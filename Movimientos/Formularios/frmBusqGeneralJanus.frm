VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBusqGeneralJanus 
   Caption         =   "Busqueda"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nueva Dimension"
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
      Left            =   105
      TabIndex        =   3
      Tag             =   "&OK"
      Top             =   4245
      Visible         =   0   'False
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
      Left            =   2520
      TabIndex        =   2
      Tag             =   "&OK"
      Top             =   4245
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
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
      Left            =   3780
      TabIndex        =   1
      Tag             =   "&Cancel"
      Top             =   4245
      Width           =   1215
   End
   Begin GridEX20.GridEX gexLista 
      Height          =   4050
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   7144
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmBusqGeneralJanus.frx":0000
      FormatStyle(2)  =   "frmBusqGeneralJanus.frx":0138
      FormatStyle(3)  =   "frmBusqGeneralJanus.frx":01E8
      FormatStyle(4)  =   "frmBusqGeneralJanus.frx":029C
      FormatStyle(5)  =   "frmBusqGeneralJanus.frx":0374
      FormatStyle(6)  =   "frmBusqGeneralJanus.frx":042C
      FormatStyle(7)  =   "frmBusqGeneralJanus.frx":050C
      ImageCount      =   0
      PrinterProperties=   "frmBusqGeneralJanus.frx":052C
   End
End
Attribute VB_Name = "frmBusqGeneralJanus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public sQuery As String
'Dim Rs_Carga As New ADODB.Recordset
Sub Cargar_Datos()
    On Error GoTo Err_CARGA_GRID

    Screen.MousePointer = vbHourglass
   
    Set gexLista.ADORecordset = CargarRecordSetDesconectado(sQuery, cConnect)
    SetGeneralGridEX gexLista, 0, 1
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
Err_CARGA_GRID:
    Screen.MousePointer = vbDefault
    MsgBox "Ocurrio un error en CARGA_GRID", vbCritical, "Mensaje"
End Sub

'Private Sub CmdNuevo_Click()
'    Load frmMantDimCaj
'    frmMantDimCaj.CARGA_GRID
'    frmMantDimCaj.Show vbModal
'    Set frmMantDimCaj = Nothing
'    Unload Me
'End Sub

Private Sub gexLista_DblClick()
On Error GoTo ErrGrilla
    'If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
    If Me.gexLista.RowCount > 0 Then
        With oParent
            .CODIGO = gexLista.Value(1)
            If gexLista.Columns.Count > 1 Then
                .DESCRIPCION = gexLista.Value(2)
                If gexLista.Columns.Count > 2 Then
                    .varCod_EstPro = gexLista.Value(3)
                End If
            End If
        End With
    End If
    Unload Me
    Exit Sub
ErrGrilla:
    Unload Me
End Sub

Private Sub gexLista_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gexLista_DblClick
    End If
End Sub

Private Sub gexlista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        gexLista_DblClick
    Else
        'Call BuscaCampo(Rs_Carga, Rs_Carga(0).Name, UCase(Chr(KeyAscii)))
        Call gexLista.Find(1, jgexContains, UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub Form_Load()
    SetGeneralGridEX gexLista, 0, 1
'If frmRealizaDespacho.Pista = 1 Or frmRealizaDespachoExc.Pista = 1 Then
'    CmdNuevo.Visible = True
'Else
'    CmdNuevo.Visible = False
'End If

End Sub

Private Sub cmdAceptar_Click()
'    If frmRealizaDespacho.Pista = 1 Or frmRealizaDespachoExc.Pista = 1 Then
'    End If
    gexLista_DblClick
    
End Sub
Private Sub cmdCancelar_Click()
    oParent.CODIGO = ""
    Unload Me
End Sub




