VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form frmBusqPartidasLote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda Partidas Lote"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAsigna 
      Height          =   1485
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   3660
      Begin VB.TextBox txtCod_OrdProv 
         Height          =   285
         Left            =   1890
         MaxLength       =   15
         TabIndex        =   7
         Top             =   510
         Width           =   1575
      End
      Begin VB.CommandButton cmdNoAsignar 
         Caption         =   "Ca&ncelar"
         Height          =   480
         Left            =   1920
         TabIndex        =   5
         Top             =   915
         Width           =   1230
      End
      Begin VB.CommandButton cmdAsignar 
         Caption         =   "A&signar"
         Height          =   495
         Left            =   435
         TabIndex        =   4
         Top             =   900
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Asignación de Orden de Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   345
         TabIndex        =   8
         Top             =   195
         Width           =   2985
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Orden de Proveedor :"
         Height          =   195
         Left            =   285
         TabIndex        =   6
         Top             =   555
         Width           =   1530
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   510
      Left            =   3885
      TabIndex        =   2
      Top             =   3270
      Width           =   1395
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   510
      Left            =   2295
      TabIndex        =   1
      Top             =   3270
      Width           =   1395
   End
   Begin GridEX20.GridEX gexLista 
      Height          =   3000
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5292
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmBusqPartidasLote.frx":0000
      Column(2)       =   "frmBusqPartidasLote.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmBusqPartidasLote.frx":016C
      FormatStyle(2)  =   "frmBusqPartidasLote.frx":02A4
      FormatStyle(3)  =   "frmBusqPartidasLote.frx":0354
      FormatStyle(4)  =   "frmBusqPartidasLote.frx":0408
      FormatStyle(5)  =   "frmBusqPartidasLote.frx":04E0
      FormatStyle(6)  =   "frmBusqPartidasLote.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmBusqPartidasLote.frx":0678
   End
End
Attribute VB_Name = "frmBusqPartidasLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public varSer_OrdComp As String
Public varCod_OrdComp As String

Public oParent As Object
Public varBusqueda As String

Sub CARGA_GRID()
    
    'Esta cadena es para devolver el Codigo de Cliente
    strSQL = "EXEC UP_SEL_PARTIDASLOTE '" & Me.varSer_OrdComp & "','" & Me.varCod_OrdComp & "'"
    
    Set gexLista.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    
    SetGeneralGridEX gexLista, 0, 1
    
    Me.gexLista.MoveLast
    Call Me.gexLista.Find(4, jgexEqual, varBusqueda)
    
    Call Configurar_Grid
    
End Sub

Public Function VALIDA_DATOS() As Boolean
    Dim varOpcion As Integer
    VALIDA_DATOS = True
    If Trim(gexLista.Value(gexLista.Columns("Cod_OrdProv").Index)) = "" Then
        varOpcion = MsgBox("El lote seleccionado no posee una Orden de Proveedor. Desea asignarle una?", vbInformation + vbYesNo + vbDefaultButton2, "Mensaje")
        If varOpcion = vbYes Then
        
            Me.varBusqueda = gexLista.Value(gexLista.Columns("ORDTRA").Index)
        
            gexLista.Enabled = False
            Me.txtCod_OrdProv.Text = ""
            fraAsigna.Visible = True
            Me.txtCod_OrdProv.SetFocus
            VALIDA_DATOS = False
        Else
            VALIDA_DATOS = False
        End If
    End If
End Function

Public Function VALIDA_COD_ORDPROV() As Boolean
On Error GoTo Err_Valida:
    VALIDA_COD_ORDPROV = True
    If Trim(Me.txtCod_OrdProv.Text) = "" Then
        VALIDA_COD_ORDPROV = False
        MsgBox "La partida del Proveedor no puede estar vacia. Sirvase verificar", vbInformation, "Mensaje"
        Exit Function
    End If
    
    strSQL = "SELECT COUNT(*) FROM TX_ORDTRA WHERE Cod_TipOrdTra = '" & gexLista.Value(gexLista.Columns("Cod_TipOrdTra").Index) & "' AND " & _
             "Cod_Proveedor = '" & gexLista.Value(gexLista.Columns("Cod_Proveedor").Index) & "' AND Cod_OrdProv = '" & Trim(Me.txtCod_OrdProv.Text) & "'"
    If DevuelveCampo(strSQL, cConnect) > 0 Then
        VALIDA_COD_ORDPROV = False
        MsgBox "La partida del Proveedor ingresada ya existe. Sirvase verificar", vbInformation, "Mensaje"
        Exit Function
    End If
    Exit Function
Err_Valida:
    VALIDA_COD_ORDPROV = False
    MsgBox "Ocurrio un error en la validación de datos. Sirvase verificar", vbCritical, "Mensaje"
End Function

Sub Salvar_Datos()
    Dim Con As New ADODB.Connection
    Dim Rs As New ADODB.Recordset
    On Error GoTo Salvar_DatosErr
    Dim strSQL As String
    
    Con.ConnectionString = cConnect
    Con.Open
    
        Con.BeginTrans

        strSQL = "EXEC UP_ACTUALIZA_Cod_OrdProv '" & _
        gexLista.Value(gexLista.Columns("Cod_TipOrdTra").Index) & "','" & _
        gexLista.Value(gexLista.Columns("Cod_Ordtra").Index) & "','" & _
        Trim(Me.txtCod_OrdProv.Text) & "'"

        Con.Execute strSQL
       
        Con.CommitTrans
        
        'Dim amensaje As New clsMessages
        'amensaje.Codigo = CodeMsg.kMSG_INF_DATA_SAVE
        'Informa "", amensaje
        
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler err, "Salvar_Datos"
End Sub

Private Sub cmdAceptar_Click()
    If VALIDA_DATOS Then
        If gexLista.RowCount > 0 Then
            With oParent
                .Codigo = gexLista.Value(gexLista.Columns("Cod_OrdProv").Index)
                .varCod_TipOrdTra = gexLista.Value(gexLista.Columns("Cod_TipOrdTra").Index)
                .varCod_OrdTra = gexLista.Value(gexLista.Columns("Cod_Ordtra").Index)
                .txtLote.Text = gexLista.Value(gexLista.Columns("Cod_OrdProv").Index)
                .Descripcion = ""
            End With
        End If
        Unload Me
    End If
End Sub

Private Sub cmdAsignar_Click()
On Error GoTo Err_Asignar
    If VALIDA_COD_ORDPROV Then
        Call Me.Salvar_Datos
        
        'Aqui ya hemos grabado la data
        With oParent
            .Codigo = Trim(Me.txtCod_OrdProv.Text)
            .varCod_TipOrdTra = gexLista.Value(gexLista.Columns("Cod_TipOrdTra").Index)
            .varCod_OrdTra = gexLista.Value(gexLista.Columns("Cod_Ordtra").Index)
            .txtLote.Text = Trim(Me.txtCod_OrdProv.Text)
            .Descripcion = ""
        End With
        
        gexLista.Enabled = True
        
        Unload Me
        
        'Call Me.CARGA_GRID
        'Call cmdAceptar_Click
    End If
Exit Sub
Err_Asignar:
    MsgBox "Hubo un error en la asignacion. Sirvase verificar", vbCritical, "Mensaje"
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdNoAsignar_Click()
    Me.fraAsigna.Visible = False
    gexLista.Enabled = True
    Me.txtCod_OrdProv.Text = ""
End Sub

Private Sub gexLista_DblClick()
    Call cmdAceptar_Click
End Sub

Private Sub gexLista_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call cmdAceptar_Click
        KeyCode = 0
    End If
End Sub

Public Sub Configurar_Grid()
    Me.gexLista.Columns("Cod_color").Visible = False
    Me.gexLista.Columns("Des_Color").Visible = False
    Me.gexLista.Columns("Cod_TipOrdTra").Visible = False
    Me.gexLista.Columns("Cod_Ordtra").Visible = False
    Me.gexLista.Columns("Cod_Proveedor").Visible = False
    Me.gexLista.Columns("Des_Proveedor").Visible = False
    
    Me.gexLista.Columns("COLOR").Caption = "Color"
    Me.gexLista.Columns("COLOR").Width = 2000
    Me.gexLista.Columns("ORDTRA").Caption = "O/T"
    Me.gexLista.Columns("ORDTRA").Width = 1300
    Me.gexLista.Columns("PROVEEDOR").Caption = "Proveedor"
    Me.gexLista.Columns("PROVEEDOR").Width = 2500
    Me.gexLista.Columns("Cod_OrdProv").Caption = "Nro. Partida"
    Me.gexLista.Columns("Cod_OrdProv").Width = 1300

End Sub

Private Sub txtCod_OrdProv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAsignar.SetFocus
    End If
End Sub
