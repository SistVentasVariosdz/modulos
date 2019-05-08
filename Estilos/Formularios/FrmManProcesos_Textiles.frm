VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmManProcesos_Textiles 
   Caption         =   "Mantenimiento "
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3765
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   4710
      Begin GridEX20.GridEX GridReg 
         Height          =   3525
         Left            =   60
         TabIndex        =   0
         Top             =   165
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   6218
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "FrmManProcesos_Textiles.frx":0000
         Column(2)       =   "FrmManProcesos_Textiles.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "FrmManProcesos_Textiles.frx":016C
         FormatStyle(2)  =   "FrmManProcesos_Textiles.frx":02A4
         FormatStyle(3)  =   "FrmManProcesos_Textiles.frx":0354
         FormatStyle(4)  =   "FrmManProcesos_Textiles.frx":0408
         FormatStyle(5)  =   "FrmManProcesos_Textiles.frx":04E0
         FormatStyle(6)  =   "FrmManProcesos_Textiles.frx":0598
         ImageCount      =   0
         PrinterProperties=   "FrmManProcesos_Textiles.frx":0678
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1245
      Left            =   0
      TabIndex        =   6
      Top             =   3840
      Width           =   4695
      Begin VB.TextBox txtDesProcesoTinto 
         Enabled         =   0   'False
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
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   4
         Top             =   840
         Width           =   2480
      End
      Begin VB.TextBox txtCodProcesoTinto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   3
         Top             =   840
         Width           =   795
      End
      Begin VB.TextBox TxtCodigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   1
         Top             =   195
         Width           =   465
      End
      Begin VB.TextBox TxtDescripcion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   2
         Top             =   480
         Width           =   3300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Proc.Tintoreria :"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   945
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
         Height          =   225
         Left            =   150
         TabIndex        =   8
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   585
         Width           =   915
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   600
      TabIndex        =   5
      Top             =   5160
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmManProcesos_Textiles.frx":0850
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "FrmManProcesos_Textiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Codigo As String
Public Descripcion As String
Public TipoAdd As String
Public tipoAdd2 As String

Dim strSQL As String
Dim sTipo As String
Public vCodigo As String


Dim posProcTinto As Integer

Private Sub GridReg_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    TxtCodigo = GridReg.Value(GridReg.Columns("Codigo").Index)
    TxtDescripcion = GridReg.Value(GridReg.Columns("Descripcion").Index)
    Me.txtCodProcesoTinto = GridReg.Value(GridReg.Columns("cod_proceso_tinto").Index)
    Me.txtDesProcesoTinto = GridReg.Value(GridReg.Columns("des_proceso_tinto").Index)
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            LIMPIA
            HABILITA_CAMPOS True
            TxtCodigo.Enabled = True
            TxtCodigo.SetFocus
        Case "MODIFICAR"
            If GridReg.RowCount = 0 Then Exit Sub
            sTipo = "U"
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            HABILITA_CAMPOS True
            TxtDescripcion.SetFocus
        Case "ELIMINAR"
            Dim vMessage As Variant
            If GridReg.RowCount = 0 Then Exit Sub
            'vMessage = MsgBox("Esta seguro de Eliminar este proceso", vbYesNo, "Eliminar")
            'If vMessage = vbYes Then
                sTipo = "D"
                SALVAR_DATOS
            'End If
            CARGA_GRID
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Case "GRABAR"
                'Dim vCodigo As String
                vCodigo = TxtCodigo
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                HABILITA_CAMPOS False
                SALVAR_DATOS
                CARGA_GRID
                TxtCodigo.Enabled = False
                sTipo = ""

                
                
                Call GridReg.Find(GridReg.Columns("Codigo").Index, jgexEqual, vCodigo)
                If posProcTinto = 1 Then
                   'Me.TxtDescripcion.Enabled = True
                   HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
                   Me.txtCodProcesoTinto.Enabled = True
                   Me.txtDesProcesoTinto.Enabled = True
                   Me.txtCodProcesoTinto.Text = ""
                   Me.txtDesProcesoTinto.Text = ""
                   Me.txtCodProcesoTinto.SetFocus
                   sTipo = "U"
                   posProcTinto = 0
                End If
        Case "DESHACER"
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            HABILITA_CAMPOS False
            CARGA_GRID
            TxtCodigo.Enabled = False
        Case "SALIR"
            Unload Me
    End Select
End Sub

Sub CARGA_GRID()
On Error GoTo ErrCargaGrid
    strSQL = "select pte.cod_proceso_tex as Codigo, pte.Descripcion,pte.cod_proceso_tinto,pti.descripcion as des_proceso_tinto from Tx_Procesos_Textiles pte, Ti_Procesos_Tintoreria pti where pte.cod_proceso_tinto*=pti.cod_proceso_tinto"
    Set GridReg.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
    CONFIGURARGRID
Exit Sub
ErrCargaGrid:
ErrorHandler Err, "Carga_Grid"
End Sub

Sub CONFIGURARGRID()
    GridReg.Columns("Codigo").Width = 700
    GridReg.Columns("Descripcion").Width = 3000
End Sub

Sub LIMPIA()
    TxtCodigo.Text = ""
    TxtDescripcion.Text = ""
    Me.txtCodProcesoTinto = ""
    Me.txtDesProcesoTinto = ""
End Sub

Sub HABILITA_CAMPOS(vBoolean As Boolean)
    TxtDescripcion.Enabled = vBoolean
    Me.txtCodProcesoTinto.Enabled = vBoolean
    Me.txtDesProcesoTinto.Enabled = vBoolean
End Sub

Sub SALVAR_DATOS()
On Error GoTo ErrSalvarDatos
    
    strSQL = "EXEC UP_MAN_Tx_Procesos_Textiles '" & sTipo & "','" & _
              TxtCodigo.Text & "','" & _
              TxtDescripcion.Text & "','" & _
              Me.txtCodProcesoTinto & "'"
    
    ExecuteCommandSQL cCONNECT, strSQL
    
Exit Sub
ErrSalvarDatos:
    If Err.Description = "Codigo Proceso Tintoreria no Existe, verifique" Then
       posProcTinto = 1
    End If
    ErrorHandler Err, "SALVAR_DATOS"
End Sub

Private Sub Form_Load()
Dim sSeguridad As String
    
    'sSeguridad = get_botones1(Me, vper, vemp, Me.Name)
    'FunctButt1.FunctionsUser = sSeguridad
    
    CARGA_GRID
    HabilitaMant MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR/SALIR"
End Sub

Private Sub txtCodProcesoTinto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BuscaProcesosTintoreria (1)
    End If
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub



Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub


Private Sub BuscaProcesosTintoreria(opcion As Integer)
Dim sField As String, iRows As Long
Dim rstAux As ADODB.Recordset

    Select Case opcion
    Case 1: strSQL = "select * from Ti_Procesos_Tintoreria where Cod_Proceso_Tinto like '%" & txtCodProcesoTinto & "%'"
    End Select
    
    txtCodProcesoTinto = ""
    txtDesProcesoTinto = ""
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = strSQL
        .Caption = "Seleccionar - Proceso Tintoreria"
        .Cargar_Datos
        
        Codigo = ""
        Descripcion = ""
        
        iRows = .DGridLista.RowCount
        Set rstAux = .DGridLista.ADORecordset
        If .DGridLista.RowCount > 1 Then
            .Show vbModal
        ElseIf .DGridLista.RowCount = 1 Then
            Codigo = .DGridLista.Value(.DGridLista.Columns("Cod_Proceso_Tinto").Index)
            Descripcion = .DGridLista.Value(.DGridLista.Columns("Descripcion").Index)
        End If
        
        If Codigo <> "" Then
            txtCodProcesoTinto = RTrim(Codigo)
            txtDesProcesoTinto = RTrim(Descripcion)
        End If
    End With
    Me.MantFunc1.SetFocus
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
End Sub

Private Sub txtDesProcesoTinto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BuscaProcesosTintoreria (2)
    End If
End Sub
