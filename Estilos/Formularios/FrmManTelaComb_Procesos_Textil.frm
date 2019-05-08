VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "mantenimientos.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmManTelaComb_Procesos_Textil 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4485
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6255
      Begin GridEX20.GridEX GridReg 
         Height          =   4125
         Left            =   60
         TabIndex        =   11
         Top             =   165
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   7276
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
         Column(1)       =   "FrmManTelaComb_Procesos_Textil.frx":0000
         Column(2)       =   "FrmManTelaComb_Procesos_Textil.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "FrmManTelaComb_Procesos_Textil.frx":016C
         FormatStyle(2)  =   "FrmManTelaComb_Procesos_Textil.frx":02A4
         FormatStyle(3)  =   "FrmManTelaComb_Procesos_Textil.frx":0354
         FormatStyle(4)  =   "FrmManTelaComb_Procesos_Textil.frx":0408
         FormatStyle(5)  =   "FrmManTelaComb_Procesos_Textil.frx":04E0
         FormatStyle(6)  =   "FrmManTelaComb_Procesos_Textil.frx":0598
         ImageCount      =   0
         PrinterProperties=   "FrmManTelaComb_Procesos_Textil.frx":0678
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
      Height          =   1605
      Left            =   0
      TabIndex        =   0
      Top             =   4560
      Width           =   6255
      Begin VB.TextBox TxtCod_Comb 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1170
         TabIndex        =   14
         Top             =   550
         Width           =   585
      End
      Begin VB.TextBox TxtDes_Comb 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1800
         TabIndex        =   13
         Top             =   550
         Width           =   3885
      End
      Begin VB.TextBox TxtCodigo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   6
         Top             =   870
         Width           =   585
      End
      Begin VB.TextBox TxtDescripcion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   5
         Top             =   870
         Width           =   3900
      End
      Begin VB.TextBox Txtcod_Tela 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1170
         TabIndex        =   4
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox TxtDes_Tela 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2280
         TabIndex        =   3
         Top             =   240
         Width           =   3400
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "..."
         Height          =   300
         Left            =   5760
         TabIndex        =   2
         Top             =   915
         Width           =   375
      End
      Begin VB.TextBox TxtSecuencia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1170
         TabIndex        =   1
         Top             =   1180
         Width           =   705
      End
      Begin VB.Label Label4 
         Caption         =   "Comb:"
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   650
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "Proceso:"
         Height          =   225
         Left            =   150
         TabIndex        =   9
         Top             =   990
         Width           =   720
      End
      Begin VB.Label Label3 
         Caption         =   "Tela:"
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   315
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Secuencia:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1275
         Width           =   810
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   1320
      TabIndex        =   12
      Top             =   6240
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmManTelaComb_Procesos_Textil.frx":0850
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "FrmManTelaComb_Procesos_Textil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Codigo As String
Public Descripcion As String
Dim strSQL As String
Dim sTipo As String
Public vCod_Tela As String
Public vCod_Comb As String

Private Sub CmdNuevo_Click()
Load FrmManProcesos_Textiles
FrmManProcesos_Textiles.CARGA_GRID
FrmManProcesos_Textiles.Show vbModal
'TxtCodigo.Text = FrmManProcesos_Textiles.vCodigo
Set FrmManProcesos_Textiles = Nothing
End Sub

Private Sub GridReg_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    TxtCodigo = GridReg.Value(GridReg.Columns("Codigo").Index)
    TxtDescripcion = GridReg.Value(GridReg.Columns("Descripcion").Index)
    TxtSecuencia = GridReg.Value(GridReg.Columns("secuencia").Index)
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
            'vMessage = MsgBox("Esta seguro de Eliminar este proceso Tela", vbYesNo, "Eliminar")
            'If vMessage = vbYes Then
                sTipo = "D"
                SALVAR_DATOS
            'End If
            CARGA_GRID
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Case "GRABAR"
                Dim vCodigo As String
                vCodigo = TxtCodigo
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                HABILITA_CAMPOS False
                SALVAR_DATOS
                CARGA_GRID
                TxtCodigo.Enabled = False
                sTipo = ""
                Call GridReg.Find(GridReg.Columns("Codigo").Index, jgexEqual, vCodigo)
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
    strSQL = "SM_Muestra_Procesos_Textiles_TelaComb '" & vCod_Tela & "','" & vCod_Comb & "'"
    Set GridReg.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
    CONFIGURARGRID
Exit Sub
ErrCargaGrid:
ErrorHandler Err, "Carga_Grid"
End Sub

Sub CONFIGURARGRID()
    GridReg.Columns("Codigo").Width = 700
    GridReg.Columns("Descripcion").Width = 3000
    GridReg.Columns("Secuencia").Width = 700
End Sub

Sub LIMPIA()
    TxtCodigo.Text = ""
    TxtDescripcion.Text = ""
    TxtSecuencia.Text = DevuelveCampo("select isnull(max(secuencia),0) + 1 from tx_telacomb_procesos_textiles where cod_tela ='" & vCod_Tela & "' and cod_comb='" & vCod_Comb & "'", cCONNECT)
End Sub

Sub HABILITA_CAMPOS(vBoolean As Boolean)
    TxtDescripcion.Enabled = vBoolean
    TxtSecuencia.Enabled = vBoolean
End Sub

Sub SALVAR_DATOS()
On Error GoTo ErrSalvarDatos
    
    strSQL = "EXEC Up_Man_Tx_Procesos_Textiles_TelaComb '" & sTipo & "','" & vCod_Tela & "','" & vCod_Comb & "','" & _
              TxtCodigo.Text & "','" & Val(TxtSecuencia) & "'"
    
    ExecuteCommandSQL cCONNECT, strSQL
    
Exit Sub
ErrSalvarDatos:
    ErrorHandler Err, "SALVAR_DATOS"
End Sub

Private Sub Form_Load()
Dim sSeguridad As String
    
    CARGA_GRID
    HabilitaMant MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR/SALIR"
End Sub



Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call BUSCAPROCESO(1)
    End If
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BUSCAPROCESO(2)
End If
End Sub

Sub BUSCAPROCESO(tipo As Integer)
Dim oTipo As New frmBusqGeneral3
Dim rs As New ADODB.Recordset

Set oTipo.oParent = Me

If tipo = 1 Then
    oTipo.sQuery = "select cod_proceso_tex as Codigo, Descripcion from tx_procesos_textiles where cod_proceso_tex like '%" & Trim(TxtCodigo.Text) & "%'"
ElseIf tipo = 2 Then
    oTipo.sQuery = "select cod_proceso_tex as Codigo, Descripcion from tx_procesos_textiles where descripcion like '%" & Trim(TxtDescripcion.Text) & "%'"
End If

oTipo.Caption = "Buscar Proceso Textil"
oTipo.Cargar_Datos

oTipo.DGridLista.Columns("Codigo").Width = 1400
oTipo.DGridLista.Columns("Descripcion").Width = 5000

If oTipo.DGridLista.RowCount > 1 Then
    oTipo.Show vbModal
Else
    Codigo = oTipo.DGridLista.Value(oTipo.DGridLista.Columns("Codigo").Index)
    Descripcion = oTipo.DGridLista.Value(oTipo.DGridLista.Columns("Descripcion").Index)
End If

If Trim(Codigo) <> "" Then
    TxtCodigo.Text = Codigo
    TxtDescripcion.Text = Descripcion
    Codigo = "": Descripcion = ""
    TxtSecuencia.SetFocus
End If

Unload oTipo
Set oTipo = Nothing
Set rs = Nothing
End Sub

Private Sub TxtSecuencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtSecuencia, KeyAscii, False)
End If
End Sub


