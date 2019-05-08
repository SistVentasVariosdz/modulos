VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmManteProcesosTelasComb 
   Caption         =   "Mantenimiento Procesos"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fralista 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      Left            =   0
      TabIndex        =   12
      Tag             =   "List"
      Top             =   0
      Width           =   6855
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2925
         Left            =   120
         TabIndex        =   13
         Top             =   255
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   5159
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   17
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Fradetalle 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      TabIndex        =   8
      Tag             =   "Detail"
      Top             =   3240
      Width           =   6855
      Begin VB.CommandButton CmdProceso 
         Caption         =   "..."
         Height          =   270
         Left            =   2100
         TabIndex        =   4
         Top             =   1320
         Width           =   270
      End
      Begin VB.TextBox TxtDEs_ProTex 
         Height          =   285
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1320
         Width           =   4335
      End
      Begin VB.TextBox txtCod_ProTex 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtSecuencia 
         BackColor       =   &H80000009&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   2
         Top             =   915
         Width           =   855
      End
      Begin VB.TextBox txtMerma 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Text            =   "0"
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   2100
         TabIndex        =   16
         Top             =   1740
         Width           =   120
      End
      Begin VB.Label LblComb 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   5505
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Comb.:"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   495
      End
      Begin VB.Label LblTela 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   5505
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tela:"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Proceso:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1395
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Secuencia:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1020
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Porc.Merma:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1755
         Width           =   900
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   1440
      TabIndex        =   7
      Top             =   5520
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmManteProcesosTelasComb.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "FrmManteProcesosTelasComb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public Codigo_tela As String, Codigo_Comb As String

Public oParent As Object
Dim sTipo As String
Dim Rs_Carga As New ADODB.Recordset

Public Codigo, Descripcion As String

Private Sub cmdFirst_Click()
If Not Rs_Carga.BOF Then
  Rs_Carga.MoveFirst
End If
End Sub
Private Sub cmdLast_Click()
If Not Rs_Carga.EOF Then
 Rs_Carga.MoveLast
End If
End Sub
Private Sub cmdNext_Click()
If Not Rs_Carga.EOF Then
 Rs_Carga.MoveNext
End If
End Sub
Private Sub cmdPrevious_Click()
If Not Rs_Carga.BOF Then
 Rs_Carga.MovePrevious
End If
End Sub
Sub Carga_Datos()
    Dim strSQL As String
    On Error GoTo Cargar_DatosErr
    strSQL = "tx_up_man_tx_telacombpro 'V','" & Codigo_tela & "','" & Codigo_Comb & "','','',0,'',''"
    Set Rs_Carga = Nothing
    Rs_Carga.ActiveConnection = cCONNECT
    Rs_Carga.CursorType = adOpenStatic
    Rs_Carga.CursorLocation = adUseClient
    Rs_Carga.LockType = adLockReadOnly
    
    Rs_Carga.Open strSQL
    Set DGridLista.DataSource = Rs_Carga
    DGridLista_RowColChange 0, 0
    If Rs_Carga.RecordCount > 0 Then
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
    Else
        LIMPIAR_DATOS
        DESHABILITA_DATOS
        HabilitaMant Me.MantFunc1, "ADICIONAR"
    End If
    Exit Sub
Cargar_DatosErr:
    Set Rs_Carga = Nothing
    ErrorHandler Err, "Cargar_Datos"
End Sub

Private Sub CmdProceso_Click()
Dim oForm As New frmBusqGeneral
Set oForm.oParent = Me
oForm.sQuery = "SELECT cod_protex as Codigo, Des_ProTEx as Descripcion FROM Tx_procesos where des_protex like '%" & TxtDEs_ProTex & "%'"
oForm.Cargar_Datos
oForm.Show 1
Set oForm = Nothing
If Codigo <> "" Then
    txtCod_ProTex.Text = Codigo
    TxtDEs_ProTex.Text = Descripcion
    Codigo = "": Descripcion = ""
    txtMerma.SetFocus
End If
End Sub

Private Sub Form_Load()
Call FormSet(Me)
FormateaGrid Me.DGridLista
'DGridLista.Columns(0).DataField = "cod_matpri"
'DGridLista.Columns(1).DataField = "des_matpri"
'Carga_Datos
MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub
Sub SALVAR_DATOS()
Dim con As New ADODB.Connection
On Error GoTo Salvar_DatosErr
con.ConnectionString = cCONNECT
con.Open

con.BeginTrans
con.Execute "tx_up_man_tx_telacombpro '" & sTipo & "','" & _
Codigo_tela & "','" & _
Codigo_Comb & "','" & _
txtSecuencia.Text & "','" & _
txtCod_ProTex.Text & "'," & _
CDbl(txtMerma.Text) & ",'" & _
vusu & "','" & ComputerName & "'"

con.CommitTrans
Dim amensaje As New clsMessages
amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
Informa "", amensaje

LIMPIAR_DATOS
RECARGAR_DATOS
Exit Sub
Salvar_DatosErr:
    con.RollbackTrans
    Set con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub
Sub ELIMINAR_DATOS()
Dim con As New ADODB.Connection
On Error GoTo Eliminar_DatosErr
con.ConnectionString = cCONNECT
con.Open
If txtSecuencia.Text <> "" Then
    con.BeginTrans
    con.Execute "tx_up_man_tx_telacombpro '" & sTipo & "','" & _
    Codigo_tela & "','" & _
    Codigo_Comb & "','" & _
    txtSecuencia.Text & "','" & _
    txtCod_ProTex.Text & "'," & _
    CDbl(txtMerma.Text) & ",'" & _
    vusu & "','" & ComputerName & "'"
    
    con.CommitTrans
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_DELETE
    Informa "", amensaje
End If
LIMPIAR_DATOS
RECARGAR_DATOS
Exit Sub
Eliminar_DatosErr:
    con.RollbackTrans
    Set con = Nothing
    ErrorHandler Err, "Eliminar_Datos"
End Sub
Sub LIMPIAR_DATOS()
    txtSecuencia.Text = ""
    txtCod_ProTex.Text = ""
    TxtDEs_ProTex.Text = ""
    txtMerma.Text = 0
End Sub
Private Sub DGridLista_Click()
If Rs_Carga.State <> 1 Then
    Exit Sub
End If
If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
    txtSecuencia.Text = Rs_Carga!num_secuencia
    txtCod_ProTex.Text = Rs_Carga!cod_protex
    TxtDEs_ProTex.Text = Rs_Carga!des_protex
    txtMerma.Text = CDbl(Rs_Carga!por_merma)
    DESHABILITA_DATOS
End If
End Sub
Sub HABILITA_DATOS()
    txtCod_ProTex.Enabled = True
    TxtDEs_ProTex.Enabled = True
    CmdProceso.Enabled = True
    txtMerma.Enabled = True
    txtCod_ProTex.SetFocus
End Sub
Sub DESHABILITA_DATOS()
    txtSecuencia.Enabled = False
    txtCod_ProTex.Enabled = False
    TxtDEs_ProTex.Enabled = False
    CmdProceso.Enabled = False
    txtMerma.Enabled = False
End Sub
Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub
Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Rs_Carga.State <> 1 Then
    Exit Sub
End If
If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
    txtSecuencia.Text = Rs_Carga!num_secuencia
    txtCod_ProTex.Text = Rs_Carga!cod_protex
    TxtDEs_ProTex.Text = Rs_Carga!des_protex
    txtMerma.Text = CDbl(Rs_Carga!por_merma)
    DESHABILITA_DATOS
End If
End Sub
Sub RECARGAR_DATOS()
Rs_Carga.Close
Carga_Datos
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set Rs_Carga = Nothing
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub
Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "ADICIONAR"
        sTipo = "I"
        LIMPIAR_DATOS
        HABILITA_DATOS
        txtCod_ProTex.SetFocus
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        DGridLista.Enabled = False
    Case "MODIFICAR"
        sTipo = "U"
        txtCod_ProTex.Enabled = True
        TxtDEs_ProTex.Enabled = True
        txtMerma.Enabled = True
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        DGridLista.Enabled = False
    Case "ELIMINAR"
        sTipo = "D"
        ELIMINAR_DATOS
    Case "GRABAR"
        If VALIDA_DATOS Then
            SALVAR_DATOS
            RECARGAR_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridLista.Enabled = True
        End If
    Case "DESHACER"
        LIMPIAR_DATOS
        RECARGAR_DATOS
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        DGridLista.Enabled = True
    Case "SALIR"
        Unload Me
End Select
End Sub
Function VALIDA_DATOS() As Boolean
Dim aMess(4)
Dim amensaje As clsMessages
Set amensaje = New clsMessages
VALIDA_DATOS = True
If Len(Trim(txtCod_ProTex.Text)) = 0 Then
   MsgBox "Ingrese Codigo Proceso", vbInformation, Me.Caption
   VALIDA_DATOS = False
End If
If Not VALIDA_DATOS Then
    LoadMessage aMess, amensaje.Codigo
    amensaje.ShowMesage (iLanguage)
End If
End Function
Private Sub txtIdMatPri_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub

Sub Busca_Proceso()
Dim Rs_busca As New ADODB.Recordset
On Error GoTo Busca_FuncionErr
B_sql = "SELECT * FROM tx_procesos " & _
"WHERE cod_protex = '" & txtCod_ProTex.Text & "'"
Rs_busca.ActiveConnection = cCONNECT
Rs_busca.CursorType = adOpenStatic
Rs_busca.Open B_sql
If Not Rs_busca.EOF Then
    TxtDEs_ProTex.Text = Rs_busca!des_protex
    'HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
    'DGridLista.Enabled = True
    txtMerma.SetFocus
End If
Rs_busca.Close
Set Rs_busca = Nothing
Exit Sub
Busca_FuncionErr:
    Set Rs_busca = Nothing
    ErrorHandler Err, "Busca_Acceso"
End Sub

Private Sub txtCod_ProTex_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(txtCod_ProTex.Text) = "" Then
        SendKeys "{TAB}"
    Else
        Call Busca_Proceso
    End If
End If
End Sub

Private Sub TxtDEs_ProTex_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call CmdProceso_Click
End If
End Sub

Private Sub txtMerma_GotFocus()
SelectionText txtMerma
End Sub

Private Sub txtMerma_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub
