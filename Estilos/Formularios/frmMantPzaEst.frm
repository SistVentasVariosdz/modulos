VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMantPzaEst 
   Caption         =   "Pieza Estilo"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5655
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Piece Style"
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
      Height          =   1200
      Left            =   45
      TabIndex        =   9
      Tag             =   "Detail"
      Top             =   3900
      Width           =   5445
      Begin VB.TextBox txtDesPzaEst 
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
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   2
         Top             =   720
         Width           =   3585
      End
      Begin VB.TextBox txtIdPzaEst 
         BackColor       =   &H80000004&
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
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
         Left            =   5010
         Top             =   180
         _cx             =   847
         _cy             =   847
         PassiveMode     =   0   'False
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Descripción :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   375
         TabIndex        =   11
         Tag             =   "Description"
         Top             =   795
         Width           =   945
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   375
         TabIndex        =   10
         Tag             =   "Code"
         Top             =   420
         Width           =   945
      End
   End
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
      Height          =   3855
      Left            =   45
      TabIndex        =   7
      Tag             =   "List"
      Top             =   0
      Width           =   5445
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   3345
         Left            =   180
         TabIndex        =   8
         Top             =   345
         Width           =   5130
         _ExtentX        =   9049
         _ExtentY        =   5900
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
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
            DataField       =   "Cod_organizacion"
            Caption         =   "Código"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Nom_organizacion"
            Caption         =   "Descripción"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3704.882
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   5145
      Width           =   1965
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "frmMantPzaEst.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Anterior"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "frmMantPzaEst.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   15
         Picture         =   "frmMantPzaEst.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Primero"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "frmMantPzaEst.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Ultimo"
         Top             =   120
         Width           =   495
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   1950
      TabIndex        =   12
      Top             =   5220
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantPzaEst.frx":05C8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmMantPzaEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Dim sTipo As String
Dim Rs_Carga As New ADODB.Recordset
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
    Dim StrSQL As String
    On Error GoTo Cargar_DatosErr
    StrSQL = "SG_Act_PzaEst '','','L'"
    Set Rs_Carga = Nothing
    Rs_Carga.ActiveConnection = cCONNECT
    Rs_Carga.CursorType = adOpenStatic
    Rs_Carga.CursorLocation = adUseClient
    Rs_Carga.LockType = adLockReadOnly
    
    Rs_Carga.Open StrSQL
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
Private Sub Form_Load()
Call FormSet(Me)
FormateaGrid Me.DGridLista
DGridLista.Columns(0).DataField = "cod_pzaest"
DGridLista.Columns(1).DataField = "des_pzaest"
Carga_Datos
MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub
Sub SALVAR_DATOS()
Dim Con As New ADODB.Connection
On Error GoTo Salvar_DatosErr
Con.ConnectionString = cCONNECT
Con.Open
If txtIdPzaEst.Text <> "" Then
    Con.BeginTrans
    Con.Execute "SG_Act_PzaEst '" & _
    txtIdPzaEst.Text & "','" & _
    txtDesPzaEst.Text & "','" & _
    sTipo & "'"
    Con.CommitTrans
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
    Informa "", amensaje
End If
LIMPIAR_DATOS
RECARGAR_DATOS
Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub
Sub ELIMINAR_DATOS()
Dim Con As New ADODB.Connection
On Error GoTo Eliminar_DatosErr
Con.ConnectionString = cCONNECT
Con.Open
If txtIdPzaEst.Text <> "" Then
    Con.BeginTrans
    Con.Execute "SG_Act_PzaEst '" & txtIdPzaEst.Text & "','','D'"
    Con.CommitTrans
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_DELETE
    Informa "", amensaje
End If
LIMPIAR_DATOS
RECARGAR_DATOS
Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Eliminar_Datos"
End Sub
Sub LIMPIAR_DATOS()
    txtIdPzaEst.Text = ""
    txtDesPzaEst.Text = ""
End Sub
Private Sub DGridLista_Click()
If Rs_Carga.State <> 1 Then
    Exit Sub
End If
If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
    txtIdPzaEst.Text = Rs_Carga!cod_pzaest
    txtDesPzaEst.Text = Rs_Carga!des_pzaest
    DESHABILITA_DATOS
End If
End Sub
Sub HABILITA_DATOS()
    txtIdPzaEst.Enabled = True
    txtDesPzaEst.Enabled = True
    txtIdPzaEst.SetFocus
End Sub
Sub DESHABILITA_DATOS()
    txtIdPzaEst.Enabled = False
    txtDesPzaEst.Enabled = False
End Sub
Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub
Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Rs_Carga.State <> 1 Then
    Exit Sub
End If
If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
    txtIdPzaEst.Text = Rs_Carga!cod_pzaest
    txtDesPzaEst.Text = Rs_Carga!des_pzaest
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
        txtIdPzaEst.SetFocus
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        DGridLista.Enabled = False
    Case "MODIFICAR"
        sTipo = "U"
        txtDesPzaEst.Enabled = True
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        DGridLista.Enabled = False
    Case "ELIMINAR"
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
If Len(Trim(txtDesPzaEst.Text)) = 0 Then
   MsgBox "Ingrese la descripcion", vbInformation, Me.Caption
   VALIDA_DATOS = False
End If
If Len(Trim(txtIdPzaEst.Text)) = 0 Then
   MsgBox "Ingrese el Codigo", vbInformation, Me.Caption
   VALIDA_DATOS = False
End If
If Not VALIDA_DATOS Then
    LoadMessage aMess, amensaje.Codigo
    amensaje.ShowMesage (iLanguage)
End If
End Function
Private Sub txtIdPzaEst_KeyDown(KeyCode As Integer, Shift As Integer)
AVANZA (KeyCode)
End Sub
Private Sub txtIdPzaEst_LostFocus()
If Len(Trim(txtIdPzaEst)) <> 0 Then
    Busca_TipComp
End If
End Sub
Sub Busca_TipComp()
Dim Rs_busca As New ADODB.Recordset
On Error GoTo Busca_FuncionErr
B_sql = "SELECT * FROM ES_PzaEst " & _
"WHERE cod_pzaest = '" & txtIdPzaEst.Text & "'"
Rs_busca.ActiveConnection = cCONNECT
Rs_busca.CursorType = adOpenStatic
Rs_busca.Open B_sql
If Not Rs_busca.EOF Then
    txtDesPzaEst.Text = Rs_busca!des_tipcompest
    DESHABILITA_DATOS
    HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
    DGridLista.Enabled = True
End If
Rs_busca.Close
Set Rs_busca = Nothing
Exit Sub
Busca_FuncionErr:
    Set Rs_busca = Nothing
    ErrorHandler Err, "Busca_Acceso"
End Sub



