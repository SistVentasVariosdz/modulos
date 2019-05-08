VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Begin VB.Form frmMantModoEmbarque 
   Caption         =   "Mantenimiento Modo Embarque"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDatos 
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   60
      TabIndex        =   1
      Top             =   3660
      Width           =   6900
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1065
         TabIndex        =   3
         Top             =   720
         Width           =   4725
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1065
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   285
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   705
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   345
         Width           =   495
      End
   End
   Begin GridEX20.GridEX gexModoEmbarque 
      Height          =   3165
      Left            =   90
      TabIndex        =   0
      Top             =   300
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5583
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
      ColumnsCount    =   2
      Column(1)       =   "frmMantModoEmbarque.frx":0000
      Column(2)       =   "frmMantModoEmbarque.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmMantModoEmbarque.frx":016C
      FormatStyle(2)  =   "frmMantModoEmbarque.frx":02A4
      FormatStyle(3)  =   "frmMantModoEmbarque.frx":0354
      FormatStyle(4)  =   "frmMantModoEmbarque.frx":0408
      FormatStyle(5)  =   "frmMantModoEmbarque.frx":04E0
      FormatStyle(6)  =   "frmMantModoEmbarque.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmMantModoEmbarque.frx":0678
   End
   Begin Mantenimientos.MantFunc MantFunc2 
      Height          =   540
      Left            =   1635
      TabIndex        =   6
      Top             =   5040
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantModoEmbarque.frx":0850
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmMantModoEmbarque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Stipo As String
Public strSQL As String
Public sql As String
Public Codigo As String
Public Descripcion As String

Private Sub Form_Load()
Call CARGA_GRID
Call INHABILITA_DATOS
End Sub

Sub CARGA_GRID()
    
    strSQL = "EXEC TG_Muestra_ModoEmbarque"
    
    Set gexModoEmbarque.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
    
    SetGeneralGridEX gexModoEmbarque, 0, 1
    
    If Me.gexModoEmbarque.RowCount > 0 Then
    Call CARGA_DATOS
    End If
    
End Sub

Public Sub LIMPIA_DATOS()
    Me.txtCodigo.Text = ""
    Me.txtDescripcion.Text = ""
End Sub

Public Sub HABILITA_DATOS()
    If Stipo = "I" Then
        Me.txtCodigo.Enabled = True
    End If
    Me.txtDescripcion.Enabled = True
End Sub

Public Sub INHABILITA_DATOS()
    Me.txtCodigo.Enabled = False
    Me.txtDescripcion.Enabled = False
End Sub

Sub SALVAR_DATOS()
    Dim strSQL As String
    On Error GoTo Salvar_DatosErr
 
    strSQL = "EXEC AS_Man_MODO_EMBARQUE '" & Stipo & "','" & Trim(txtCodigo.Text) & "','" & Trim(txtDescripcion.Text) & "'"
      
    ExecuteCommandSQL cCONNECT, strSQL

    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMESSAGE_INF_DATA_SAVE
    Informa "", amensaje
    
    Exit Sub
Salvar_DatosErr:
    ErrorHandler Err, "Salvar_Datos"
End Sub

Public Sub CARGA_DATOS()
    If Me.gexModoEmbarque.RowCount > 0 Then
        Me.txtCodigo.Text = gexModoEmbarque.Value(gexModoEmbarque.Columns("COD_EMBARQUE").Index)
        Me.txtDescripcion.Text = gexModoEmbarque.Value(gexModoEmbarque.Columns("DESCRIPCION").Index)
        
    End If
    
    gexModoEmbarque.Columns("COD_EMBARQUE").Width = 1500
    gexModoEmbarque.Columns("DESCRIPCION").Width = 4000
End Sub

Private Sub MantFunc2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
  Dim ELIMINAR As Integer
    Dim vRow As Long
    Select Case ActionName
        Case "ADICIONAR"
            Stipo = "I"
            Call LIMPIA_DATOS
            Call HABILITA_DATOS
            txtDescripcion.SetFocus
            HabilitaMant Me.MantFunc2, "GRABAR/DESHACER"
        Case "MODIFICAR"
            Stipo = "U"
            Call HABILITA_DATOS
            HabilitaMant Me.MantFunc2, "GRABAR/DESHACER"
        Case "ELIMINAR"
            ELIMINAR = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Combinación-Detalle")
            If ELIMINAR = vbYes Then
                Stipo = "D"
                Call SALVAR_DATOS
                Call CARGA_GRID
                Stipo = ""
            End If
        Case "GRABAR"
            Call SALVAR_DATOS
            Call CARGA_GRID
            Call INHABILITA_DATOS
            HabilitaMant Me.MantFunc2, "ADICIONAR/MODIFICAR/ELIMINAR"
            Stipo = ""
        Case "DESHACER"
            Call LIMPIA_DATOS
            Call CARGA_DATOS
            Call INHABILITA_DATOS
            HabilitaMant Me.MantFunc2, "ADICIONAR/MODIFICAR/ELIMINAR"
            Stipo = ""
         Case "SALIR"
            Unload Me
      End Select
End Sub

Private Sub gexModoEmbarque_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    Call CARGA_DATOS
End Sub


