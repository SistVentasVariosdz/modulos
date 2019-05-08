VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Begin VB.Form frmMantAgenteAduana 
   Caption         =   "Agente de Aduanas"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCargos 
      Caption         =   "Agente de Aduana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3705
      Left            =   135
      TabIndex        =   7
      Top             =   135
      Width           =   7125
      Begin GridEX20.GridEX gexAgenteAduana 
         Height          =   3345
         Left            =   210
         TabIndex        =   8
         Top             =   225
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   5900
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
         Column(1)       =   "frmMantAgenteAduana.frx":0000
         Column(2)       =   "frmMantAgenteAduana.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmMantAgenteAduana.frx":016C
         FormatStyle(2)  =   "frmMantAgenteAduana.frx":02A4
         FormatStyle(3)  =   "frmMantAgenteAduana.frx":0354
         FormatStyle(4)  =   "frmMantAgenteAduana.frx":0408
         FormatStyle(5)  =   "frmMantAgenteAduana.frx":04E0
         FormatStyle(6)  =   "frmMantAgenteAduana.frx":0598
         ImageCount      =   0
         PrinterProperties=   "frmMantAgenteAduana.frx":0678
      End
   End
   Begin Mantenimientos.MantFunc MantFunct1 
      Height          =   540
      Left            =   1590
      TabIndex        =   0
      Top             =   5670
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantAgenteAduana.frx":0850
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
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
      Height          =   1620
      Left            =   360
      TabIndex        =   1
      Top             =   3975
      Width           =   6900
      Begin VB.TextBox txtContacto 
         Height          =   315
         Left            =   1065
         TabIndex        =   4
         Top             =   1125
         Width           =   4770
      End
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
      Begin VB.Label Label2 
         Caption         =   "Contacto"
         Height          =   210
         Left            =   165
         TabIndex        =   9
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   765
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   345
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmMantAgenteAduana"
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
    
    strSQL = "EXEC TG_Muestra_AgenteAduana"
    
    Set gexAgenteAduana.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
    
    SetGeneralGridEX gexAgenteAduana, 0, 1
    
    If Me.gexAgenteAduana.RowCount > 0 Then
    Call CARGA_DATOS
    End If
    
End Sub

Public Sub LIMPIA_DATOS()
    Me.txtCodigo.Text = ""
    Me.txtDescripcion.Text = ""
    Me.txtContacto.Text = ""
End Sub

Public Sub HABILITA_DATOS()
    If Stipo = "I" Then
        Me.txtCodigo.Enabled = True
    End If
        Me.txtDescripcion.Enabled = True
        Me.txtContacto.Enabled = True
End Sub

Public Sub INHABILITA_DATOS()
    Me.txtCodigo.Enabled = False
    Me.txtDescripcion.Enabled = False
    Me.txtContacto.Enabled = False
End Sub

Sub SALVAR_DATOS()
    Dim strSQL As String
    On Error GoTo Salvar_DatosErr
 
    strSQL = "EXEC AS_Man_AGENTE_ADUANA '" & Stipo & "','" & Trim(txtCodigo.Text) & "','" & Trim(txtDescripcion.Text) & "','" & Trim(txtContacto.Text) & "'"
      
    ExecuteCommandSQL cCONNECT, strSQL

    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMESSAGE_INF_DATA_SAVE
    Informa "", amensaje
    
    Exit Sub
Salvar_DatosErr:
    ErrorHandler Err, "Salvar_Datos"
End Sub

Public Sub CARGA_DATOS()
    If Me.gexAgenteAduana.RowCount > 0 Then
        Me.txtCodigo.Text = gexAgenteAduana.Value(gexAgenteAduana.Columns("Codigo").Index)
        Me.txtDescripcion.Text = gexAgenteAduana.Value(gexAgenteAduana.Columns("descripcion").Index)
        Me.txtContacto.Text = gexAgenteAduana.Value(gexAgenteAduana.Columns("contacto").Index)
        
    End If
    
   gexAgenteAduana.Columns("Codigo").Width = 1500
   gexAgenteAduana.Columns("descripcion").Width = 2500
   gexAgenteAduana.Columns("contacto").Width = 2500
End Sub

Private Sub MantFunct1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
  Dim ELIMINAR As Integer
    Dim vRow As Long
    Select Case ActionName
        Case "ADICIONAR"
            Stipo = "I"
            Call LIMPIA_DATOS
            Call HABILITA_DATOS
            txtDescripcion.SetFocus
            HabilitaMant Me.MantFunct1, "GRABAR/DESHACER"
        Case "MODIFICAR"
            Stipo = "U"
            Call HABILITA_DATOS
            HabilitaMant Me.MantFunct1, "GRABAR/DESHACER"
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
            HabilitaMant Me.MantFunct1, "ADICIONAR/MODIFICAR/ELIMINAR"
            Stipo = ""
        Case "DESHACER"
            Call LIMPIA_DATOS
            Call CARGA_DATOS
            Call INHABILITA_DATOS
            HabilitaMant Me.MantFunct1, "ADICIONAR/MODIFICAR/ELIMINAR"
            Stipo = ""
         Case "SALIR"
            Unload Me
      End Select
End Sub

Private Sub gexAgenteAduana_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    Call CARGA_DATOS
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  Me.txtCodigo.SetFocus
End If
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  Me.txtContacto.SetFocus
End If
End Sub
