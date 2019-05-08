VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Begin VB.Form frmMantAlmacenAduana 
   Caption         =   "Mantenimiento de Almacen Aduana"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   7125
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
      Height          =   2265
      Left            =   120
      TabIndex        =   6
      Top             =   3525
      Width           =   6900
      Begin VB.TextBox txtContacto 
         Height          =   300
         Left            =   1035
         TabIndex        =   4
         Top             =   1875
         Width           =   2370
      End
      Begin VB.TextBox txtRuc 
         Height          =   285
         Left            =   1050
         TabIndex        =   3
         Top             =   1500
         Width           =   2295
      End
      Begin VB.TextBox txtDireccion 
         Height          =   300
         Left            =   1035
         TabIndex        =   2
         Top             =   1110
         Width           =   4725
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   330
         Width           =   630
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1035
         TabIndex        =   1
         Top             =   720
         Width           =   4725
      End
      Begin VB.Label Label5 
         Caption         =   "Contacto"
         Height          =   270
         Left            =   150
         TabIndex        =   12
         Top             =   1860
         Width           =   720
      End
      Begin VB.Label Label3 
         Caption         =   "RUC"
         Height          =   285
         Left            =   150
         TabIndex        =   11
         Top             =   1500
         Width           =   585
      End
      Begin VB.Label Label2 
         Caption         =   "Direccion"
         Height          =   225
         Left            =   135
         TabIndex        =   10
         Top             =   1155
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   390
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   750
         Width           =   840
      End
   End
   Begin GridEX20.GridEX gexAlmacenAduana 
      Height          =   3165
      Left            =   135
      TabIndex        =   5
      Top             =   135
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
      Column(1)       =   "frmMantAlmacenAduana.frx":0000
      Column(2)       =   "frmMantAlmacenAduana.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmMantAlmacenAduana.frx":016C
      FormatStyle(2)  =   "frmMantAlmacenAduana.frx":02A4
      FormatStyle(3)  =   "frmMantAlmacenAduana.frx":0354
      FormatStyle(4)  =   "frmMantAlmacenAduana.frx":0408
      FormatStyle(5)  =   "frmMantAlmacenAduana.frx":04E0
      FormatStyle(6)  =   "frmMantAlmacenAduana.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmMantAlmacenAduana.frx":0678
   End
   Begin Mantenimientos.MantFunc MantFunc2 
      Height          =   540
      Left            =   1575
      TabIndex        =   9
      Top             =   5955
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantAlmacenAduana.frx":0850
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmMantAlmacenAduana"
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
    
    strSQL = "EXEC TG_Muestra_AlamacenAduana"
    
    Set gexAlmacenAduana.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
    
    SetGeneralGridEX gexAlmacenAduana, 0, 1
    
    If Me.gexAlmacenAduana.RowCount > 0 Then
    Call CARGA_DATOS
    End If
    
End Sub

Public Sub LIMPIA_DATOS()
    Me.txtCodigo.Text = ""
    Me.txtDescripcion.Text = ""
    Me.txtDireccion.Text = ""
    Me.txtRuc.Text = ""
    Me.txtContacto.Text = ""
End Sub

Public Sub HABILITA_DATOS()
    If Stipo = "I" Then
        Me.txtCodigo.Enabled = True
    End If
        Me.txtDescripcion.Enabled = True
        Me.txtDireccion.Enabled = True
        Me.txtRuc.Enabled = True
        Me.txtContacto.Enabled = True
End Sub

Public Sub INHABILITA_DATOS()
    Me.txtCodigo.Enabled = False
    Me.txtDescripcion.Enabled = False
    Me.txtDireccion.Enabled = False
    Me.txtRuc.Enabled = False
    Me.txtContacto.Enabled = False
End Sub

Sub SALVAR_DATOS()
    Dim strSQL As String
    On Error GoTo Salvar_DatosErr
 
    strSQL = "EXEC AS_Man_ALMACEN_ADUANA '" & Stipo & "','" & Trim(txtCodigo.Text) & "','" & Trim(txtDescripcion.Text) & "','" & Trim(txtDireccion.Text) & "','" & Trim(txtRuc.Text) & "','" & Trim(txtContacto.Text) & "'"
      
    ExecuteCommandSQL cCONNECT, strSQL

    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMESSAGE_INF_DATA_SAVE
    Informa "", amensaje
    
    Exit Sub
Salvar_DatosErr:
    ErrorHandler Err, "Salvar_Datos"
End Sub

Public Sub CARGA_DATOS()
    If Me.gexAlmacenAduana.RowCount > 0 Then
        Me.txtCodigo.Text = gexAlmacenAduana.Value(gexAlmacenAduana.Columns("CODIGO").Index)
        Me.txtDescripcion.Text = gexAlmacenAduana.Value(gexAlmacenAduana.Columns("NOMBRE_ALMACEN").Index)
        Me.txtDireccion.Text = gexAlmacenAduana.Value(gexAlmacenAduana.Columns("DIRECCION").Index)
        Me.txtRuc.Text = gexAlmacenAduana.Value(gexAlmacenAduana.Columns("RUC").Index)
        Me.txtContacto.Text = gexAlmacenAduana.Value(gexAlmacenAduana.Columns("CONTACTO").Index)
        
    End If
    
    gexAlmacenAduana.Columns("CODIGO").Width = 800
    gexAlmacenAduana.Columns("NOMBRE_ALMACEN").Width = 3500
    gexAlmacenAduana.Columns("DIRECCION").Width = 3800
    gexAlmacenAduana.Columns("RUC").Width = 2000
    gexAlmacenAduana.Columns("CONTACTO").Width = 2500
    
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

Private Sub gexAlmacenAduana_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    Call CARGA_DATOS
End Sub



Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtDescripcion.SetFocus
End If
End Sub


Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtDireccion.SetFocus
End If
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtRuc.SetFocus
End If

End Sub


Private Sub txtRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AVANZA 13
        txtContacto.SetFocus
    Else
        Call SoloNumeros(txtRuc, KeyAscii, False)
    End If
End Sub
