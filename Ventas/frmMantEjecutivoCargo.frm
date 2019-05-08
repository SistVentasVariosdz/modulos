VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "mantenimientos.ocx"
Begin VB.Form frmMantEjecutivoCarga 
   Caption         =   "Mantenimiento Ejecutivo Carga"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAgenteCarga_Des 
      Height          =   300
      Left            =   2580
      TabIndex        =   8
      Top             =   180
      Width           =   2700
   End
   Begin VB.TextBox txtAgenteCarga 
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      Top             =   150
      Width           =   915
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
      Height          =   1230
      Left            =   45
      TabIndex        =   3
      Top             =   4365
      Width           =   6900
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   720
         Width           =   4725
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1065
         Locked          =   -1  'True
         TabIndex        =   10
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
   Begin VB.Frame fraCargos 
      Caption         =   "Ejecutivo Carga"
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
      Left            =   165
      TabIndex        =   1
      Top             =   615
      Width           =   6675
      Begin GridEX20.GridEX gexEjecutivoCarga 
         Height          =   3345
         Left            =   120
         TabIndex        =   2
         Top             =   210
         Width           =   6435
         _ExtentX        =   11351
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
         Column(1)       =   "frmMantEjecutivoCargo.frx":0000
         Column(2)       =   "frmMantEjecutivoCargo.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmMantEjecutivoCargo.frx":016C
         FormatStyle(2)  =   "frmMantEjecutivoCargo.frx":02A4
         FormatStyle(3)  =   "frmMantEjecutivoCargo.frx":0354
         FormatStyle(4)  =   "frmMantEjecutivoCargo.frx":0408
         FormatStyle(5)  =   "frmMantEjecutivoCargo.frx":04E0
         FormatStyle(6)  =   "frmMantEjecutivoCargo.frx":0598
         ImageCount      =   0
         PrinterProperties=   "frmMantEjecutivoCargo.frx":0678
      End
   End
   Begin Mantenimientos.MantFunc MantFunc2 
      Height          =   540
      Left            =   1455
      TabIndex        =   6
      Top             =   5640
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantEjecutivoCargo.frx":0850
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin FunctionsButtons.FunctButt FBBuscar 
      Height          =   495
      Left            =   5595
      TabIndex        =   9
      Top             =   105
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Label Label2 
      Caption         =   "Agente Carga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   255
      TabIndex        =   7
      Top             =   195
      Width           =   1230
   End
End
Attribute VB_Name = "frmMantEjecutivoCarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim Stipo As String
Public Codigo As String
Public Descripcion As String
Public TipoAdd As String
Dim rstAux As ADODB.Recordset
Public oParent As Object
Public sAccion As String
Public sCOD As String, sDES As String


Private Sub FBBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Screen.MousePointer = 11


strSQL = "tg_muestra_AgenteCarga_Ejecutivo '" & txtAgenteCarga.Text & "'"

Set gexEjecutivoCarga.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

gexEjecutivoCarga.Columns("CODIGO").Width = 1110
gexEjecutivoCarga.Columns("NOMBRE_EJECUTIVO").Width = 4500

Screen.MousePointer = 0

End Sub

Private Sub Form_Load()
txtAgenteCarga.Text = sCOD
txtAgenteCarga_Des.Text = sDES
 Call INHABILITA_DATOS
 Call CARGA_GRID
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

Private Sub txtAgenteCarga_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    BuscaAgenteCarga 1
    If Me.txtAgenteCarga <> "" Then
        FBBuscar.SetFocus
    End If
End If
End Sub

Public Sub BuscaAgenteCarga(Tipo As Integer)

    strSQL = "SELECT Cod_AgenteCarga , Des_AgenteCarga FROM TG_AGENTECARGA WHERE "
    Select Case Tipo
    
        Case 1: strSQL = strSQL & "Cod_AgenteCarga  like '%" & txtAgenteCarga & "%'"
        Case 2: strSQL = strSQL & "Des_AgenteCarga LIKE '%" & txtAgenteCarga_Des & "%'"
    End Select

    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
        
    frmBusqGeneral3.gexLista.Columns("Cod_AgenteCarga").Width = 570
    frmBusqGeneral3.gexLista.Columns("des_AgenteCarga").Width = 2370
    
    frmBusqGeneral3.gexLista.Columns("Cod_AgenteCarga").Caption = "CODIGO AGENTE CARGA"
    frmBusqGeneral3.gexLista.Columns("des_AgenteCarga").Caption = "DESCRIPCION AGENTE"
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtAgenteCarga = ""
    txtAgenteCarga_Des = ""
    
    If Codigo <> "" Then
        txtAgenteCarga = Codigo
        txtAgenteCarga_Des = Descripcion
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    Codigo = ""
    Descripcion = ""
        
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
 
    strSQL = "EXEC AS_Man_AGENTE_CARGA_EJECUTIVO '" & Stipo & "','" & Trim(txtAgenteCarga.Text) & "','" & Trim(txtCodigo.Text) & "','" & Trim(txtDescripcion.Text) & "'"
      
    ExecuteCommandSQL cCONNECT, strSQL

    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMESSAGE_INF_DATA_SAVE
    Informa "", amensaje
    
    Exit Sub
Salvar_DatosErr:
    ErrorHandler err, "Salvar_Datos"
End Sub

Sub CARGA_GRID()
    
    strSQL = "EXEC tg_muestra_AgenteCarga_Ejecutivo '" & txtAgenteCarga.Text & "'"
    
    Set gexEjecutivoCarga.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
    
    SetGeneralGridEX gexEjecutivoCarga, 0, 1
    
    If Me.gexEjecutivoCarga.RowCount > 0 Then
        Call CARGA_DATOS
    End If
End Sub

Public Sub CARGA_DATOS()
    If Me.gexEjecutivoCarga.RowCount > 0 Then
        Me.txtCodigo.Text = gexEjecutivoCarga.Value(gexEjecutivoCarga.Columns("CODIGO").Index)
        Me.txtDescripcion.Text = gexEjecutivoCarga.Value(gexEjecutivoCarga.Columns("NOMBRE_EJECUTIVO").Index)
        
    End If

End Sub

Private Sub gexEjecutivoCarga_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    Call CARGA_DATOS
End Sub


