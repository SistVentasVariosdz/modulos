VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMantItemCombDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Combinaciones"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Hilados"
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   4550
      Width           =   1965
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "frmMantItemCombDet.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Ultimo"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   15
         Picture         =   "frmMantItemCombDet.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Primero"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "frmMantItemCombDet.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "frmMantItemCombDet.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Anterior"
         Top             =   120
         Width           =   495
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
      Height          =   3270
      Left            =   90
      TabIndex        =   2
      Tag             =   "List"
      Top             =   105
      Width           =   6015
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2925
         Left            =   210
         TabIndex        =   3
         Top             =   255
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   5159
         _Version        =   393216
         Enabled         =   -1  'True
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
            DataField       =   "Des_Secuencia"
            Caption         =   "Secuencia"
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
            DataField       =   "Des_CombDet"
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
            MarqueeStyle    =   3
            BeginProperty Column00 
               ColumnWidth     =   1980.284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3240
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
      Height          =   1110
      Left            =   90
      TabIndex        =   0
      Tag             =   "Detail"
      Top             =   3405
      Width           =   6015
      Begin VB.TextBox txtDes_CombDet 
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   12
         Top             =   640
         Width           =   4335
      End
      Begin VB.ComboBox cboNum_Secuencia 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Descripción :"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Secuencia :"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   345
         Width           =   855
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2520
      TabIndex        =   4
      Top             =   4620
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantItemCombDet.frx":05C8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmMantItemCombDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public Codigo, Descripcion As String
Dim sTipo As String
Dim Rs_Grid As New ADODB.Recordset
Dim StrSQL As String
Public Codigo_item As String
Public Codigo_Comb As String

'Private Sub cboNum_Secuencia_Click()
'    strSQL = "SELECT Cod_HilTel FROM tx_hilostel WHERE Cod_Tela ='" & Codigo_item & "' AND Num_Secuencia='" & cboNum_Secuencia.Text & "'"
'    txtCod_HilTel.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
'    If Trim(txtCod_HilTel.Text) <> "" Then
'        Call txtCod_HilTel_KeyPress(13)
'    End If
'End Sub

'Private Sub cmdBusColor_Click()
'    Dim oTipo As New frmBusqGeneral
'    Dim rs As New ADODB.Recordset
'    Set oTipo.oParent = Me
'    oTipo.sQuery = "SELECT Cod_Color as Código, Des_Color as Descripción FROM LB_Color"
'    oTipo.Cargar_Datos
'    oTipo.Show 1
'    If Codigo <> "" Then
'        txtCod_Color.Text = Codigo
'        txtDes_Color.Text = Descripcion
'    End If
'    Set oTipo = Nothing
'    Set rs = Nothing
'End Sub

'Private Sub cmdBusHilo_Click()
'    Dim oTipo As New frmBusqGeneral
'    Dim rs As New ADODB.Recordset
'    Set oTipo.oParent = Me
'    oTipo.sQuery = "SELECT Cod_HilTel as Codigo, des_HilTel as Descripcion FROM IT_Hilado"
'    oTipo.Cargar_Datos
'    oTipo.Show 1
'    If Codigo <> "" Then
'        txtCod_HilTel.Text = Codigo
'        txtDes_HilTel.Text = Descripcion
'    End If
'    Set oTipo = Nothing
'    Set rs = Nothing
'End Sub

Private Sub cmdFirst_Click()
    If Not Rs_Grid.BOF Then
        Rs_Grid.MoveFirst
    End If
End Sub
Private Sub cmdLast_Click()
    If Not Rs_Grid.EOF Then
        Rs_Grid.MoveLast
    End If
End Sub
Private Sub cmdNext_Click()
    If Not Rs_Grid.EOF Then
        Rs_Grid.MoveNext
    End If
End Sub
Private Sub cmdPrevious_Click()
    If Not Rs_Grid.BOF Then
        Rs_Grid.MovePrevious
    End If
End Sub

Private Sub Form_Load()
    Call FormSet(Me)
    FormateaGrid Me.DGridLista
    HabilitaMant Me.MantFunc1, ""
    DESHABILITA_DATOS
    Call CargaCombos
    MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub

Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub
Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Rs_Grid.State <> 1 Then
        Exit Sub
    End If
    If Not Rs_Grid.EOF And Not Rs_Grid.BOF Then
        Call Carga_Datos
        DESHABILITA_DATOS
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Rs_Grid = Nothing
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Eliminar As Integer
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            LIMPIAR_DATOS
            HABILITA_DATOS
            cboNum_Secuencia.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "MODIFICAR"
            sTipo = "U"
            HABILITA_DATOS
            cboNum_Secuencia.Enabled = False
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "ELIMINAR"
            Eliminar = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Combinación-Detalle")
            If Eliminar = vbYes Then
                sTipo = "D"
                ELIMINAR_DATOS
                RECARGAR_DATOS
            End If
        Case "GRABAR"
            If VALIDA_DATOS Then
                SALVAR_DATOS
                RECARGAR_DATOS
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                DGridLista.Enabled = True
                sTipo = ""
            End If
            
        Case "DESHACER"
            LIMPIAR_DATOS
            RECARGAR_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridLista.Enabled = True
            sTipo = ""
         Case "SALIR"
            Unload Me
    End Select
End Sub

Sub LIMPIAR_DATOS()
    cboNum_Secuencia.ListIndex = -1
    txtDes_CombDet.Text = ""
    'txtCod_Color.Text = ""
    'txtDes_Color.Text = ""
    'txtCod_HilTel.Text = ""
    'txtDes_HilTel.Text = ""
End Sub


Public Sub CargaCombos()
   
    'Llena Combo de Secuencias
    StrSQL = "SELECT Des_Secuencia + SPACE(100) + Num_Secuencia FROM LG_ITEMPARTES WHERE Cod_Item ='" & Codigo_item & "'"
    Call LlenaCombo(cboNum_Secuencia, StrSQL, cCONNECT)
    
End Sub

Sub Carga_Datos()

    If Not Rs_Grid.EOF Then
       
        'txtNum_Secuencia.Text = Trim(Rs_Grid("Num_Secuencia").Value)
        txtDes_CombDet.Text = Trim(Rs_Grid("Des_CombDet").Value)
        
        'txtCod_Color.Text = Trim(Rs_Grid("Cod_Color").Value)
        'Call txtCod_Color_KeyPress(13)
        
        'txtCod_HilTel.Text = Trim(Rs_Grid("Cod_HilTel").Value)
        'Call txtCod_HilTel_KeyPress(13)
        
        Call BuscaCombo(Rs_Grid("Num_Secuencia").Value, 2, cboNum_Secuencia)
      
    End If

End Sub

Sub RECARGAR_DATOS()
    
    Rs_Grid.Close
    CARGA_GRID
    
End Sub

Public Sub CARGA_GRID()
    Dim StrSQL As String
    Set Rs_Grid = New ADODB.Recordset
    Rs_Grid.ActiveConnection = cCONNECT
    Rs_Grid.CursorType = adOpenStatic
    Rs_Grid.CursorLocation = adUseClient
    Rs_Grid.LockType = adLockReadOnly
    
    'Esta cadena es para devolver el Codigo de Cliente
    StrSQL = "EXEC UP_SEL_ITEMCOMBDET '" & Codigo_item & "','" & Codigo_Comb & "'"
    
    Rs_Grid.Open StrSQL
    Set DGridLista.DataSource = Rs_Grid
    DGridLista.Refresh

    If Rs_Grid.RecordCount > 0 Then
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Call Carga_Datos
    Else
        HabilitaMant Me.MantFunc1, "ADICIONAR"
        Call LIMPIAR_DATOS
    End If
End Sub

Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
    
    If sTipo = "I" Then
        If cboNum_Secuencia.Text = "" Then
            MsgBox ("Usted no ha seleccionado un Numero de Secuencia. Sirvase Verificar")
            VALIDA_DATOS = False
            Exit Function
        End If
    
        StrSQL = "SELECT Num_Secuencia FROM LG_ITEMCOMBDET WHERE Cod_Item='" & Codigo_item & "' AND Cod_Comb='" & Codigo_Comb & "' AND Num_Secuencia='" & Right(cboNum_Secuencia.Text, 2) & "'"
        If DevuelveCampo(StrSQL, cCONNECT) <> "" Then
            MsgBox "Ese registro ya se encuentra ingresado. Sirvase verificar", vbCritical, "Detalle-Combinación"
            VALIDA_DATOS = False
            Exit Function
        End If
    End If
    
    If sTipo <> "D" Then
        'Verifica Color del Hilado
        If Trim(txtDes_CombDet.Text) = "" Then
            MsgBox "La descripción no puede estar vacia. Sirvase verficar", vbInformation, "Detalle-Combinación"
            VALIDA_DATOS = False
            txtDes_CombDet.Text = ""
            txtDes_CombDet.SetFocus
        End If

'        If Trim(txtCod_HilTel.Text) <> "" Then
'            strSQL = "SELECT Cod_HilTel, des_HilTel FROM IT_Hilado WHERE Cod_HilTel='" & Trim(txtCod_HilTel.Text) & "'"
'            If DevuelveCampo(strSQL, cCONNECT) = "" Then
'                MsgBox ("El código de hilo ingresado no existe, Sirvase verficar")
'                    VALIDA_DATOS = False
'                    txtCod_HilTel.SetFocus
'            End If
'        End If
    End If
End Function

Sub HABILITA_DATOS()
   
    cboNum_Secuencia.Enabled = True
    txtDes_CombDet.Enabled = True
    
    'txtCod_Color.Enabled = True
    'txtDes_Color.Enabled = True
    'txtCod_HilTel.Enabled = True
    'txtDes_HilTel.Enabled = True
    'cmdBusHilo.Enabled = True
    'cmdBusColor.Enabled = True
    
End Sub

Sub DESHABILITA_DATOS()

    'txtDes_Comb.Enabled = False
    cboNum_Secuencia.Enabled = False
    txtDes_CombDet.Enabled = False
    
    'txtCod_Color.Enabled = False
    'txtDes_Color.Enabled = False
    'txtCod_HilTel.Enabled = False
    'txtDes_HilTel.Enabled = False
    'cmdBusHilo.Enabled = False
    'cmdBusColor.Enabled = False
End Sub

Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Dim StrSQL As String
    
    Con.ConnectionString = cCONNECT
    Con.Open
    
        Con.BeginTrans

        StrSQL = "EXEC UP_MAN_ITEMCOMBDET '" & _
        sTipo & "','" & _
        Codigo_item & "','" & _
        Codigo_Comb & "','" & _
        Right(cboNum_Secuencia.Text, 2) & "','" & _
        txtDes_CombDet.Text & "'"
        
        Con.Execute StrSQL

        Con.CommitTrans
        Dim amensaje As New clsMessages
        amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
        Informa "", amensaje
        
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
    Con.BeginTrans
       
        StrSQL = "EXEC UP_MAN_ITEMCOMBDET '" & _
        sTipo & "','" & _
        Codigo_item & "','" & _
        Codigo_Comb & "','" & _
        Right(cboNum_Secuencia.Text, 2) & "','" & _
        txtDes_CombDet.Text & "'"
        
        Con.Execute StrSQL
        
    Con.CommitTrans
    
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_DELETE
    Informa "", amensaje
    
Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Eliminar_Datos"

End Sub

'Private Sub txtDes_Color_Change()
'
'End Sub

'Private Sub txtCod_Color_KeyPress(KeyAscii As Integer)
'    Dim strSQL As String
'    If KeyAscii = 13 Then
'        If Trim(txtCod_Color.Text) = "" Then
'            cmdBusColor_Click
'        Else
'            strSQL = "SELECT Des_Color FROM LB_Color WHERE Cod_Color='" & txtCod_Color.Text & "'"
'            txtDes_Color.Text = DevuelveCampo(strSQL, cCONNECT)
'        End If
'    End If
'End Sub

'Private Sub txtCod_HilTel_KeyPress(KeyAscii As Integer)
'    Dim strSQL As String
'    If KeyAscii = 13 Then
'        If Trim(txtCod_HilTel.Text) = "" Then
'            cmdBusHilo_Click
'        Else
'            strSQL = "SELECT des_HilTel FROM IT_Hilado WHERE Cod_HilTel='" & txtCod_HilTel.Text & "'"
'            txtDes_HilTel.Text = DevuelveCampo(strSQL, cCONNECT)
'        End If
'    End If
'End Sub
'
'Private Sub txtDes_Color_KeyPress(KeyAscii As Integer)
'    Dim strSQL As String
'    If KeyAscii = 13 Then
'        If Len(txtDes_Color.Text) < 5 Then
'            Call MsgBox("La descripción debe tener como mínimo 5 caracteres. Sirvase verificar", vbInformation)
'            Exit Sub
'        Else
'            strSQL = "SELECT Cod_Color FROM LB_Color WHERE Des_Color LIKE '" & txtCod_Color.Text & "%'"
'            txtCod_Color.Text = DevuelveCampo(strSQL, cCONNECT)
'        End If
'    End If
'End Sub


'Private Sub txtDes_HilTel_KeyPress(KeyAscii As Integer)
'    Dim strSQL As String
'    If KeyAscii = 13 Then
'        If Len(txtDes_HilTel.Text) < 5 Then
'            Call MsgBox("La descripción debe tener como mínimo 5 caracteres. Sirvase verificar", vbInformation)
'            Exit Sub
'        Else
'            strSQL = "SELECT Cod_HilTel FROM IT_Hilado WHERE Des_HilTel LIKE '" & txtDes_HilTel.Text & "%'"
'            txtCod_HilTel.Text = DevuelveCampo(strSQL, cCONNECT)
'        End If
'    End If
'End Sub

