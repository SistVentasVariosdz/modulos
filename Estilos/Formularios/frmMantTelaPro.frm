VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMantTelaPro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesos"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Hilados"
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   4980
      Width           =   1965
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "frmMantTelaPro.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Ultimo"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   15
         Picture         =   "frmMantTelaPro.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Primero"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "frmMantTelaPro.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "frmMantTelaPro.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   9
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
      TabIndex        =   5
      Tag             =   "List"
      Top             =   105
      Width           =   6135
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2925
         Left            =   210
         TabIndex        =   6
         Top             =   255
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   5159
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "Num_Secuencia"
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
            DataField       =   "Des_ProTex"
            Caption         =   "Proceso"
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
         BeginProperty Column02 
            DataField       =   "Por_Merma"
            Caption         =   "Porcentaje"
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
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2700.284
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1140.095
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
      Height          =   1470
      Left            =   90
      TabIndex        =   0
      Tag             =   "Detail"
      Top             =   3405
      Width           =   6135
      Begin VB.TextBox txtPorcentaje 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Text            =   "0"
         Top             =   980
         Width           =   975
      End
      Begin VB.TextBox txtDes_Tela 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cboCod_ProTex 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtNum_Secuencia 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "% Merma"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1035
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Proceso"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tela"
         Height          =   195
         Left            =   3120
         TabIndex        =   14
         Top             =   345
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Secuencia"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   340
         Width           =   765
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2640
      TabIndex        =   7
      Top             =   5040
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantTelaPro.frx":05C8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmMantTelaPro"
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
Public Codigo_tela As String

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
    Call CargaCombo
    Call DESHABILITA_DATOS
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
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            LIMPIAR_DATOS
            HABILITA_DATOS
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "MODIFICAR"
            sTipo = "U"
            HABILITA_DATOS
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "ELIMINAR"
            sTipo = "D"
            ELIMINAR_DATOS
            RECARGAR_DATOS
        Case "GRABAR"
            If VALIDA_DATOS Then
                SALVAR_DATOS
                RECARGAR_DATOS
                DESHABILITA_DATOS
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                DGridLista.Enabled = True
            End If
            sTipo = ""
        Case "DESHACER"
            LIMPIAR_DATOS
            RECARGAR_DATOS
            DESHABILITA_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridLista.Enabled = True
            sTipo = ""
         Case "SALIR"
            Unload Me
    End Select
End Sub

Sub LIMPIAR_DATOS()

    txtNum_Secuencia.Text = ""
    txtNum_Secuencia.Text = ""
    txtPorcentaje.Text = 0
    'txtLong_Malla.Text = ""
    'txtDes_Tela.Text = ""
    
    cboCod_ProTex.ListIndex = -1

End Sub

Public Sub CargaCombo()
        
    'Combo Motivo Preproduccion
    StrSQL = "SELECT Des_ProTex + space(100) + cod_ProTex FROM TX_PROCESOS"
    Call LlenaCombo(cboCod_ProTex, StrSQL, cCONNECT)
   
End Sub


Sub Carga_Datos()

    If Not Rs_Grid.EOF Then
    
        txtNum_Secuencia.Text = Trim(Rs_Grid("Num_Secuencia").Value)
        txtPorcentaje.Text = Trim(Rs_Grid("Por_Merma").Value)
        'txtLong_Malla.Text = Trim(Rs_Grid("Long_Malla").Value)
        Call BuscaCombo(Rs_Grid("Cod_ProTex"), 2, cboCod_ProTex)
    
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
    StrSQL = "EXEC UP_SEL_TelaPro '" & Codigo_tela & "'"
    
    Rs_Grid.Open StrSQL
    Set DGridLista.DataSource = Rs_Grid
    DGridLista.Refresh

    If Rs_Grid.RecordCount > 0 Then
        'HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Call Carga_Datos
    Else
        'HabilitaMant Me.MantFunc1, "ADICIONAR"
        Call LIMPIAR_DATOS
    End If
End Sub

Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
    If cboCod_ProTex.Text = "" Then
        MsgBox ("Debe seleccionar como mínimo un Proceso para realizar la operación")
        VALIDA_DATOS = False
    End If
End Function

Sub HABILITA_DATOS()
    
    txtPorcentaje.Enabled = True
    'txtLong_Malla.Enabled = True
    cboCod_ProTex.Enabled = True
    
End Sub

Sub DESHABILITA_DATOS()

    txtNum_Secuencia.Enabled = False
    txtPorcentaje.Enabled = False
    'txtLong_Malla.Enabled = False
    'txtDes_Tela.Enabled = False
    
    cboCod_ProTex.Enabled = False
    
End Sub

Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Dim StrSQL As String
    
    Con.ConnectionString = cCONNECT
    Con.Open
    
        Con.BeginTrans

        StrSQL = "EXEC UP_MAN_TELAPRO '" & _
        sTipo & "','" & _
        Codigo_tela & "','" & _
        txtNum_Secuencia.Text & "','" & _
        Right(cboCod_ProTex.Text, 2) & "'," & _
        txtPorcentaje.Text & ",'" & vusu & "'"
        
        
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
       
        StrSQL = "EXEC UP_MAN_TELAPRO '" & _
        sTipo & "','" & _
        Codigo_tela & "','" & _
        txtNum_Secuencia.Text & "','" & _
        Right(cboCod_ProTex.Text, 2) & "'," & _
        txtPorcentaje.Text & ",'" & vusu & "'"
        
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

'Private Sub txtLong_Malla_KeyPress(KeyAscii As Integer)
'    SoloNumeros txtLong_Malla, KeyAscii, True, 4, 2
'End Sub
'
'Private Sub txtPorcentaje_KeyPress(KeyAscii As Integer)
'    SoloNumeros txtPorcentaje, KeyAscii, True, 3, 2
'End Sub
