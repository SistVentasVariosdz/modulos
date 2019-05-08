VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMantHilosItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Composición"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Hilados"
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
      Width           =   6135
      Begin VB.TextBox txtDes_Secuencia 
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   650
         Width           =   3135
      End
      Begin VB.TextBox txtDes_Item 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtNum_Secuencia 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descripción :"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Item :"
         Height          =   195
         Left            =   3120
         TabIndex        =   12
         Top             =   345
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Secuencia"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   340
         Width           =   765
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   4500
      Width           =   1965
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "frmMantHilosItem.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Ultimo"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   15
         Picture         =   "frmMantHilosItem.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Primero"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "frmMantHilosItem.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "frmMantHilosItem.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   7
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
      TabIndex        =   3
      Tag             =   "List"
      Top             =   105
      Width           =   6135
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2925
         Left            =   210
         TabIndex        =   4
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
         ColumnCount     =   2
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
            DataField       =   "Des_Secuencia"
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
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4155.024
            EndProperty
         EndProperty
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2640
      TabIndex        =   5
      Top             =   4560
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantHilosItem.frx":05C8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmMantHilosItem"
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
    'MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
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
            txtDes_Secuencia.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "MODIFICAR"
            sTipo = "U"
            HABILITA_DATOS
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "ELIMINAR"
            Eliminar = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Combinación-Detalle")
            If Eliminar = vbYes Then
                sTipo = "D"
                If VALIDA_DATOS Then
                    ELIMINAR_DATOS
                    RECARGAR_DATOS
                End If
            End If
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

Sub LIMPIAR_DATOS()

    txtNum_Secuencia.Text = ""
    txtDes_Secuencia.Text = ""
    'txtPorcentaje.Text = ""
    'txtLong_Malla.Text = ""
    'txtDes_Tela.Text = ""
    
    'cboCod_HilTel.ListIndex = -1

End Sub

Public Sub CargaCombo()
        
    'Combo Motivo Preproduccion
    'strSQL = "SELECT cod_hiltel + ' - ' + des_hiltel FROM IT_Hilado ORDER BY cod_hiltel"
    'Call LlenaCombo(cboCod_HilTel, strSQL, cCONNECT)
   
End Sub


Sub Carga_Datos()

    If Not Rs_Grid.EOF Then
    
        txtNum_Secuencia.Text = Trim(Rs_Grid("Num_Secuencia").Value)
        txtDes_Secuencia.Text = Trim(Rs_Grid("Des_Secuencia").Value)
        'txtPorcentaje.Text = Trim(Rs_Grid("Porcentaje").Value)
        'txtLong_Malla.Text = Trim(Rs_Grid("Long_Malla").Value)
        'txtCod_Tela.Text = Trim(Rs_Grid("Cod_Tela").Value)
        'Call BuscaCombo(Rs_Grid("Cod_HilTel"), 1, cboCod_HilTel)
    
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
    StrSQL = "EXEC UP_SEL_HILOSITEM '" & Codigo_item & "'"
    
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
       
    If sTipo <> "D" Then
    
        If Trim(txtDes_Secuencia.Text) = "" Then
            MsgBox "La Descripción no puede estar vacia. Sirvase verificar", vbInformation, "Composición"
            VALIDA_DATOS = False
            Exit Function
        End If
    
        'If cboCod_HilTel.Text = "" Then
        '    MsgBox ("Debe seleccionar como mínimo un Hilado para realizar la operación")
        '    VALIDA_DATOS = False
        '    Exit Function
        'End If
    Else
        StrSQL = "SELECT COUNT(Cod_Comb) FROM LG_ITEMCOMBDET WHERE Cod_Item='" & Trim(Rs_Grid("Cod_Item").Value) & "' AND Num_Secuencia='" & Trim(Rs_Grid("Num_Secuencia").Value) & "'"
        If DevuelveCampo(StrSQL, cCONNECT) > 0 Then
            Call MsgBox("No se puede eliminar el Registro por que posee registros relacionados", vbCritical)
            VALIDA_DATOS = False
            Exit Function
        End If
    End If
End Function

Sub HABILITA_DATOS()

    txtNum_Secuencia.Enabled = True
    txtDes_Secuencia.Enabled = True
    'txtPorcentaje.Enabled = True
    'txtLong_Malla.Enabled = True
    'txtCod_Tela.Enabled = True
    
    'cboCod_HilTel.Enabled = True
    
End Sub

Sub DESHABILITA_DATOS()

    txtNum_Secuencia.Enabled = False
    txtDes_Secuencia.Enabled = False
    'txtPorcentaje.Enabled = False
    'txtLong_Malla.Enabled = False
    'txtDes_Tela.Enabled = False
    
    'cboCod_HilTel.Enabled = False
    
End Sub

Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Dim StrSQL As String
    
    Con.ConnectionString = cCONNECT
    Con.Open
    
        Con.BeginTrans

        StrSQL = "EXEC UP_MAN_HILOSITEM '" & _
        sTipo & "','" & _
        Codigo_item & "','" & _
        txtNum_Secuencia.Text & "','" & _
        txtDes_Secuencia & "'"
       
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
       
        StrSQL = "EXEC UP_MAN_HILOSITEM '" & _
        sTipo & "','" & _
        Codigo_item & "','" & _
        Rs_Grid("Num_Secuencia").Value & "','" & _
        Rs_Grid("Des_Secuencia").Value & "'"
        
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
'    SoloNumeros txtLong_Malla, KeyAscii, True, 2, 4
'End Sub

'Private Sub txtLong_Malla_LostFocus()
'    If Trim(txtLong_Malla.Text) = "" Then
'        txtLong_Malla.Text = 0
'    End If
'End Sub

'Private Sub txtPorcentaje_KeyPress(KeyAscii As Integer)
'    SoloNumeros txtPorcentaje, KeyAscii, True, 2, 3
'End Sub
