VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMantHilosTel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hilados"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   12825
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Hilados"
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   5880
      Width           =   1965
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "frmMantHilosTel.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Ultimo"
         Top             =   30
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "frmMantHilosTel.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Siguiente"
         Top             =   30
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "frmMantHilosTel.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Anterior"
         Top             =   30
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   15
         Picture         =   "frmMantHilosTel.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Primero"
         Top             =   30
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
      Left            =   0
      TabIndex        =   13
      Tag             =   "List"
      Top             =   120
      Width           =   12765
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2925
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   12555
         _ExtentX        =   22146
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
         ColumnCount     =   6
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
            DataField       =   "Des_HilTel"
            Caption         =   "Hilo"
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
            DataField       =   "Porcentaje"
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
         BeginProperty Column03 
            DataField       =   "Long_Malla"
            Caption         =   "L. Malla"
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
         BeginProperty Column04 
            DataField       =   "Nom_Malla"
            Caption         =   "D.Malla"
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
         BeginProperty Column05 
            DataField       =   "cod_hilado_Estructurado"
            Caption         =   "Hilado Estructurado"
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
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   5790.047
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1709.858
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
      Height          =   2475
      Left            =   0
      TabIndex        =   12
      Tag             =   "Detail"
      Top             =   3405
      Width           =   12735
      Begin VB.TextBox TxtNum_Alimentadores 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         TabIndex        =   8
         Text            =   "0"
         Top             =   1730
         Width           =   780
      End
      Begin VB.TextBox TxtNum_Agujas 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Text            =   "0"
         Top             =   1730
         Width           =   780
      End
      Begin VB.ComboBox CboTorsion 
         Height          =   315
         Left            =   10800
         TabIndex        =   10
         Top             =   1680
         Width           =   975
      End
      Begin VB.ComboBox CboParafinado 
         Height          =   315
         Left            =   10800
         TabIndex        =   9
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox TxtCod_HiladoEstructurado 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1215
         TabIndex        =   6
         Top             =   1365
         Width           =   1515
      End
      Begin VB.TextBox txtDes_Tela 
         Enabled         =   0   'False
         Height          =   285
         Left            =   10770
         TabIndex        =   1
         Top             =   225
         Width           =   1830
      End
      Begin VB.TextBox txtPorcentaje 
         Height          =   285
         Left            =   1215
         TabIndex        =   4
         Text            =   "100"
         Top             =   990
         Width           =   975
      End
      Begin VB.TextBox txtNom_Malla 
         Enabled         =   0   'False
         Height          =   285
         Left            =   10770
         MaxLength       =   50
         TabIndex        =   3
         Top             =   600
         Width           =   1830
      End
      Begin VB.TextBox txtLong_Malla 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10800
         TabIndex        =   5
         Text            =   "0"
         Top             =   990
         Width           =   900
      End
      Begin VB.ComboBox cboCod_HilTel 
         Height          =   315
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   8925
      End
      Begin VB.TextBox txtNum_Secuencia 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1215
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Num. Alimentadores :"
         Height          =   195
         Left            =   2280
         TabIndex        =   30
         Top             =   1815
         Width           =   1500
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Num. Agujas :"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   1815
         Width           =   990
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Torsion :"
         Height          =   195
         Left            =   9600
         TabIndex        =   28
         Top             =   1695
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Parafinado :"
         Height          =   195
         Left            =   9600
         TabIndex        =   27
         Top             =   1335
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Hilado Antiguo"
         Height          =   195
         Left            =   105
         TabIndex        =   26
         Top             =   1410
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tela"
         Height          =   195
         Left            =   10200
         TabIndex        =   25
         Top             =   270
         Width           =   315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Longitud Malla :"
         Height          =   195
         Left            =   9585
         TabIndex        =   24
         Top             =   1005
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje"
         Height          =   195
         Left            =   135
         TabIndex        =   23
         Top             =   1050
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Hilado Nuevo"
         Height          =   255
         Left            =   135
         TabIndex        =   22
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Malla"
         Height          =   195
         Left            =   10215
         TabIndex        =   21
         Top             =   630
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Secuencia"
         Height          =   195
         Left            =   135
         TabIndex        =   20
         Top             =   345
         Width           =   765
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   3120
      TabIndex        =   11
      Top             =   5880
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantHilosTel.frx":05C8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmMantHilosTel"
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

Private Sub cboCod_HilTel_Click()

TxtCod_HiladoEstructurado = Trim(Right(cboCod_HilTel.Text, 10))
End Sub

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
    Dim Eliminar As Integer
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
    txtNum_Secuencia.Text = ""
    txtPorcentaje.Text = "100"
    txtLong_Malla.Text = "0"
    'txtDes_Tela.Text = ""
    txtNom_Malla = ""
    TxtCod_HiladoEstructurado = ""
    cboCod_HilTel.ListIndex = -1
    CboParafinado.ListIndex = -1
    CboTorsion.ListIndex = -1
    TxtNum_Agujas.Text = 0
    txtNum_Alimentadores.Text = 0
End Sub

Public Sub CargaCombo()
        
    'Combo Motivo Preproduccion
    'StrSQL = "SELECT COD_HILADO_ESTRUCTURADO + ' - ' + des_hiltel + space(100) + cod_hiltel FROM IT_Hilado WHERE COD_HILADO_ESTRUCTURADO<>'' ORDER BY COD_HILADO_ESTRUCTURADO"
    
    StrSQL = "SM_Muestra_It_Hilado"
    Call LlenaCombo(cboCod_HilTel, StrSQL, cCONNECT)
   
    CboParafinado.Clear
    CboParafinado.AddItem ("S")
    CboParafinado.AddItem ("N")
    CboParafinado.ListIndex = 0
    
    CboTorsion.Clear
    CboTorsion.AddItem ("S")
    CboTorsion.AddItem ("Z")
    CboTorsion.ListIndex = 0
    
End Sub


Sub Carga_Datos()

    If Not Rs_Grid.EOF Then
        txtNum_Secuencia.Text = Trim(Rs_Grid("Num_Secuencia").Value)
        txtPorcentaje.Text = Trim(Rs_Grid("Porcentaje").Value)
        txtLong_Malla.Text = Trim(Rs_Grid("Long_Malla").Value)
        txtNom_Malla.Text = Trim(Rs_Grid("Nom_Malla").Value)
        'TxtCod_HiladoEstructurado.Text = Trim(Rs_Grid("cod_hilado_estructurado").Value)
        'Call BuscaCombo(Rs_Grid("Cod_HilTel"), 1, cboCod_HilTel)
        TxtCod_HiladoEstructurado.Text = Trim(Rs_Grid("Cod_HilTel"))
        TxtNum_Agujas.Text = CDbl(Rs_Grid("Num_Agujas").Value)
        txtNum_Alimentadores.Text = CDbl(Rs_Grid("Num_Alimentadores").Value)
        Call BuscaCombo(Rs_Grid("cod_hilado_estructurado"), 1, cboCod_HilTel)
        Call BuscaCombo(Rs_Grid("Parafinado"), 1, CboParafinado)
        Call BuscaCombo(Rs_Grid("Torsion"), 1, CboTorsion)
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
    StrSQL = "EXEC UP_SEL_HILOSTEL '" & Codigo_tela & "'"
    
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
    If sTipo <> "D" Then
        If cboCod_HilTel.Text = "" Then
            MsgBox ("Debe seleccionar como mínimo un Hilado para realizar la operación")
            cboCod_HilTel.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
    Else
        StrSQL = "SELECT COUNT(Cod_Comb) FROM TX_TELACOMBDET WHERE Cod_Tela='" & Trim(Rs_Grid("Cod_Tela").Value) & "' AND Num_Secuencia='" & Trim(Rs_Grid("Num_Secuencia").Value) & "'"
        If DevuelveCampo(StrSQL, cCONNECT) > 0 Then
            Call MsgBox("No se puede eliminar el Registro por que posee registros relacionados", vbCritical)
            VALIDA_DATOS = False
            Exit Function
        End If
    End If
End Function

Sub HABILITA_DATOS()

    'txtNum_Secuencia.Enabled = True
    'txtNum_Secuencia.Enabled = True
    txtPorcentaje.Enabled = True
    txtLong_Malla.Enabled = True
    'txtCod_Tela.Enabled = True
    txtNom_Malla.Enabled = True
    
    cboCod_HilTel.Enabled = True
    CboParafinado.Enabled = True
    CboTorsion.Enabled = True
    
    TxtNum_Agujas.Enabled = True
    txtNum_Alimentadores.Enabled = True
End Sub

Sub DESHABILITA_DATOS()

    txtNum_Secuencia.Enabled = False
    txtPorcentaje.Enabled = False
    txtLong_Malla.Enabled = False
    'txtDes_Tela.Enabled = False
    txtNom_Malla.Enabled = False
    
    cboCod_HilTel.Enabled = False
    CboParafinado.Enabled = False
    CboTorsion.Enabled = False
    
    TxtNum_Agujas.Enabled = False
    txtNum_Alimentadores.Enabled = False
End Sub

Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Dim StrSQL As String
    
    Con.ConnectionString = cCONNECT
    Con.Open
    
        Con.BeginTrans

        StrSQL = "EXEC UP_MAN_HILOSTEL '" & _
                sTipo & "','" & _
                Codigo_tela & "','" & _
                txtNum_Secuencia.Text & "','" & _
                TxtCod_HiladoEstructurado & "','" & _
                txtPorcentaje.Text & "','" & _
                txtLong_Malla.Text & "','" & txtNom_Malla.Text & "','" & vusu & "','" & _
                CboParafinado.Text & "','" & CboTorsion.Text & "'," & Val(TxtNum_Agujas) & "," & Val(txtNum_Alimentadores)
        
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
       
        StrSQL = "EXEC UP_MAN_HILOSTEL '" & _
        sTipo & "','" & _
        Codigo_tela & "','" & _
        txtNum_Secuencia.Text & "','" & _
        TxtCod_HiladoEstructurado & "','" & _
        txtPorcentaje.Text & "','" & _
        txtLong_Malla.Text & "','" & txtNom_Malla.Text & "','" & vusu & "','" & _
        CboParafinado.Text & "','" & CboTorsion.Text & "'," & Val(TxtNum_Agujas) & "," & Val(txtNum_Alimentadores)
        
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

Private Sub txtLong_Malla_KeyPress(KeyAscii As Integer)
    SoloNumeros txtLong_Malla, KeyAscii, True, 2, 4
End Sub

Private Sub txtLong_Malla_LostFocus()
    If Trim(txtLong_Malla.Text) = "" Then
        txtLong_Malla.Text = "0"
    End If
End Sub

Private Sub TxtNum_Agujas_GotFocus()
SelectionText TxtNum_Agujas
End Sub

Private Sub TxtNum_Agujas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtNum_Agujas, KeyAscii, False)
End If
End Sub

Private Sub TxtNum_Alimentadores_GotFocus()
SelectionText txtNum_Alimentadores
End Sub

Private Sub TxtNum_Alimentadores_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(txtNum_Alimentadores, KeyAscii, False)
End If
End Sub

Private Sub txtPorcentaje_GotFocus()
SelectionText txtPorcentaje
End Sub

Private Sub txtPorcentaje_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(txtPorcentaje, KeyAscii, True, 2)
End If
End Sub

Private Sub txtPorcentaje_LostFocus()
    If Trim(txtPorcentaje.Text) = "" Then
        txtPorcentaje.Text = "0"
    End If
End Sub
