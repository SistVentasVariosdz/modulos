VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmProcesos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesos"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   45
      TabIndex        =   15
      Top             =   5925
      Width           =   1965
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "frmProcesos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Anterior"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "frmProcesos.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   15
         Picture         =   "frmProcesos.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Primero"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "frmProcesos.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Ultimo"
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame FraLista 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   45
      TabIndex        =   13
      Top             =   -15
      Width           =   5760
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2985
         Left            =   120
         TabIndex        =   14
         Top             =   225
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   5265
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "Cod_ProTex"
            Caption         =   "Código"
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
            DataField       =   "Des_ProTex"
            Caption         =   "Descripción"
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
         BeginProperty Column02 
            DataField       =   "Rep_Predolar"
            Caption         =   "Precio ($)"
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
         BeginProperty Column03 
            DataField       =   "tip_precio"
            Caption         =   "Costeo"
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
            MarqueeStyle    =   2
            BeginProperty Column00 
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2715.024
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDetalle 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   45
      TabIndex        =   1
      Top             =   3450
      Width           =   5760
      Begin VB.TextBox txtCod_ProTex 
         Height          =   315
         Left            =   1185
         MaxLength       =   2
         TabIndex        =   8
         Top             =   300
         Width           =   750
      End
      Begin VB.TextBox TxtPrecio 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Text            =   "0"
         Top             =   1260
         Width           =   765
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   1320
         TabIndex        =   4
         Top             =   1680
         Width           =   2775
         Begin VB.OptionButton OpTejido 
            Caption         =   "Post Tejido"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton OpTenido 
            Caption         =   "Post Teñido"
            Height          =   255
            Left            =   1440
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.TextBox txtDes_ProTex 
         Height          =   315
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   3
         Top             =   780
         Width           =   4305
      End
      Begin VB.ComboBox CboCosteo 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1260
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   195
         TabIndex        =   12
         Top             =   315
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción :"
         Height          =   195
         Left            =   195
         TabIndex        =   11
         Top             =   795
         Width           =   930
      End
      Begin VB.Label Label3 
         Caption         =   "Precio Unit. $"
         Height          =   285
         Left            =   210
         TabIndex        =   10
         Top             =   1290
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Costeo (K/M) :"
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2115
      TabIndex        =   0
      Top             =   6015
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmProcesos.frx":05C8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs_Lista As ADODB.Recordset
Dim StrSQL As String
Dim sTipo As String

Private Sub cmdFirst_Click()
    If Not Rs_Lista.BOF Then
        Rs_Lista.MoveFirst
    End If
End Sub
Private Sub cmdLast_Click()
    If Not Rs_Lista.EOF Then
        Rs_Lista.MoveLast
    End If
End Sub
Private Sub cmdNext_Click()
    If Not Rs_Lista.EOF Then
        Rs_Lista.MoveNext
    End If
End Sub
Private Sub cmdPrevious_Click()
    If Not Rs_Lista.BOF Then
        Rs_Lista.MovePrevious
    End If
End Sub


Sub LIMPIAR_DATOS()
    txtCod_ProTex.Text = ""
    txtDes_ProTex.Text = ""
    TxtPrecio.Text = 0
    CboCosteo.ListIndex = -1
    OpTejido.Value = False
    OpTenido.Value = True
End Sub

Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
    If sTipo <> "D" Then
    
        If sTipo = "I" Then
            If ExisteCampo("Cod_ProTex", "TX_PROCESOS", Trim(txtCod_ProTex.Text), cCONNECT, True) Then
                MsgBox "El código ingresado ya se encuentra registrado. Sirvase verficar", vbInformation, "Procesos Textíles"
                txtCod_ProTex.SetFocus
                VALIDA_DATOS = False
                Exit Function
            End If
        End If
        
        If Trim(txtCod_ProTex.Text) = "" Then
            MsgBox "El código de condición de venta no puede estar vacío. Sirvase verificar", vbInformation, "Procesos Textíles"
            txtCod_ProTex.Text = ""
            txtCod_ProTex.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If

        If Trim(txtDes_ProTex.Text) = "" Then
            MsgBox "La descripción de condición de venta no puede estar vacío. Sirvase verificar", vbInformation, "Procesos Textíles"
            txtDes_ProTex.Text = ""
            txtDes_ProTex.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
    
        If Trim(Right(CboCosteo.Text, 2)) = "" Then
            MsgBox "El costeo no puede estar vacío. Sirvase verificar", vbInformation, "Procesos Textíles"
            CboCosteo.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
    Else
        'Aqui se valida que no tenga registros dependientes
        StrSQL = "SELECT COUNT(*) FROM TX_TELAPRO WHERE Cod_ProTex = '" & Trim(txtCod_ProTex.Text) & "'"
        If DevuelveCampo(StrSQL, cCONNECT) > 0 Then
            MsgBox "El registro seleccionado posee elementos relacionados. Sírvase verificar", vbInformation, "Procesos Textíles"
            VALIDA_DATOS = False
            Exit Function
        End If
    End If
End Function

Sub Carga_Datos()

    If Not Rs_Lista.EOF And Not Rs_Lista.BOF Then
        txtCod_ProTex.Text = Rs_Lista("Cod_ProTex").Value
        txtDes_ProTex.Text = Rs_Lista("Des_ProTex").Value
        TxtPrecio.Text = Rs_Lista("rep_predolar").Value
        Call BuscaCombo(Rs_Lista("tip_precio").Value, 2, CboCosteo)
        If Rs_Lista("flg_tejten").Value = "J" Then
            OpTejido.Value = True
            OpTenido.Value = False
        Else
            OpTejido.Value = False
            OpTenido.Value = True
        End If
    End If
End Sub

Sub HABILITA_DATOS()
    txtCod_ProTex.Enabled = True
    txtDes_ProTex.Enabled = True
    TxtPrecio.Enabled = True
    Frame2.Enabled = True
    CboCosteo.Enabled = True
End Sub

Sub INHABILITA_DATOS()
    txtCod_ProTex.Enabled = False
    txtDes_ProTex.Enabled = False
    TxtPrecio.Enabled = False
    Frame2.Enabled = False
    CboCosteo.Enabled = False
End Sub

Sub CARGA_GRID()
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.ActiveConnection = cCONNECT
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    'Esta cadena es para devolver el Codigo de Cliente
    StrSQL = "EXEC UP_SEL_PROCESOS "
    
    Rs_Lista.Open StrSQL
    Set DGridLista.DataSource = Rs_Lista
    DGridLista.Refresh

    If Rs_Lista.RecordCount > 0 Then
        DGridLista.Enabled = True
        'HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Call Carga_Datos
    Else
        DGridLista.Enabled = False
        'HabilitaMant Me.MantFunc1, "ADICIONAR"
        Call LIMPIAR_DATOS
    End If
End Sub

Sub Llena_Combo()
    CboCosteo.AddItem "Kilos                                K", 0
    CboCosteo.AddItem "Metros                               M", 1
End Sub

Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Dim StrSQL As String
    
    Con.ConnectionString = cCONNECT
    Con.Open
    
        Con.BeginTrans

        StrSQL = "EXEC UP_MAN_PROCESOS '" & _
        sTipo & "','" & _
        Trim(txtCod_ProTex.Text) & "','" & _
        Trim(txtDes_ProTex.Text) & "',''," & Val(TxtPrecio.Text) & ",'" & IIf(OpTenido.Value = True, "T", "J") & "','" & Trim(Right(CboCosteo.Text, 2)) & "'"
        
        
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
       
        StrSQL = "EXEC UP_MAN_PROCESOS '" & _
        sTipo & "','" & _
        Trim(txtCod_ProTex.Text) & "','" & _
        Trim(txtDes_ProTex.Text) & "',''," & Val(TxtPrecio.Text) & ",''"
        
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

Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call Carga_Datos
End Sub

Private Sub Form_Load()
'cCONNECT = "Provider=sqloledb;Server=SERVIDOR;Database=lives;UID=sa;pwd=;"
'cSEGURIDAD = "Provider=sqloledb;Server=servidor;Database=seguridad;UID=sa;pwd=;"
    Call FormSet(Me)
    FormateaGrid Me.DGridLista
    Call INHABILITA_DATOS
    Call CARGA_GRID
    Call Llena_Combo
    MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Eliminar As Integer
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
            HABILITA_DATOS
            txtCod_ProTex.Enabled = False
            txtDes_ProTex.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "ELIMINAR"
            Eliminar = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Combinación-Detalle")
            If Eliminar = vbYes Then
                sTipo = "D"
                If VALIDA_DATOS Then
                    Call ELIMINAR_DATOS
                    Call CARGA_GRID
                End If
                Call INHABILITA_DATOS
            End If
        Case "GRABAR"
            If VALIDA_DATOS Then
                Call SALVAR_DATOS
                Call CARGA_GRID
                Call INHABILITA_DATOS
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                DGridLista.Enabled = True
            End If
        Case "DESHACER"
            Call LIMPIAR_DATOS
            Call Carga_Datos
            Call INHABILITA_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridLista.Enabled = True
         Case "SALIR"
            Unload Me
    End Select
End Sub


Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    Call SoloNumeros(TxtPrecio, KeyAscii, True, 4)
End Sub


