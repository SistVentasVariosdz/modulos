VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Begin VB.Form frmClaOrdComp 
   Caption         =   "Clases de Orden de Compra"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2295
      TabIndex        =   2
      Top             =   5190
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmClaOrdComp.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
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
      Height          =   1650
      Left            =   60
      TabIndex        =   1
      Top             =   3420
      Width           =   5910
      Begin VB.ComboBox CmbPresentacion 
         Height          =   315
         ItemData        =   "frmClaOrdComp.frx":0160
         Left            =   4185
         List            =   "frmClaOrdComp.frx":016A
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1080
         Width           =   1680
      End
      Begin VB.ComboBox CmbTipReq 
         Height          =   315
         ItemData        =   "frmClaOrdComp.frx":0174
         Left            =   1455
         List            =   "frmClaOrdComp.frx":017E
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1080
         Width           =   1470
      End
      Begin VB.ComboBox cboFlg_Requerimiento 
         Height          =   315
         ItemData        =   "frmClaOrdComp.frx":0188
         Left            =   1470
         List            =   "frmClaOrdComp.frx":0192
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   690
         Width           =   810
      End
      Begin VB.TextBox txtTip_Item 
         Height          =   285
         Left            =   4185
         TabIndex        =   11
         Top             =   690
         Width           =   945
      End
      Begin VB.TextBox txtDes_ClaOrdComp 
         Height          =   285
         Left            =   2130
         MaxLength       =   50
         TabIndex        =   10
         Top             =   300
         Width           =   3705
      End
      Begin VB.TextBox txtCod_ClaOrdComp 
         Height          =   285
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   9
         Top             =   300
         Width           =   660
      End
      Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
         Left            =   2490
         Top             =   630
         _cx             =   847
         _cy             =   847
         PassiveMode     =   0   'False
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Presentacion"
         Height          =   195
         Index           =   2
         Left            =   3150
         TabIndex        =   19
         Top             =   1125
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Requerim:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   17
         Top             =   1125
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Item"
         Height          =   195
         Left            =   3165
         TabIndex        =   14
         Top             =   705
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Requerimiento :"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   13
         Top             =   735
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "O. Compra :"
         Height          =   195
         Left            =   195
         TabIndex        =   12
         Top             =   315
         Width           =   840
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
      TabIndex        =   0
      Top             =   15
      Width           =   5925
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2985
         Left            =   180
         TabIndex        =   8
         Top             =   225
         Width           =   5640
         _ExtentX        =   9948
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "Cod_ClaOrdComp"
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
            DataField       =   "Des_ClaOrdComp"
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
            DataField       =   "Flg_Requerimiento"
            Caption         =   "Flag. Reque."
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
            DataField       =   "Tip_Item"
            Caption         =   "Tipo Item"
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
         BeginProperty Column04 
            DataField       =   "Des_TipRequ"
            Caption         =   "Tipo Req"
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
            DataField       =   "Des_Presentacion"
            Caption         =   "Presentacion"
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
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2954.835
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   915.024
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   180
      TabIndex        =   3
      Top             =   5115
      Width           =   1965
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "frmClaOrdComp.frx":019C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Ultimo"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   15
         Picture         =   "frmClaOrdComp.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Primero"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "frmClaOrdComp.frx":0480
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "frmClaOrdComp.frx":05F2
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Anterior"
         Top             =   120
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmClaOrdComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs_Lista As ADODB.Recordset
Dim StrSql As String
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
    txtCod_ClaOrdComp.Text = ""
    txtDes_ClaOrdComp.Text = ""
    cboFlg_Requerimiento.ListIndex = 0
    txtTip_Item.Text = ""
    
    Me.CmbPresentacion.ListIndex = -1
    Me.CmbTipReq.ListIndex = -1

End Sub

Function VALIDA_DATOS() As Boolean
On Error GoTo hand
    VALIDA_DATOS = True
    If sTipo <> "D" Then
    
        If sTipo = "I" Then
            If ExisteCampo("Cod_ClaOrdComp", "Lg_ClaOrdComp", Trim(txtCod_ClaOrdComp.Text), cConnect, True) Then
                MsgBox "El código de la clase de orden de compra ya se encuentra registrado. Sirvase verificar", vbInformation, "Clase de Orden de Compra"
                txtCod_ClaOrdComp.SetFocus
                VALIDA_DATOS = False
                Exit Function
            End If
        End If
    
        If Trim(txtCod_ClaOrdComp.Text) = "" Then
            MsgBox "El código de la clase de orden de compra no puede estar vacío. Sirvase verificar", vbInformation, "Clase de Orden de Compra"
            txtCod_ClaOrdComp.Text = ""
            txtCod_ClaOrdComp.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If

        If Trim(txtDes_ClaOrdComp.Text) = "" Then
            MsgBox "La descripción de la Orden de Compra no puede estar vacío. Sirvase verificar", vbInformation, "Clase de Orden de Compra"
            txtDes_ClaOrdComp.Text = ""
            txtDes_ClaOrdComp.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
        
        If Trim(txtTip_Item.Text) = "" Then
            MsgBox "El tipo de Item no puede estar vacío. Sirvase verificar", vbInformation, "Clase de Orden de Compra"
            txtTip_Item.Text = ""
            txtTip_Item.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
        
        If Me.txtTip_Item = "H" Or Me.txtTip_Item = "T" Then
            If Me.CmbPresentacion = "" Then
                MsgBox "Debe seleccionar una presentacion", vbInformation
                VALIDA_DATOS = False
                Exit Function
            End If
        End If
        
    Else
        'Aqui se valida que no tenga registros dependientes
        'Strsql = "SELECT COUNT(Cod_Comb) FROM TX_TELACOMBDET WHERE Cod_Tela='" & Trim(Rs_Grid("Cod_Tela").Value) & "' AND Num_Secuencia='" & Trim(Rs_Grid("Num_Secuencia").Value) & "'"
        'If DevuelveCampo(Strsql, cCONNECT) > 0 Then
        '    MsgBox "Debe seleccionar como mínimo un Hilado para realizar la operación", vbInformation, "Condición de Venta"
        '    VALIDA_DATOS = False
        '    Exit Function
        'End If
    End If
Exit Function
hand:
ErrorHandler Err, "VALIDA_DATOS"
End Function

Sub CARGA_DATOS()

    If Not Rs_Lista.EOF And Not Rs_Lista.BOF Then
        txtCod_ClaOrdComp = Rs_Lista("Cod_ClaOrdComp").Value
        txtDes_ClaOrdComp = Rs_Lista("Des_ClaOrdComp").Value
        txtTip_Item = Rs_Lista("Tip_Item").Value
        Call BuscaCombo(Rs_Lista("Flg_Requerimiento").Value, 1, cboFlg_Requerimiento)
        BuscaCombo Rs_Lista("Des_TipRequ"), 1, Me.CmbTipReq
        BuscaCombo Rs_Lista("Des_Presentacion"), 1, Me.CmbPresentacion
    End If
End Sub

Sub HABILITA_DATOS()

    txtCod_ClaOrdComp.Enabled = True
    txtDes_ClaOrdComp.Enabled = True
    txtTip_Item.Enabled = True
    cboFlg_Requerimiento.Enabled = True
    Me.CmbPresentacion.Enabled = True
    Me.CmbTipReq.Enabled = True
    
End Sub

Sub INHABILITA_DATOS()
    txtCod_ClaOrdComp.Enabled = False
    txtDes_ClaOrdComp.Enabled = False
    txtTip_Item.Enabled = False
    cboFlg_Requerimiento.Enabled = False
    Me.CmbPresentacion.Enabled = False
    Me.CmbTipReq.Enabled = False
End Sub

Sub CARGA_GRID()
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.ActiveConnection = cConnect
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    'Esta cadena es para devolver el Codigo de Cliente
    StrSql = "EXEC UP_SEL_CLAORDCOMP "
    
    Rs_Lista.Open StrSql
    Set DGridLista.DataSource = Rs_Lista
    DGridLista.Refresh

    If Rs_Lista.RecordCount > 0 Then
        DGridLista.Enabled = True
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Call CARGA_DATOS
    Else
        DGridLista.Enabled = False
        HabilitaMant Me.MantFunc1, "ADICIONAR"
        Call LIMPIAR_DATOS
    End If
End Sub

Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Dim StrSql As String
    
    Con.ConnectionString = cConnect
    Con.Open
    
        Con.BeginTrans

        StrSql = "EXEC UP_MAN_CLAORDCOMP '" & _
        sTipo & "','" & _
        Trim(txtCod_ClaOrdComp.Text) & "','" & _
        Trim(txtDes_ClaOrdComp.Text) & "','" & _
        Trim(cboFlg_Requerimiento.Text) & "','" & _
        Trim(txtTip_Item.Text) & "','" & _
        Trim(Right(Me.CmbTipReq, 4)) & "','" & _
        Trim(Right(Me.CmbPresentacion, 2)) & "'"
        Con.Execute StrSql

        Con.CommitTrans
        Dim amensaje As New clsMessages
        amensaje.Codigo = CodeMsg.kMESSAGE_INF_DATA_SAVE
        'Informa "", amensaje
        
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub
Sub ELIMINAR_DATOS()
 Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
   
    Con.ConnectionString = cConnect
    Con.Open
    Con.BeginTrans
       
        StrSql = "EXEC UP_MAN_CLAORDCOMP '" & _
        sTipo & "','" & _
        Trim(txtCod_ClaOrdComp.Text) & "','" & _
        Trim(txtDes_ClaOrdComp.Text) & "','" & _
        Trim(cboFlg_Requerimiento.Text) & "','" & _
        Trim(txtTip_Item.Text) & "'"
        
        Con.Execute StrSql
    
    Con.CommitTrans
    
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMESSAGE_INF_DATA_DELETE
    'Informa "", amensaje
    
Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Eliminar_Datos"

End Sub

Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call CARGA_DATOS
End Sub

Private Sub Form_Load()
'cCONNECT = "Provider=sqloledb;Server=SERVIDOR;Database=lives;UID=sa;pwd=;"
'cSEGURIDAD = "Provider=sqloledb;Server=servidor;Database=seguridad;UID=sa;pwd=;"

    LlenaCombo CmbPresentacion, "select Des_Presentacion  +space(100)+Tip_Presentacion  from lg_presentacion order by 1", cConnect
    LlenaCombo Me.CmbTipReq, "select Des_TipRequ  +space(100)+Cod_TipRequ  from lg_tipreq order by 1", cConnect
    
    Call FormateaGrid(DGridLista)
    Call INHABILITA_DATOS
    Call CARGA_GRID
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo hand
    Dim eliminar As Integer
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            LIMPIAR_DATOS
            HABILITA_DATOS
            txtCod_ClaOrdComp.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "MODIFICAR"
            sTipo = "U"
            HABILITA_DATOS
            txtCod_ClaOrdComp.Enabled = False
            txtDes_ClaOrdComp.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "ELIMINAR"
            eliminar = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Combinación-Detalle")
            If eliminar = vbYes Then
                sTipo = "D"
                If VALIDA_DATOS Then
                    Call ELIMINAR_DATOS
                    Call CARGA_GRID
                End If
            End If
        Case "GRABAR"
            If VALIDA_DATOS Then
                Call SALVAR_DATOS
                Call CARGA_GRID
                Call INHABILITA_DATOS
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                DGridLista.Enabled = True
            End If
            Call INHABILITA_DATOS
        Case "DESHACER"
            Call LIMPIAR_DATOS
            Call CARGA_DATOS
            Call INHABILITA_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridLista.Enabled = True
         Case "SALIR"
            Unload Me
    End Select
Exit Sub

hand:
ErrorHandler Err, "MantFunc1_ActionClick"
End Sub
