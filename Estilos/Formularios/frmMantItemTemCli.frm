VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmMantItemTemCli 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Temporada Cliente"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Style Component"
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
      Height          =   2835
      Left            =   60
      TabIndex        =   9
      Tag             =   "Detail"
      Top             =   3480
      Width           =   5445
      Begin VB.ComboBox cboCod_TemCli 
         Height          =   315
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   2895
      End
      Begin VB.ComboBox cboCod_Cliente 
         Height          =   315
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   2895
      End
      Begin VB.ComboBox cboCod_Item 
         Height          =   315
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
      Begin VB.ComboBox cboFlg_Status 
         Height          =   315
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2040
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpFec_Aprobacion 
         Height          =   315
         Left            =   1275
         TabIndex        =   6
         Top             =   2400
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   23658499
         CurrentDate     =   37209
      End
      Begin VB.TextBox txtComentario 
         Height          =   675
         Left            =   1275
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1320
         Width           =   3855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Temporada :"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cliente : "
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Item : "
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F. Aprobación :"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   2520
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Comentario"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Status :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Tag             =   "Type:"
         Top             =   2160
         Width           =   555
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
      Height          =   3495
      Left            =   60
      TabIndex        =   7
      Tag             =   "List"
      Top             =   0
      Width           =   5445
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2985
         Left            =   180
         TabIndex        =   8
         Top             =   345
         Width           =   5130
         _ExtentX        =   9049
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Nom_Cliente"
            Caption         =   "Cliente"
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
            DataField       =   "Nom_TemCli"
            Caption         =   "Temporada"
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
            DataField       =   "Flg_Status"
            Caption         =   "Status"
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
            DataField       =   "Fec_Aprobacion"
            Caption         =   "Aprobacion"
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
            DataField       =   "Comentario"
            Caption         =   "Comentario"
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
            MarqueeStyle    =   2
            BeginProperty Column00 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   6360
      Width           =   1965
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "frmMantItemTemCli.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Anterior"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "frmMantItemTemCli.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   15
         Picture         =   "frmMantItemTemCli.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Primero"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "frmMantItemTemCli.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Ultimo"
         Top             =   120
         Width           =   495
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   1920
      TabIndex        =   20
      Top             =   6435
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantItemTemCli.frx":05C8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmMantItemTemCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public Codigo, Descripcion As String
Public Codigo_item As String
Dim Fecha_Aprobacion As String
Dim sTipo As String
Dim Rs_Carga As New ADODB.Recordset

Private Sub cboCod_Cliente_Click()
    Dim strSQL As String
    'Combo de Temporadas
    strSQL = "SELECT Nom_TemCli + space(100) + Cod_TemCli FROM TG_TemCli WHERE Cod_Cliente='" & Right(cboCod_Cliente, 5) & "'"
    Call LlenaCombo(cboCod_TemCli, strSQL, cCONNECT)
End Sub

Private Sub cmdFirst_Click()
    If Not Rs_Carga.BOF Then
        Rs_Carga.MoveFirst
    End If
End Sub
Private Sub cmdLast_Click()
    If Not Rs_Carga.EOF Then
        Rs_Carga.MoveLast
    End If
End Sub
Private Sub cmdNext_Click()
    If Not Rs_Carga.EOF Then
        Rs_Carga.MoveNext
        If Rs_Carga.EOF Then
            Rs_Carga.MoveLast
        End If
    End If
End Sub
Private Sub cmdPrevious_Click()
    If Not Rs_Carga.BOF Then
        Rs_Carga.MovePrevious
        If Rs_Carga.BOF Then
            Rs_Carga.MoveFirst
        End If
    End If
End Sub
Sub Carga_Datos()
    Dim strSQL As String
    On Error GoTo Cargar_DatosErr
    
    Call BuscaCombo(Codigo_item, 2, cboCod_Item)
    
    strSQL = "EXEC UP_MAN_ITEMTEMCLI 'S','" & Codigo_item & "','','','','',''"
    Set Rs_Carga = Nothing
    Rs_Carga.ActiveConnection = cCONNECT
    Rs_Carga.CursorType = adOpenStatic
    Rs_Carga.CursorLocation = adUseClient
    Rs_Carga.LockType = adLockReadOnly
    Rs_Carga.Open strSQL
    Set DGridLista.DataSource = Rs_Carga
    'DGridLista_RowColChange 0, 0
    If Rs_Carga.RecordCount > 0 Then
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
    Else
        LIMPIAR_DATOS
        DESHABILITA_DATOS
        HabilitaMant Me.MantFunc1, "ADICIONAR"
    End If
    
    'Esta rutina es para cargar los datos
    If Rs_Carga.State <> 1 Then
        Exit Sub
    End If
    If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
    
        txtComentario.Text = Trim(Rs_Carga("Comentario"))
        If IsNull(Rs_Carga("Fec_Aprobacion")) Then
            dtpFec_Aprobacion.CustomFormat = " "
        Else
            dtpFec_Aprobacion.CustomFormat = "dd/MM/yyyy"
        End If
        
        dtpFec_Aprobacion = Trim(Rs_Carga("Fec_Aprobacion"))
           
        Call BuscaCombo(Rs_Carga("Cod_Item"), 2, cboCod_Item)
        Call BuscaCombo(Rs_Carga("Cod_Cliente"), 2, cboCod_Cliente)
        Call BuscaCombo(Rs_Carga("Cod_TemCli"), 2, cboCod_TemCli)
        Call BuscaCombo(Rs_Carga("Flg_Status"), 2, cboFlg_Status)
        DESHABILITA_DATOS
    End If
    
    
    Exit Sub
Cargar_DatosErr:
    Set Rs_Carga = Nothing
    ErrorHandler Err, "Cargar_Datos"
End Sub


Private Sub dtpFec_Aprobacion_Change()
    If IsNull(dtpFec_Aprobacion.Value) Then
        dtpFec_Aprobacion.CustomFormat = " "
        Fecha_Aprobacion = ""
    Else
        dtpFec_Aprobacion.CustomFormat = "dd/MM/yyyy"
        Fecha_Aprobacion = Format(dtpFec_Aprobacion.Value, "yyyyMMdd")
    End If
End Sub

Private Sub dtpFec_Aprobacion_Click()
    If IsNull(dtpFec_Aprobacion.Value) Then
        dtpFec_Aprobacion.CustomFormat = " "
        Fecha_Aprobacion = ""
    Else
        dtpFec_Aprobacion.CustomFormat = "dd/MM/yyyy"
        Fecha_Aprobacion = Format(dtpFec_Aprobacion.Value, "yyyyMMdd")
    End If
End Sub

Private Sub Form_Load()
    Call FormSet(Me)
    FormateaGrid Me.DGridLista
    Call CARGA_COMBOS
    Call DESHABILITA_DATOS
    MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub
Sub SALVAR_DATOS()
    Dim con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Dim strSQL As String
    con.ConnectionString = cCONNECT
    con.Open
    
        con.BeginTrans

        strSQL = "EXEC UP_MAN_ITEMTEMCLI '" & _
        sTipo & "','" & _
        Codigo_item & "','" & _
        Right(cboCod_Cliente.Text, 5) & "','" & _
        Right(cboCod_TemCli.Text, 3) & "','" & _
        Right(cboFlg_Status.Text, 1) & "','" & _
        Fecha_Aprobacion & "','" & _
        txtComentario.Text & "'"

        con.Execute strSQL

        con.CommitTrans
        Dim amensaje As New clsMessages
        amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
        Informa "", amensaje
    
    LIMPIAR_DATOS
    RECARGAR_DATOS
    Exit Sub
Salvar_DatosErr:
    con.RollbackTrans
    Set con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub
Sub ELIMINAR_DATOS()
    Dim con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
    Dim strSQL As String
    con.ConnectionString = cCONNECT
    con.Open
    
    con.BeginTrans
    
        strSQL = "EXEC UP_MAN_ITEMTEMCLI '" & _
        sTipo & "','" & _
        Codigo_item & "','" & _
        Right(cboCod_Cliente.Text, 5) & "','" & _
        Right(cboCod_TemCli.Text, 3) & "','" & _
        Right(cboFlg_Status.Text, 1) & "','" & _
        Format(dtpFec_Aprobacion.Value, "yyyyMMdd") & "','" & _
        txtComentario.Text & "'"
    
    con.Execute strSQL
   
    con.CommitTrans
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_DELETE
    Informa "", amensaje

    LIMPIAR_DATOS
    RECARGAR_DATOS
    
    Exit Sub
    
Eliminar_DatosErr:
    con.RollbackTrans
    Set con = Nothing
    ErrorHandler Err, "Eliminar_Datos"
End Sub
Sub LIMPIAR_DATOS()

    'cboCod_Item.ListIndex = -1
    cboCod_Cliente.ListIndex = -1
    cboCod_TemCli.ListIndex = -1
    cboFlg_Status.ListIndex = -1
    txtComentario.Text = ""
    dtpFec_Aprobacion = Now()
    
End Sub

Sub HABILITA_DATOS()

    'cboCod_Item.Enabled = True
    cboCod_Cliente.Enabled = True
    cboCod_TemCli.Enabled = True
    cboFlg_Status.Enabled = True
    txtComentario.Enabled = True
    dtpFec_Aprobacion.Enabled = True
    
End Sub
Sub DESHABILITA_DATOS()

    cboCod_Item.Enabled = False
    cboCod_Cliente.Enabled = False
    cboCod_TemCli.Enabled = False
    cboFlg_Status.Enabled = False
    txtComentario.Enabled = False
    dtpFec_Aprobacion.Enabled = False

End Sub
Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub
Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Rs_Carga.State <> 1 Then
        Exit Sub
    End If
    If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
    
        txtComentario.Text = Trim(Rs_Carga("Comentario"))
        If IsNull(Rs_Carga("Fec_Aprobacion")) Then
            dtpFec_Aprobacion.CustomFormat = " "
        Else
            dtpFec_Aprobacion.CustomFormat = "dd/MM/yyyy"
        End If
        
        dtpFec_Aprobacion = Trim(Rs_Carga("Fec_Aprobacion"))
           
        Call BuscaCombo(Rs_Carga("Cod_Item"), 2, cboCod_Item)
        Call BuscaCombo(Rs_Carga("Cod_Cliente"), 2, cboCod_Cliente)
        Call BuscaCombo(Rs_Carga("Cod_TemCli"), 2, cboCod_TemCli)
        Call BuscaCombo(Rs_Carga("Flg_Status"), 2, cboFlg_Status)
      
        DESHABILITA_DATOS
    End If
End Sub
Sub RECARGAR_DATOS()
    Rs_Carga.Close
    Carga_Datos
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Rs_Carga = Nothing
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
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
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            txtComentario.Enabled = True
            cboFlg_Status.Enabled = True
            dtpFec_Aprobacion.Enabled = True
            txtComentario.SetFocus
            DGridLista.Enabled = False
        Case "ELIMINAR"
            sTipo = "D"
            ELIMINAR_DATOS
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
Function VALIDA_DATOS() As Boolean
    Dim mensaje As String
    Dim strSQL As String
    Dim rsValida_Datos As New ADODB.Recordset
    
    VALIDA_DATOS = True
    
    mensaje = "Es necesario llenar los campos: "
    If Len(Trim(cboCod_Item.Text)) = 0 Then
        mensaje = mensaje & "Item"
        VALIDA_DATOS = False
    End If
    If Len(Trim(cboCod_Cliente.Text)) = 0 Then
        If VALIDA_DATOS = False Then
            mensaje = mensaje & ", Cliente"
        Else
            mensaje = mensaje & "Cliente"
        End If
        VALIDA_DATOS = False
    End If
    If Len(Trim(cboCod_TemCli.Text)) = 0 Then
        If VALIDA_DATOS = False Then
            mensaje = mensaje & ", Temporada"
        Else
            mensaje = mensaje & "Temporada"
        End If
        VALIDA_DATOS = False
    End If
    
    If VALIDA_DATOS = False Then
        MsgBox (mensaje)
    Else
        
        If sTipo = "I" Then
            strSQL = "SELECT Cod_Item FROM LG_ItemTemCli WHERE Cod_Item='" & Right(cboCod_Item.Text, 8) & "' AND Cod_Cliente='" & Right(cboCod_Cliente.Text, 5) & "' AND Cod_TemCli='" & Right(cboCod_TemCli.Text, 3) & "'"
        
            Set rsValida_Datos = New ADODB.Recordset
            rsValida_Datos.ActiveConnection = cCONNECT
            rsValida_Datos.CursorType = adOpenStatic
            rsValida_Datos.CursorLocation = adUseClient
            rsValida_Datos.LockType = adLockReadOnly
    
            rsValida_Datos.Open strSQL

            If rsValida_Datos.RecordCount > 0 Then
                MsgBox ("Ya existe un registro con los mismos datos. Sirvase verificar")
                VALIDA_DATOS = False
            End If
        End If
    End If
    
End Function

Public Sub CARGA_COMBOS()
    Dim strSQL As String
    
    'Combo de Clientes
    strSQL = "SELECT des_item + space(100) + Cod_item FROM LG_ITEM"
    Call LlenaCombo(cboCod_Item, strSQL, cCONNECT)
    
    'Combo de Clientes
    strSQL = "SELECT nom_cliente + space(100) + Cod_Cliente FROM TG_Cliente"
    Call LlenaCombo(cboCod_Cliente, strSQL, cCONNECT)
        
    'Combo Flag Estatus
    strSQL = "SELECT des_status + space(100) + flg_status  FROM TG_StaDes"
    Call LlenaCombo(cboFlg_Status, strSQL, cCONNECT)
End Sub

