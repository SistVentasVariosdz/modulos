VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMantFamItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Familia de Item"
   ClientHeight    =   9300
   ClientLeft      =   300
   ClientTop       =   570
   ClientWidth     =   8685
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   620
   ScaleMode       =   0  'User
   ScaleWidth      =   579
   Tag             =   "Componente Hilado"
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
      Height          =   5160
      Left            =   120
      TabIndex        =   18
      Tag             =   "Detail"
      Top             =   3285
      Width           =   8445
      Begin VB.TextBox txt_son 
         Height          =   285
         Left            =   2925
         MaxLength       =   1
         TabIndex        =   47
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox txtMerma_Local0 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   25
         Text            =   "0"
         Top             =   4320
         Width           =   1155
      End
      Begin VB.TextBox txtMerma_Importada0 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   28
         Text            =   "0"
         Top             =   4680
         Width           =   1155
      End
      Begin VB.TextBox txtMerma_Importada1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3225
         TabIndex        =   29
         Text            =   "0"
         Top             =   4680
         Width           =   1155
      End
      Begin VB.TextBox txtMerma_Importada2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4965
         TabIndex        =   30
         Text            =   "0"
         Top             =   4680
         Width           =   1155
      End
      Begin VB.TextBox txtMerma_Local1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3225
         TabIndex        =   26
         Text            =   "0"
         Top             =   4320
         Width           =   1155
      End
      Begin VB.TextBox txtMerma_Local2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4965
         TabIndex        =   27
         Text            =   "0"
         Top             =   4320
         Width           =   1155
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Left            =   120
         TabIndex        =   39
         Top             =   3870
         Width           =   6255
      End
      Begin VB.TextBox txtDes_ctacont 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2910
         TabIndex        =   38
         Top             =   1095
         Width           =   3225
      End
      Begin VB.TextBox txtCod_CtaConCIF 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   1875
         Width           =   1305
      End
      Begin VB.TextBox txtDes_CtaConCIF 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2925
         TabIndex        =   33
         Top             =   1860
         Width           =   3225
      End
      Begin VB.TextBox txtCod_CtaConDerechosAduaneros 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1545
         TabIndex        =   8
         Top             =   2250
         Width           =   1305
      End
      Begin VB.TextBox txtDes_CtaConDerechosAduaneros 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2925
         TabIndex        =   32
         Top             =   2250
         Width           =   3210
      End
      Begin VB.TextBox txtCod_CtaConGastosDespacho 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1545
         TabIndex        =   9
         Top             =   2640
         Width           =   1305
      End
      Begin VB.TextBox txtDes_CtaConGastosDespacho 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2925
         TabIndex        =   31
         Top             =   2640
         Width           =   3210
      End
      Begin VB.TextBox txtCod_CtaConAbono 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1545
         TabIndex        =   10
         Top             =   3030
         Width           =   1305
      End
      Begin VB.TextBox txtDes_CtaConAbono 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2925
         TabIndex        =   24
         Top             =   3030
         Width           =   3210
      End
      Begin VB.ComboBox cboCod_TipFamItem 
         Height          =   315
         Left            =   1545
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   2115
      End
      Begin VB.TextBox txtPor_Mermacnf 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4305
         TabIndex        =   6
         Text            =   "0"
         Top             =   1440
         Width           =   1155
      End
      Begin VB.TextBox txtDes_FamItem 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2325
         MaxLength       =   50
         TabIndex        =   2
         Top             =   360
         Width           =   3360
      End
      Begin VB.TextBox txtPor_MermaLog 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1545
         TabIndex        =   5
         Text            =   "0"
         Top             =   1455
         Width           =   1155
      End
      Begin VB.TextBox txtcod_ctacont 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1545
         MaxLength       =   14
         TabIndex        =   4
         Top             =   1095
         Width           =   1290
      End
      Begin VB.TextBox txtCod_FamItem 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1545
         MaxLength       =   2
         TabIndex        =   1
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lbl_son 
         Caption         =   "Se Incluye en Consolidados Contables"
         Height          =   495
         Left            =   120
         TabIndex        =   46
         Top             =   3405
         Width           =   2655
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "<= 10 Mil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1575
         TabIndex        =   45
         Top             =   4080
         Width           =   795
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Exportacion"
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
         Left            =   240
         TabIndex        =   44
         Tag             =   "Porcentaje :"
         Top             =   4725
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Local"
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
         Left            =   240
         TabIndex        =   43
         Tag             =   "Porcentaje :"
         Top             =   4365
         Width           =   390
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   ">=  20 Mil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5040
         TabIndex        =   42
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Mermas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   41
         Tag             =   "Porcentaje :"
         Top             =   4080
         Width           =   690
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   ">= 10 Mil y < 20 Mil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3000
         TabIndex        =   40
         Top             =   4080
         Width           =   1680
      End
      Begin VB.Label Label7 
         Caption         =   "Cta Cont. CIF"
         Height          =   300
         Left            =   90
         TabIndex        =   37
         Top             =   1890
         Width           =   1260
      End
      Begin VB.Label Label4 
         Caption         =   "Cta Cont. Derechos Aduaneros"
         Height          =   420
         Left            =   90
         TabIndex        =   36
         Top             =   2220
         Width           =   1410
      End
      Begin VB.Label Label5 
         Caption         =   "Cta Cont. Gastos Despacho"
         Height          =   420
         Left            =   90
         TabIndex        =   35
         Top             =   2640
         Width           =   1230
      End
      Begin VB.Label Label6 
         Caption         =   "Cta Cont. Abono"
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   3120
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Familia"
         Height          =   315
         Left            =   90
         TabIndex        =   23
         Top             =   795
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "% Merma Conf :"
         Height          =   195
         Left            =   3105
         TabIndex        =   22
         Top             =   1515
         Width           =   1110
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Cta Contable:"
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
         Left            =   90
         TabIndex        =   21
         Tag             =   "Mat. Prima :"
         Top             =   1170
         Width           =   960
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Familia :"
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
         Index           =   0
         Left            =   150
         TabIndex        =   20
         Tag             =   "Hilado :"
         Top             =   435
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "% Merma Log:"
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
         Left            =   90
         TabIndex        =   19
         Tag             =   "Porcentaje :"
         Top             =   1515
         Width           =   1035
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
      Height          =   3135
      Left            =   120
      TabIndex        =   14
      Tag             =   "List"
      Top             =   75
      Width           =   8565
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2775
         Left            =   180
         TabIndex        =   16
         Top             =   345
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   4895
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "cod_famitem"
            Caption         =   "Código"
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
            DataField       =   "des_famitem"
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
         BeginProperty Column02 
            DataField       =   "cod_ctacont"
            Caption         =   "Cta.Contable"
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
            DataField       =   "por_mermalog"
            Caption         =   "Merma Log."
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
            DataField       =   "por_mermacnf"
            Caption         =   "Merma Cnf"
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
            BeginProperty Column00 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3720.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1065.26
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   495
      TabIndex        =   0
      Top             =   8610
      Width           =   1965
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "frmMantFamItem.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Anterior"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "frmMantFamItem.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Siguiente"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   0
         Picture         =   "frmMantFamItem.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Primero"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "frmMantFamItem.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Ultimo"
         Top             =   0
         Width           =   495
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2505
      TabIndex        =   11
      Top             =   8550
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantFamItem.frx":05C8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmMantFamItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public Codigo, Descripcion As String
Dim sTipo As String
Dim StrSQL As String
Dim Rs_Carga As New ADODB.Recordset
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
    End If
End Sub
Private Sub cmdPrevious_Click()
    If Not Rs_Carga.BOF Then
        Rs_Carga.MovePrevious
    End If
End Sub
Private Sub Cargar_Datos()
    On Error GoTo Cargar_DatosErr
    StrSQL = "EXEC UP_SEL_FAMITE"
    Set Rs_Carga = Nothing
    Rs_Carga.ActiveConnection = cCONNECT
    Rs_Carga.CursorType = adOpenStatic
    Rs_Carga.CursorLocation = adUseClient
    Rs_Carga.LockType = adLockReadOnly
    Rs_Carga.Open StrSQL
    Set DGridLista.DataSource = Rs_Carga
    DGridLista_RowColChange 0, 0
    If Rs_Carga.RecordCount > 0 Then
        'HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
    Else
        LIMPIAR_DATOS
        DESHABILITA_DATOS
        'HabilitaMant Me.MantFunc1, "ADICIONAR"
    End If
    Exit Sub
Cargar_DatosErr:
    Set Rs_Carga = Nothing
    ErrorHandler Err, "Cargar_Datos"
End Sub

Private Sub Form_Load()
Dim sql As String
    Call FormSet(Me)
    
    sql = "SELECT des_tipfam + SPACE(100) + cod_tipfam FROM lg_tipfam"
    Call LlenaCombo(cboCod_TipFamItem, sql, cCONNECT)
    
    FormateaGrid Me.DGridLista
    Call Cargar_Datos
    MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub

Private Sub SALVAR_DATOS()
 Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Dim StrSQL As String
    
    Con.ConnectionString = cCONNECT
    Con.Open
    
        Con.BeginTrans

        StrSQL = "EXEC UP_MAN_FAMITE '" & _
        sTipo & "','" & _
        txtCod_FamItem.Text & "','" & _
        txtdes_famitem.Text & "','" & _
        txtCod_CtaCont.Text & "'," & _
        txtPor_MermaLog.Text & "," & _
        txtPor_Mermacnf.Text & ",'" & _
        Right(cboCod_TipFamItem.Text, 1) & "','" & _
        txtCod_CtaConCIF & "','" & _
        txtCod_CtaConDerechosAduaneros & "','" & _
        txtCod_CtaConGastosDespacho & "','" & _
        txtCod_CtaConAbono & "'," & _
        txtMerma_Local1 & "," & _
        txtMerma_Local2 & "," & _
        txtMerma_Importada1 & "," & _
        txtMerma_Importada2 & "," & _
        txtMerma_Local0 & "," & _
        txtMerma_Importada0 & ",'" & txt_son & "'"

        
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
Private Sub ELIMINAR_DATOS()
 Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
   
    Con.ConnectionString = cCONNECT
    Con.Open
    Con.BeginTrans
       
        StrSQL = "EXEC UP_MAN_FAMITE '" & _
        sTipo & "','" & _
        txtCod_FamItem.Text & "','" & _
        txtdes_famitem.Text & "','" & _
        txtCod_CtaCont.Text & "'," & _
        txtPor_MermaLog.Text & "," & _
        txtPor_Mermacnf.Text & ",'','" & _
        txtCod_CtaConCIF & "','" & _
        txtCod_CtaConDerechosAduaneros & "','" & _
        txtCod_CtaConGastosDespacho & "','" & _
        txtCod_CtaConAbono & "'," & _
        txtMerma_Local1 & "," & _
        txtMerma_Local2 & "," & _
        txtMerma_Importada1 & "," & _
        txtMerma_Importada2
        
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
Private Sub LIMPIAR_DATOS()
Dim sql As String
    txtCod_FamItem.Text = ""
    txtdes_famitem.Text = ""
    txtCod_CtaCont.Text = ""
    txtDes_ctacont.Text = ""
    txtPor_MermaLog.Text = "0"
    txtPor_Mermacnf.Text = "0"
    sql = "Select cod_tipfam FROM lg_tipfam where flg_default='*'"
    Call BuscaCombo(DevuelveCampo(sql, cCONNECT), 2, cboCod_TipFamItem)
    txtCod_CtaConCIF.Text = ""
    txtDes_CtaConCIF.Text = ""
    txtCod_CtaConDerechosAduaneros.Text = ""
    txtDes_CtaConDerechosAduaneros.Text = ""
    txtCod_CtaConGastosDespacho.Text = ""
    txtDes_CtaConGastosDespacho.Text = ""
    txtCod_CtaConAbono.Text = ""
    txtDes_CtaConAbono.Text = ""
    txtMerma_Importada0.Text = "0"
    txtMerma_Importada1.Text = "0"
    txtMerma_Importada2.Text = "0"
    txtMerma_Local0.Text = "0"
    txtMerma_Local1.Text = "0"
    txtMerma_Local2.Text = "0"
End Sub
Private Sub DGridLista_Click()
    If Rs_Carga.State <> 1 Then
        Exit Sub
    End If
    If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
        txtCod_FamItem.Text = Trim(Rs_Carga("Cod_FamItem").Value)
        txtdes_famitem.Text = Trim(Rs_Carga("Des_FamItem").Value)
        txtCod_CtaCont.Text = Trim(Rs_Carga("Cod_CtaCont").Value)
        txtDes_ctacont.Text = Trim(Rs_Carga("Des_CtaCont").Value)
        txtPor_MermaLog.Text = Trim(Rs_Carga("Por_MermaLog").Value)
        txtPor_Mermacnf.Text = Trim(Rs_Carga("Por_Mermacnf").Value)
        Call BuscaCombo(Trim(Rs_Carga("cod_tipfam").Value), 2, cboCod_TipFamItem)
        txtCod_CtaConCIF = FixNulos(Trim(Rs_Carga("Cod_CtaConCIF").Value), vbString)
        txtDes_CtaConCIF = FixNulos(Trim(Rs_Carga("DES_CtaConCIF").Value), vbString)
        txtCod_CtaConDerechosAduaneros = FixNulos(Trim(Rs_Carga("Cod_CtaConDerechosAduaneros").Value), vbString)
        txtDes_CtaConDerechosAduaneros = FixNulos(Trim(Rs_Carga("DES_CtaConDerechosAduaneros").Value), vbString)
        txtCod_CtaConGastosDespacho = FixNulos(Trim(Rs_Carga("Cod_CtaConGastosDespacho").Value), vbString)
        txtDes_CtaConGastosDespacho = FixNulos(Trim(Rs_Carga("DES_CtaConGastosDespacho").Value), vbString)
        txtCod_CtaConAbono = FixNulos(Trim(Rs_Carga("Cod_CtaConAbono").Value), vbString)
        txtDes_CtaConAbono = FixNulos(Trim(Rs_Carga("DES_CtaConAbono").Value), vbString)
        txtMerma_Local0 = FixNulos(Trim(Rs_Carga("Por_Merma_Local_0").Value), vbString)
        txtMerma_Local1 = FixNulos(Trim(Rs_Carga("Por_Merma_Local_1").Value), vbString)
        txtMerma_Local2 = FixNulos(Trim(Rs_Carga("Por_Merma_Local_2").Value), vbString)
        txtMerma_Importada0 = FixNulos(Trim(Rs_Carga("Por_Merma_Importado_0").Value), vbString)
        txtMerma_Importada1 = FixNulos(Trim(Rs_Carga("Por_Merma_Importado_1").Value), vbString)
        txtMerma_Importada2 = FixNulos(Trim(Rs_Carga("Por_Merma_Importado_2").Value), vbString)
        txt_son.Text = Trim(Rs_Carga("flg_incluir_consolidados").Value)
        DESHABILITA_DATOS
    End If
End Sub
Private Sub HABILITA_DATOS()
    txtCod_FamItem.Enabled = True
    txtdes_famitem.Enabled = True
    txtCod_CtaCont.Enabled = True
    txtDes_ctacont.Enabled = True
    txtPor_MermaLog.Enabled = True
    txtPor_Mermacnf.Enabled = True
    cboCod_TipFamItem.Enabled = True
    txtCod_CtaConCIF.Enabled = True
    txtDes_CtaConCIF.Enabled = True
    txtCod_CtaConDerechosAduaneros.Enabled = True
    txtDes_CtaConDerechosAduaneros.Enabled = True
    txtCod_CtaConGastosDespacho.Enabled = True
    txtDes_CtaConGastosDespacho.Enabled = True
    txtCod_CtaConAbono.Enabled = True
    txtDes_CtaConAbono.Enabled = True
    txtMerma_Local0.Enabled = True
    txtMerma_Local1.Enabled = True
    txtMerma_Local2.Enabled = True
    txtMerma_Importada0.Enabled = True
    txtMerma_Importada1.Enabled = True
    txtMerma_Importada2.Enabled = True
    txt_son.Enabled = True
End Sub
Private Sub DESHABILITA_DATOS()
    txtCod_FamItem.Enabled = False
    txtdes_famitem.Enabled = False
    txtCod_CtaCont.Enabled = False
    txtDes_ctacont.Enabled = False
    txtPor_MermaLog.Enabled = False
    txtPor_Mermacnf.Enabled = False
    cboCod_TipFamItem.Enabled = False
    txtCod_CtaConCIF.Enabled = False
    txtDes_CtaConCIF.Enabled = False
    txtCod_CtaConDerechosAduaneros.Enabled = False
    txtDes_CtaConDerechosAduaneros.Enabled = False
    txtCod_CtaConGastosDespacho.Enabled = False
    txtDes_CtaConGastosDespacho.Enabled = False
    txtCod_CtaConAbono.Enabled = False
    txtDes_CtaConAbono.Enabled = False
    txtMerma_Local0.Enabled = False
    txtMerma_Local1.Enabled = False
    txtMerma_Local2.Enabled = False
    txtMerma_Importada0.Enabled = False
    txtMerma_Importada1.Enabled = False
    txtMerma_Importada2.Enabled = False
    txt_son.Enabled = False
End Sub
Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub
Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Rs_Carga.State <> 1 Then
        Exit Sub
    End If
    If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
        txtCod_FamItem.Text = Trim(Rs_Carga("Cod_FamItem").Value)
        txtdes_famitem.Text = Trim(Rs_Carga("Des_FamItem").Value)
        txtCod_CtaCont.Text = Trim(Rs_Carga("Cod_CtaCont").Value)
        txtDes_ctacont.Text = Trim(Rs_Carga("Des_CtaCont").Value)
        txtPor_MermaLog.Text = Trim(Rs_Carga("Por_MermaLog").Value)
        txtPor_Mermacnf.Text = Trim(Rs_Carga("Por_Mermacnf").Value)
        Call BuscaCombo(Trim(Rs_Carga("cod_tipfam").Value), 2, cboCod_TipFamItem)
        txtCod_CtaConCIF = FixNulos(Trim(Rs_Carga("Cod_CtaConCIF").Value), vbString)
        txtDes_CtaConCIF = FixNulos(Trim(Rs_Carga("DES_CtaConCIF").Value), vbString)
        txtCod_CtaConDerechosAduaneros = FixNulos(Trim(Rs_Carga("Cod_CtaConDerechosAduaneros").Value), vbString)
        txtDes_CtaConDerechosAduaneros = FixNulos(Trim(Rs_Carga("DES_CtaConDerechosAduaneros").Value), vbString)
        txtCod_CtaConGastosDespacho = FixNulos(Trim(Rs_Carga("Cod_CtaConGastosDespacho").Value), vbString)
        txtDes_CtaConGastosDespacho = FixNulos(Trim(Rs_Carga("DES_CtaConGastosDespacho").Value), vbString)
        txtCod_CtaConAbono = FixNulos(Trim(Rs_Carga("Cod_CtaConAbono").Value), vbString)
        txtDes_CtaConAbono = FixNulos(Trim(Rs_Carga("DES_CtaConAbono").Value), vbString)
        txtMerma_Local0 = FixNulos(Trim(Rs_Carga("Por_Merma_Local_0").Value), vbString)
        txtMerma_Local1 = FixNulos(Trim(Rs_Carga("Por_Merma_Local_1").Value), vbString)
        txtMerma_Local2 = FixNulos(Trim(Rs_Carga("Por_Merma_Local_2").Value), vbString)
        txtMerma_Importada0 = FixNulos(Trim(Rs_Carga("Por_Merma_Importado_0").Value), vbString)
        txtMerma_Importada1 = FixNulos(Trim(Rs_Carga("Por_Merma_Importado_1").Value), vbString)
        txtMerma_Importada2 = FixNulos(Trim(Rs_Carga("Por_Merma_Importado_2").Value), vbString)
        txt_son = FixNulos(Trim(Rs_Carga("flg_incluir_consolidados").Value), vbString)
        DESHABILITA_DATOS
    End If
End Sub
Private Sub RECARGAR_DATOS()
    Rs_Carga.Close
    Cargar_Datos
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Rs_Carga = Nothing
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Eliminar As Integer
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            LIMPIAR_DATOS
            HABILITA_DATOS
            txtCod_FamItem.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "MODIFICAR"
            sTipo = "U"
            HABILITA_DATOS
            txtCod_FamItem.Enabled = False
            txtdes_famitem.Enabled = False
            cboCod_TipFamItem.Enabled = False
            txtCod_CtaCont.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "ELIMINAR"
            sTipo = "D"
            If VALIDA_DATOS Then
                Eliminar = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Familia")
                If Eliminar = vbYes Then
                    ELIMINAR_DATOS
                    RECARGAR_DATOS
                    sTipo = ""
                End If
            End If
        Case "GRABAR"
            If VALIDA_DATOS Then
                SALVAR_DATOS
                RECARGAR_DATOS
                'HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
                DGridLista.Enabled = True
                'cmdBuscaMatPri.Enabled = False
                sTipo = ""
            End If
        Case "DESHACER"
            LIMPIAR_DATOS
            RECARGAR_DATOS
            'HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
            DGridLista.Enabled = True
            sTipo = ""
        Case "SALIR"
            Unload Me
    End Select
End Sub

Function VALIDA_DATOS() As Boolean
    Dim aMess(4)
    Dim amensaje As clsMessages
    Set amensaje = New clsMessages
    VALIDA_DATOS = True
    
    If sTipo = "I" Then
        If Trim(txtCod_FamItem.Text) = "" Then
            VALIDA_DATOS = False
            MsgBox ("Sirvase ingresar el código de familia")
            txtCod_FamItem.SetFocus
        Else
            StrSQL = "SELECT cod_famitem FROM LG_FAMITE WHERE cod_famitem ='" & Trim(txtCod_FamItem.Text) & "'"
            If DevuelveCampo(StrSQL, cCONNECT) <> "" Then
                VALIDA_DATOS = False
                MsgBox ("El código ingresado ya se encuentra registrado. Sirvase ingresar otro")
                txtCod_FamItem.SetFocus
            End If
        End If
    End If
    
    If sTipo = "D" Then
        StrSQL = "SELECT * FROM LG_ITEM WHERE Cod_FamItem ='" & txtCod_FamItem.Text & "'"
        If DevuelveCampo(StrSQL, cCONNECT) <> "" Then
            VALIDA_DATOS = False
            MsgBox ("No se puede eliminar la familia, por que tiene items relacionados")
        End If
    End If
    
    If cboCod_TipFamItem.ListIndex = -1 Then
        VALIDA_DATOS = False
        MsgBox ("No ha seleccionado la familia de item")
        cboCod_TipFamItem.SetFocus
    End If
End Function

Private Sub txtCod_CtaConAbono_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And RTrim(ActiveControl) <> "" Then
        If RTrim(txtCod_CtaConAbono.Text) = "" Then
            BUSCA_CUENTACONTABLE 3, txtCod_CtaConAbono, txtDes_CtaConAbono
        Else
            BUSCA_CUENTACONTABLE 1, txtCod_CtaConAbono, txtDes_CtaConAbono
        End If
    
    ElseIf RTrim(ActiveControl) = "" And KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtCod_CtaConCIF_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And RTrim(ActiveControl) <> "" Then
        If RTrim(txtCod_CtaConCIF.Text) = "" Then
            BUSCA_CUENTACONTABLE 3, txtCod_CtaConCIF, txtDes_CtaConCIF
        Else
            BUSCA_CUENTACONTABLE 1, txtCod_CtaConCIF, txtDes_CtaConCIF
        End If
    
    ElseIf RTrim(ActiveControl) = "" And KeyAscii = vbKeyReturn Then
        txt_son.SetFocus
    End If

End Sub

Private Sub txtCod_CtaConDerechosAduaneros_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And RTrim(ActiveControl) <> "" Then
        If RTrim(txtCod_CtaConDerechosAduaneros.Text) = "" Then
            BUSCA_CUENTACONTABLE 3, txtCod_CtaConDerechosAduaneros, txtDes_CtaConDerechosAduaneros
        Else
            BUSCA_CUENTACONTABLE 1, txtCod_CtaConDerechosAduaneros, txtDes_CtaConDerechosAduaneros
        End If
    
    ElseIf RTrim(ActiveControl) = "" And KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtCod_CtaConGastosDespacho_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And RTrim(ActiveControl) <> "" Then
        If RTrim(txtCod_CtaConGastosDespacho.Text) = "" Then
            BUSCA_CUENTACONTABLE 3, txtCod_CtaConGastosDespacho, txtDes_CtaConGastosDespacho
        Else
            BUSCA_CUENTACONTABLE 1, txtCod_CtaConGastosDespacho, txtDes_CtaConGastosDespacho
        End If
    
    ElseIf RTrim(ActiveControl) = "" And KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtcod_ctacont_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub

Private Sub txtcod_ctacont_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And RTrim(ActiveControl) <> "" Then
        If RTrim(txtCod_CtaCont.Text) = "" Then
            BUSCA_CUENTACONTABLE 3, txtCod_CtaCont, txtDes_ctacont
        Else
            BUSCA_CUENTACONTABLE 1, txtCod_CtaCont, txtDes_ctacont
        End If
    
    ElseIf RTrim(ActiveControl) = "" And KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDes_CtaConAbono_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        BUSCA_CUENTACONTABLE 2, txtCod_CtaConAbono, txtDes_CtaConAbono
    End If
End Sub

Private Sub txtDes_CtaConCIF_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        BUSCA_CUENTACONTABLE 2, txtCod_CtaConCIF, txtDes_CtaConCIF
    End If
End Sub

Private Sub txtDes_CtaConDerechosAduaneros_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        BUSCA_CUENTACONTABLE 2, txtCod_CtaConDerechosAduaneros, txtDes_CtaConDerechosAduaneros
    End If
End Sub

Private Sub txtDes_CtaConGastosDespacho_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        BUSCA_CUENTACONTABLE 2, txtCod_CtaConGastosDespacho, txtDes_CtaConGastosDespacho
    End If
End Sub

Private Sub txtDes_ctacont_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        BUSCA_CUENTACONTABLE 2, txtCod_CtaCont, txtDes_ctacont
    End If
End Sub

Private Sub txtult_numgen_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub

Private Sub txtMerma_Importada0_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub

Private Sub txtMerma_Importada1_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub

Private Sub txtMerma_Importada1_KeyPress(KeyAscii As Integer)
    SoloNumeros txtMerma_Importada1, KeyAscii, True, 2, 5
End Sub

Private Sub txtMerma_Importada1_LostFocus()
    If Trim(txtMerma_Importada1.Text) = "" Then
        txtMerma_Importada1.Text = "0"
    End If
End Sub

Private Sub txtMerma_Importada2_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub

Private Sub txtMerma_Importada2_KeyPress(KeyAscii As Integer)
    SoloNumeros txtMerma_Importada2, KeyAscii, True, 2, 5
End Sub

Private Sub txtMerma_Importada2_LostFocus()
    If Trim(txtMerma_Importada2.Text) = "" Then
        txtMerma_Importada2.Text = "0"
    End If
End Sub

Private Sub txtMerma_Local0_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub

Private Sub txtMerma_Local1_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub

Private Sub txtMerma_Local1_KeyPress(KeyAscii As Integer)
    SoloNumeros txtMerma_Local1, KeyAscii, True, 2, 5
End Sub


Private Sub txtMerma_Local1_LostFocus()
    If Trim(txtMerma_Local1.Text) = "" Then
        txtMerma_Local1.Text = "0"
    End If
End Sub

Private Sub txtMerma_Local2_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub

Private Sub txtMerma_Local2_KeyPress(KeyAscii As Integer)
    SoloNumeros txtMerma_Local2, KeyAscii, True, 2, 5
End Sub

Private Sub txtMerma_Local2_LostFocus()
    If Trim(txtMerma_Local2.Text) = "" Then
        txtMerma_Local2.Text = "0"
    End If
End Sub

Private Sub txtPor_Mermacnf_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub

Private Sub txtPor_Mermacnf_KeyPress(KeyAscii As Integer)
    SoloNumeros txtPor_Mermacnf, KeyAscii, True, 2, 5
End Sub

Private Sub txtPor_Mermacnf_LostFocus()
    If Trim(txtPor_Mermacnf.Text) = "" Then
        txtPor_Mermacnf.Text = "0"
    End If
End Sub

Private Sub txtPor_MermaLog_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub

Private Sub txtPor_MermaLog_KeyPress(KeyAscii As Integer)
    SoloNumeros txtPor_MermaLog, KeyAscii, True, 2, 5
End Sub


Private Sub txtPor_MermaLog_LostFocus()
    If Trim(txtPor_MermaLog.Text) = "" Then
        txtPor_MermaLog.Text = "0"
    End If
End Sub

Private Sub BUSCA_CUENTACONTABLE(tipo As Integer, ByRef Cod_Cta As TextBox, ByRef Des_Cta As TextBox)
On Error GoTo errx
Dim StrSQL  As String

    Select Case tipo
    Case 1:
        StrSQL = "SELECT Cuenta as 'Código', Descripcion as 'Descripción' " & _
                 "FROM CN_PLAN WHERE CUENTA like '" & Trim(Cod_Cta) & "%' ORDER BY Cuenta "
                
    Case 2, 3:
            StrSQL = "SELECT Cuenta AS 'Código', " & _
            " Descripcion as 'Descripción' " & _
            "FROM CN_PLAN " & _
            "WHERE DESCRIPCION LIKE '%" & Trim(Des_Cta) _
            & "%' AND DATALENGTH(RTRIM(CUENTA)) = 8 ORDER BY 2"
    End Select
    
    With frmBusqGeneral3
        .Caption = "Buscar Cuenta"
        .sQuery = StrSQL
        Set .oParent = Me
        .Cargar_Datos
        
        
        .DGridLista.Columns("Código").Caption = "Código"
        .DGridLista.Columns("Descripción").Caption = "Desc. Cuenta"
        
        .DGridLista.Columns("Descripción").Width = 4800
                
        If .DGridLista.RowCount > 1 Then
            .Show vbModal
        Else
            .bCancel = False
            Codigo = .DGridLista.Value(.DGridLista.Columns("Código").Index)
            Descripcion = .DGridLista.Value(.DGridLista.Columns("Descripción").Index)
        End If
            
        
        If Not .bCancel Then
            Cod_Cta = Codigo
            Des_Cta = Descripcion
            Codigo = "": Descripcion = ""
            SendKeys "{TAB}"
        End If
    End With
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    Exit Sub
    
errx:
    errores Err.Number
End Sub





