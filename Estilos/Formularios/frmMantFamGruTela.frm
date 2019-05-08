VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMantFamGruTela 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupos de Tela"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1170
      Left            =   120
      TabIndex        =   11
      Tag             =   "Detail"
      Top             =   3405
      Width           =   5580
      Begin VB.TextBox txtCod_GruTela 
         Height          =   315
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   2
         Top             =   720
         Width           =   795
      End
      Begin VB.TextBox txtDes_GruTela 
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
         Left            =   2100
         MaxLength       =   50
         TabIndex        =   3
         Top             =   720
         Width           =   3360
      End
      Begin VB.TextBox txtCod_FamTela 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
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
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   1
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Grupo :"
         Height          =   195
         Left            =   210
         TabIndex        =   13
         Top             =   820
         Width           =   525
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
         Left            =   210
         TabIndex        =   12
         Tag             =   "Hilado :"
         Top             =   435
         Width           =   570
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
      Height          =   3255
      Left            =   120
      TabIndex        =   7
      Tag             =   "List"
      Top             =   75
      Width           =   5595
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2775
         Left            =   180
         TabIndex        =   9
         Top             =   345
         Width           =   5265
         _ExtentX        =   9287
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "cod_grutela"
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
            DataField       =   "des_grutela"
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
            BeginProperty Column00 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3734.929
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   4740
      Width           =   1965
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "frmMantFamGruTela.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Anterior"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "frmMantFamGruTela.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Siguiente"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   0
         Picture         =   "frmMantFamGruTela.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Primero"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "frmMantFamGruTela.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Ultimo"
         Top             =   0
         Width           =   495
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2160
      TabIndex        =   4
      Top             =   4680
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantFamGruTela.frx":05C8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmMantFamGruTela"
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
Public Sub Cargar_Datos()
    On Error GoTo Cargar_DatosErr
    StrSQL = "EXEC UP_SEL_GRUTELA '" & txtCod_FamTela.Text & "'"
    Set Rs_Carga = Nothing
    Rs_Carga.ActiveConnection = cCONNECT
    Rs_Carga.CursorType = adOpenStatic
    Rs_Carga.CursorLocation = adUseClient
    Rs_Carga.LockType = adLockReadOnly
    Rs_Carga.Open StrSQL
    Set DGridLista.DataSource = Rs_Carga
    DGridLista_RowColChange 0, 0
    If Rs_Carga.RecordCount > 0 Then
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
    Else
        LIMPIAR_DATOS
        DESHABILITA_DATOS
        HabilitaMant Me.MantFunc1, "ADICIONAR"
    End If
    Exit Sub
Cargar_DatosErr:
    Set Rs_Carga = Nothing
    ErrorHandler Err, "Cargar_Datos"
End Sub

Private Sub Form_Load()
    Call FormSet(Me)
    FormateaGrid Me.DGridLista
    'Call Cargar_Datos
    MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub

Private Sub SALVAR_DATOS()
 Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Dim StrSQL As String
    
    Con.ConnectionString = cCONNECT
    Con.Open
    
        Con.BeginTrans

        StrSQL = "EXEC UP_MAN_GRUTELA '" & _
        sTipo & "','" & _
        txtCod_FamTela.Text & "','" & _
        txtCod_GruTela.Text & "','" & _
        txtDes_GruTela.Text & "'"
        
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
       
        StrSQL = "EXEC UP_MAN_GRUTELA '" & _
        sTipo & "','" & _
        txtCod_FamTela.Text & "','" & _
        txtCod_GruTela.Text & "','" & _
        txtDes_GruTela.Text & "'"

        Con.Execute StrSQL
        
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
  
    txtCod_GruTela.Text = ""
    txtDes_GruTela.Text = ""
End Sub
Private Sub DGridLista_Click()
    If Rs_Carga.State <> 1 Then
        Exit Sub
    End If
    If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
        txtCod_FamTela.Text = Trim(Rs_Carga("Cod_FamTela").Value)
        txtCod_GruTela.Text = Trim(Rs_Carga("Cod_GruTela").Value)
        txtDes_GruTela.Text = Trim(Rs_Carga("Des_GruTela").Value)
        DESHABILITA_DATOS
    End If
End Sub
Private Sub HABILITA_DATOS()
    txtCod_GruTela.Enabled = True
    txtDes_GruTela.Enabled = True
End Sub
Private Sub DESHABILITA_DATOS()
    txtCod_FamTela.Enabled = False
    txtCod_GruTela.Enabled = False
    txtDes_GruTela.Enabled = False
End Sub
Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub
Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Rs_Carga.State <> 1 Then
        Exit Sub
    End If
    If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
        txtCod_FamTela.Text = Trim(Rs_Carga("Cod_FamTela").Value)
        txtCod_GruTela.Text = Trim(Rs_Carga("Cod_GruTela").Value)
        txtDes_GruTela.Text = Trim(Rs_Carga("Des_GruTela").Value)
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
            'txtCod_GruTela.Enabled = False
            'txtDes_GruTela.SetFocus
            txtCod_GruTela.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "MODIFICAR"
            sTipo = "U"
            HABILITA_DATOS
            txtCod_FamTela.Enabled = False
            txtCod_GruTela.Enabled = False
            txtDes_GruTela.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "ELIMINAR"
            sTipo = "D"
            If VALIDA_DATOS Then
                Eliminar = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Familia-Grupo")
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

Function VALIDA_DATOS() As Boolean
    Dim aMess(4)
    Dim amensaje As clsMessages
    Set amensaje = New clsMessages
    VALIDA_DATOS = True
    If sTipo = "D" Then
        StrSQL = "SELECT * FROM LG_ITEM WHERE Cod_GruItem ='" & txtCod_GruTela.Text & "'"
        If DevuelveCampo(StrSQL, cCONNECT) <> "" Then
            VALIDA_DATOS = False
            MsgBox ("No se puede eliminar el grupo, por que tiene items relacionados")
        End If
    Else
        If Trim(txtDes_GruTela.Text) = "" Then
            VALIDA_DATOS = False
            Call MsgBox("La descripción del grupo no puede ser vacia. Sirvase verificar", vbInformation)
        End If
    End If
End Function

Private Sub txtcod_ctacont_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub



