VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBuscaGrupo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busca Grupo"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   510
      Left            =   3795
      TabIndex        =   6
      Top             =   3120
      Width           =   1395
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   510
      Left            =   2205
      TabIndex        =   5
      Top             =   3120
      Width           =   1395
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
      Height          =   2310
      Left            =   30
      TabIndex        =   3
      Top             =   735
      Width           =   5265
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   1920
         Left            =   90
         TabIndex        =   4
         Top             =   240
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   3387
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Cod_Grupo"
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
            DataField       =   "Des_Grupo"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1379.906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3209.953
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraBuscar 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5280
      Begin VB.TextBox txtCod_GrupoTex 
         Height          =   315
         Left            =   1005
         TabIndex        =   2
         Top             =   240
         Width           =   2325
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Grupo"
         Height          =   195
         Left            =   300
         TabIndex        =   1
         Top             =   300
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmBuscaGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Strsql As String
Dim Rs_Lista As ADODB.Recordset
Public oParent As Object
Public varTipo As String

Sub CARGA_GRID()
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.ActiveConnection = cConnect
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    If varTipo = "1" Then
        'Esta cadena es para devolver el Codigo de Cliente
        Strsql = "SELECT Cod_GrupoTex as Cod_Grupo, Des_Grupo FROM ES_GRUPOTEX WHERE Cod_GrupoTex LIKE '" & Trim(Me.txtCod_GrupoTex.Text) & "%'"
    Else
        'Esta cadena es para devolver el Codigo de Cliente
        Strsql = "SELECT Cod_GrupoLog as Cod_Grupo, Des_Grupo FROM ES_GRUPOLOG WHERE Cod_GrupoLog LIKE '" & Trim(Me.txtCod_GrupoTex.Text) & "%'"
    End If
    
    Rs_Lista.Open Strsql
    Set DGridLista.DataSource = Rs_Lista

    Me.txtCod_GrupoTex.SelStart = Len(Me.txtCod_GrupoTex.Text)

End Sub

Private Sub cmdAceptar_Click()
    If varTipo = "1" Then
        If Rs_Lista.RecordCount > 0 Then
            oParent.Codigo = Rs_Lista("Cod_Grupo").Value
            oParent.Descripcion = Rs_Lista("Des_Grupo").Value
        Else
            oParent.Codigo = ""
            oParent.Descripcion = ""
        End If
    Else
        If Rs_Lista.RecordCount > 0 Then
            oParent.Codigo = Rs_Lista("Cod_Grupo").Value
            oParent.Descripcion = Rs_Lista("Des_Grupo").Value
        Else
            oParent.Codigo = ""
            oParent.Descripcion = ""
        End If
    End If
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub DGridLista_DblClick()
    Call cmdAceptar_Click
End Sub

Private Sub DGridLista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdAceptar_Click
End Sub

Private Sub Form_Load()
    Call FormateaGrid(DGridLista)
End Sub

Private Sub txtCod_GrupoTex_Change()
    Call CARGA_GRID
End Sub

Private Sub txtCod_GrupoTex_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAceptar_Click
    End If
End Sub
