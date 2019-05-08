VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBusqGeneralGrupo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2265
      TabIndex        =   1
      Tag             =   "&OK"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3510
      TabIndex        =   0
      Tag             =   "&Cancel"
      Top             =   3600
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DGridLista 
      Height          =   3180
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   5609
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   17
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
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBusqGeneralGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oParent As Object
Public sQuery As String
Dim Rs_Carga As ADODB.Recordset
Sub Cargar_Datos()
Set Rs_Carga = New ADODB.Recordset
On Error GoTo Cargar_DatosErr
Rs_Carga.ActiveConnection = cCONNECT
Rs_Carga.CursorType = adOpenStatic
Rs_Carga.CursorLocation = adUseClient
Rs_Carga.LockType = adLockReadOnly
Rs_Carga.Open sQuery
Set DGridLista.DataSource = Rs_Carga
Exit Sub
Cargar_DatosErr:
    Set Rs_Carga = Nothing
    ErrorHandler Err, "Cargar_Datos"
End Sub
Private Sub DGridlista_DblClick()
If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
    With oParent
        .Codigo = Rs_Carga(0)
        .txtCod_OrdPro = Rs_Carga(0)
        If Rs_Carga.Fields.Count > 1 Then
            .Descripcion = Rs_Carga(1)
            .txtDes_estpro = Rs_Carga(1)
        End If
    End With
    
    If oParent.VALIDA_ANADE_ORDPRO Then
        oParent.ANADE_ORDPRO
        oParent.CARGA_ORDPRO
        'sTipo = ""
        'fraOrdPro.Visible = False
    End If
    Me.Cargar_Datos
End If
'Unload Me
End Sub

Private Sub DGridLista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar_Click
    End If
End Sub

Private Sub DGridlista_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If DGridLista.RowContaining(y) >= 1 And DGridLista.RowContaining(y) <= Rs_Carga.RecordCount Then
    DGridLista.Bookmark = DGridLista.RowBookmark(DGridLista.RowContaining(y))
End If
End Sub
Private Sub Form_Load()
Call FormSet(Me)
FormateaGrid DGridLista
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Rs_Carga = Nothing
End Sub
Private Sub cmdAceptar_Click()
    DGridlista_DblClick
End Sub
Private Sub cmdCancelar_Click()
oParent.Codigo = ""
oParent.txtCod_OrdPro = ""
oParent.txtDes_estpro = ""

Unload Me
End Sub


