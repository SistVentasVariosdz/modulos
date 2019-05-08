VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmListaGeneral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Find"
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
      Left            =   3915
      TabIndex        =   2
      Tag             =   "&Cancel"
      Top             =   4515
      Width           =   1185
   End
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
      Left            =   2655
      TabIndex        =   1
      Tag             =   "&OK"
      Top             =   4500
      Width           =   1185
   End
   Begin MSDataGridLib.DataGrid dGridLista 
      Height          =   4170
      Left            =   135
      TabIndex        =   0
      Top             =   195
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   7355
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
Attribute VB_Name = "frmListaGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public SQUERY As String
Dim Rs_Carga As New ADODB.RecordSet
Sub CARGAR_DATOS()
On Error GoTo Cargar_DatosErr
Rs_Carga.ActiveConnection = cCONNECT
Rs_Carga.CursorType = adOpenStatic
Rs_Carga.CursorLocation = adUseClient
Rs_Carga.LockType = adLockReadOnly
Rs_Carga.Open SQUERY
Set dGridLista.DataSource = Rs_Carga
Exit Sub
Cargar_DatosErr:
    Set Rs_Carga = Nothing
    ErrorHandler Err, "Cargar_Datos"
End Sub
Private Sub cmdAceptar_Click()
    DGridlista_DblClick
End Sub
Private Sub cmdCancelar_Click()
    oParent.Codigo = ""
    Unload Me
End Sub
Private Sub DGridlista_DblClick()
If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
    With oParent
    .Codigo = Rs_Carga!Codigo
    .Descripcion = Rs_Carga!Descripcion
    End With
End If
Unload Me
End Sub
Private Sub DGridlista_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If dGridLista.RowContaining(y) >= 1 And dGridLista.RowContaining(y) <= Rs_Carga.RecordCount Then
    dGridLista.Bookmark = dGridLista.RowBookmark(dGridLista.RowContaining(y))
End If
End Sub
Private Sub Form_Load()
Call FormSet(Me)
FormateaGrid dGridLista
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Rs_Carga.Close
    Set Rs_Carga = Nothing
End Sub
