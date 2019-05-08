VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmTiposCobranzaAdd 
   Caption         =   "Mantenimiento Tipo Cobranza"
   ClientHeight    =   2580
   ClientLeft      =   2445
   ClientTop       =   1485
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   9795
   Begin VB.Frame frFacturas 
      ForeColor       =   &H8000000D&
      Height          =   7095
      Left            =   120
      TabIndex        =   6
      Top             =   7320
      Visible         =   0   'False
      Width           =   9705
      Begin VB.TextBox TxtMonto2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   7905
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   6645
         Width           =   1455
      End
      Begin VB.TextBox TxtMonto1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   7905
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton CmdNext 
         BackColor       =   &H8000000D&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2910
         TabIndex        =   10
         Top             =   3390
         Width           =   675
      End
      Begin VB.CommandButton CmdNextAll 
         BackColor       =   &H8000000D&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3630
         TabIndex        =   11
         Top             =   3390
         Width           =   675
      End
      Begin VB.CommandButton CmdBack 
         BackColor       =   &H8000000D&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4710
         TabIndex        =   12
         Top             =   3390
         Width           =   675
      End
      Begin VB.CommandButton cmdBackAll 
         BackColor       =   &H8000000D&
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5430
         TabIndex        =   14
         Top             =   3390
         Width           =   675
      End
      Begin VB.TextBox txtNumeroPendiente 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4590
         TabIndex        =   8
         Text            =   "0"
         Top             =   240
         Width           =   1245
      End
      Begin VB.TextBox txtNumeroxCancelar 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   960
         TabIndex        =   9
         Text            =   "0"
         Top             =   3480
         Width           =   1245
      End
      Begin GridEX20.GridEX gexGrid2 
         Height          =   2595
         Left            =   195
         TabIndex        =   13
         Top             =   3960
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   4577
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         BackColorBkg    =   -2147483628
         ColumnHeaderHeight=   285
         IntProp1        =   0
         ColumnsCount    =   3
         Column(1)       =   "frmTiposCobranzaAdd.frx":0000
         Column(2)       =   "frmTiposCobranzaAdd.frx":00F4
         Column(3)       =   "frmTiposCobranzaAdd.frx":01E0
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmTiposCobranzaAdd.frx":02AC
         FormatStyle(2)  =   "frmTiposCobranzaAdd.frx":03E4
         FormatStyle(3)  =   "frmTiposCobranzaAdd.frx":0494
         FormatStyle(4)  =   "frmTiposCobranzaAdd.frx":0548
         FormatStyle(5)  =   "frmTiposCobranzaAdd.frx":0620
         FormatStyle(6)  =   "frmTiposCobranzaAdd.frx":06D8
         ImageCount      =   0
         PrinterProperties=   "frmTiposCobranzaAdd.frx":07B8
      End
      Begin GridEX20.GridEX gexGrid1 
         Height          =   2595
         Left            =   195
         TabIndex        =   24
         Top             =   660
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   4577
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         BackColorBkg    =   -2147483628
         ColumnHeaderHeight=   285
         IntProp1        =   0
         ColumnsCount    =   3
         Column(1)       =   "frmTiposCobranzaAdd.frx":0990
         Column(2)       =   "frmTiposCobranzaAdd.frx":0A84
         Column(3)       =   "frmTiposCobranzaAdd.frx":0B70
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmTiposCobranzaAdd.frx":0C3C
         FormatStyle(2)  =   "frmTiposCobranzaAdd.frx":0D74
         FormatStyle(3)  =   "frmTiposCobranzaAdd.frx":0E24
         FormatStyle(4)  =   "frmTiposCobranzaAdd.frx":0ED8
         FormatStyle(5)  =   "frmTiposCobranzaAdd.frx":0FB0
         FormatStyle(6)  =   "frmTiposCobranzaAdd.frx":1068
         ImageCount      =   0
         PrinterProperties=   "frmTiposCobranzaAdd.frx":1148
      End
      Begin FunctionsButtons.FunctButt fncBuscar 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   165
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   661
         Custom          =   $"frmTiposCobranzaAdd.frx":1320
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   350
         ControlSeparator=   110
      End
      Begin VB.Label Label17 
         Caption         =   "Total Pendiente :"
         Height          =   285
         Left            =   6585
         TabIndex        =   18
         Top             =   3495
         Width           =   1305
      End
      Begin VB.Label Label13 
         Caption         =   "Total a Cancelar :"
         Height          =   255
         Left            =   6225
         TabIndex        =   17
         Top             =   6660
         Width           =   1395
      End
      Begin VB.Label Label14 
         Caption         =   "Numero :"
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         Top             =   270
         Width           =   675
      End
      Begin VB.Label Label15 
         Caption         =   "Numero :"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3510
         Width           =   675
      End
   End
   Begin VB.Frame frTransacciones 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      TabIndex        =   19
      Top             =   -240
      Width           =   9735
      Begin VB.Frame Frame3 
         Height          =   1815
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   9495
         Begin VB.TextBox txtCod_Cobranza1 
            Height          =   285
            Left            =   2640
            MaxLength       =   4
            TabIndex        =   3
            Top             =   1320
            Width           =   600
         End
         Begin VB.TextBox txtDes_Cobranza1 
            Height          =   285
            Left            =   3360
            TabIndex        =   4
            Top             =   1320
            Width           =   5745
         End
         Begin VB.TextBox txtCod_Cobranza 
            Height          =   285
            Left            =   2640
            MaxLength       =   4
            TabIndex        =   1
            Top             =   960
            Width           =   600
         End
         Begin VB.TextBox txtDes_Cobranza 
            Height          =   285
            Left            =   3360
            TabIndex        =   2
            Top             =   960
            Width           =   5745
         End
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Left            =   2640
            TabIndex        =   25
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtDescripcion 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2640
            TabIndex        =   0
            Top             =   600
            Width           =   5265
         End
         Begin VB.Label Label2 
            Caption         =   "Cod. Concepto Cobranza Haber :"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1335
            Width           =   2415
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. Concepto Cobranza Debe :"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   975
            Width           =   2415
         End
         Begin VB.Label Label4 
            Caption         =   "Código :"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   270
            Width           =   645
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Descripción :"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   645
            Width           =   930
         End
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   3720
         TabIndex        =   5
         Top             =   2160
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmTiposCobranzaAdd.frx":13AD
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
   End
End
Attribute VB_Name = "frmTiposCobranzaAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public codigo As String, Descripcion As String, StrOption As String, strCod_Anxo As String, lfSalvar As Boolean
Public sCod_Tipcobranza As String
Dim RsGrid1 As ADODB.Recordset
Dim RsGrid2 As ADODB.Recordset
Dim intTransaccion As Integer, vrTotalTransaccion As Double
Dim strSQL As String, intCancel As Integer

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim sMessage As Long
Dim strSQL As String

On Error GoTo dprError

Select Case ActionName
  Case "GRABAR"
  
  If txtDescripcion.Text = "" Then
    MsgBox "Debe ingresar la Descripción"
    Exit Sub
  End If
  
    If MsgBox("Esta seguro de Generar un Nuevo Tipo de Cobranza ", vbYesNo, "IMPORTANTE") = vbYes Then
      Salvar_Datos
        intCancel = 0
        Unload Me
      End If
    
    
  Case "CANCELAR"
      lfSalvar = False
      intCancel = 0
      Unload Me

End Select

Exit Sub

dprError:

errores err.Number
End Sub

Sub Salvar_Datos()
On Error GoTo ErrSalvarDatos
    strSQL = "exec Ventas_Mantenimiento_TipoCobranza '" & txtcodigo.Text & "','" & txtDescripcion.Text & "','" & StrOption & "','" & txtCod_Cobranza.Text & "','" & txtCod_Cobranza1.Text & "'"
    ExecuteSQL cCONNECT, strSQL
    
        
Exit Sub
ErrSalvarDatos:
    ErrorHandler err, "SALVAR_DATOS"
End Sub


Private Sub txtCod_Cobranza_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion3("Cod_Concepto_Cobranza", "Descripcion", "Cn_Ventas_Conceptos_Cobranza Where ", txtCod_Cobranza, txtDes_Cobranza, 1, Me)
    txtCod_Cobranza1.SetFocus
  End If
End Sub

Private Sub txtDes_Cobranza_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion3("Cod_Concepto_Cobranza", "Descripcion", "Cn_Ventas_Conceptos_Cobranza Where ", txtCod_Cobranza, txtDes_Cobranza, 2, Me)
End Sub

Private Sub txtCod_Cobranza1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  Call Busca_Opcion3("Cod_Concepto_Cobranza", "Descripcion", "Cn_Ventas_Conceptos_Cobranza Where ", txtCod_Cobranza1, txtDes_Cobranza1, 1, Me)
  FunctButt1.SetFocus
  End If
End Sub

Private Sub txtDes_Cobranza1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion3("Cod_Concepto_Cobranza", "Descripcion", "Cn_Ventas_Conceptos_Cobranza Where ", txtCod_Cobranza1, txtDes_Cobranza1, 2, Me)
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCod_Cobranza.SetFocus
End If
End Sub
