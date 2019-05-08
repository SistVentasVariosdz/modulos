VERSION 5.00
Begin VB.Form FrmIngresoEstilo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Estilo Propio"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   525
      Left            =   2370
      TabIndex        =   10
      Top             =   2565
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   525
      Left            =   480
      TabIndex        =   9
      Top             =   2565
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Height          =   2220
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   3825
      Begin VB.ComboBox cboCod_UsuPre 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1710
         Width           =   2415
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Top             =   300
         Width           =   2415
      End
      Begin VB.ComboBox cmbTipPre 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   750
         Width           =   2415
      End
      Begin VB.ComboBox CmdGruTal 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1230
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usu. Pren :"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   1785
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion"
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Prenda"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Grupo Talla"
         Height          =   315
         Index           =   2
         Left            =   150
         TabIndex        =   5
         Top             =   1260
         Width           =   1425
      End
   End
End
Attribute VB_Name = "FrmIngresoEstilo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Estilo As String
Public Descripcion As String
'Dim Reg As New ADODB.Recordset
Public Papa As Object
Dim Strsql As String
Sub Limpia()
Me.TxtDescripcion = ""
Me.cmbTipPre.ListIndex = -1
Me.CmdGruTal.ListIndex = -1
Me.cboCod_UsuPre.ListIndex = -1
End Sub


Private Sub Command1_Click()
Dim varCod_EstPro As String
On Error GoTo hand
'Set Reg = Nothing

'Reg.CursorLocation = adUseClient
''Reg.Open "UP_Es_EsPro 'I','" & Cliente & "','" & Estilo & "','" & Temporada & "','" & EstProp & "','" & Descripcion & "','" & Right(Me.cmbTipPre, 3) & "','" & Right(Me.CmdGruTal, 3) & "','',''", cCONNECT
'Reg.Open "UP_Es_EsPro 'i','','','','','" & Me.TxtDescripcion & "','" & Right(Me.cmbTipPre, 3) & "','" & Right(Me.CmdGruTal, 3) & "','',''", cCONNECT

'Set Reg = Nothing

varCod_EstPro = DevuelveCampo("EXEC UP_Es_EsPro 'i','','','','','" & Me.TxtDescripcion & "','" & Right(Me.cmbTipPre, 3) & "','" & Right(Me.CmdGruTal, 3) & "','','','" & Right(Me.cboCod_UsuPre.Text, 3) & "'", cCONNECT)
If varCod_EstPro <> "" Then
    Strsql = "select Des_estpro from es_estpro where Cod_EstPro='" & varCod_EstPro & "' and Cod_tippre='" & Right(Me.cmbTipPre, 3) & "' and Cod_GruTal='" & Right(Me.CmdGruTal, 3) & "'"
    Call MsgBox("Se creo el estilo " & Trim(DevuelveCampo(Strsql, cCONNECT)) & " con el siguiente numero :" & varCod_EstPro, vbInformation)
End If

With Papa
    .txtCod_EstPro = varCod_EstPro
    .txtDes_estpro = DevuelveCampo("select Des_estpro from es_estpro where Cod_EstPro='" & varCod_EstPro & "' and Cod_tippre='" & Right(Me.cmbTipPre, 3) & "' and Cod_GruTal='" & Right(Me.CmdGruTal, 3) & "'", cCONNECT)
End With
Limpia
Unload Me
Exit Sub
hand:
ErrorHandler Err, "Aceptar"
With Papa
    .txtCod_EstPro = DevuelveCampo("select Cod_EstPro from es_estpro where Des_estpro='" & Me.TxtDescripcion & "' and Cod_tippre='" & Right(Me.cmbTipPre, 3) & "' and Cod_GruTal='" & Right(Me.CmdGruTal, 3) & "'", cCONNECT)
    .txtDes_estpro = DevuelveCampo("select Des_estpro from es_estpro where Des_estpro='" & Me.TxtDescripcion & "' and Cod_tippre='" & Right(Me.cmbTipPre, 3) & "' and Cod_GruTal='" & Right(Me.CmdGruTal, 3) & "'", cCONNECT)
End With
'Set Reg = Nothing
Limpia
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Form_Load()
LlenaCombo Me.cmbTipPre, "Select Des_TipPre  + space(100) +Cod_TipPre from tg_tippre order by Des_TipPre  ", cCONNECT
LlenaCombo Me.CmdGruTal, "Select  des_grutal + space(100)+ cod_grutal  from es_tallas  order by des_grutal ", cCONNECT
LlenaCombo Me.cboCod_UsuPre, "Select  Des_UsuPre + space(100)+ Cod_UsuPre  from TG_USUPRENDAS", cCONNECT
TxtDescripcion = Me.Descripcion
End Sub


