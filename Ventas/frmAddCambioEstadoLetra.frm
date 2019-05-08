VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddCambioEstadoLetra 
   Caption         =   "Entrega Proveedor"
   ClientHeight    =   1365
   ClientLeft      =   5880
   ClientTop       =   4485
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1365
   ScaleWidth      =   3975
   Begin MSComCtl2.DTPicker dtpFecEstado 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   180
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   23986177
      CurrentDate     =   38371
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmAddCambioEstadoLetra.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha  Entrega Proveedor:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1920
   End
End
Attribute VB_Name = "frmAddCambioEstadoLetra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Num_Corre As String

Private Sub Form_Load()
  dtpFecEstado = Date
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo hand
Dim sSql As String

Select Case ActionName
  Case "ACEPTAR"
  
    sSql = "UP_CN_DOCUM_ENTREGADO_PROVEEDOR '" & Num_Corre & "'," & IIf(IsNull(dtpFecEstado), "Null", "'" & dtpFecEstado & "'")
    ExecuteSQL cCONNECT, sSql
  
    Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
    
    Unload Me
  
  Case "CANCELAR"
    Unload Me
End Select

Exit Sub
hand:
ErrorHandler Err, "ELIMINAR_LETRA"
End Sub

