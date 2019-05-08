VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmNuevoMot_Rechazo 
   Caption         =   "Nuevo Motivo Rechazo"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSeleccionar 
      Caption         =   "Seleccionar"
      Height          =   540
      Left            =   3990
      TabIndex        =   6
      Top             =   3150
      Width           =   1065
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   0
      TabIndex        =   5
      Top             =   3150
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmNuevoMot_Rechazo.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin VB.Frame Frame2 
      Height          =   2400
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5070
      Begin GridEX20.GridEX gexMotivo_Rechazo 
         Height          =   2040
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   3598
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "FrmNuevoMot_Rechazo.frx":0160
         Column(2)       =   "FrmNuevoMot_Rechazo.frx":0228
         FormatStylesCount=   6
         FormatStyle(1)  =   "FrmNuevoMot_Rechazo.frx":02CC
         FormatStyle(2)  =   "FrmNuevoMot_Rechazo.frx":0404
         FormatStyle(3)  =   "FrmNuevoMot_Rechazo.frx":04B4
         FormatStyle(4)  =   "FrmNuevoMot_Rechazo.frx":0568
         FormatStyle(5)  =   "FrmNuevoMot_Rechazo.frx":0640
         FormatStyle(6)  =   "FrmNuevoMot_Rechazo.frx":06F8
         ImageCount      =   0
         PrinterProperties=   "FrmNuevoMot_Rechazo.frx":07D8
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   2415
      Width           =   5055
      Begin VB.TextBox TxtDes_MotRechazo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1260
         TabIndex        =   1
         Top             =   210
         Width           =   3585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   210
         TabIndex        =   2
         Top             =   280
         Width           =   840
      End
   End
End
Attribute VB_Name = "FrmNuevoMot_Rechazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim sTipo As String

Private Sub CmdSeleccionar_Click()
With frmDetDatosTecnicos
            .TxtCod_MotRechazo = gexMotivo_Rechazo.Value(gexMotivo_Rechazo.Columns("Cod_MotRechazo").Index)
            .TxtDes_MotRechazo = gexMotivo_Rechazo.Value(gexMotivo_Rechazo.Columns("Des_MotRechazo").Index)
End With
Unload Me
End Sub

Private Sub Form_Load()
CARGA_GRID
End Sub

Sub CARGA_DATOS()
On Error GoTo hand
    TxtDes_MotRechazo.Text = Trim(gexMotivo_Rechazo.Value(gexMotivo_Rechazo.Columns("Des_MotRechazo").Index))
Exit Sub
hand:
ErrorHandler err, "CARGA_DATOS"
End Sub

Sub CARGA_GRID()
On Error GoTo hand

strSQL = "EXEC TX__MAN_MOTIVORECHAZO 'S','',''"
Set Me.gexMotivo_Rechazo.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)

Exit Sub
hand:
ErrorHandler err, "CARGA_GRID"
End Sub

Private Sub gexMotivo_Rechazo_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
If gexMotivo_Rechazo.RowCount = 0 Then Exit Sub
    CARGA_DATOS
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo hand
Select Case ActionName
    Case "ADICIONAR"
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Limpia
        HABILITA_CAMPOS True
        gexMotivo_Rechazo.Enabled = False
        TxtDes_MotRechazo.SetFocus
        sTipo = "I"
    Case "MODIFICAR"
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        sTipo = "U"
        HABILITA_CAMPOS True
        gexMotivo_Rechazo.Enabled = False
        TxtDes_MotRechazo.SetFocus
    Case "ELIMINAR"
        sTipo = "D"
        SALVAR_DATOS
        Limpia
        CARGA_GRID
        HABILITA_CAMPOS False
        sTipo = ""
    Case "GRABAR"
        SALVAR_DATOS
        Limpia
        HABILITA_CAMPOS False
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        gexMotivo_Rechazo.Enabled = True
        CARGA_GRID
        sTipo = ""
    Case "DESHACER"
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Limpia
        gexMotivo_Rechazo.Enabled = True
        CARGA_GRID
        HABILITA_CAMPOS False
        sTipo = ""
    Case "SALIR"
        
        Unload Me
End Select

Exit Sub
hand:
ErrorHandler err, "MantFunc1_ActionClick"
End Sub

Sub HABILITA_CAMPOS(vEstado As Boolean)
    TxtDes_MotRechazo.Enabled = vEstado
End Sub

Sub Limpia()
    TxtDes_MotRechazo.Text = ""
End Sub

Sub SALVAR_DATOS()
On Error GoTo hand
If TxtDes_MotRechazo = "" Then
    MsgBox "Ingrese Descripcion", vbInformation
    Exit Sub
End If

strSQL = "TX__MAN_MOTIVORECHAZO '" & sTipo & "','" & gexMotivo_Rechazo.Value(gexMotivo_Rechazo.Columns("Cod_MotRechazo").Index) & "','" & TxtDes_MotRechazo & "'"
    
Call ExecuteSQL(cConnect, strSQL)
Exit Sub
hand:
    ErrorHandler err, "SALVAR_DATOS"
End Sub

