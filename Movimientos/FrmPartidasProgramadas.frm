VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmPartidasProgramadas 
   Caption         =   "Partidas Programadas"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBuscar 
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
      Height          =   540
      Left            =   7035
      TabIndex        =   3
      Top             =   240
      Width           =   1380
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detalle"
      Height          =   4395
      Left            =   0
      TabIndex        =   7
      Top             =   1050
      Width           =   8805
      Begin GridEX20.GridEX gexList 
         Height          =   4110
         Left            =   90
         TabIndex        =   8
         Top             =   195
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   7250
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         SelectionStyle  =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "FrmPartidasProgramadas.frx":0000
         Column(2)       =   "FrmPartidasProgramadas.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "FrmPartidasProgramadas.frx":016C
         FormatStyle(2)  =   "FrmPartidasProgramadas.frx":02A4
         FormatStyle(3)  =   "FrmPartidasProgramadas.frx":0354
         FormatStyle(4)  =   "FrmPartidasProgramadas.frx":0408
         FormatStyle(5)  =   "FrmPartidasProgramadas.frx":04E0
         FormatStyle(6)  =   "FrmPartidasProgramadas.frx":0598
         ImageCount      =   0
         PrinterProperties=   "FrmPartidasProgramadas.frx":0678
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   555
      Left            =   2940
      TabIndex        =   4
      Top             =   5460
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   979
      Custom          =   "0~0~IMPRIMIR~True~True~&Imprimir~0~0~1~~0~False~False~&Imprimir~~1~0~SALIR~True~True~&Salir~0~0~2~~0~False~False~&Salir~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1300
      ControlHeigth   =   530
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6420
      Begin MSComCtl2.DTPicker DTInicio 
         Height          =   330
         Left            =   1995
         TabIndex        =   1
         Top             =   630
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   57737217
         CurrentDate     =   38001
      End
      Begin VB.OptionButton OptFecha 
         Caption         =   "Por Rango Fechas"
         Height          =   330
         Left            =   210
         TabIndex        =   9
         Top             =   630
         Width           =   1695
      End
      Begin VB.TextBox TxtPartida 
         Height          =   330
         Left            =   1995
         TabIndex        =   0
         Top             =   210
         Width           =   2220
      End
      Begin VB.OptionButton OptPartida 
         Caption         =   "Por Partida"
         Height          =   330
         Left            =   210
         TabIndex        =   6
         Top             =   210
         Value           =   -1  'True
         Width           =   1170
      End
      Begin MSComCtl2.DTPicker DTFin 
         Height          =   330
         Left            =   4305
         TabIndex        =   2
         Top             =   630
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   57737217
         CurrentDate     =   38001
      End
   End
End
Attribute VB_Name = "FrmPartidasProgramadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim vCod_Ordprov As String
Dim Tipo As String


Sub CARGA_GRID()
On Error GoTo hand
If Tipo = 1 Then
    strSQL = "EXEC sm_muestra_partidas_programadas '" & Tipo & "','" & vCod_Ordprov & "',NULL,NULL"
Else
    strSQL = "EXEC sm_muestra_partidas_programadas '" & Tipo & "','" & vCod_Ordprov & "','" & DTInicio.Value & "','" & DTFin.Value & "'"
End If
VB.Screen.MousePointer = 11
Set Me.gexList.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
VB.Screen.MousePointer = 0

If OptPartida.Value = True Then ConfigurarGrid
Exit Sub
hand:
ErrorHandler Err, "CARGA_GRID"
End Sub

Sub ConfigurarGrid()
'    gexList.Columns("Cod.Avio").Width = 1200
'    gexList.Columns("Descripcion").Width = 2500
'    gexList.Columns("UN").Width = 700
'    gexList.Columns("Origen").Width = 700
'    gexList.Columns("Requerida").Width = 1000
'    gexList.Columns("Comprada").Width = 1000
'    gexList.Columns("Recibida").Width = 1000
End Sub

Private Sub CmdBuscar_Click()
If OptPartida.Value = True Then
    vCod_Ordprov = RTrim(TxtPartida)
    If vCod_Ordprov = "" Then
        MsgBox "Ingrese Partida a Buscar", vbInformation
        Exit Sub
    End If
    Tipo = "1"
Else
    Tipo = "2"
End If
Call CARGA_GRID
End Sub

Private Sub Form_Load()
DTInicio.Value = Date
DTFin.Value = Date
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
    Call Reporte
Case "SALIR"
    Unload Me
End Select
End Sub

Private Sub Reporte()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String

    Screen.MousePointer = 11
    Ruta = vRuta & "\RptTelasGeneradas.xlt"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    
   oo.Run "Reporte", gexList.ADORecordset, vCod_Ordprov, cConnect
    Set oo = Nothing
    Screen.MousePointer = 0
Exit Sub
hand:
    Screen.MousePointer = 0
    ErrorHandler Err, "GeneraReportes"
    Set oo = Nothing
End Sub

Private Sub OptFecha_Click()
If OptPartida.Value = True Then
    TxtPartida.Enabled = True
    DTInicio.Enabled = False
    DTFin.Enabled = False
ElseIf OptFecha.Value = True Then
    DTInicio.Enabled = True
    DTFin.Enabled = True
    TxtPartida.Enabled = False
End If
End Sub

Private Sub OptPartida_Click()
If OptPartida.Value = True Then
    TxtPartida.Enabled = True
    DTInicio.Enabled = False
    DTFin.Enabled = False
ElseIf OptFecha.Value = True Then
    DTInicio.Enabled = True
    DTFin.Enabled = True
    TxtPartida.Enabled = False
End If
End Sub

Private Sub TxtPartida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdBuscar.SetFocus
End If
End Sub
