VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConosHilosCoserEnvProv 
   Caption         =   "Conos de Hilos de Coser Enviados Por Proveedor"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1680
      TabIndex        =   5
      Top             =   960
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmConosHilosCoserEnvProv.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin MSComCtl2.DTPicker DTPInicio 
         Height          =   255
         Left            =   1200
         TabIndex        =   1
         Top             =   290
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   23986177
         CurrentDate     =   38449
      End
      Begin MSComCtl2.DTPicker DTPFin 
         Height          =   255
         Left            =   3960
         TabIndex        =   2
         Top             =   290
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   23986177
         CurrentDate     =   38449
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fin"
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
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmConosHilosCoserEnvProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
DTPInicio.Value = Date
Me.DTPFin.Value = Date
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
  Case "IMPRIMIR"
      Call Reporte
  Case "CANCELAR"
      Unload Me
End Select
End Sub
Sub Reporte()
On Error GoTo ErrorImpresion
Dim oo As Object

    'strSQL = "Planeamiento_Analisis_Abastecimiento_Avios_Semanal '" & DTPInicio & "','" & DTPFin & "'"
    
    Set oo = CreateObject("excel.application")
    
    oo.Workbooks.Open vRuta & "\RptConosEnviados.XLT"
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "reporte", DTPInicio, DTPFin, cConnect
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub

