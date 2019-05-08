VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_FactExpoSujetas 
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3615
      Begin MSComCtl2.DTPicker dtpAnoMes 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM yyyy"
         Format          =   62652419
         CurrentDate     =   37887
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes - Año: "
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
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   690
      TabIndex        =   0
      Top             =   1155
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"Frm_FactExpoSujetas.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   120
      Top             =   1065
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "Frm_FactExpoSujetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
dtpAnoMes.Value = Date
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "IMPRIMIR"
         
                    Call Reporte
            
        Case "SALIR"
            Unload Me
    End Select
End Sub


Public Sub Reporte()
Dim oo As Object

On Error GoTo ErrorImpresion

    VB.Screen.MousePointer = vbHourglass
    
    strSQL = "CN_VENTAS_FACTURAS_EXPO_SUJETAS_DRAW_BACK '" & Format(dtpAnoMes.Value, "YYYY") & "','" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2) & "'"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\RptVenFacExpSujetas.XLT"

    oo.Visible = True
    oo.Run "REPORTE", strSQL, Format(dtpAnoMes.Value, "YYYY") & "-" & Right("00" & Format(dtpAnoMes.Value, "MM"), 2), cCONNECT
        
    Screen.MousePointer = vbNormal
    
    Set oo = Nothing
    
    Exit Sub
ErrorImpresion:
    Screen.MousePointer = vbNormal
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte" & err.Description, vbCritical, "Impresion"
End Sub

