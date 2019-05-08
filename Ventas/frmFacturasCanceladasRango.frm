VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmFacturasCanceladasRango 
   Caption         =   "Facturas Canceladas Segun Rango de Fechas"
   ClientHeight    =   1650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   1650
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   90767363
         CurrentDate     =   39018
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   285
         Left            =   4200
         TabIndex        =   3
         Top             =   240
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   90767363
         CurrentDate     =   39018
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta:"
         Height          =   210
         Left            =   3480
         TabIndex        =   4
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Desde:"
         Height          =   210
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   540
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   615
      Left            =   1920
      TabIndex        =   5
      Top             =   1080
      Width           =   3240
      _ExtentX        =   5556
      _ExtentY        =   1111
      Custom          =   $"frmFacturasCanceladasRango.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1500
      ControlHeigth   =   600
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   360
      Top             =   2160
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmFacturasCanceladasRango"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strSQL As String
Public vopcion As String
Public sNomOpcion As String

Private Sub Form_Load()
dtpInicio.Value = Date
dtpFin.Value = Date
End Sub


 Public Sub ImprimirReporte()
 On Error GoTo ErrorImpresion
 Dim oo As Object
  
Dim Adors1 As Object
Set Adors1 = CreateObject("ADODB.Recordset")
Dim rutaLogo As String
rutaLogo = DevuelveCampo("select ruta_logo=isNUll(ruta_logo,'') from seguridad..seg_empresas where cod_empresa='" & vemp & "'", cCONNECT)





strSQL = " EXEC cn_ventas_muestra_facturas_Canceladas '" & dtpInicio.Value & "','" & dtpFin.Value & "','1'"

Set Adors1 = CargarRecordSetDesconectado(strSQL, cCONNECT)

Set oo = CreateObject("Excel.Application")
    oo.Workbooks.Open vRuta & "\Rpt_Facturas_Canceladas.xlt"
    oo.Visible = True
    oo.displayalerts = False
    oo.Run "Reporte", rutaLogo, Adors1, dtpInicio.Value, dtpFin.Value, sNomOpcion
Set oo = Nothing
 

Exit Sub
ErrorImpresion:

   Screen.MousePointer = vbNormal
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte" & err.Description, vbCritical, "Impresion"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "IMPRIMIR"
        ImprimirReporte
        
    Case "SALIR"
        Unload Me
End Select
End Sub
