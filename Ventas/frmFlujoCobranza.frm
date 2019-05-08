VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form frmFlujoCobranza 
   Caption         =   "Registro de Letras"
   ClientHeight    =   2820
   ClientLeft      =   285
   ClientTop       =   720
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2760
      Left            =   90
      TabIndex        =   3
      Top             =   -15
      Width           =   6045
      Begin VB.Frame frmano 
         Height          =   615
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   5535
         Begin VB.TextBox txtAno 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   840
            MaxLength       =   4
            TabIndex        =   8
            Top             =   240
            Width           =   660
         End
         Begin VB.TextBox txtMes 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   2280
            MaxLength       =   2
            TabIndex        =   10
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label8 
            Caption         =   "Mes :"
            Height          =   255
            Left            =   1800
            TabIndex        =   11
            Top             =   255
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Año :"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   255
            Width           =   495
         End
      End
      Begin VB.OptionButton Opt2 
         Caption         =   "Facturación Mensual"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   600
         Width           =   4695
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Saldos por Cobrar Pendiente"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.TextBox TxtDes_Banco 
         Height          =   285
         Left            =   1815
         TabIndex        =   1
         Top             =   1005
         Width           =   3855
      End
      Begin VB.TextBox TxtCod_Banco 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   1005
         Width           =   615
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   525
         Left            =   1965
         TabIndex        =   2
         Top             =   2040
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   926
         Custom          =   $"frmFlujoCobranza.frx":0000
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1075
         ControlHeigth   =   500
         ControlSeparator=   75
      End
      Begin VB.Label Label1 
         Caption         =   "Origen : "
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   735
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   4920
      Top             =   960
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmFlujoCobranza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public codigo As String, Descripcion As String
Public sTipoBusq As String
Public sopcion As Integer


Sub Reporte()
On Error GoTo ErrorImpresion
Dim oo As Object, lvSql As String
Dim strSQL As String
Dim sEmpresa As String, Ruta As String
    
    strSQL = "SELECT DES_EMPRESA FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA='" & vemp & "'"
    sEmpresa = DevuelveCampo(strSQL, cCONNECT)
    
    If MsgBox("Imprimir usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir") = vbYes Then
        Set oo = CreateObject("excel.application")
        oo.Workbooks.Open vRuta & "\RptFlujoCobranza.XLT"
        oo.Visible = True
        oo.DisplayAlerts = False
        
        oo.Run "reporte", TxtCod_Banco.Text, TxtDes_Banco.Text, sopcion, txtAno.Text, txtMes.Text, cCONNECT, sEmpresa
    Else
        Ruta = vRuta & "\RptFlujoCobranza.OTS"
        Set oo = CreateObject("ooBusiness.Calc")
        oo.OfficeTemplateSheet = Ruta
        oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
        oo.MacroLibraryName = "Library1"
        oo.MacroModuleName = "Module1"
        oo.MacroName = "Reporte"
        
        oo.Run TxtCod_Banco.Text, TxtDes_Banco.Text, sopcion, txtAno.Text, txtMes.Text, cCONNECT, sEmpresa
    End If
    
    Set oo = Nothing
Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte " & err.Description, vbCritical, "Impresion"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim varSecuencia As Integer

On Error GoTo hand

Select Case ActionName

  Case "IMPRIMIR"
    If Opt1.Value = True Then
        sopcion = 1
    Else
        sopcion = 2
    End If
      Reporte
      
  Case "SALIR"
    Unload Me
End Select

Exit Sub
Resume
hand:

errores err.Number

End Sub

Private Sub Opt1_Click()
    frmano.Visible = False
    TxtCod_Banco.SetFocus
End Sub

Private Sub Opt2_Click()
    frmano.Visible = True
    TxtCod_Banco.SetFocus
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
         txtMes.SetFocus

  End If
End Sub

Private Sub TxtCod_Banco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  Call Busca_Opcion3("Origen", "Des_Origen", "CN_Origen where ", TxtCod_Banco, TxtDes_Banco, 1, Me)
    If Opt1.Value = True Then
        FunctButt1.SetFocus
    Else
        txtAno.SetFocus
    End If
  End If
End Sub

Private Sub TxtDes_Banco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion3("Origen", "Des_Origen", "CN_Origen where ", TxtCod_Banco, TxtDes_Banco, 2, Me)
        If Opt1.Value = True Then
        FunctButt1.SetFocus
    Else
        txtAno.SetFocus
    End If
  End If
End Sub

Private Sub txtMes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     FunctButt1.SetFocus

  End If
End Sub



