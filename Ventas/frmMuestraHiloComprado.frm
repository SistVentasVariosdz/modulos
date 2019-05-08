VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Begin VB.Form frmMuestraHiloComprado 
   Caption         =   "Ventas de Hilo Comprado"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Imprimir"
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
      Left            =   960
      TabIndex        =   4
      Tag             =   "&OK"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
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
      Left            =   2400
      TabIndex        =   5
      Tag             =   "&Cancel"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   4335
      Begin VB.TextBox TxtMes 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   2
         Top             =   360
         Width           =   660
      End
      Begin VB.TextBox TxtAno 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   1
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label2 
         Caption         =   "Mes"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Año"
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   375
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   4080
      Top             =   1920
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmMuestraHiloComprado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdAceptar_Click()
Call ImprimirReporte
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtAno.Text = Format(Date, "yyyy")
txtMes.Text = Format(Date, "MM")

End Sub


 Public Sub ImprimirReporte()
 On Error GoTo ErrorImpresion
 Dim oo As Object
 
Dim Adors1 As Object
Set Adors1 = CreateObject("ADODB.Recordset")
Dim rutaLogo As String
rutaLogo = DevuelveCampo("select ruta_logo=isNUll(ruta_logo,'') from seguridad..seg_empresas where cod_empresa='" & vemp & "'", cCONNECT)

strSQL = " EXEC VENTAS_MUESTRA_HILO_COMPRADO '" & txtAno.Text & "','" & txtMes.Text & "'"

Set Adors1 = CargarRecordSetDesconectado(strSQL, cCONNECT)

Set oo = CreateObject("Excel.Application")
    oo.Workbooks.Open vRuta & "\Rpt_Muestra_Hilo_Comprado1.XLT"
    oo.Visible = True
    oo.displayalerts = False
    oo.Run "Reporte", Adors1, rutaLogo, txtAno.Text, txtMes.Text
Set oo = Nothing
 

Exit Sub
ErrorImpresion:

   Screen.MousePointer = vbNormal
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte" & err.Description, vbCritical, "Impresion"
End Sub



Private Sub txtAno_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   txtMes.SetFocus
 Else

   Call SoloNumeros(txtAno, KeyAscii, False, 0, 4)
 End If
End Sub

Private Sub txtMes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
   cmdAceptar.SetFocus
  Else
  Call SoloNumeros(txtMes, KeyAscii, False, 0, 2)
 End If
End Sub
