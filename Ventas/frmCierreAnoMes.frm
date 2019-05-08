VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCierreAnoMes 
   Caption         =   "Cierre Año Mes"
   ClientHeight    =   1845
   ClientLeft      =   3765
   ClientTop       =   2310
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1845
   ScaleWidth      =   3285
   Begin VB.Frame frMensual 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin MSComCtl2.DTPicker DTAnoMes 
         Height          =   330
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MM / yyyy"
         Format          =   94109699
         CurrentDate     =   37987
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Año/Mes :"
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   315
         Width           =   750
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmCierreAnoMes.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmCierreAnoMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  DTAnoMes = "01/" & DevuelveCampo("select Ultimo_Mes_Cerrado from Cn_Control_Ventas", cCONNECT) & "/" & DevuelveCampo("select Ultimo_Ano_Cerrado from Cn_Control_Ventas", cCONNECT)
  DTAnoMes = DTAnoMes + 31
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

On Error GoTo dprDepurar

Select Case ActionName

Case Is = "GRABAR"
  If MsgBox("Desea Grabar este Producto " & txtCod_Producto, vbYesNo, "AVISO") = vbYes Then
    Grabar
    MsgBox "El Proceso ha terminado satisfactoriamente", vbInformation, "AVISO"
    Unload Me
  End If
Case Is = "CANCELAR"
  Unload Me
End Select

Exit Sub

dprDepurar:

errores err.Number


End Sub


Sub Grabar()
 
Dim RS As Object
Set RS = CreateObject("ADODB.Recordset")
 
strSQL = "Ventas_Cierre_Ano_Mes '" & Format(DTAnoMes, "yyyy") & "','" & Format(DTAnoMes, "mm") & "','" & vusu & "','" & ComputerName & "'"

ExecuteCommandSQL cCONNECT, strSQL

       
End Sub


