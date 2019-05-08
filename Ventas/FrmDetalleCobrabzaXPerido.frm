VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form FrmDetalleCobrabzaXPerido 
   Caption         =   "Detalle de Cobranza por Periodo"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   4080
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   3
      Top             =   -60
      Width           =   3930
      Begin VB.TextBox txtexterior 
         Height          =   285
         Left            =   2355
         TabIndex        =   8
         Top             =   1440
         Width           =   1395
      End
      Begin VB.TextBox txtmes 
         Height          =   300
         Left            =   2355
         TabIndex        =   0
         Top             =   360
         Width           =   1425
      End
      Begin VB.TextBox txtcobranza 
         Height          =   300
         Left            =   2385
         TabIndex        =   1
         Top             =   720
         Width           =   1380
      End
      Begin VB.TextBox txtabonadas 
         Height          =   285
         Left            =   2355
         TabIndex        =   2
         Top             =   1080
         Width           =   1395
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   810
         TabIndex        =   4
         Top             =   1920
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"FrmDetalleCobrabzaXPerido.frx":0000
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Imp. Dol. Cobranza Exterior :"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   2010
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Imp Dol Letras Abonadas :"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1875
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Imp. Dol. Parte Cobranzas : "
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes : "
         Height          =   195
         Left            =   1755
         TabIndex        =   5
         Top             =   360
         Width           =   435
      End
   End
End
Attribute VB_Name = "FrmDetalleCobrabzaXPerido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CODIGO, Descripcion, snum_corre As String
Public stipo_trabajador, scod_trabajador As String
Public sTip_Concepto, sCod_Concepto As String
Public sano As String
Dim strSQL As String
Public Stipo, Saccion As String
Public smes As String
Public sCobranzas As String
Public sAbonadas As String
Public sExterior As String


Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "GRABAR"
        Grabar
Case "CANCELAR"
    Unload Me
End Select

End Sub
Sub Grabar()
Dim sql As String
On Error GoTo hand


 sql = "RH_DETALLE_COBRANZA_X_PERIODO  '" & Saccion & "','" & sano & "','" & Trim(txtMes.Text) & "','" & Trim(txtcobranza.Text) & "','" & Trim(txtabonadas.Text) & "','" & Trim(txtexterior.Text) & "'"

 Call ExecuteSQL(cCONNECT, sql)
 
 Unload Me
Exit Sub
hand:
    SALVAR_CABECERA = False
    ErrorHandler Err, "Grabar Conceptos"
End Sub

Private Sub txtano_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    TxtNumero.SetFocus
End If
End Sub

Private Sub txtnumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtImporte.SetFocus
End If
End Sub
