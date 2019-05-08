VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form Frm_Adicionar_Utilidad 
   Caption         =   "Adicionar"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2565
      Left            =   0
      TabIndex        =   3
      Top             =   -60
      Width           =   4170
      Begin VB.TextBox txtano 
         Height          =   300
         Left            =   2355
         TabIndex        =   0
         Top             =   360
         Width           =   1425
      End
      Begin VB.TextBox txtnumero 
         Height          =   300
         Left            =   2385
         TabIndex        =   1
         Text            =   "365"
         Top             =   720
         Width           =   1380
      End
      Begin VB.TextBox txtimporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2355
         TabIndex        =   2
         Text            =   "0"
         Top             =   1080
         Width           =   1395
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   960
         TabIndex        =   4
         Top             =   1650
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"Frm_Adicionar_Utilidad.frx":0000
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Importe de Utlidad en Soles"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1950
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Número de Dias Efectivos : "
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año : "
         Height          =   195
         Left            =   1755
         TabIndex        =   5
         Top             =   360
         Width           =   420
      End
   End
End
Attribute VB_Name = "Frm_Adicionar_Utilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CODIGO, dESCRIPCION, snum_corre As String
Public stipo_trabajador, scod_trabajador As String
Public sTip_Concepto, sCod_Concepto As String
Public scod_fabrica As String
Dim strSQL As String
Public Stipo, Saccion As String


Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "GRABAR"
If Trim(txtnumero.Text) <= 365 Then
        Grabar
Else
        MsgBox "El Número de dias tiene que ser menor o igual que 365"
End If
Case "CANCELAR"
    Unload Me
End Select

End Sub
Sub Grabar()
Dim sql As String
On Error GoTo hand


 sql = "RH_agregar_Utilidades  '" & Saccion & "','" & Trim(scod_fabrica) & "','" & Trim(txtano.Text) & "','" & Trim(txtnumero.Text) & "','" & Trim(txtimporte.Text) & "','" & vusu & "'"

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
    txtnumero.SetFocus
End If
End Sub

Private Sub txtnumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtimporte.SetFocus
End If
End Sub
