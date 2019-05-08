VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmPesosBal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Secuencia de Pesos"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   420
      Left            =   1950
      TabIndex        =   1
      Top             =   3525
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   741
      Custom          =   $"frmPesosBal.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   400
      ControlSeparator=   110
   End
   Begin VB.ListBox lstPesos 
      Height          =   2985
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   7455
   End
   Begin MSCommLib.MSComm mscCentral 
      Left            =   180
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Total Kg :"
      Height          =   225
      Left            =   5460
      TabIndex        =   3
      Top             =   3150
      Width           =   915
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      TabIndex        =   2
      Top             =   3105
      Width           =   1005
   End
End
Attribute VB_Name = "frmPesosBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sTit As String, sErr As String

Private Sub Form_Load()
On Error GoTo ErrLoad
    lstPesos.Clear
    CalculaTotal
    mscCentral.CommPort = 1
    mscCentral.Settings = "9600,S,7,1"
    mscCentral.PortOpen = True
Exit Sub
ErrLoad:
    MsgBox err.Description, vbCritical + vbOKOnly, "Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.Visible Then
        mscCentral.PortOpen = False
        Me.Hide
        Cancel = 200
    End If
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo ErrPeso
Dim iRow As Long, sPeso As String
    
    Select Case ActionName
    Case "CAPTURAR"
        sTit = "Capturar Peso"
        iRow = 1
        If lstPesos.ListCount > 0 Then _
        iRow = CLng(Left(lstPesos.List(lstPesos.ListCount - 1), 2)) + 1
        
        Close #1
        Open "C:\pesos.txt" For Output As #1
        
        sPeso = mscCentral.Input
        
        Print #1, sPeso
        
        'El Input menos el "kg" + 2 chr(13) final (menos 4 caracteres)
        sPeso = Left(sPeso, Len(sPeso) - 4)
        Print #1, "menos 4 carac. finales"
        Print #1, sPeso
        
        'El Input menos el chr(13) anterior al ultimo peso
        sPeso = Mid(sPeso, InStrRev(sPeso, vbLf) + 1)
        Print #1, "desde el ultimo chr(13) + 1"
        Print #1, sPeso
        
'        'La cadena sin chr(13) se le quita el 'KG'
'        sPeso = UCase(sPeso)
'        sPeso = Replace(sPeso, "KG", "")
        
        'Es un numero? si no que inicie el error
        Print #1, "Es un numero? si no que inicie el error"
        Print #1, sPeso
        Close #1
        sPeso = CDbl(sPeso)
        lstPesos.AddItem Format(iRow, "00") & Space(6) & Format(sPeso, "0.00")
        
        CalculaTotal
    Case "ELIMINAR"
        sTit = "Eliminar Peso"
        If lstPesos.ListIndex < 0 Then Exit Sub
        lstPesos.RemoveItem lstPesos.ListIndex
        CalculaTotal
    Case "SALIR"
        sTit = "Salir"
        Unload Me
    End Select
Exit Sub
ErrPeso:
    Close #1
    MsgBox err.Description & ", peso no detectado", vbCritical + vbOKOnly, sTit
End Sub

Private Sub CalculaTotal()
Dim iRow As Long, dTotal As Double
    
    dTotal = 0
    For iRow = 0 To lstPesos.ListCount - 1
        dTotal = dTotal + CDbl(Trim(Mid(lstPesos.List(iRow), 3)))
    Next iRow
    lblTotal = Format(dTotal, "0.00")
End Sub
