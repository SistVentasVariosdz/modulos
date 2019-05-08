VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form FrmModReq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requerimiento"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1530
      TabIndex        =   26
      Top             =   4275
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "0~0~ACEPTAR~True~True~&Aceptar~0~0~1~~0~False~False~&Aceptar~~1~0~CANCELAR~True~True~&Cancelar~0~0~2~~0~False~False~&Cancelar~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.TextBox TxtReqNew 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4305
      TabIndex        =   25
      Text            =   "0"
      Top             =   3870
      Width           =   1095
   End
   Begin VB.TextBox TxtReqAnt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4305
      TabIndex        =   23
      Text            =   "0"
      Top             =   3540
      Width           =   1080
   End
   Begin VB.TextBox TxtEstilo 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1185
      TabIndex        =   21
      Top             =   3075
      Width           =   4200
   End
   Begin VB.TextBox TxtDestino 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1170
      TabIndex        =   20
      Top             =   2745
      Width           =   4215
   End
   Begin VB.TextBox TxtCompEst 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1185
      TabIndex        =   17
      Top             =   1095
      Width           =   1320
   End
   Begin VB.TextBox TxtPresent 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1170
      TabIndex        =   15
      Top             =   765
      Width           =   4230
   End
   Begin VB.TextBox TxtOrdPro 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1185
      TabIndex        =   13
      Top             =   435
      Width           =   1320
   End
   Begin VB.TextBox TxtSecuencia 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4725
      TabIndex        =   11
      Top             =   105
      Width           =   660
   End
   Begin VB.TextBox TxtColor 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1170
      TabIndex        =   9
      Top             =   2085
      Width           =   4230
   End
   Begin VB.TextBox TxtTalla 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1170
      TabIndex        =   7
      Top             =   2415
      Width           =   1350
   End
   Begin VB.TextBox TxtComb 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1170
      TabIndex        =   5
      Top             =   1755
      Width           =   4215
   End
   Begin VB.TextBox TxtOrdComp 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1185
      TabIndex        =   3
      Top             =   105
      Width           =   1320
   End
   Begin VB.TextBox TxtItem 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1185
      TabIndex        =   2
      Top             =   1425
      Width           =   4200
   End
   Begin VB.Label Label13 
      Caption         =   "Cantidad Requerida Nueva :"
      Height          =   225
      Left            =   1890
      TabIndex        =   24
      Top             =   3915
      Width           =   2235
   End
   Begin VB.Label Label12 
      Caption         =   "Cantidad Requerida Anterior :"
      Height          =   225
      Left            =   1890
      TabIndex        =   22
      Top             =   3600
      Width           =   2235
   End
   Begin VB.Label Label11 
      Caption         =   "Estilo"
      Height          =   225
      Left            =   135
      TabIndex        =   19
      Top             =   3135
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Destino"
      Height          =   210
      Left            =   150
      TabIndex        =   18
      Top             =   2805
      Width           =   960
   End
   Begin VB.Label Label9 
      Caption         =   "Composicion"
      Height          =   210
      Left            =   105
      TabIndex        =   16
      Top             =   1170
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Presentacion"
      Height          =   210
      Left            =   120
      TabIndex        =   14
      Top             =   855
      Width           =   1065
   End
   Begin VB.Label Label7 
      Caption         =   "O.P."
      Height          =   195
      Left            =   150
      TabIndex        =   12
      Top             =   510
      Width           =   990
   End
   Begin VB.Label Label6 
      Caption         =   "Secuencia"
      Height          =   195
      Left            =   3780
      TabIndex        =   10
      Top             =   150
      Width           =   900
   End
   Begin VB.Label Label5 
      Caption         =   "Color"
      Height          =   210
      Left            =   135
      TabIndex        =   8
      Top             =   2190
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Talla"
      Height          =   210
      Left            =   135
      TabIndex        =   6
      Top             =   2490
      Width           =   945
   End
   Begin VB.Label Label3 
      Caption         =   "Combinacion"
      Height          =   210
      Left            =   135
      TabIndex        =   4
      Top             =   1845
      Width           =   945
   End
   Begin VB.Label Label2 
      Caption         =   "Item"
      Height          =   210
      Left            =   135
      TabIndex        =   1
      Top             =   1500
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Nº O. C."
      Height          =   240
      Left            =   165
      TabIndex        =   0
      Top             =   165
      Width           =   990
   End
End
Attribute VB_Name = "FrmModReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sSer_OrdComp As String
Public sCod_OrdComp As String
Public sSec_OrdComp As String
Public sCod_Fabrica As String
Public sCod_OrdPro As String
Public sCod_Present As Integer
Public sCod_CompEst As String
Public sCod_Item As String
Public sCod_Comb As String
Public sCod_Color As String
Public sCod_Talla As String
Public sCod_Destino As String
Public sCod_EstCli As String

Public sTipoItem As String
Dim StrSql As String

Private Sub Form_Load()
    TxtReqNew.SelStart = 0
    TxtReqNew.SelLength = Len(TxtReqNew.Text)
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            If Trim(TxtReqAnt.Text) = "" Then TxtReqAnt.Text = 0
            If CDbl(TxtReqAnt.Text) < CDbl(TxtReqNew.Text) Then
                MsgBox "La cantidad Requerida Nueva no puede ser Mayor a la cantidad Requerida Anterior", vbInformation, Me.Caption
                TxtReqNew.SetFocus
                TxtReqNew.SelStart = 0
                TxtReqNew.SelLength = Len(TxtReqNew.Text)
                Exit Sub
            End If
            SALVAR_DATOS
            Unload Me
        Case "CANCELAR"
            Unload Me
    End Select
End Sub

Private Sub TxtReqNew_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FunctButt1.SetFocus
    Else
        Call SoloNumeros(TxtReqNew, KeyAscii, True, 7)
    End If
End Sub

Sub SALVAR_DATOS()
 On Error GoTo Salvar_DatosErr
   
        StrSql = "EXEC UP_MAN_ORDCOMPITEMREQ_MODI '" & _
        sSer_OrdComp & "','" & _
        sCod_OrdComp & "','" & _
        sSec_OrdComp & "','" & _
        sCod_Fabrica & "','" & _
        sCod_OrdPro & "','" & _
        sCod_Present & "','" & _
        sCod_CompEst & "','" & _
        sCod_Item & "','" & _
        sCod_Comb & "','" & _
        sCod_Color & "','" & _
        sCod_Talla & "','" & _
        sCod_Destino & "','" & _
        sCod_EstCli & "'," & _
        TxtReqAnt.Text & "," & _
        TxtReqNew.Text
        
        Call ExecuteSQL(cConnect, StrSql)
        
Exit Sub
Salvar_DatosErr:
    ErrorHandler Err, "Salvar_Datos"

End Sub


