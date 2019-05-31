VERSION 5.00
Begin VB.Form frmCambioSinglePrePack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio Single/PrePack"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   870
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.OptionButton optPrePackNo 
         Caption         =   "No"
         Height          =   255
         Left            =   2040
         TabIndex        =   1
         Top             =   300
         Width           =   615
      End
      Begin VB.OptionButton optPrePackSi 
         Caption         =   "Si"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pre Pack :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Width           =   750
      End
   End
   Begin VB.Shape Shape1 
      Height          =   630
      Left            =   3240
      Top             =   90
      Width           =   2775
   End
End
Attribute VB_Name = "frmCambioSinglePrePack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sCodigoCliente       As String

Public sCodigoPurOrd        As String

Public sCodigoLotePurOrd    As String

Public sCodigoEstadoCliente As String

Private Function UpdateDatGen() As Boolean

    On Error GoTo Error
        
    Dim oCn   As New Connection

    Dim oCmd  As New Command

    Dim sFlag As String * 1
    
    sFlag = "S": If optPrePackNo.value = True Then sFlag = "N"
    oCn.ConnectionString = cCONNECT: oCn.Open

    With oCmd
        .ActiveConnection = oCn
        .CommandType = adCmdStoredProc
        .CommandText = "SM_TG_LotEstUpdate_PrePack"
        .Parameters.Append .CreateParameter("@Cod_Cliente", adChar, adParamInput, 5, sCodigoCliente)
        .Parameters.Append .CreateParameter("@Cod_PurOrd", adChar, adParamInput, 20, sCodigoPurOrd)
        .Parameters.Append .CreateParameter("@Cod_LotPurOrd", adChar, adParamInput, 3, sCodigoLotePurOrd)
        .Parameters.Append .CreateParameter("@Cod_EstCli", adChar, adParamInput, 20, sCodigoEstadoCliente)
        .Parameters.Append .CreateParameter("@flg_prepack", adChar, adParamInput, 1, sFlag)
        .Execute
    End With

    Set oCn = Nothing
    Set oCmd = Nothing
    Unload Me

    Exit Function

Error:
    Set oCn = Nothing: Set oCmd = Nothing
    ErrorHandler Err, Err.Description
End Function

Private Sub cmdAceptar_Click()
    UpdateDatGen
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub
