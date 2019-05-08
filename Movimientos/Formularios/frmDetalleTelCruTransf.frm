VERSION 5.00
Begin VB.Form frmDetalleTelCruTransf 
   Caption         =   "Transferencia de Stock /Tela-Comb"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   555
      Left            =   4935
      TabIndex        =   18
      Top             =   2895
      Width           =   1650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   555
      Left            =   2940
      TabIndex        =   17
      Top             =   2910
      Width           =   1650
   End
   Begin VB.Frame Fradetalle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   90
      TabIndex        =   0
      Tag             =   "Detail"
      Top             =   90
      Width           =   9450
      Begin VB.ComboBox CmbCombinacion 
         Height          =   315
         Left            =   1185
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1215
         Width           =   3285
      End
      Begin VB.TextBox TxtDesitem 
         Height          =   315
         Left            =   2130
         TabIndex        =   5
         Top             =   840
         Width           =   2325
      End
      Begin VB.TextBox TxtItem 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1185
         MaxLength       =   8
         TabIndex        =   4
         Top             =   840
         Width           =   945
      End
      Begin VB.TextBox TxtLote 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1170
         MaxLength       =   15
         TabIndex        =   3
         Top             =   180
         Width           =   1935
      End
      Begin VB.TextBox TxtProveedor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1170
         TabIndex        =   2
         Top             =   510
         Width           =   3285
      End
      Begin VB.CommandButton cmdGetInfo 
         Height          =   285
         Left            =   3210
         Picture         =   "frmDetalleTelCruTransf.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Seleccionar Datos por Tela"
         Top             =   180
         Width           =   375
      End
      Begin VB.Label Label2 
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
         Left            =   1170
         TabIndex        =   15
         Top             =   1590
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Calidad:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   1590
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tela:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   945
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lote:"
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   12
         Top             =   255
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   11
         Top             =   585
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comb:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   1260
         Width           =   450
      End
      Begin VB.Label Label5 
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
         Left            =   1185
         TabIndex        =   9
         Top             =   1950
         Width           =   2355
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Medida:"
         Height          =   195
         Index           =   5
         Left            =   225
         TabIndex        =   8
         Top             =   1950
         Width           =   570
      End
      Begin VB.Label lblCantidad1 
         AutoSize        =   -1  'True
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
         Left            =   8010
         TabIndex        =   7
         Top             =   270
         Width           =   75
      End
      Begin VB.Label lblCantidad2 
         AutoSize        =   -1  'True
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
         Left            =   8010
         TabIndex        =   6
         Top             =   600
         Width           =   75
      End
   End
End
Attribute VB_Name = "frmDetalleTelCruTransf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public xcod_almacen As String
Public xNum_MovStk As String
Public xlote As String
Public xCod_Proveedor As String
Public xCod_Tela As String
Public xCod_Comb As String
Public xCod_Calidad As String
Public xCod_Medida As String
Public bOk As Boolean

Public Sub LlenaDatos()
  LlenaCombo Me.CmbCombinacion, "select Des_Comb+space(100)+Cod_Comb from TX_TELACOMB where COD_TELA ='" & Me.xCod_Tela & "'", cConnect
End Sub

Private Sub Command1_Click()
    If CmbCombinacion.ListIndex = -1 Then
        xCod_Comb = ""
    Else
        xCod_Comb = Right(CmbCombinacion.Text, 3)
    End If
    
    bOk = True
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
