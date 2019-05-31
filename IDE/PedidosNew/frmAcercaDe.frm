VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Object = "{7B0D986D-3A03-4634-828F-D16994E0941A}#1.0#0"; "ECNVB6WINCTRL.ocx"
Begin VB.Form frmAcercaDe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de Sistema de Gestión Textil"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   Icon            =   "frmAcercaDe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAcercaDe.frx":058A
   ScaleHeight     =   3720
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ChamaleonButton.ChameleonBtn cmdAceptar 
      Height          =   435
      Left            =   5100
      TabIndex        =   4
      ToolTipText     =   "Aceptar"
      Top             =   3240
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "&Aceptar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAcercaDe.frx":3084B
      PICN            =   "frmAcercaDe.frx":30867
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00869391&
      Height          =   1305
      Left            =   1860
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmAcercaDe.frx":30E01
      Top             =   1860
      Width           =   4515
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6435
      Begin ECNVB6WINCTRL.ucLabel lblSistema 
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   503
         Caption         =   "ACERCA DE"
         ForeColor       =   0
         BackColor       =   16777215
         ShadowColor     =   6710886
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ECNVB6WINCTRL.ucLabel ucLabel1 
         Height          =   285
         Left            =   450
         TabIndex        =   3
         Top             =   630
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   503
         Caption         =   "SISTEMA DE GESTIÓN TEXTIL"
         ForeColor       =   0
         BackColor       =   16777215
         ShadowColor     =   6710886
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line1 
         X1              =   30
         X2              =   6380
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Image Image3 
         Height          =   1050
         Left            =   5070
         Picture         =   "frmAcercaDe.frx":30F23
         Top             =   0
         Width           =   1320
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Un sistema informático creado para obtener velocidad, facilidad de uso y seguridad al realizar procesos textiles."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00474747&
      Height          =   465
      Left            =   1860
      TabIndex        =   5
      Top             =   1290
      Width           =   4515
   End
   Begin VB.Image Image2 
      Height          =   1695
      Left            =   60
      Picture         =   "frmAcercaDe.frx":32144
      Top             =   1290
      Width           =   1755
   End
End
Attribute VB_Name = "frmAcercaDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    Unload Me
End Sub
