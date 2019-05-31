VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pictureMenuAccesos 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00B4C6C3&
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
      ForeColor       =   &H80000008&
      Height          =   7740
      Left            =   0
      ScaleHeight     =   516
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   0
      Top             =   0
      Width           =   3555
      Begin VB.TextBox txtBuscarOpcion 
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
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Left            =   60
         TabIndex        =   1
         Text            =   " Ingrese descripción de la opción a buscar"
         ToolTipText     =   "Ingrese descripción de la opción y luego hacer clic en el ícono de buscar."
         Top             =   60
         Width           =   3135
      End
      Begin MSComctlLib.TreeView tvMenuAccesos 
         Height          =   5835
         Left            =   60
         TabIndex        =   2
         Top             =   360
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   10292
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TreeView tvMenuAccesosBusqueda 
         Height          =   5835
         Left            =   60
         TabIndex        =   3
         Top             =   6330
         Visible         =   0   'False
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   10292
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   30
         Top             =   30
         Width           =   3495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Left            =   3270
         Picture         =   "frmTest.frx":0000
         ToolTipText     =   "Use la caja de texto para buscar una opción del sistema."
         Top             =   60
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList imageListFondo 
      Left            =   4650
      Top             =   3270
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   563
      ImageHeight     =   550
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":37CD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":A2032
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":149B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":2281B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":333ED3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":45D18D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   4650
      Top             =   2310
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      Tools           =   "frmTest.frx":6D1DC5
      ToolBars        =   "frmTest.frx":6D1DDD
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3750
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6D1DF5
            Key             =   "mancli"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6D1E53
            Key             =   "manfab"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6D1EB1
            Key             =   "manOrg"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6D1F0F
            Key             =   "mantra"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6D1F6D
            Key             =   "mancomisin"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6D1FCB
            Key             =   "manBan"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6D2029
            Key             =   "mandestino"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6D2087
            Key             =   "mantippre"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":6D20E5
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   5250
      Top             =   2250
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   4800
      Top             =   1530
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
