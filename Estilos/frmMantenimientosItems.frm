VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmMantenimientosItems 
   Caption         =   " Mantenimientos Items"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   LinkTopic       =   "frmMantenimientosItems"
   ScaleHeight     =   8910
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraBuscar 
      Caption         =   "Buscar Por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4050
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   11505
      Begin VB.TextBox txtNom_Cliente 
         Height          =   285
         Left            =   3270
         TabIndex        =   26
         Top             =   750
         Width           =   1695
      End
      Begin VB.TextBox txtNom_TemCli 
         Height          =   285
         Left            =   6990
         TabIndex        =   25
         Top             =   750
         Width           =   1455
      End
      Begin VB.CommandButton cmdBusTemporada 
         Caption         =   "..."
         Height          =   330
         Left            =   6630
         TabIndex        =   24
         Top             =   720
         Width           =   360
      End
      Begin VB.TextBox txttemporada 
         Height          =   285
         Left            =   5910
         TabIndex        =   23
         Top             =   750
         Width           =   735
      End
      Begin VB.TextBox txtcliente 
         Height          =   285
         Left            =   2190
         MaxLength       =   5
         TabIndex        =   22
         Top             =   750
         Width           =   765
      End
      Begin VB.CommandButton cmdBusCliente 
         Caption         =   "..."
         Height          =   330
         Left            =   2910
         TabIndex        =   21
         Tag             =   "..."
         Top             =   720
         Width           =   360
      End
      Begin VB.OptionButton optcliente 
         Caption         =   "Cliente"
         Height          =   300
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   1065
      End
      Begin VB.Frame Fraitem 
         Height          =   640
         Left            =   1200
         TabIndex        =   15
         Top             =   2880
         Width           =   7455
         Begin VB.TextBox txtcod_item 
            Height          =   285
            Left            =   1590
            MaxLength       =   8
            TabIndex        =   18
            Top             =   270
            Width           =   1005
         End
         Begin VB.TextBox txtdes_item 
            Height          =   285
            Left            =   2880
            TabIndex        =   17
            Top             =   240
            Width           =   4200
         End
         Begin VB.CommandButton cmdBusItem 
            Caption         =   "..."
            Height          =   330
            Left            =   2520
            TabIndex        =   16
            Tag             =   "..."
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label2 
            Caption         =   "Item"
            Height          =   240
            Left            =   360
            TabIndex        =   19
            Top             =   330
            Width           =   690
         End
      End
      Begin VB.Frame Fracliente 
         Height          =   885
         Left            =   1200
         TabIndex        =   14
         Top             =   1920
         Width           =   7455
      End
      Begin VB.Frame Frafamilia 
         Height          =   885
         Left            =   960
         TabIndex        =   5
         Top             =   1200
         Width           =   7575
         Begin VB.TextBox txtgrupo 
            Height          =   285
            Left            =   4605
            TabIndex        =   11
            Top             =   270
            Width           =   735
         End
         Begin VB.CommandButton cmdBusgrupo 
            Caption         =   "..."
            Height          =   330
            Left            =   5325
            TabIndex        =   10
            Top             =   240
            Width           =   360
         End
         Begin VB.TextBox txtfamilia 
            Height          =   285
            Left            =   1605
            MaxLength       =   2
            TabIndex        =   9
            Top             =   240
            Width           =   525
         End
         Begin VB.CommandButton cmdBusFamItem 
            Caption         =   "..."
            Height          =   330
            Left            =   2085
            TabIndex        =   8
            Tag             =   "..."
            Top             =   240
            Width           =   360
         End
         Begin VB.TextBox txtdes_famitem 
            Height          =   285
            Left            =   2400
            TabIndex        =   7
            Top             =   270
            Width           =   1575
         End
         Begin VB.TextBox txtdes_famgruite 
            Height          =   285
            Left            =   5680
            MaxLength       =   50
            TabIndex        =   6
            Top             =   270
            Width           =   1695
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Grupo"
            Height          =   195
            Left            =   4125
            TabIndex        =   13
            Top             =   345
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Familia de Item"
            Height          =   195
            Left            =   360
            TabIndex        =   12
            Top             =   315
            Width           =   1050
         End
      End
      Begin VB.Frame fraoptions 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   120
         Width           =   6135
         Begin VB.OptionButton optitem 
            Caption         =   "Item"
            Height          =   300
            Left            =   2760
            TabIndex        =   3
            Top             =   0
            Width           =   1425
         End
         Begin VB.OptionButton optfamitem 
            Caption         =   "Familia de Item"
            Height          =   330
            Left            =   480
            TabIndex        =   2
            Top             =   0
            Value           =   -1  'True
            Width           =   1550
         End
      End
      Begin FunctionsButtons.FunctButt FunctBuscar 
         Height          =   495
         Left            =   8880
         TabIndex        =   4
         Top             =   360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~True~True~&Buscar~0~0~1~~0~False~False~&Buscar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Temporada"
         Height          =   195
         Left            =   5070
         TabIndex        =   28
         Top             =   780
         Width           =   810
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   1440
         TabIndex        =   27
         Top             =   840
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmMantenimientosItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
