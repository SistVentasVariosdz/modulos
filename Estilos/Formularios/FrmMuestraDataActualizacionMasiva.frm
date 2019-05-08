VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmMuestraDataActualizacionMasiva 
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Aceptar"
      Height          =   540
      Left            =   2835
      TabIndex        =   1
      Top             =   2940
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Height          =   2850
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3795
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2550
         Left            =   105
         TabIndex        =   2
         Top             =   210
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   4498
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   17
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmMuestraDataActualizacionMasiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command2_Click()
Unload Me
End Sub

Sub formato_grid()
DGridLista.Columns(0).Width = "1100"
DGridLista.Columns(1).Width = "2100"


End Sub

