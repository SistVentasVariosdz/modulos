VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{196346A1-12A8-4652-B4FB-010B924E2704}#2.0#0"; "prjKEXPCheck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComCtl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0F0877EF-2A93-4AE6-8BA8-4129832C32C3}#230.0#0"; "SmartMenuXP.ocx"
Object = "*\A..\..\ECNVB6WINCTRL\ECNVB6WINCTRL.vbp"
Begin VB.Form frm002_ECNSQLQRYDESIGN 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diseñador de Consultas SQL"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ECNSQLQRYDESIGN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   19635
   StartUpPosition =   2  'CenterScreen
   Begin ECNVB6WINCTRL.ucScrollContainer ecnScrollContainerDg 
      Height          =   8055
      Left            =   0
      TabIndex        =   23
      Top             =   360
      Width           =   6375
      _extentx        =   11245
      _extenty        =   14208
      backcolor       =   8421504
      Begin VB.PictureBox pctDiseño 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7755
         Left            =   90
         ScaleHeight     =   7755
         ScaleWidth      =   6045
         TabIndex        =   24
         Top             =   150
         Width           =   6045
         Begin VB.PictureBox pctTabla 
            Appearance      =   0  'Flat
            BackColor       =   &H00F2A240&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   360
            Picture         =   "ECNSQLQRYDESIGN.frx":2372
            ScaleHeight     =   225
            ScaleWidth      =   255
            TabIndex        =   26
            Top             =   4140
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton cmdTabla 
            BackColor       =   &H00C0C0C0&
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   9
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   25
            Tag             =   "0"
            Top             =   3780
            Visible         =   0   'False
            Width           =   285
         End
         Begin prjKEXPCheck.KEXPCheck chkTabla 
            Height          =   315
            Index           =   0
            Left            =   360
            TabIndex        =   27
            Top             =   3390
            Visible         =   0   'False
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
            CaptionBackColor=   -2147483633
            Caption         =   "* (Todas las Columnas)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckStyle      =   1
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshTabla 
            DragIcon        =   "ECNSQLQRYDESIGN.frx":40C8
            Height          =   2865
            Index           =   0
            Left            =   120
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   5054
            _Version        =   393216
            Rows            =   4
            FixedRows       =   3
            FixedCols       =   0
            BackColorFixed  =   15901248
            BackColorBkg    =   16777215
            GridColor       =   16777215
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Line linRelacionPFK 
            BorderColor     =   &H00404040&
            BorderStyle     =   3  'Dot
            Index           =   0
            Visible         =   0   'False
            X1              =   1410
            X2              =   3150
            Y1              =   4110
            Y2              =   4110
         End
         Begin VB.Image imgRelacion_PK 
            Height          =   200
            Index           =   0
            Left            =   1230
            Picture         =   "ECNSQLQRYDESIGN.frx":686A
            Stretch         =   -1  'True
            Top             =   3990
            Visible         =   0   'False
            Width           =   200
         End
         Begin VB.Image imgRelacion_FK 
            Height          =   195
            Index           =   0
            Left            =   3150
            Picture         =   "ECNSQLQRYDESIGN.frx":68F8
            Stretch         =   -1  'True
            Top             =   3990
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblTabla 
            AutoSize        =   -1  'True
            Caption         =   "LBL : NOMBRE DE TABLA MAS ALEAS PARA EL TRATAMIENTO DE REPETICIONES DE TABLAS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   30
            TabIndex        =   29
            Tag             =   "0"
            Top             =   3150
            Visible         =   0   'False
            Width           =   7365
         End
      End
   End
   Begin MSComctlLib.ImageList imgL2 
      Left            =   3960
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECNSQLQRYDESIGN.frx":69A4
            Key             =   ""
            Object.Tag             =   "SORT_ASC"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECNSQLQRYDESIGN.frx":86F7
            Key             =   ""
            Object.Tag             =   "SORT_DESC"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECNSQLQRYDESIGN.frx":A449
            Key             =   ""
            Object.Tag             =   "COLUMN"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECNSQLQRYDESIGN.frx":C1A7
            Key             =   ""
            Object.Tag             =   "TEXT"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECNSQLQRYDESIGN.frx":DEFD
            Key             =   ""
            Object.Tag             =   "FX"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECNSQLQRYDESIGN.frx":FC4E
            Key             =   ""
            Object.Tag             =   "ZIGMA"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECNSQLQRYDESIGN.frx":119AF
            Key             =   ""
            Object.Tag             =   "SQL_LEFT"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECNSQLQRYDESIGN.frx":136FB
            Key             =   ""
            Object.Tag             =   "SQL_IINNER"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECNSQLQRYDESIGN.frx":15447
            Key             =   ""
            Object.Tag             =   "SQL_RIGHT"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECNSQLQRYDESIGN.frx":17194
            Key             =   ""
            Object.Tag             =   "SQL_FROM"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECNSQLQRYDESIGN.frx":18EC8
            Key             =   ""
            Object.Tag             =   "UNION"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab sstConfiguraciones 
      Height          =   8070
      Left            =   6360
      TabIndex        =   0
      Top             =   360
      Width           =   13290
      _ExtentX        =   23442
      _ExtentY        =   14235
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   564
      TabMaxWidth     =   4586
      BackColor       =   14737632
      TabCaption(0)   =   " Configuracion de Campos"
      TabPicture(0)   =   "ECNSQLQRYDESIGN.frx":1AC19
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "mshDiseñoDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cboTipSQLJoin(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cboTipoHAVING(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dcCampo(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboTipCampo(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboTipOrden(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboOperWHERE(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboOperHAVING(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboTipoWHERE(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "dcTabla(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "pctMSHCab(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "pctMSHCab(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "pctMSHCab(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "pctMSHCab(3)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "pctMSHCab(4)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "pctMSHCab(5)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TxT(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cboOrdenDeCampos(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "pctDiseñoDatos"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   " Script SQL resultado"
      TabPicture(1)   =   "ECNSQLQRYDESIGN.frx":1AD73
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "rtfDiseñoQuerySQL"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox pctDiseñoDatos 
         BackColor       =   &H00FFFFFF&
         Height          =   2355
         Left            =   810
         ScaleHeight     =   2295
         ScaleWidth      =   11475
         TabIndex        =   20
         Top             =   3240
         Visible         =   0   'False
         Width           =   11535
         Begin ECNVB6WINCTRL.ucProgressCircular ecnPbrCir_Wait 
            Height          =   735
            Left            =   5850
            TabIndex        =   21
            Top             =   1350
            Width           =   735
            _extentx        =   1296
            _extenty        =   1296
            forecolor       =   15567120
            backcolor       =   16777215
            linestart       =   1
            lineend         =   1
            interval        =   0
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshFinta 
            Height          =   1575
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   11490
            _ExtentX        =   20267
            _ExtentY        =   2778
            _Version        =   393216
            Rows            =   32
            Cols            =   24
            FixedRows       =   2
            BackColorFixed  =   14737632
            BackColorBkg    =   16777215
            GridColor       =   15987699
            WordWrap        =   -1  'True
            TextStyleFixed  =   2
            ScrollBars      =   2
            AllowUserResizing=   3
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   24
         End
      End
      Begin VB.ComboBox cboOrdenDeCampos 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         ItemData        =   "ECNSQLQRYDESIGN.frx":1AECD
         Left            =   5010
         List            =   "ECNSQLQRYDESIGN.frx":1AEEC
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   6180
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.TextBox TxT 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5EAD6&
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   0
         Left            =   3090
         TabIndex        =   17
         Top             =   6180
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.PictureBox pctMSHCab 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   5
         Left            =   9060
         Picture         =   "ECNSQLQRYDESIGN.frx":1AF1A
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   16
         Tag             =   "ORDER"
         Top             =   6690
         Width           =   255
      End
      Begin VB.PictureBox pctMSHCab 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   8760
         Picture         =   "ECNSQLQRYDESIGN.frx":1B0AE
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   15
         Tag             =   "HAVING"
         Top             =   6690
         Width           =   255
      End
      Begin VB.PictureBox pctMSHCab 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   8460
         Picture         =   "ECNSQLQRYDESIGN.frx":1B481
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   14
         Tag             =   "GROUP"
         Top             =   6690
         Width           =   255
      End
      Begin VB.PictureBox pctMSHCab 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   8160
         Picture         =   "ECNSQLQRYDESIGN.frx":1B6D0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   13
         Tag             =   "WHERE"
         Top             =   6690
         Width           =   255
      End
      Begin VB.PictureBox pctMSHCab 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   7830
         Picture         =   "ECNSQLQRYDESIGN.frx":1BAA3
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   12
         Tag             =   "CAMPO"
         Top             =   6690
         Width           =   255
      End
      Begin VB.PictureBox pctMSHCab 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   7530
         Picture         =   "ECNSQLQRYDESIGN.frx":1BD01
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   11
         Tag             =   "TABLA"
         Top             =   6690
         Width           =   255
      End
      Begin MSDataListLib.DataCombo dcTabla 
         Height          =   285
         Index           =   0
         Left            =   210
         TabIndex        =   9
         Top             =   6630
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cboTipoWHERE 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Index           =   0
         ItemData        =   "ECNSQLQRYDESIGN.frx":1BE4B
         Left            =   5010
         List            =   "ECNSQLQRYDESIGN.frx":1BE55
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   6630
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.ComboBox cboOperHAVING 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         ItemData        =   "ECNSQLQRYDESIGN.frx":1BE62
         Left            =   3120
         List            =   "ECNSQLQRYDESIGN.frx":1BE81
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   7020
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.ComboBox cboOperWHERE 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         ItemData        =   "ECNSQLQRYDESIGN.frx":1BEAF
         Left            =   3120
         List            =   "ECNSQLQRYDESIGN.frx":1BECE
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   6630
         Visible         =   0   'False
         Width           =   1725
      End
      Begin MSComctlLib.ImageCombo cboTipOrden 
         Height          =   330
         Index           =   0
         Left            =   1710
         TabIndex        =   3
         Top             =   6630
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtfDiseñoQuerySQL 
         Height          =   7995
         Left            =   -74970
         TabIndex        =   2
         Top             =   30
         Width           =   12915
         _ExtentX        =   22781
         _ExtentY        =   14102
         _Version        =   393217
         TextRTF         =   $"ECNSQLQRYDESIGN.frx":1BEFC
      End
      Begin MSComctlLib.ImageCombo cboTipCampo 
         Height          =   330
         Index           =   0
         Left            =   1710
         TabIndex        =   4
         Top             =   7020
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcCampo 
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   7020
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cboTipoHAVING 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Index           =   0
         ItemData        =   "ECNSQLQRYDESIGN.frx":1BF78
         Left            =   5010
         List            =   "ECNSQLQRYDESIGN.frx":1BF82
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   6960
         Visible         =   0   'False
         Width           =   1665
      End
      Begin MSComctlLib.ImageCombo cboTipSQLJoin 
         Height          =   330
         Index           =   0
         Left            =   1710
         TabIndex        =   19
         Top             =   6180
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDiseñoDatos 
         Height          =   3015
         Left            =   60
         TabIndex        =   1
         Top             =   30
         Width           =   13200
         _ExtentX        =   23283
         _ExtentY        =   5318
         _Version        =   393216
         Rows            =   32
         Cols            =   24
         FixedRows       =   2
         BackColorFixed  =   14737632
         BackColorBkg    =   16777215
         GridColor       =   15987699
         WordWrap        =   -1  'True
         TextStyleFixed  =   2
         ScrollBars      =   2
         AllowUserResizing=   3
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   24
      End
   End
   Begin VBSmartXPMenu.SmartMenuXP mnPopPup_DgGrafico 
      Height          =   375
      Left            =   8460
      Top             =   30
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VBSmartXPMenu.SmartMenuXP mnPopPup_Tabla 
      Height          =   375
      Left            =   11730
      Top             =   30
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VBSmartXPMenu.SmartMenuXP MNMAIN 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BackColor       =   14737632
      SelBackColor    =   14737632
      BorderStyle     =   3
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
   Begin VBSmartXPMenu.SmartMenuXP mnPopPup_DgMatricial 
      Height          =   375
      Left            =   9420
      Top             =   30
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgL3 
      Left            =   5340
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECNSQLQRYDESIGN.frx":1BF8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECNSQLQRYDESIGN.frx":1C329
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECNSQLQRYDESIGN.frx":1C6C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECNSQLQRYDESIGN.frx":1CA5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECNSQLQRYDESIGN.frx":1CDF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECNSQLQRYDESIGN.frx":1D191
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECNSQLQRYDESIGN.frx":1D52B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECNSQLQRYDESIGN.frx":1D8C5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imgL 
      Left            =   2370
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   24
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":1DC5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":1DD61
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":1E283
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":1E545
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":1E807
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":1EB59
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":1EEAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":1F1FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":1F54F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":1F6F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":1FA43
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":1FD65
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":200B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":20409
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":2075B
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":20AAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":20DFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":21151
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":214A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":217F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":21B47
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":21E99
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":221EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNSQLQRYDESIGN.frx":2253D
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm002_ECNSQLQRYDESIGN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public PU_SW_DISEÑO_CON_TABLAS As Boolean
Public PU_SW_DISEÑO_CON_RELACIONES As Boolean

Private Const Kdbl_COLOR_TABLA As Double = &HF2A240
Private Const Kdbl_COLOR_VISTA As Double = &H2CC561

Private Const Ki_FRM_WIDTH_MIN As Integer = 19815
Private Const Ki_FRM_HEIGHT_MIN As Integer = 9000
Private Const Ki_FRM_TOP_MIN As Integer = 0
Private Const Ki_FRM_LEFT_MIN As Integer = 0

Private Const Ks_CBO_TABLA_NULL_COD As String = "NULL"
Private Const Ks_CBO_TABLA_NULL_DES As String = "{Sin Tabla}"

Private Const Ki_MSH_CAMPOS_00_FIXED__COL As Integer = 0
Private Const Ki_MSH_CAMPOS_01_FIXED__COL As Integer = 1
Private Const Ki_MSH_CAMPOS_02_TABLA_NOMB As Integer = 2
Private Const Ki_MSH_CAMPOS_03_TABLA_JOIN As Integer = 3
Private Const Ki_MSH_CAMPOS_04_TABLA_CODI As Integer = 4
Private Const Ki_MSH_CAMPOS_04_BARRA_0001 As Integer = 4
Private Const Ki_MSH_CAMPOS_05_SELEC_TIPO As Integer = 5
Private Const Ki_MSH_CAMPOS_06_SELEC_NOMB As Integer = 6
Private Const Ki_MSH_CAMPOS_07_SELEC_ALEA As Integer = 7
Private Const Ki_MSH_CAMPOS_08_SELEC_ACTI As Integer = 8
Private Const Ki_MSH_CAMPOS_09_SELEC_ACTV As Integer = 9
Private Const Ki_MSH_CAMPOS_10_BARRA_0002 As Integer = 10
Private Const Ki_MSH_CAMPOS_11_WHERE_CHKI As Integer = 11
Private Const Ki_MSH_CAMPOS_12_WHERE_CHKV As Integer = 12
Private Const Ki_MSH_CAMPOS_13_WHERE_TIPO As Integer = 13
Private Const Ki_MSH_CAMPOS_14_WHERE_OPER As Integer = 14
Private Const Ki_MSH_CAMPOS_15_WHERE_CRI1 As Integer = 15
Private Const Ki_MSH_CAMPOS_16_WHERE_CRI2 As Integer = 16
Private Const Ki_MSH_CAMPOS_17_BARRA_0003 As Integer = 17
Private Const Ki_MSH_CAMPOS_18_GROUP_CHKI As Integer = 18
Private Const Ki_MSH_CAMPOS_19_GROUP_CHKV As Integer = 19
Private Const Ki_MSH_CAMPOS_20_GROUP_NUME As Integer = 20
Private Const Ki_MSH_CAMPOS_21_BARRA_0004 As Integer = 21
Private Const Ki_MSH_CAMPOS_22_HAVIN_CHKI As Integer = 22
Private Const Ki_MSH_CAMPOS_23_HAVIN_CHKV As Integer = 23
Private Const Ki_MSH_CAMPOS_24_HAVIN_TIPO As Integer = 24
Private Const Ki_MSH_CAMPOS_25_HAVIN_OPER As Integer = 25
Private Const Ki_MSH_CAMPOS_26_HAVIN_CRI1 As Integer = 26
Private Const Ki_MSH_CAMPOS_27_HAVIN_CRI2 As Integer = 27
Private Const Ki_MSH_CAMPOS_28_BARRA_0005 As Integer = 28
Private Const Ki_MSH_CAMPOS_29_ORDER_CHKI As Integer = 29
Private Const Ki_MSH_CAMPOS_30_ORDER_CHKV As Integer = 30
Private Const Ki_MSH_CAMPOS_31_ORDER_NUME As Integer = 31
Private Const Ki_MSH_CAMPOS_32_ORDER_TIPO As Integer = 32
Private Const Ki_MSH_CAMPOS_33_BARRA_0006 As Integer = 33

Private Const Ki_MSH_CAMPOS_11_WHERE_CHKI_WIDTH As Integer = 300
Private Const Ki_MSH_CAMPOS_12_WHERE_CHKV_WIDTH As Integer = 0
Private Const Ki_MSH_CAMPOS_13_WHERE_TIPO_WIDTH As Integer = 400
Private Const Ki_MSH_CAMPOS_14_WHERE_OPER_WIDTH As Integer = 850
Private Const Ki_MSH_CAMPOS_15_WHERE_CRI1_WIDTH As Integer = 850
Private Const Ki_MSH_CAMPOS_16_WHERE_CRI2_WIDTH As Integer = 850

Private Const Ki_MSH_CAMPOS_18_GROUP_CHKI_WIDTH As Integer = 300
Private Const Ki_MSH_CAMPOS_19_GROUP_CHKV_WIDTH As Integer = 0
Private Const Ki_MSH_CAMPOS_20_GROUP_NUME_WIDTH As Integer = 300

Private Const Ki_MSH_CAMPOS_22_HAVIN_CHKI_WIDTH As Integer = 300
Private Const Ki_MSH_CAMPOS_23_HAVIN_CHKV_WIDTH As Integer = 0
Private Const Ki_MSH_CAMPOS_24_HAVIN_TIPO_WIDTH As Integer = 400
Private Const Ki_MSH_CAMPOS_25_HAVIN_OPER_WIDTH As Integer = 850
Private Const Ki_MSH_CAMPOS_26_HAVIN_CRI1_WIDTH As Integer = 800
Private Const Ki_MSH_CAMPOS_27_HAVIN_CRI2_WIDTH As Integer = 800

Private Const Ki_TABLA_Col_Codigo As Integer = 0
Private Const Ki_TABLA_Col_Descripcion As Integer = 1
Private Const Ki_TABLA_Col_IsPK As Integer = 2
Private Const Ki_TABLA_Col_IsUK As Integer = 3
Private Const Ki_TABLA_Col_IsFK As Integer = 4
Private Const Ki_TABLA_Col_CheckIco As Integer = 5
Private Const Ki_TABLA_Col_CheckVal As Integer = 6

Private Const Ki_Ico_UnChk As Integer = 1
Private Const Ki_Ico_Check As Integer = 2
Private Const Ki_Ico_Tabla As Integer = 3
Private Const Ki_Ico_Vista As Integer = 4
Private Const Ki_Ico_ColPK As Integer = 5
Private Const Ki_Ico_ColFK As Integer = 6
Private Const Ki_Ico_ColUK As Integer = 7
Private Const Ki_Ico_TableAdd As Integer = 8
Private Const Ki_Ico_TableMin As Integer = 9
Private Const Ki_Ico_TableDel As Integer = 10
Private Const Ki_Ico_DiagramDel As Integer = 11
Private Const Ki_Ico_TableMax As Integer = 12
Private Const Ki_Ico_Tick As Integer = 13
Private Const Ki_Ico_TableSave As Integer = 14
Private Const Ki_Ico_TableInsert As Integer = 15
Private Const Ki_Ico_TableDelete As Integer = 16
Private Const Ki_Ico_TableReplace As Integer = 17
Private Const Ki_Ico_TableSort As Integer = 18
Private Const Ki_Ico_TablePaint As Integer = 19
Private Const Ki_Ico_Comentarios As Integer = 20
Private Const Ki_Ico_AlignTop As Integer = 21
Private Const Ki_Ico_AlignBottom As Integer = 22
Private Const Ki_Ico_BorderTop As Integer = 23
Private Const Ki_Ico_BorderBottom As Integer = 24

Private Const Ki_Ico_ImgL2_SortAsc As Integer = 1
Private Const Ki_Ico_ImgL2_SortDes As Integer = 2
Private Const Ki_Ico_ImgL2_Column As Integer = 3
Private Const Ki_Ico_ImgL2_Text As Integer = 4
Private Const Ki_Ico_ImgL2_Fx As Integer = 5
Private Const Ki_Ico_ImgL2_Zigma As Integer = 6
Private Const Ki_Ico_ImgL2_LeftJoin As Integer = 7
Private Const Ki_Ico_ImgL2_InnerJoin As Integer = 8
Private Const Ki_Ico_ImgL2_RightJoin As Integer = 9
Private Const Ki_Ico_ImgL2_From As Integer = 10
Private Const Ki_Ico_ImgL2_Union As Integer = 11

Private Const Ki_Ico_ImgL3_Rotate1 As Integer = 1
Private Const Ki_Ico_ImgL3_Rotate2 As Integer = 2
Private Const Ki_Ico_ImgL3_Rotate3 As Integer = 3
Private Const Ki_Ico_ImgL3_Rotate4 As Integer = 4
Private Const Ki_Ico_ImgL3_Rotate5 As Integer = 5
Private Const Ki_Ico_ImgL3_Rotate6 As Integer = 6
Private Const Ki_Ico_ImgL3_Rotate7 As Integer = 7
Private Const Ki_Ico_ImgL3_Rotate8 As Integer = 8

Private Const Ks_CarPK As String * 1 = "P"
Private Const Ks_CarFK As String * 1 = "F"
Private Const Ks_CarSepPKFK As String * 1 = ":"

Private Const Kdbl_Espacio_Entre_Tablas As Double = 800
Private Const Ki_MargenDragKeyPermitido As Integer = 20

Private Const Ki_TABLA_ROWHEIGHT_FX1 As Integer = 350
Private Const Ki_TABLA_ROWHEIGHT_FX2 As Integer = 380
Private Const Ki_TABLA_ROWHEIGHT_FX3 As Integer = 50
Private Const Ki_TABLA_ROWHEIGHT As Integer = 300
Private Const Ki_TABLA_ROWS_MAX As Integer = 7
Private Const Ki_TABLA_LEFT_INICIAL As Integer = 150
Private Const Ki_TABLA_HEIGHT_MIN As Integer = 1005
Private Const Ki_TABLA_HEIGHT_MAX As Integer = 2865

Private Const Ki_MNPOPU_PCTDISEÑO_ID2_ADDTABLE = 2
Private Const Ki_MNPOPU_PCTDISEÑO_ID4_DELTBALL = 4
Private Const Ki_MNPOPU_PCTDISEÑO_ID5_MINTABLE = 5
Private Const Ki_MNPOPU_PCTDISEÑO_ID6_MAXTABLE = 6

Private Const Ki_MNPOPU_MSHTABLA_ID2_DELTABLE = 2

Private Const Ks_MNMAIN_00_01___ As String = "DIAGRAMA"
Private Const Ks_MNMAIN_00_01_01 As String = "ADDTABLE"
Private Const Ks_MNMAIN_00_01_02 As String = "DELETALL"
Private Const Ks_MNMAIN_00_01_03 As String = "MINIMALL"
Private Const Ks_MNMAIN_00_01_04 As String = "MAXIMALL"

Private Const Ks_MNMAIN_00_02___ As String = "GRILLA"
Private Const Ks_MNMAIN_00_02_01 As String = "SEC_WHERE"
Private Const Ks_MNMAIN_00_02_02 As String = "SEC_GROUP_BY"
Private Const Ks_MNMAIN_00_02_03 As String = "SEC_HAVING"

Private Const Ks_MNMAIN_00_02_04 As String = "TABLE_SAVE"
Private Const Ks_MNMAIN_00_02_05 As String = "TABLE_INSERT"
Private Const Ks_MNMAIN_00_02_05_01 As String = "TABLE_INSERT_INICIO"
Private Const Ks_MNMAIN_00_02_05_02 As String = "TABLE_INSERT_ARRIBA_DE_FILA"
Private Const Ks_MNMAIN_00_02_05_03 As String = "TABLE_INSERT_DEBAJO_DE_FILA"
Private Const Ks_MNMAIN_00_02_05_04 As String = "TABLE_INSERT_FINAL"
Private Const Ks_MNMAIN_00_02_06 As String = "TABLE_DELETE"
Private Const Ks_MNMAIN_00_02_07 As String = "TABLE_REPLACE"
Private Const Ks_MNMAIN_00_02_08 As String = "TABLE_SORT"
Private Const Ks_MNMAIN_00_02_09 As String = "TABLE_ERASE"

Private Const Ks_MNMAIN_00_03___ As String = "TABLA"
Private Const Ks_MNMAIN_00_03_01 As String = "DELTABLE"

Private Const Ks_MNMAIN_00_04___ As String = "SALIR"

Private oECNLIB01_FUNSUB As New ECNVB6LIB.ECNLIB01_FUNSUB
Private oECNLIB03_WINEVE As New ECNVB6LIB.ECNLIB03_WINEVE

Private iValorWindowState As Integer
Private iIndiceTablaSEL As Integer

Private sngDragX_Tabla As Single
Private sngDragY_Tabla As Single
Private iDragInd_Tabla As Integer

Private sngDragX_ImgPK As Single
Private sngDragY_ImgPK As Single
Private iDragInd_ImgPK As Integer

Private sngDragX_ImgFK As Single
Private sngDragY_ImgFK As Single
Private iDragInd_ImgFK As Integer

Private Const Ki_Vector_ValorNULL As Integer = -1

Private IND_VECTOR_DE_TABLAS As Integer
Private IND_VECTOR_DE_RELACIONES As Integer

Private blSW_Load As Boolean

Private aVectorDeIndicesDeTablasDelDiagrama() As Integer
Private aVectorDeIndicesRelacionesDelDiagrama() As Integer

Private objECN_HandlerFlatCboTipSJOIN  As New ECNVB6LIB.ECNLIB03_WINEVE_FLAT_CTRL
Private objECN_HandlerFlatCboTipCAMPO  As New ECNVB6LIB.ECNLIB03_WINEVE_FLAT_CTRL
Private objECN_HandlerFlatCboTipWHERE  As New ECNVB6LIB.ECNLIB03_WINEVE_FLAT_CTRL
Private objECN_HandlerFlatCboOpeWHERE  As New ECNVB6LIB.ECNLIB03_WINEVE_FLAT_CTRL
Private objECN_HandlerFlatCboTipHAVIN  As New ECNVB6LIB.ECNLIB03_WINEVE_FLAT_CTRL
Private objECN_HandlerFlatCboOpeHAVIN  As New ECNVB6LIB.ECNLIB03_WINEVE_FLAT_CTRL
Private objECN_HandlerFlatCboTipORDER  As New ECNVB6LIB.ECNLIB03_WINEVE_FLAT_CTRL
Private objECN_HandlerFlatCboOrdCampo  As New ECNVB6LIB.ECNLIB03_WINEVE_FLAT_CTRL

Private objBallonTooT_MSHDiseño As New ECNVB6LIB.ECNLIB03_WINEVE_TOOT_BALLOON
Private objBallonTooT_PctMSHCabTabla As New ECNVB6LIB.ECNLIB03_WINEVE_TOOT_BALLOON
Private objBallonTooT_PctMSHCabCampo As New ECNVB6LIB.ECNLIB03_WINEVE_TOOT_BALLOON
Private objBallonTooT_PctMSHCabWhere As New ECNVB6LIB.ECNLIB03_WINEVE_TOOT_BALLOON
Private objBallonTooT_PctMSHCabGroup As New ECNVB6LIB.ECNLIB03_WINEVE_TOOT_BALLOON
Private objBallonTooT_PctMSHCabHavin As New ECNVB6LIB.ECNLIB03_WINEVE_TOOT_BALLOON
Private objBallonTooT_PctMSHCabOrder As New ECNVB6LIB.ECNLIB03_WINEVE_TOOT_BALLOON

Private iCOL_MOUSE_MOVE_MSH_DISEÑO As Integer

Private Sub Form_Load()
    blSW_Load = True
    Call oECNLIB01_FUNSUB.AsignarValoresDeUnFormulario(Me, GO_002_RUTA_INI_PARAM_WIN)
    Call ConfiguraMSHCampos
    Call CrearMenusDeLaApp
    
    IND_VECTOR_DE_TABLAS = 0
    Call RealizarEfectosECNLib
    blSW_Load = False
End Sub

Sub RealizarEfectosECNLib()
    '--+------------------------------------------------------------------------------------------------------------------------+--
    '=> BALLOON TOOL TIP EN EL MSH DEL DISEÑO
    '--+------------------------------------------------------------------------------------------------------------------------+--
    Dim sTitulo As String
    Dim sMessag As String
    
    With objBallonTooT_MSHDiseño
        Call .CreateBalloon(mshDiseñoDatos, "Diagrama Matricial", Me.Caption, ECNVB6LIB.tipiconinfo)
        .Active = False
    End With

    sTitulo = Me.Caption & " : Diagrama Matricial"
    
    With objBallonTooT_PctMSHCabTabla
        sMessag = "Sección [TABLA],Columna [TIPO]" _
                         & vbNewLine _
                         & "{<FROM>|<INNER/LEFT/RIGHT JOIN> T1,T2,T3 ...}" _
                         & vbNewLine & Chr(13) _
                         & "Aquí se indican las tablas que serán consideradas en el Query SQL. Estas tablas deben" _
                         & vbNewLine _
                         & "pertenecer al diagrama realizados. Indicar tambien la orientacion de coincidencia que" _
                         & vbNewLine _
                         & "debe tener el SELECT dentro del Query SQL." _
                         & vbNewLine & Chr(13) & "--+-----------------------------------------------------------------------------------+--" _
                         & vbNewLine & "Columnas dentro de la sección :" _
                         & vbNewLine & "--+-----------------------------------------------------------------------------------+--" _
                         & vbNewLine _
                         & "* NOMBRE : es el nombre de la tabla (previamente agregada al Diagrama de" & vbNewLine _
                         & "                     diseño), de ésta, se eligirá el campo a configurar" & vbNewLine _
                         & "* TP            : indica el tipo de union que realizará la tabla : Inner, Left ó Right" _
                         & vbNewLine & "--+-----------------------------------------------------------------------------------+--"
        Call .CreateBalloon(pctMSHCab(0), sMessag, sTitulo, 1)
    End With
    With objBallonTooT_PctMSHCabCampo
        sMessag = "Sección [CAMPO]" _
                & vbNewLine _
                & "{SELECT C1,C2,C3 ...}" _
                & vbNewLine & Chr(13) _
                & "Aquí se indican los campos que serán parte del Query SQL. Se considera el tipo, nombre y " _
                & vbNewLine _
                & "aleas. Para poder visualizar el campo como resultado éste debe ser confirmado mediante un" _
                & vbNewLine & "check." _
                & vbNewLine & Chr(13) & "--+-------------------------------------------------------------------------------------------+--" & vbNewLine _
                & "Columnas dentro de la sección" _
                & vbNewLine & "--+-------------------------------------------------------------------------------------------+--" & vbNewLine _
                & "* NOMBRE : solo habilitado para el tipo CAMPO, es un campo perteneciente de la tabla" & vbNewLine _
                & "* ALEAS     : Para los tipos FX,TXT y AGREGADO, un aleas es obligatorio" & vbNewLine _
                & "* TP            : tipo de campo, tenemos :" & vbNewLine & Chr(13) _
                & Space(10) & "- CAMPO,perteneciente a la tabla seleccionada" & vbNewLine _
                & Space(10) & "- FX, es una función, la cual puede permite cualquier sentencia T-SQL" & vbNewLine _
                & Space(10) & "- TXT, es una cadena alfanumérica" & vbNewLine _
                & Space(10) & "- AGREGADO, es una función de agregado, trabaja solo si se ha" _
                & vbNewLine & "--+-------------------------------------------------------------------------------------------+--"

        Call .CreateBalloon(pctMSHCab(1), sMessag, sTitulo, 1)
    End With
    With objBallonTooT_PctMSHCabWhere
        sMessag = "Sección [FILTRO]" _
                & vbNewLine _
                & "{WHERE Condicion 1 <and/or> Condicion_2 <and/or> ...}" _
                & vbNewLine & Chr(13) _
                & "Aquí se indican las condiciones que el campo debe cumplir dentro de la" _
                & vbNewLine _
                & "ejecución del Query SQL." _
                & vbNewLine & Chr(13) & "--+-----------------------------------------------------------------------+--" _
                & vbNewLine & "Columnas dentro de la sección" _
                & vbNewLine & "--+-----------------------------------------------------------------------+--" _
                & vbNewLine & "* .?         : indica si el campo tendrá un filtro" _
                & vbNewLine & "* TP       : tipo de conjunción para el filtro {AND}, {OR}" _
                & vbNewLine & "* OPE     : operador que se empleará para la aplicación del filtro" _
                & vbNewLine & "* CRI.01 : valor por el cual se condicionará el filtro" _
                & vbNewLine & "* CRI.02 : segundo valor, solo permitido para el operador BETWEEN" _
                & vbNewLine & "--+-------------------------------------------------------------------------+--"
        Call .CreateBalloon(pctMSHCab(2), sMessag, sTitulo, 1)
    End With
    With objBallonTooT_PctMSHCabGroup
         sMessag = "Sección [GRUPO]" _
                & vbNewLine _
                & "{GROUP BY C1,C2,C3...}" _
                & vbNewLine & Chr(13) _
                & "En esta sección se indica si el Query SQL debe realizar un agrupamiento por" _
                & vbNewLine _
                & "el campo indicado en el orden indicado." _
                & vbNewLine & Chr(13) & "--+-----------------------------------------------------------------------+--" _
                & vbNewLine & "Columnas dentro de la sección" _
                & vbNewLine & "--+-----------------------------------------------------------------------+--" _
                & vbNewLine & "* .?   : indica si el campo será un criterio de agrupamiento" _
                & vbNewLine & "* N° : indica en que orden de los agrupamientos se llevará a cabo" _
                & vbNewLine & "--+-------------------------------------------------------------------------+--"
                
        Call .CreateBalloon(pctMSHCab(3), sMessag, sTitulo, 1)
    End With
    With objBallonTooT_PctMSHCabHavin
         sMessag = "Sección [FILTRO DEL GRUPO]" _
                & vbNewLine _
                & "{HAVING Condicion 1 <and/or> Condicion_2 <and/or> ...}" _
                & vbNewLine & Chr(13) _
                & "Aquí se indican las condiciones que el campo debe cumplir dentro de los" _
                & vbNewLine _
                & "agrupamientos configurados." _
                & vbNewLine & Chr(13) & "--+-----------------------------------------------------------------------+--" _
                & vbNewLine & "Columnas dentro de la sección" _
                & vbNewLine & "--+-----------------------------------------------------------------------+--" _
                & vbNewLine & "* .?         : indica si al grupo se le aplicará un un filtro" _
                & vbNewLine & "* TP       : tipo de conjunción para el filtro {AND}, {OR}" _
                & vbNewLine & "* OPE     : operador que se empleará para la aplicación del filtro" _
                & vbNewLine & "* CRI.01 : valor por el cual se condicionará el filtro" _
                & vbNewLine & "* CRI.02 : segundo valor, solo permitido para el operador BETWEEN" _
                & vbNewLine & "--+-------------------------------------------------------------------------+--"
        Call .CreateBalloon(pctMSHCab(4), sMessag, sTitulo, 1)
    End With
    With objBallonTooT_PctMSHCabOrder
        sMessag = "Sección [ORDEN]" _
                & vbNewLine & Chr(13) _
                & "{ORDER BY C1,C2,C3...}" _
                & vbNewLine _
                & "En esta sección se indica si el Query SQL debe realizar un ordenamiento por" _
                & vbNewLine _
                & "el campo indicado en el orden indicado" _
                & vbNewLine & Chr(13) & "--+-----------------------------------------------------------------------+--" _
                & vbNewLine & "Columnas dentro de la sección" _
                & vbNewLine & "--+-----------------------------------------------------------------------+--" _
                & vbNewLine & "* .?   : indica si el campo es un criterio de ordenamiento" _
                & vbNewLine & "* N° : indica el N° de orden que lleva el campo dentro del ordenamiento" _
                & vbNewLine & "--+-------------------------------------------------------------------------+--"
        Call .CreateBalloon(pctMSHCab(5), sMessag, sTitulo, 1)
        
    End With
    '--+------------------------------------------------------------------------------------------------------------------------+--
    '=> CONTROLES CON ESTILO FLAT
    '--+------------------------------------------------------------------------------------------------------------------------+--
    objECN_HandlerFlatCboTipSJOIN.Attach cboTipSQLJoin(0)
    objECN_HandlerFlatCboTipCAMPO.Attach cboTipCampo(0)
    objECN_HandlerFlatCboTipWHERE.Attach cboTipoWHERE(0)
    objECN_HandlerFlatCboOpeWHERE.Attach cboOperWHERE(0)
    objECN_HandlerFlatCboTipHAVIN.Attach cboTipoHAVING(0)
    objECN_HandlerFlatCboOpeHAVIN.Attach cboOperHAVING(0)
    objECN_HandlerFlatCboTipORDER.Attach cboTipOrden(0)
    objECN_HandlerFlatCboOrdCampo.Attach cboOrdenDeCampos(0)
End Sub

Private Sub Form_Activate()
    iValorWindowState = Me.WindowState
    If GO_002_SW_LOAD_DESIGN Then
        Call AgregarTabla
    End If
    Call UbicaImgDeMSHCab
    iCOL_MOUSE_MOVE_MSH_DISEÑO = -1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    Me.Left = 0
  
    If Me.WindowState <> iValorWindowState Then iValorWindowState = Me.WindowState
    If Me.Width < Ki_FRM_WIDTH_MIN Then Me.Width = Ki_FRM_WIDTH_MIN
    If Me.Height < Ki_FRM_HEIGHT_MIN Then Me.Height = Ki_FRM_HEIGHT_MIN
    '--+-----------------------------------------------------------------------------------+--
    '--> ANCHO DE LA VENTANA
    '--+-----------------------------------------------------------------------------------+--
    ecnScrollContainerDg.Top = MNMAIN.Height + 1
    With sstConfiguraciones
        .Top = ecnScrollContainerDg.Top
        .Left = Me.Width - .Width - 120
    End With
    
    ecnScrollContainerDg.Width = sstConfiguraciones.Left
    mshDiseñoDatos.Top = 10
    '--+-----------------------------------------------------------------------------------+--
    '--> ALTO DE LA VENTANA
    '--+-----------------------------------------------------------------------------------+--
    ecnScrollContainerDg.Height = Me.ScaleHeight - MNMAIN.Height '+ 30
    sstConfiguraciones.Height = ecnScrollContainerDg.Height
    
    mshDiseñoDatos.Height = sstConfiguraciones.Height - 380
    rtfDiseñoQuerySQL.Height = mshDiseñoDatos.Height
    
    With pctDiseñoDatos
        .Left = mshDiseñoDatos.Left
        .Width = mshDiseñoDatos.Width - 360
         mshDiseñoDatos.Row = mshDiseñoDatos.FixedRows
        .Top = mshDiseñoDatos.Top
        .Height = mshDiseñoDatos.Height
        
        ecnPbrCir_Wait.Top = (.Height / 2) - (ecnPbrCir_Wait.Height / 2)
        ecnPbrCir_Wait.Left = (.Width / 2) - (ecnPbrCir_Wait.Width / 2)
        
        mshFinta.Width = .Width - 20
        mshFinta.Height = .Height - 40
    End With
    '--+-----------------------------------------------------------------------------------+--
    '--> SCROLLS Y AREA DE DISEÑO
    '--+-----------------------------------------------------------------------------------+--
    Select Case PU_SW_DISEÑO_CON_TABLAS
        Case True
        Case False
            With pctDiseño
                .Left = 15
                .Top = 30
                .Width = ecnScrollContainerDg.Width - 20
                .Height = ecnScrollContainerDg.Height - 20
            End With
    End Select
    Call oECNLIB01_FUNSUB.DeGradarColoRPictuRe(pctDiseño, 129, 148, 192, 170, 240, 240, 0, pctDiseño.ScaleHeight, ECNVB6LIB.eRuDeGrada_VERTICAL)
    Call UbicaControlesMSH
End Sub

Private Sub mnPopPup_DgMatricial_Click(ByVal ID As Long)
    Call EjecucionMenuDgMatricial(mnPopPup_DgMatricial, ID)
End Sub

Private Sub mshDiseñoDatos_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13, _
             vbKeySpace
            Call mshDiseñoDatos_Click
    End Select
End Sub

Private Sub mshDiseñoDatos_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    On Error Resume Next
    If Button <> vbRightButton Then Exit Sub
    With mnPopPup_DgMatricial
        .PopupMenu .MenuItems.Key2ID(Ks_MNMAIN_00_02___), _
                   .ClientToScreenX(mshDiseñoDatos.hWnd, X), _
                   .ClientToScreenY(mshDiseñoDatos.hWnd, y)
    End With
End Sub

Private Sub mshDiseñoDatos_Click()
    If PU_SW_DISEÑO_CON_TABLAS = False Then Exit Sub
    On Error Resume Next

    Dim sValor As String
    Dim iColVal As Integer

    iColVal = -1
    With mshDiseñoDatos
        Select Case .Col
            Case Ki_MSH_CAMPOS_11_WHERE_CHKI: iColVal = Ki_MSH_CAMPOS_12_WHERE_CHKV
            Case Ki_MSH_CAMPOS_18_GROUP_CHKI: iColVal = Ki_MSH_CAMPOS_19_GROUP_CHKV
            Case Ki_MSH_CAMPOS_22_HAVIN_CHKI: iColVal = Ki_MSH_CAMPOS_23_HAVIN_CHKV
            Case Ki_MSH_CAMPOS_29_ORDER_CHKI: iColVal = Ki_MSH_CAMPOS_30_ORDER_CHKV
            Case Ki_MSH_CAMPOS_08_SELEC_ACTI: iColVal = Ki_MSH_CAMPOS_09_SELEC_ACTV
            Case Else
                Exit Sub
        End Select

        If .TextMatrix(.Row, Ki_MSH_CAMPOS_04_TABLA_CODI) = Ks_CBO_TABLA_NULL_COD And _
           .Col <> Ki_MSH_CAMPOS_08_SELEC_ACTI Then
           Call SetearValoresPorTablaNull(.Row)
           .SetFocus
            Exit Sub
        End If
        
        Dim iIndICO As Integer

        Select Case iColVal
            Case -1
            Case Else
                sValor = .TextMatrix(.Row, iColVal)

                Select Case sValor
                    Case GO_ECNLIB00_CONST.VAL_UNCHK: sValor = GO_ECNLIB00_CONST.VAL_CHECK
                    Case GO_ECNLIB00_CONST.VAL_CHECK: sValor = GO_ECNLIB00_CONST.VAL_UNCHK
                End Select

                If sValor = GO_ECNLIB00_CONST.VAL_CHECK Then
                    If Len(Trim(.TextMatrix(.Row, Ki_MSH_CAMPOS_02_TABLA_NOMB))) = 0 Then _
                        Exit Sub
                End If

                .TextMatrix(.Row, iColVal) = sValor

                iIndICO = -1
                Select Case sValor
                    Case GO_ECNLIB00_CONST.VAL_UNCHK
                        Select Case iColVal
                            Case Ki_MSH_CAMPOS_09_SELEC_ACTV
                            Case Else: iIndICO = Ki_Ico_UnChk
                        End Select
                    Case GO_ECNLIB00_CONST.VAL_CHECK
                        Select Case iColVal
                            Case Ki_MSH_CAMPOS_09_SELEC_ACTV: iIndICO = Ki_Ico_Tick
                            Case Else: iIndICO = Ki_Ico_Check
                        End Select
                End Select

                If iIndICO = -1 Then Set .CellPicture = Nothing _
                Else Set .CellPicture = imgL.ListImages(iIndICO).Picture

                If sValor = GO_ECNLIB00_CONST.VAL_UNCHK Then
                    Select Case .Col
                        Case Ki_MSH_CAMPOS_11_WHERE_CHKI
                            .TextMatrix(.Row, Ki_MSH_CAMPOS_13_WHERE_TIPO) = Empty
                            .TextMatrix(.Row, Ki_MSH_CAMPOS_14_WHERE_OPER) = Empty
                            .TextMatrix(.Row, Ki_MSH_CAMPOS_15_WHERE_CRI1) = Empty
                            .TextMatrix(.Row, Ki_MSH_CAMPOS_16_WHERE_CRI2) = Empty
                        Case Ki_MSH_CAMPOS_18_GROUP_CHKI
                            .TextMatrix(.Row, Ki_MSH_CAMPOS_20_GROUP_NUME) = Empty
                        Case Ki_MSH_CAMPOS_22_HAVIN_CHKI
                            .TextMatrix(.Row, Ki_MSH_CAMPOS_24_HAVIN_TIPO) = Empty
                            .TextMatrix(.Row, Ki_MSH_CAMPOS_25_HAVIN_OPER) = Empty
                            .TextMatrix(.Row, Ki_MSH_CAMPOS_26_HAVIN_CRI1) = Empty
                            .TextMatrix(.Row, Ki_MSH_CAMPOS_27_HAVIN_CRI2) = Empty
                        Case Ki_MSH_CAMPOS_29_ORDER_CHKI
                            .TextMatrix(.Row, Ki_MSH_CAMPOS_31_ORDER_NUME) = Empty
                            .TextMatrix(.Row, Ki_MSH_CAMPOS_32_ORDER_TIPO) = Empty
                            .Col = Ki_MSH_CAMPOS_32_ORDER_TIPO
                            Set .CellPicture = Nothing
                    End Select
                End If
                If .Col = Ki_MSH_CAMPOS_08_SELEC_ACTI Then Call CargarDatosOrdenDeCampos
        End Select
        .Refresh
        Call SeteaFocoEditEnMSHDIseño
    End With
End Sub

Private Sub mshDiseñoDatos_RowColChange()
    Dim F As Integer
    With mshDiseñoDatos
        For F = .FixedRows To .Rows - 1
            .TextMatrix(F, Ki_MSH_CAMPOS_01_FIXED__COL) = Empty
        Next F
        .TextMatrix(.Row, Ki_MSH_CAMPOS_01_FIXED__COL) = GO_ECNLIB00_CONST.CARESP_WEB_DERECHA
    End With
End Sub

Private Sub MuestraControlesDGMatricial(ByVal blSW As Boolean)
    dcTabla(0).Visible = blSW
    cboTipSQLJoin(0).Visible = blSW
    dcCampo(0).Visible = blSW
    cboTipCampo(0).Visible = blSW
    cboTipoWHERE(0).Visible = blSW
    cboOperWHERE(0).Visible = blSW
    cboTipoHAVING(0).Visible = blSW
    cboOperHAVING(0).Visible = blSW
    cboTipOrden(0).Visible = blSW
    TxT(0).Visible = blSW
End Sub



Private Sub mshDiseñoDatos_EnterCell()
    If PU_SW_DISEÑO_CON_TABLAS = False Then Exit Sub
    objBallonTooT_MSHDiseño.Active = False
    
    Dim xControl As Control
    Dim iColVal As Integer
    
    Call MuestraControlesDGMatricial(False)
SALTO_REGRESO:
    iColVal = -1
    With mshDiseñoDatos
        Select Case .Col
            Case Ki_MSH_CAMPOS_02_TABLA_NOMB
                Set xControl = dcTabla(0)
                iColVal = -99
            Case Ki_MSH_CAMPOS_03_TABLA_JOIN
                Set xControl = cboTipSQLJoin(0)
                iColVal = Ki_MSH_CAMPOS_02_TABLA_NOMB
            Case Ki_MSH_CAMPOS_06_SELEC_NOMB
                If .TextMatrix(.Row, Ki_MSH_CAMPOS_05_SELEC_TIPO) = GO_002_Ks_TIPO_DE_CAMPO_FD Then
                    Set xControl = dcCampo(0)
                Else
                    Set xControl = TxT(0)
                End If
                iColVal = Ki_MSH_CAMPOS_02_TABLA_NOMB
            Case Ki_MSH_CAMPOS_05_SELEC_TIPO
                Set xControl = cboTipCampo(0)
                iColVal = Ki_MSH_CAMPOS_02_TABLA_NOMB
            Case Ki_MSH_CAMPOS_07_SELEC_ALEA
                Set xControl = TxT(0)
                iColVal = Ki_MSH_CAMPOS_02_TABLA_NOMB
            Case Ki_MSH_CAMPOS_13_WHERE_TIPO
                Set xControl = cboTipoWHERE(0)
                iColVal = Ki_MSH_CAMPOS_12_WHERE_CHKV
            Case Ki_MSH_CAMPOS_14_WHERE_OPER
                Set xControl = cboOperWHERE(0)
                iColVal = Ki_MSH_CAMPOS_12_WHERE_CHKV
            Case Ki_MSH_CAMPOS_24_HAVIN_TIPO
                Set xControl = cboTipoHAVING(0)
                iColVal = Ki_MSH_CAMPOS_23_HAVIN_CHKV
            Case Ki_MSH_CAMPOS_15_WHERE_CRI1, _
                 Ki_MSH_CAMPOS_16_WHERE_CRI2
                Set xControl = TxT(0)
                iColVal = Ki_MSH_CAMPOS_12_WHERE_CHKV
            Case Ki_MSH_CAMPOS_25_HAVIN_OPER
                Set xControl = cboOperHAVING(0)
                iColVal = Ki_MSH_CAMPOS_23_HAVIN_CHKV
            Case Ki_MSH_CAMPOS_20_GROUP_NUME
                Set xControl = cboOrdenDeCampos(0)
                iColVal = Ki_MSH_CAMPOS_19_GROUP_CHKV
            Case Ki_MSH_CAMPOS_26_HAVIN_CRI1, _
                 Ki_MSH_CAMPOS_27_HAVIN_CRI2
                Set xControl = TxT(0)
                iColVal = Ki_MSH_CAMPOS_23_HAVIN_CHKV
            Case Ki_MSH_CAMPOS_31_ORDER_NUME
                Set xControl = cboOrdenDeCampos(0)
                iColVal = Ki_MSH_CAMPOS_30_ORDER_CHKV
            Case Ki_MSH_CAMPOS_32_ORDER_TIPO
                Set xControl = cboTipOrden(0)
                iColVal = Ki_MSH_CAMPOS_30_ORDER_CHKV
        End Select
        
        If Not xControl Is Nothing Then
            Select Case iColVal
                Case -99
                Case Ki_MSH_CAMPOS_02_TABLA_NOMB
                    If Len(Trim(.TextMatrix(.Row, iColVal))) = 0 Then _
                        Exit Sub
                    If .Col = Ki_MSH_CAMPOS_06_SELEC_NOMB Then
                        Select Case .TextMatrix(.Row, Ki_MSH_CAMPOS_05_SELEC_TIPO)
                            Case GO_002_Ks_TIPO_DE_CAMPO_FD
                            Case GO_002_Ks_TIPO_DE_CAMPO_TX
                            Case GO_002_Ks_TIPO_DE_CAMPO_FX
                            Case GO_002_Ks_TIPO_DE_CAMPO_AD
                        End Select
                    ElseIf .Col = Ki_MSH_CAMPOS_03_TABLA_JOIN Then
                        If .TextMatrix(.Row, Ki_MSH_CAMPOS_04_TABLA_CODI) = Ks_CBO_TABLA_NULL_COD Then
                            Call SetearValoresPorTablaNull(.Row)
                            .Col = Ki_MSH_CAMPOS_05_SELEC_TIPO
                            GoTo SALTO_REGRESO
                            Exit Sub
                        End If
                    End If
                Case Else
                    Dim sValor As String
                    sValor = .TextMatrix(.Row, iColVal)
                    If sValor = GO_ECNLIB00_CONST.VAL_UNCHK Then _
                        Exit Sub
            End Select

            xControl.Top = .Top + .CellTop
            xControl.Left = .Left + .CellLeft
            xControl.Width = .CellWidth
            xControl.Visible = True
            xControl.ZOrder 0
            
            Select Case UCase(Trim(TypeName(xControl)))
                Case "TEXTBOX"
                    xControl.Height = .CellHeight
                Case "COMBOBOX", _
                     "IMAGECOMBO", _
                     "DATACOMBO"
                    Call GO_ECNLIB01_FUNSUB.CambiarComboHeight(xControl, .CellHeight)
                    Call GO_ECNLIB01_FUNSUB.DesplegarCb(xControl, True)
            End Select
            
            xControl.SetFocus
        End If
    End With
End Sub

Private Sub SeteaFocoEditEnMSHDIseño()
    On Error Resume Next
    Select Case mshDiseñoDatos.Col
        Case Ki_MSH_CAMPOS_02_TABLA_NOMB: mshDiseñoDatos.Col = Ki_MSH_CAMPOS_03_TABLA_JOIN
        Case Ki_MSH_CAMPOS_03_TABLA_JOIN: mshDiseñoDatos.Col = Ki_MSH_CAMPOS_05_SELEC_TIPO
        Case Ki_MSH_CAMPOS_04_TABLA_CODI:
        Case Ki_MSH_CAMPOS_05_SELEC_TIPO: mshDiseñoDatos.Col = Ki_MSH_CAMPOS_06_SELEC_NOMB
        Case Ki_MSH_CAMPOS_06_SELEC_NOMB: mshDiseñoDatos.Col = Ki_MSH_CAMPOS_07_SELEC_ALEA
        Case Ki_MSH_CAMPOS_07_SELEC_ALEA: mshDiseñoDatos.Col = Ki_MSH_CAMPOS_08_SELEC_ACTI
        Case Ki_MSH_CAMPOS_08_SELEC_ACTI: mshDiseñoDatos.Col = Ki_MSH_CAMPOS_11_WHERE_CHKI
        Case Ki_MSH_CAMPOS_11_WHERE_CHKI: mshDiseñoDatos.Col = Ki_MSH_CAMPOS_13_WHERE_TIPO
        Case Ki_MSH_CAMPOS_13_WHERE_TIPO: mshDiseñoDatos.Col = Ki_MSH_CAMPOS_14_WHERE_OPER
        Case Ki_MSH_CAMPOS_14_WHERE_OPER: mshDiseñoDatos.Col = Ki_MSH_CAMPOS_15_WHERE_CRI1
        Case Ki_MSH_CAMPOS_15_WHERE_CRI1: mshDiseñoDatos.Col = Ki_MSH_CAMPOS_16_WHERE_CRI2
        Case Ki_MSH_CAMPOS_16_WHERE_CRI2: mshDiseñoDatos.Col = Ki_MSH_CAMPOS_18_GROUP_CHKI
        Case Ki_MSH_CAMPOS_18_GROUP_CHKI: mshDiseñoDatos.Col = Ki_MSH_CAMPOS_20_GROUP_NUME
        Case Ki_MSH_CAMPOS_20_GROUP_NUME: mshDiseñoDatos.Col = Ki_MSH_CAMPOS_22_HAVIN_CHKI
        Case Ki_MSH_CAMPOS_22_HAVIN_CHKI: mshDiseñoDatos.Col = Ki_MSH_CAMPOS_24_HAVIN_TIPO
        Case Ki_MSH_CAMPOS_24_HAVIN_TIPO: mshDiseñoDatos.Col = Ki_MSH_CAMPOS_25_HAVIN_OPER
        Case Ki_MSH_CAMPOS_25_HAVIN_OPER: mshDiseñoDatos.Col = Ki_MSH_CAMPOS_26_HAVIN_CRI1
        Case Ki_MSH_CAMPOS_26_HAVIN_CRI1: mshDiseñoDatos.Col = Ki_MSH_CAMPOS_27_HAVIN_CRI2
        Case Ki_MSH_CAMPOS_27_HAVIN_CRI2: mshDiseñoDatos.Col = Ki_MSH_CAMPOS_29_ORDER_CHKI
        Case Ki_MSH_CAMPOS_29_ORDER_CHKI: mshDiseñoDatos.Col = Ki_MSH_CAMPOS_31_ORDER_NUME
        Case Ki_MSH_CAMPOS_31_ORDER_NUME: mshDiseñoDatos.Col = Ki_MSH_CAMPOS_32_ORDER_TIPO
        Case Ki_MSH_CAMPOS_32_ORDER_TIPO:
            With mshDiseñoDatos
                If .Row < .Rows - 1 Then .Row = .Row + 1
                .Col = Ki_MSH_CAMPOS_02_TABLA_NOMB
            End With
    End Select
    mshDiseñoDatos.SetFocus
    Call mshDiseñoDatos_RowColChange
    Call mshDiseñoDatos_EnterCell
End Sub



Private Sub TxT_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            If blSW_Load = True Then Exit Sub
                    
            With mshDiseñoDatos
                .TextMatrix(.Row, .Col) = TxT(0).Text
            End With
            TxT(0).Text = Empty
            TxT(0).Visible = False
            Call SeteaFocoEditEnMSHDIseño
        Case vbKeyEscape
            TxT(0).Text = Empty
            TxT(0).Visible = False
            mshDiseñoDatos.SetFocus
    End Select
End Sub

Private Sub dcTabla_Click(Index As Integer, Area As Integer)
    If blSW_Load = True Then Exit Sub
    If Area <> 2 Then Exit Sub
    
    With mshDiseñoDatos
        If .TextMatrix(.Row, Ki_MSH_CAMPOS_04_TABLA_CODI) <> dcTabla(0).BoundText Then
            .TextMatrix(.Row, Ki_MSH_CAMPOS_06_SELEC_NOMB) = Empty
        End If
        .TextMatrix(.Row, Ki_MSH_CAMPOS_02_TABLA_NOMB) = dcTabla(0).Text
        .TextMatrix(.Row, Ki_MSH_CAMPOS_04_TABLA_CODI) = dcTabla(0).BoundText
    End With
    dcTabla(0).Visible = False
    Call CargarDatosDeCboDeCamposPorTabla
    Call SeteaFocoEditEnMSHDIseño
End Sub

Private Sub dcTabla_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    Select Case KeyAscii
        Case 13
            Call dcTabla_Click(Index, 2)
        Case vbKeyEscape
            dcTabla(0).Visible = False
            mshDiseñoDatos.SetFocus
    End Select
End Sub

Private Sub dcCampo_Click(Index As Integer, Area As Integer)
    If blSW_Load = True Then Exit Sub
    If Area <> 2 Then Exit Sub
    
    With mshDiseñoDatos
        .TextMatrix(.Row, Ki_MSH_CAMPOS_06_SELEC_NOMB) = dcCampo(0).Text
    End With
    dcCampo(0).Visible = False
    Call CargarDatosOrdenDeCampos
    Call SeteaFocoEditEnMSHDIseño
End Sub

Private Sub dcCampo_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    Select Case KeyAscii
        Case 13
            Call dcCampo_Click(Index, 2)
        Case vbKeyEscape
            dcCampo(0).Visible = False
            mshDiseñoDatos.SetFocus
    End Select
End Sub

Private Sub CargarDatosOrdenDeCampos()
    On Error Resume Next
    
    Dim F As Integer
    Dim i As Integer
    Dim blSW As Boolean
    Dim INDICE As Integer
    
    Dim aCamposValidosParaOrdenar() As Variant
    
    cboOrdenDeCampos(0).Clear
    INDICE = 0
    With mshDiseñoDatos
        For F = .FixedRows To .Rows - 1
            If .TextMatrix(F, Ki_MSH_CAMPOS_09_SELEC_ACTV) = GO_ECNLIB00_CONST.VAL_CHECK And _
               .TextMatrix(F, Ki_MSH_CAMPOS_05_SELEC_TIPO) = GO_002_Ks_TIPO_DE_CAMPO_FD Then
                
                blSW = False
                If INDICE > 0 Then
                    For i = 1 To UBound(aCamposValidosParaOrdenar())
                        If .TextMatrix(F, Ki_MSH_CAMPOS_06_SELEC_NOMB) = aCamposValidosParaOrdenar(i) Then
                            blSW = True
                            Exit For
                        End If
                    Next i
                End If
                If blSW = False Then
                    INDICE = INDICE + 1
                    ReDim Preserve aCamposValidosParaOrdenar(1 To INDICE)
                    aCamposValidosParaOrdenar(INDICE) = .TextMatrix(F, Ki_MSH_CAMPOS_06_SELEC_NOMB)
                End If
            End If
        Next F
        
        If INDICE > 0 Then
            For i = 1 To UBound(aCamposValidosParaOrdenar())
                cboOrdenDeCampos(0).AddItem CStr(i) & "°"
            Next i
        End If
    End With
    Call GO_ECNLIB01_FUNSUB.CambiarComboListaAncho(cboOrdenDeCampos(0), 30)
End Sub

Private Sub CargarDatosDeCboDeCamposPorTabla()
    On Error Resume Next
    Dim oRs As New ADODB.Recordset
    
    Set oRs = New ADODB.Recordset
    With oRs
        .CursorLocation = adUseClient
        .Fields.Append "CODIGO", adVarChar, 100
        .Fields.Append "NOMBRE", adVarChar, 100
        .Open
    End With
    
    Dim i  As Integer
    Dim INDICE As Integer
    
    Set dcCampo(0).RowSource = Nothing
    dcCampo(0).Visible = False
    
    For i = LBound(aVectorDeIndicesDeTablasDelDiagrama()) To _
            UBound(aVectorDeIndicesDeTablasDelDiagrama())
        INDICE = aVectorDeIndicesDeTablasDelDiagrama(i)
        If UCase(Trim(mshTabla(INDICE).Tag)) = UCase(mshDiseñoDatos.TextMatrix(mshDiseñoDatos.Row, Ki_MSH_CAMPOS_04_TABLA_CODI)) Then _
            Exit For
    Next i
    
    With mshTabla(INDICE)
        For i = .FixedRows To .Rows - 1
            With oRs
                .AddNew
                .Fields("CODIGO") = UCase(mshTabla(INDICE).TextMatrix(i, Ki_TABLA_Col_Codigo))
                .Fields("NOMBRE") = UCase(mshTabla(INDICE).TextMatrix(i, Ki_TABLA_Col_Descripcion))
                .Update
                .MoveLast
            End With
        Next i
    End With
    
    With dcCampo(0)
        Set .RowSource = oRs
        .ListField = "NOMBRE"
        .BoundColumn = "CODIGO"
        .Refresh
    End With
    
'    Call GO_ECNLIB01_FUNSUB.CambiarComboListaAncho(dcTabla(0), 900)
'    Call GO_ECNLIB01_FUNSUB.CambiarComboListaLargo(dcTabla(0), 300)
End Sub

Private Sub cboTipSQLJoin_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Call cboTipSQLJoin_Click(Index)
        Case vbKeyEscape
            cboTipSQLJoin(0).Text = Empty
            cboTipSQLJoin(0).Visible = False
            mshDiseñoDatos.SetFocus
    End Select
End Sub

Private Sub cboTipSQLJoin_Click(Index As Integer)
    If blSW_Load = True Then Exit Sub
    
    With mshDiseñoDatos
        Dim oCboItem As ComboItem
        
        Set oCboItem = cboTipSQLJoin(0).SelectedItem
        If oCboItem Is Nothing Then Exit Sub
        .TextMatrix(.Row, Ki_MSH_CAMPOS_03_TABLA_JOIN) = oCboItem.Text
        Set .CellPicture = imgL2.ListImages(oCboItem.Image).Picture
    End With
    cboTipSQLJoin(0).Visible = False
    
    Dim sTabla As String
    Dim F As Integer
    Dim iFilaRow As Integer
    
    With mshDiseñoDatos
        sTabla = .TextMatrix(.Row, Ki_MSH_CAMPOS_02_TABLA_NOMB)
        iFilaRow = .Row
        
        For F = .FixedRows To .Rows - 1
            If .TextMatrix(F, Ki_MSH_CAMPOS_02_TABLA_NOMB) = sTabla Then
                .Row = F
                Set .CellPicture = imgL2.ListImages(oCboItem.Image).Picture
                .TextMatrix(.Row, Ki_MSH_CAMPOS_03_TABLA_JOIN) = oCboItem.Key
            End If
        Next F
        
        .Row = iFilaRow
    End With
    Call SeteaFocoEditEnMSHDIseño
End Sub

Private Sub cboTipCampo_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Call cboTipCampo_Click(Index)
        Case vbKeyEscape
            cboTipCampo(0).Text = Empty
            cboTipCampo(0).Visible = False
            mshDiseñoDatos.SetFocus
    End Select
End Sub

Private Sub cboTipCampo_Click(Index As Integer)
    If blSW_Load = True Then Exit Sub
    
    With mshDiseñoDatos
         If Len(Trim(.TextMatrix(.Row, Ki_MSH_CAMPOS_02_TABLA_NOMB))) = 0 Then
            cboTipCampo(0).Visible = False
            Exit Sub
        End If
        Dim oCboItem As ComboItem
        
        Set oCboItem = cboTipCampo(0).SelectedItem
        If Not oCboItem Is Nothing Then
            If .TextMatrix(.Row, Ki_MSH_CAMPOS_04_TABLA_CODI) = Ks_CBO_TABLA_NULL_COD Then
                If oCboItem.Key = GO_002_Ks_TIPO_DE_CAMPO_FD Then
                    .TextMatrix(.Row, Ki_MSH_CAMPOS_05_SELEC_TIPO) = Empty
                    .Col = Ki_MSH_CAMPOS_05_SELEC_TIPO
                    Set .CellPicture = Nothing
                    .TextMatrix(.Row, Ki_MSH_CAMPOS_06_SELEC_NOMB) = Empty
                    
                    MsgBox "Tipo de campo no válido para la tabla seleccionada dentro de la sección [TABLA], columna [NOMBRE]", vbCritical, Me.Caption
                    cboTipCampo(Index).SetFocus
                    Call GO_ECNLIB01_FUNSUB.DesplegarCb(cboTipCampo(0), True)
                    Exit Sub
                End If
            Else
                If oCboItem.Key <> GO_002_Ks_TIPO_DE_CAMPO_FD Then
                    MsgBox "Tipo de campo no válido para la tabla seleccionada dentro de la sección [TABLA], columna [NOMBRE]", vbCritical, Me.Caption
                    cboTipCampo(Index).SetFocus
                    Call GO_ECNLIB01_FUNSUB.DesplegarCb(cboTipCampo(0), True)
                    Exit Sub
                End If
            End If
            .TextMatrix(.Row, Ki_MSH_CAMPOS_05_SELEC_TIPO) = oCboItem.Key
            Set .CellPicture = imgL2.ListImages(oCboItem.Image).Picture
        End If
    End With
    cboTipCampo(0).Visible = False
    Call SeteaFocoEditEnMSHDIseño
End Sub

Private Sub cboTipOrden_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Call cboTipOrden_Click(Index)
        Case vbKeyEscape
            cboTipOrden(0).Text = Empty
            cboTipOrden(0).Visible = False
            mshDiseñoDatos.SetFocus
    End Select
End Sub

Private Sub cboTipOrden_Click(Index As Integer)
    If blSW_Load = True Then Exit Sub
    
    With mshDiseñoDatos
        If .TextMatrix(.Row, Ki_MSH_CAMPOS_30_ORDER_CHKV) = GO_ECNLIB00_CONST.VAL_UNCHK Then
            cboTipOrden(0).Visible = False
            Exit Sub
        End If
        Dim oCboItem As ComboItem
        
        Set oCboItem = cboTipOrden(0).SelectedItem
        If oCboItem Is Nothing Then Exit Sub
        .TextMatrix(.Row, Ki_MSH_CAMPOS_32_ORDER_TIPO) = oCboItem.Key
        Set .CellPicture = imgL2.ListImages(oCboItem.Image).Picture
    End With
    cboTipOrden(0).Visible = False
    Call SeteaFocoEditEnMSHDIseño
End Sub

Private Sub cboTipoWHERE_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    Select Case KeyAscii
        Case 13
            Call cboTipoWHERE_Click(Index)
        Case vbKeyEscape
            cboTipoWHERE(0).Text = Empty
            cboTipoWHERE(0).Visible = False
            mshDiseñoDatos.SetFocus
    End Select
End Sub

Private Sub cboTipoWHERE_Click(Index As Integer)
    If blSW_Load = True Then Exit Sub
    With mshDiseñoDatos
        If .TextMatrix(.Row, Ki_MSH_CAMPOS_12_WHERE_CHKV) = GO_ECNLIB00_CONST.VAL_UNCHK Then
            cboTipoWHERE(0).Visible = False
            Exit Sub
        End If
        
        .TextMatrix(.Row, Ki_MSH_CAMPOS_13_WHERE_TIPO) = cboTipoWHERE(0).List(cboTipoWHERE(0).ListIndex)
    End With
    cboTipoWHERE(0).Visible = False
    Call SeteaFocoEditEnMSHDIseño
End Sub

Private Sub cboTipoHAVING_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    Select Case KeyAscii
        Case 13
            Call cboTipoHAVING_Click(Index)
        Case vbKeyEscape
            cboTipoHAVING(0).Text = Empty
            cboTipoHAVING(0).Visible = False
            mshDiseñoDatos.SetFocus
    End Select
End Sub

Private Sub cboTipoHAVING_Click(Index As Integer)
    If blSW_Load = True Then Exit Sub
    With mshDiseñoDatos
        .TextMatrix(.Row, Ki_MSH_CAMPOS_24_HAVIN_TIPO) = cboTipoHAVING(0).List(cboTipoHAVING(0).ListIndex)
    End With
    cboTipoHAVING(0).Visible = False
    Call SeteaFocoEditEnMSHDIseño
End Sub

Private Sub cboOrdenDeCampos_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    Select Case KeyAscii
        Case 13
            Call cboOrdenDeCampos_Click(Index)
        Case vbKeyEscape
            cboOrdenDeCampos(0).Text = Empty
            cboOrdenDeCampos(0).Visible = False
            mshDiseñoDatos.SetFocus
    End Select
End Sub

Private Sub cboOrdenDeCampos_Click(Index As Integer)
    If blSW_Load = True Then Exit Sub
    With mshDiseñoDatos
        .TextMatrix(.Row, .Col) = cboOrdenDeCampos(0).List(cboOrdenDeCampos(0).ListIndex)
    End With
    cboOrdenDeCampos(0).Visible = False
    Call SeteaFocoEditEnMSHDIseño
End Sub

Private Sub cboOperWHERE_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    Select Case KeyAscii
        Case 13
            Call cboOperWHERE_Click(Index)
        Case vbKeyEscape
            cboOperWHERE(0).Text = Empty
            cboOperWHERE(0).Visible = False
            mshDiseñoDatos.SetFocus
    End Select
End Sub

Private Sub cboOperWHERE_Click(Index As Integer)
    If blSW_Load = True Then Exit Sub
    With mshDiseñoDatos
        If .TextMatrix(.Row, Ki_MSH_CAMPOS_12_WHERE_CHKV) = GO_ECNLIB00_CONST.VAL_UNCHK Then
            cboOperWHERE(0).Visible = False
            Exit Sub
        End If
        .TextMatrix(.Row, Ki_MSH_CAMPOS_14_WHERE_OPER) = cboOperWHERE(0).List(cboOperWHERE(0).ListIndex)
    End With
    cboOperWHERE(0).Visible = False
    Call SeteaFocoEditEnMSHDIseño
End Sub

Private Sub cboOperHAVING_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    Select Case KeyAscii
        Case 13
            Call cboOperHAVING_Click(Index)
        Case vbKeyEscape
            cboOperHAVING(0).Text = Empty
            cboOperHAVING(0).Visible = False
            mshDiseñoDatos.SetFocus
    End Select
End Sub

Private Sub cboOperHAVING_Click(Index As Integer)
    If blSW_Load = True Then Exit Sub
    With mshDiseñoDatos
         If .TextMatrix(.Row, Ki_MSH_CAMPOS_23_HAVIN_CHKV) = GO_ECNLIB00_CONST.VAL_UNCHK Then
            cboOperHAVING(0).Visible = False
            Exit Sub
        End If
        .TextMatrix(.Row, Ki_MSH_CAMPOS_25_HAVIN_OPER) = cboOperHAVING(0).List(cboOperHAVING(0).ListIndex)
    End With
    cboOperHAVING(0).Visible = False
    Call SeteaFocoEditEnMSHDIseño
End Sub

Private Sub rtfDiseñoQuerySQL_Change()
    Call oECNLIB03_WINEVE.FormateaSQL_TXT(rtfDiseñoQuerySQL)
End Sub

Private Sub mshTabla_Click(Index As Integer)
    On Error Resume Next
   
    With mshTabla(Index)
        Select Case .Col
            Case Ki_TABLA_Col_CheckIco
                Dim sValor As String
                                
                sValor = .TextMatrix(.Row, Ki_TABLA_Col_CheckVal)
                Select Case sValor
                    Case GO_ECNLIB00_CONST.VAL_UNCHK: sValor = GO_ECNLIB00_CONST.VAL_CHECK
                    Case GO_ECNLIB00_CONST.VAL_CHECK: sValor = GO_ECNLIB00_CONST.VAL_UNCHK
                End Select
                
                .TextMatrix(.Row, Ki_TABLA_Col_CheckVal) = sValor
                
                Select Case sValor
                    Case GO_ECNLIB00_CONST.VAL_UNCHK: Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
                    Case GO_ECNLIB00_CONST.VAL_CHECK: Set .CellPicture = imgL.ListImages(Ki_Ico_Check).Picture
                End Select
        End Select
        .Refresh
    End With
End Sub

Private Sub mshTabla_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    Select Case Button
        Case vbLeftButton
            mshTabla(Index).Drag vbBeginDrag
            Call MostrarControlesDeTabla(Index, False)
            iDragInd_Tabla = Index
            sngDragX_Tabla = X
            sngDragY_Tabla = y
        
        Case vbRightButton
            iIndiceTablaSEL = Index
            With mnPopPup_Tabla
                .PopupMenu .MenuItems.Key2ID(Ks_MNMAIN_00_03___), _
                    .ClientToScreenX(mshTabla(Index).hWnd, X), _
                    .ClientToScreenY(mshTabla(Index).hWnd, y)
            End With
    End Select
End Sub

Private Sub imgRelacion_FK_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button <> vbLeftButton Then Exit Sub
    imgRelacion_FK(Index).Drag vbBeginDrag
    iDragInd_ImgFK = Index
    sngDragX_ImgFK = X
    sngDragY_ImgFK = y
End Sub

Private Sub imgRelacion_PK_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button <> vbLeftButton Then Exit Sub
    imgRelacion_PK(Index).Drag vbBeginDrag
    iDragInd_ImgPK = Index
    sngDragX_ImgPK = X
    sngDragY_ImgPK = y
End Sub

Private Sub pctDiseño_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    On Error Resume Next
    If Button <> vbRightButton Then Exit Sub
    With mnPopPup_DgGrafico.MenuItems
        .Enabled(Ki_MNPOPU_PCTDISEÑO_ID4_DELTBALL) = PU_SW_DISEÑO_CON_TABLAS
        .Enabled(Ki_MNPOPU_PCTDISEÑO_ID5_MINTABLE) = PU_SW_DISEÑO_CON_TABLAS
        .Enabled(Ki_MNPOPU_PCTDISEÑO_ID6_MAXTABLE) = PU_SW_DISEÑO_CON_TABLAS
    End With
    With mnPopPup_DgGrafico
        .PopupMenu .MenuItems.Key2ID(Ks_MNMAIN_00_01___), _
            .ClientToScreenX(pctDiseño.hWnd, X), _
            .ClientToScreenY(pctDiseño.hWnd, y)
    End With
End Sub

Private Sub mnPopPup_DgGrafico_Click(ByVal ID As Long)
    Call EjecucionMenuDgGrafico(mnPopPup_DgGrafico, ID)
End Sub

Private Sub mnPopPup_Tabla_Click(ByVal ID As Long)
    On Error Resume Next
    Select Case ID
        '--------------------------------------------
        '=> ELIMINAR TABLA
        '--------------------------------------------
        Case Ki_MNPOPU_MSHTABLA_ID2_DELTABLE
            Call EliminaTabla(iIndiceTablaSEL)
            iIndiceTablaSEL = -1
    End Select
End Sub

Private Sub EliminaTabla(ByVal iIndTabla As Integer)
    If iIndTabla < 0 Then Exit Sub
    If IndiceEsParteDelArregloDeControlesDeTablas(iIndTabla) = False Then Exit Sub
    If SeteaNullPorValorEnVectorDeIndicesDeTablasEnDiagrama(iIndTabla) = False Then Exit Sub
    
    
    Select Case iIndTabla
        Case 0
            Call ConfiguracionBase_Tablas
        Case Else
            Unload mshTabla(iIndTabla)
            Unload lblTabla(iIndTabla)
            Unload chkTabla(iIndTabla)
            Unload cmdTabla(iIndTabla)
            Unload pctTabla(iIndTabla)
    End Select
    
    Dim i As Integer
    Dim iIndREL As Integer
    Dim iConDel As Integer
    Dim aDelRel() As Integer
       
    iConDel = 0
    For i = LBound(aVectorDeIndicesRelacionesDelDiagrama()) To _
            UBound(aVectorDeIndicesRelacionesDelDiagrama())
        iIndREL = aVectorDeIndicesRelacionesDelDiagrama(i)
        If iIndREL <> Ki_Vector_ValorNULL Then
            If CInt(Val(imgRelacion_PK(iIndREL).Tag)) = iIndTabla Or _
               CInt(Val(imgRelacion_FK(iIndREL).Tag)) = iIndTabla Then
                iConDel = iConDel + 1
                ReDim Preserve aDelRel(1 To iConDel) As Integer
                aDelRel(iConDel) = i
            End If
        End If
    Next i
    
    If iConDel > 0 Then
        For i = LBound(aDelRel()) To _
                UBound(aDelRel())
            Select Case aVectorDeIndicesRelacionesDelDiagrama(aDelRel(i))
                Case 0
                    Call ConfiguracionBase_Relaciones
                Case Else
                    Unload linRelacionPFK(aVectorDeIndicesRelacionesDelDiagrama(aDelRel(i)))
                    Unload imgRelacion_PK(aVectorDeIndicesRelacionesDelDiagrama(aDelRel(i)))
                    Unload imgRelacion_FK(aVectorDeIndicesRelacionesDelDiagrama(aDelRel(i)))
                    
                    aVectorDeIndicesRelacionesDelDiagrama(aDelRel(i)) = Ki_Vector_ValorNULL
            End Select
        Next i
    End If
    '--+-----------------------------------------------------------------------------------+--
    '=> DEFINO LOS NUEVOS VALORES DEL INDICE
    '--+-----------------------------------------------------------------------------------+--
    Dim iMaximo As Integer
    Dim iValor As Integer
    
    iMaximo = IND_VECTOR_DE_RELACIONES
    For i = LBound(aVectorDeIndicesRelacionesDelDiagrama()) To _
            UBound(aVectorDeIndicesRelacionesDelDiagrama())
        iValor = aVectorDeIndicesRelacionesDelDiagrama(i)
        If iValor > iMaximo Then iMaximo = iValor
    Next i
    IND_VECTOR_DE_RELACIONES = iMaximo + 1
    
    iMaximo = IND_VECTOR_DE_TABLAS
    For i = LBound(aVectorDeIndicesDeTablasDelDiagrama()) To _
            UBound(aVectorDeIndicesDeTablasDelDiagrama())
        iValor = aVectorDeIndicesDeTablasDelDiagrama(i)
        If iValor > iMaximo Then iMaximo = iValor
    Next i
    IND_VECTOR_DE_TABLAS = iMaximo + 1
End Sub


Private Sub pctDiseño_DragDrop(Source As Control, X As Single, y As Single)
     On Error Resume Next
     Select Case UCase(Trim(Source.Name))
        '--+------------------------------------------------------------------------------------------------------------------------------+--
        ' DRAG DROP A LA TABLA
        '--+------------------------------------------------------------------------------------------------------------------------------+--
        Case UCase(Trim("mshTabla"))
            If iDragInd_Tabla < 0 Then Exit Sub
    
            Source.Move (X - sngDragX_Tabla), (y - sngDragY_Tabla)
            'Source.Move (X - Source.Width / 2), (Y - Source.Height / 2)
            With Source
                If .Top + .Height + Kdbl_Espacio_Entre_Tablas > ecnScrollContainerDg.Height Then
                    ecnScrollContainerDg.Height = .Top + .Height + Kdbl_Espacio_Entre_Tablas
                End If
                If .Left + .Width + Kdbl_Espacio_Entre_Tablas > ecnScrollContainerDg.Width Then
                    ecnScrollContainerDg.Width = .Left + .Width + Kdbl_Espacio_Entre_Tablas
                End If
            End With
            
            Call UbicaControlesDeTabla(iDragInd_Tabla)
            Call ReUbicaRelaciones(iDragInd_Tabla)
            Call Form_Resize
            iDragInd_Tabla = -1
        '--+------------------------------------------------------------------------------------------------------------------------------+--
        ' DRAG DROP A LA IMAGEN PK
        '--+------------------------------------------------------------------------------------------------------------------------------+--
        Case UCase(Trim("imgRelacion_PK"))
            If iDragInd_ImgPK < 0 Then Exit Sub
            
            Dim iIndTablaPK As Integer
            iIndTablaPK = CInt(Val(imgRelacion_PK(iDragInd_ImgPK).Tag))
            If iIndTablaPK < 0 Then Exit Sub

            If X > (mshTabla(iIndTablaPK).Left + mshTabla(iIndTablaPK).Width + imgRelacion_PK(iDragInd_ImgPK).Width) Then
                X = mshTabla(iIndTablaPK).Left + mshTabla(iIndTablaPK).Width - Ki_MargenDragKeyPermitido
            End If
            If X < (mshTabla(iIndTablaPK).Left - (imgRelacion_PK(iDragInd_ImgPK).Width)) Then
                X = mshTabla(iIndTablaPK).Left - imgRelacion_PK(iDragInd_ImgPK).Width + Ki_MargenDragKeyPermitido
            End If
                        
            If y > (mshTabla(iIndTablaPK).Top + mshTabla(iIndTablaPK).Height + imgRelacion_PK(iDragInd_ImgPK).Height) Then
                y = mshTabla(iIndTablaPK).Top + mshTabla(iIndTablaPK).Height - Ki_MargenDragKeyPermitido
            End If
            If y < (mshTabla(iIndTablaPK).Top - imgRelacion_PK(iDragInd_ImgPK).Height) Then
                y = mshTabla(iIndTablaPK).Top - imgRelacion_PK(iDragInd_ImgPK).Height + Ki_MargenDragKeyPermitido
            End If
            
            Source.Move X, y
            'Source.Move (X - Source.Width / 2), (Y - Source.Height / 2)
            'Source.Move (X - Source.Width), (Y - Source.Height)
            With linRelacionPFK(iDragInd_ImgPK)
                .X1 = imgRelacion_PK(iDragInd_ImgPK).Left + imgRelacion_PK(iDragInd_ImgPK).Width
                .Y1 = imgRelacion_PK(iDragInd_ImgPK).Top + (imgRelacion_PK(iDragInd_ImgPK).Height / 2)
            End With
            iDragInd_ImgPK = -1
        '--+------------------------------------------------------------------------------------------------------------------------------+--
        ' DRAG DROP A LA IMAGEN FK
        '--+------------------------------------------------------------------------------------------------------------------------------+--
        Case UCase(Trim("imgRelacion_FK"))
            If iDragInd_ImgFK < 0 Then Exit Sub
            
            Dim iIndTablaFK As Integer
            iIndTablaFK = EncuentraIndiceDeTablaPorRelacionFK(iDragInd_ImgFK)
            If iIndTablaFK < 0 Then Exit Sub
            
            If X > (mshTabla(iIndTablaFK).Left + mshTabla(iIndTablaFK).Width + imgRelacion_FK(iDragInd_ImgFK).Width) Then
                X = mshTabla(iIndTablaFK).Left + mshTabla(iIndTablaFK).Width - Ki_MargenDragKeyPermitido
            End If
            If X < (mshTabla(iIndTablaFK).Left - imgRelacion_PK(iDragInd_ImgFK).Width) Then
                X = mshTabla(iIndTablaFK).Left - imgRelacion_PK(iDragInd_ImgFK).Width + Ki_MargenDragKeyPermitido
            End If
            
            If y > (mshTabla(iIndTablaFK).Top + mshTabla(iIndTablaFK).Height + imgRelacion_FK(iDragInd_ImgFK).Height) Then
                y = mshTabla(iIndTablaFK).Top + mshTabla(iIndTablaFK).Height - Ki_MargenDragKeyPermitido
            End If
            If y < (mshTabla(iIndTablaFK).Top - imgRelacion_PK(iDragInd_ImgFK).Height) Then
                y = mshTabla(iIndTablaFK).Top - imgRelacion_PK(iDragInd_ImgFK).Height + Ki_MargenDragKeyPermitido
            End If
            
            Source.Move X, y
            With linRelacionPFK(iDragInd_ImgFK)
                .X2 = imgRelacion_FK(iDragInd_ImgFK).Left
                .Y2 = imgRelacion_FK(iDragInd_ImgFK).Top + (imgRelacion_FK(iDragInd_ImgFK).Height / 2)
            End With
            iDragInd_ImgFK = -1
    End Select
End Sub

Private Sub cmdTabla_Click(Index As Integer)
    On Error Resume Next
    Dim blSW_Minimazado As Boolean
    Dim iRowHeight As Integer
    Dim iMshHeight As Integer
    Dim sCapButton As String
    
    blSW_Minimazado = CBool(Val(cmdTabla(Index).Tag))
    Select Case blSW_Minimazado
        Case True
            iRowHeight = 300
            iMshHeight = Ki_TABLA_HEIGHT_MAX
            sCapButton = GO_ECNLIB00_CONST.CARESP_WEB_RESTAURADO
        Case False
            iRowHeight = 0
            iMshHeight = Ki_TABLA_HEIGHT_MIN
            sCapButton = GO_ECNLIB00_CONST.CARESP_WEB_MINIMIZADO
    End Select
    
    Dim F As Integer
    With mshTabla(Index)
        For F = .FixedRows To .Rows - 1
            .Row = F
            .RowHeight(F) = iRowHeight
        Next F
    End With
    
    cmdTabla(Index).Tag = IIf(blSW_Minimazado, GO_ECNLIB00_CONST.VAL_UNCHK, GO_ECNLIB00_CONST.VAL_CHECK)
    cmdTabla(Index).Caption = sCapButton
    mshTabla(Index).Height = iMshHeight
    
    If blSW_Minimazado = True Then Call AjustaTamañoTabla(Index)
    
    mshTabla(Index).Refresh
End Sub

Private Sub chkTabla_Click(Index As Integer)
    On Error Resume Next
    '--+----------------------------------------------------------------------------------------------------------+--
    '=> EFECTO CHECKBOX
    '--+----------------------------------------------------------------------------------------------------------+--
    With mshTabla(Index)
        Dim F As Integer
        Dim i As Integer
        Dim sValor As String
        
        sValor = GO_ECNLIB00_CONST.VAL_UNCHK
        If chkTabla(Index).Value = Checked Then sValor = GO_ECNLIB00_CONST.VAL_CHECK
        .Col = Ki_TABLA_Col_CheckIco
        For F = .FixedRows To .Rows - 1
            If sValor <> .TextMatrix(F, Ki_TABLA_Col_CheckVal) Then
                .TextMatrix(F, Ki_TABLA_Col_CheckVal) = sValor
                .Row = F
                Select Case sValor
                    Case GO_ECNLIB00_CONST.VAL_UNCHK: Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
                    Case GO_ECNLIB00_CONST.VAL_CHECK: Set .CellPicture = imgL.ListImages(Ki_Ico_Check).Picture
                End Select
            End If
        Next F
        .Refresh
    End With
    '--+----------------------------------------------------------------------------------------------------------+--
    '=> AGREGO LA COLUMNA AL DIAGRAMA MATRICIAL
    '--+----------------------------------------------------------------------------------------------------------+--
    Dim sNomTabla As String
    Dim sNomCampo As String
    Dim sKeyField As String, sKeyField_Fnd As String
    Dim sTJoinSQL As String
    Dim blSW_Find As Boolean
    Dim iConRepet As Integer
    
    Dim aTabla() As String
    Dim aJnSQL() As String
    Dim aCampo() As String
    Dim aAleas() As String
    
    
    Dim aTabla_Rep() As String
    Dim aJnSQL_Rep() As String
    Dim aCampo_Rep() As String
    Dim aNumDeRepe() As String
    Dim sMsgRepetido As String
    
    Dim iIndItem As Integer
    Dim iIndRepe As Integer
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' EN VECTORES PARALELOS AGREGO LOS CAMPOS (CON SUS DATOS RESPECTIVOS) QUE AGREGARE AL DIAGRAMA MATRICIAL
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If sValor = GO_ECNLIB00_CONST.VAL_CHECK Then
        With mshTabla(Index)
            sNomTabla = mshTabla(Index).Tag
            For F = .FixedRows To .Rows - 1
                sNomCampo = .TextMatrix(F, Ki_TABLA_Col_Descripcion)
                sKeyField = sNomTabla _
                          & sNomCampo
                With mshDiseñoDatos
                    blSW_Find = False
                    iConRepet = 0
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(i, Ki_MSH_CAMPOS_05_SELEC_TIPO) = GO_002_Ks_TIPO_DE_CAMPO_FD Then
                            sKeyField_Fnd = .TextMatrix(i, Ki_MSH_CAMPOS_02_TABLA_NOMB) _
                                          & .TextMatrix(i, Ki_MSH_CAMPOS_06_SELEC_NOMB)
                            If sKeyField = sKeyField_Fnd Then
                                blSW_Find = True
                                iConRepet = iConRepet + 1
                            End If
                        End If
                    Next i
                    If blSW_Find = True Then
                        iIndRepe = iIndRepe + 1
                        ReDim Preserve aTabla_Rep(1 To iIndRepe) As String
                        ReDim Preserve aJnSQL_Rep(1 To iIndRepe) As String
                        ReDim Preserve aCampo_Rep(1 To iIndRepe) As String
                        ReDim Preserve aNumDeRepe(1 To iIndRepe) As String
                        
                        aTabla_Rep(iIndRepe) = sNomTabla
                        aJnSQL_Rep(iIndItem) = .TextMatrix(i, Ki_MSH_CAMPOS_05_SELEC_TIPO)
                        aCampo_Rep(iIndRepe) = sNomCampo
                        aNumDeRepe(iIndRepe) = iConRepet
                        
                        sMsgRepetido = Trim(sMsgRepetido) _
                                     & vbNewLine & Space(5) & "- [" & sNomTabla & "] {" & sNomCampo & "}"
                    Else
                        iIndItem = iIndItem + 1
                        ReDim Preserve aTabla(1 To iIndItem) As String
                        ReDim Preserve aCampo(1 To iIndItem) As String
                        ReDim Preserve aJnSQL(1 To iIndItem) As String
                        ReDim Preserve aAleas(1 To iIndItem) As String
                        
                        aTabla(iIndItem) = sNomTabla
                        aJnSQL(iIndItem) = .TextMatrix(i, Ki_MSH_CAMPOS_05_SELEC_TIPO)
                        aCampo(iIndItem) = sNomCampo
                        aAleas(iIndRepe) = Empty
                    End If
                End With
            Next F
        End With
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' SI HE ENCONTRADO CAMPO A AGREGAR REPETIDOS ES DECIR QUE YA SE ENCUENTRAN EN EL DIAGRAMA MATRICIAL, PREGUNTO SI
    ' DE TODOS MODOS EL USUARIO DESEA AGREGARLOS AL DIAGRAMA MATRICIAL, SI ES ASI  AGREGO ESTOS VALORES A LOS VECTO-
    ' RES PARALELOS CON UN NOMBRE DE ALEAS RESPECTIVO
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Len(Trim(sMsgRepetido)) > 0 Then
        sMsgRepetido = "Los siguientes datos ya se encuentran en el diagrama matricial, desea agregarlos de todos modos" _
                     & vbNewLine _
                     & sMsgRepetido
        If MsgBox(sMsgRepetido, vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            For i = LBound(aCampo_Rep()) To UBound(aCampo_Rep())
                iIndItem = iIndItem + 1
                ReDim Preserve aTabla(1 To iIndItem) As String
                ReDim Preserve aCampo(1 To iIndItem) As String
                ReDim Preserve aJnSQL(1 To iIndItem) As String
                ReDim Preserve aAleas(1 To iIndItem) As String
                
                aTabla(iIndItem) = aTabla_Rep(i)
                aJnSQL(iIndItem) = aJnSQL_Rep(i)
                aCampo(iIndItem) = aCampo_Rep(i)
                aAleas(iIndItem) = aCampo_Rep(i) & "_" & CStr(aNumDeRepe(i))
            Next i
        End If
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' LEO LOS VECTORES PARALELOS LLENO DE LOS CAMPOS QUE VOY A AGREGAR AL DIAGRAMA MATRICIAL Y LOS AGREGO AL MISMO
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If iIndItem = 0 Then Exit Sub
    Dim blSW_FilasAgotadas As Boolean
    Dim iFila As String
    
    For i = LBound(aCampo()) To UBound(aCampo())
        With mshDiseñoDatos
            '------------------------------------------------------------------------
            '=> BUSCO SI EXISTE UNA FILA DISPONIBLE DONDE AGREGAR LOS CAMPOS
            '   CARGADOS EN LOS VECTORES PARALELOS
            '------------------------------------------------------------------------
            blSW_FilasAgotadas = True
            For F = .FixedRows To .Rows - 1
                If Len(Trim(.TextMatrix(F, Ki_MSH_CAMPOS_02_TABLA_NOMB))) = 0 Then
                    blSW_FilasAgotadas = False
                    iFila = F
                    Exit For
                End If
            Next F
            '------------------------------------------------------------------------
            '=> SI NO HAY FILAS DISPONIBLES ENTONCES AGREGO UNA NUEVA FILA
            '   Y APLICO EL FORMATO RESPECTIVO
            '------------------------------------------------------------------------
            If blSW_FilasAgotadas = True Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
                iFila = .Row
                Call ConfiguraMSHCampos(iFila)
            End If
            '------------------------------------------------------------------------
            '=> AGREGO LOS DATOS CARGADOS EN LOS VECTORES PARALELOS AL
            '   DIAGRAMA MATRICIAL
            '------------------------------------------------------------------------
            .TextMatrix(iFila, Ki_MSH_CAMPOS_02_TABLA_NOMB) = aTabla(i)
            .TextMatrix(iFila, Ki_MSH_CAMPOS_03_TABLA_JOIN) = aJnSQL(i)
            .TextMatrix(iFila, Ki_MSH_CAMPOS_06_SELEC_NOMB) = aCampo(i)
            .TextMatrix(iFila, Ki_MSH_CAMPOS_05_SELEC_TIPO) = GO_002_Ks_TIPO_DE_CAMPO_FD
            .TextMatrix(iFila, Ki_MSH_CAMPOS_07_SELEC_ALEA) = aAleas(i)
            
            .Row = iFila
            .Col = Ki_MSH_CAMPOS_05_SELEC_TIPO
            Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_Column).Picture
        End With
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call oECNLIB01_FUNSUB.GrabarValoresDeUnFormulario(Me, GO_002_RUTA_INI_PARAM_WIN)
    Set oECNLIB01_FUNSUB = Nothing
    Set oECNLIB03_WINEVE = Nothing
    
    Set objBallonTooT_MSHDiseño = Nothing
    
    Set objECN_HandlerFlatCboTipSJOIN = Nothing
    Set objECN_HandlerFlatCboTipCAMPO = Nothing
    Set objECN_HandlerFlatCboTipWHERE = Nothing
    Set objECN_HandlerFlatCboOpeWHERE = Nothing
    Set objECN_HandlerFlatCboTipHAVIN = Nothing
    Set objECN_HandlerFlatCboOpeHAVIN = Nothing
    Set objECN_HandlerFlatCboTipORDER = Nothing
    Set objECN_HandlerFlatCboOrdCampo = Nothing
End Sub

Private Sub DiseñarTablas()
    On Error Resume Next
    If GO_002_RS_TABLAS Is Nothing Then Exit Sub
    If GO_002_RS_TABLAS.RecordCount = 0 Then Exit Sub
    
    Dim sCodTabla As String
    Dim sDesTabla As String
    Dim sNomTabla As String
    Dim sTipTabla As String
    Dim oRsColumns As New ADODB.Recordset
    
    Dim blSW_TablaExisteEnDiagrama As Boolean
    Dim i As Integer
    Dim INDICE As Integer
    Dim INDICE_CERCANO As Integer
    Dim iConRepeticiones As Integer
   
    With GO_002_RS_TABLAS
        .MoveFirst
        Do While Not .EOF
            If IND_VECTOR_DE_TABLAS <> 0 Then
                Load chkTabla(IND_VECTOR_DE_TABLAS)
                Load cmdTabla(IND_VECTOR_DE_TABLAS)
                Load pctTabla(IND_VECTOR_DE_TABLAS)
                Load lblTabla(IND_VECTOR_DE_TABLAS)
                
                Load mshTabla(IND_VECTOR_DE_TABLAS)
                
                INDICE_CERCANO = FindIndMasCercanoEnVectorDeTablas(IND_VECTOR_DE_TABLAS - 1)
                With mshTabla(IND_VECTOR_DE_TABLAS)
                    .Top = mshTabla(INDICE_CERCANO).Top
                    .Left = mshTabla(INDICE_CERCANO).Left + mshTabla(INDICE_CERCANO).Width + Kdbl_Espacio_Entre_Tablas
                    If .Left + .Width > pctDiseño.Width Then
                        .Left = Ki_TABLA_LEFT_INICIAL
                        .Top = mshTabla(INDICE_CERCANO).Top + mshTabla(INDICE_CERCANO).Height + Kdbl_Espacio_Entre_Tablas
                    End If
                    
                    If .Top + .Height > pctDiseño.Height Then
                        pctDiseño.Height = .Top + .Height + Kdbl_Espacio_Entre_Tablas
                    End If
                End With
            End If
            sCodTabla = UCase(.Fields(GO_002_Ks_TABLAS_CAMPO_CODIGO).Value)
            sDesTabla = UCase(.Fields(GO_002_Ks_TABLAS_CAMPO_DESCRI).Value)
            sTipTabla = CStr(.Fields(GO_002_Ks_TABLAS_CAMPO_TIPO).Value)
            
            mshTabla(IND_VECTOR_DE_TABLAS).Tag = sDesTabla
            chkTabla(IND_VECTOR_DE_TABLAS).Tag = sTipTabla
            pctTabla(IND_VECTOR_DE_TABLAS).Tag = sCodTabla
            
            blSW_TablaExisteEnDiagrama = False
            If IND_VECTOR_DE_TABLAS > 0 Then
                For i = LBound(aVectorDeIndicesDeTablasDelDiagrama()) To _
                        UBound(aVectorDeIndicesDeTablasDelDiagrama())
                    INDICE = aVectorDeIndicesDeTablasDelDiagrama(i)
                    Select Case INDICE
                        Case Ki_Vector_ValorNULL, IND_VECTOR_DE_TABLAS
                        Case Else
                            If mshTabla(INDICE).Tag = sDesTabla Then
                                blSW_TablaExisteEnDiagrama = True
                                Exit For
                            End If
                    End Select
                Next i
            End If
            
            sNomTabla = sDesTabla
            If blSW_TablaExisteEnDiagrama Then
                iConRepeticiones = 0
                For i = LBound(aVectorDeIndicesDeTablasDelDiagrama()) To _
                        UBound(aVectorDeIndicesDeTablasDelDiagrama())
                    INDICE = aVectorDeIndicesDeTablasDelDiagrama(i)
                    Select Case INDICE
                        Case Ki_Vector_ValorNULL, IND_VECTOR_DE_TABLAS
                        Case Else
                            If mshTabla(INDICE).Tag = sDesTabla Then
                                iConRepeticiones = iConRepeticiones + 1
                            End If
                    End Select
                Next i
                sNomTabla = sDesTabla _
                          & Ks_CarSepPKFK _
                          & CStr(iConRepeticiones)
            End If
            lblTabla(IND_VECTOR_DE_TABLAS).Caption = sNomTabla
                        
            Set mshTabla(IND_VECTOR_DE_TABLAS).DataSource = PU_002_CargarColumnas(sCodTabla)
            mshTabla(IND_VECTOR_DE_TABLAS).Visible = True
            PU_SW_DISEÑO_CON_TABLAS = True
            Call ConfiguraGrillaTabla(IND_VECTOR_DE_TABLAS)
            
            ReDim Preserve aVectorDeIndicesDeTablasDelDiagrama(1 To (IND_VECTOR_DE_TABLAS + 1)) As Integer
            aVectorDeIndicesDeTablasDelDiagrama(IND_VECTOR_DE_TABLAS + 1) = IND_VECTOR_DE_TABLAS
            
            .MoveNext
            IND_VECTOR_DE_TABLAS = IND_VECTOR_DE_TABLAS + 1
        Loop
    End With
    Call AjustaTamañoTabla
    Call DiseñarRelaciones
    Call CargarDatosDeCboDeTablasParaMSH
    pctDiseño.SetFocus
End Sub

Private Sub CargarDatosDeCboDeTablasParaMSH()
    On Error Resume Next
    Dim oRs As New ADODB.Recordset
    
    Set oRs = New ADODB.Recordset
    With oRs
        .CursorLocation = adUseClient
        .Fields.Append "CODIGO", adVarChar, 100
        .Fields.Append "NOMBRE", adVarChar, 100
        .Open
    End With
    
    Dim oRsCboTablas As New ADODB.Recordset
    Dim blSWExisNull As Boolean
    With oRsCboTablas
        Set .DataSource = dcTabla(0).RowSource
        blSWExisNull = False
        If GO_ECNLIB02_VALIDA.RsEsValidoParaLectura(oRsCboTablas) Then
            .MoveFirst
            Do While .EOF
                If .Fields("CODIGO").Value = Ks_CBO_TABLA_NULL_COD Then
                    blSWExisNull = True
                    Exit Do
                End If
                oRsCboTablas.MoveNext
            Loop
        End If
    End With
    
    If blSWExisNull = False Then
        With oRs
            .AddNew
            .Fields("CODIGO") = Ks_CBO_TABLA_NULL_COD
            .Fields("NOMBRE") = Ks_CBO_TABLA_NULL_DES
            .Update
            .MoveLast
        End With
    End If
    
    Dim i  As Integer
    Dim INDICE As Integer
    
    Set dcTabla(0).RowSource = Nothing
    dcTabla(0).Visible = False
    
    For i = LBound(aVectorDeIndicesDeTablasDelDiagrama()) To _
            UBound(aVectorDeIndicesDeTablasDelDiagrama())
        INDICE = aVectorDeIndicesDeTablasDelDiagrama(i)
        With oRs
            .AddNew
            .Fields("CODIGO") = UCase(mshTabla(INDICE).Tag)
            .Fields("NOMBRE") = UCase(lblTabla(INDICE).Caption)
            .Update
            .MoveLast
        End With
    Next i
    
    With dcTabla(0)
        Set .RowSource = oRs
        .ListField = "NOMBRE"
        .BoundColumn = "CODIGO"
        .Refresh
    End With
    
'    Call GO_ECNLIB01_FUNSUB.CambiarComboListaAncho(dcTabla(0), 100)
'    Call GO_ECNLIB01_FUNSUB.CambiarComboListaLargo(dcTabla(0), 300)
End Sub

Private Sub AjustaTamañoTabla(Optional ByVal iIndTabla As Integer = -1)
    On Error Resume Next
    Dim i As Integer
    Dim F As Integer
    Dim iIni As Integer
    Dim iFin As Integer
    Dim iConRow As Integer
    Dim INDICE As Integer
    
    If iIndTabla >= 0 Then
        iIni = 1
        iFin = 1
    Else
        iIni = LBound(aVectorDeIndicesDeTablasDelDiagrama())
        iFin = UBound(aVectorDeIndicesDeTablasDelDiagrama())
    End If
  
    For i = iIni To iFin
        If iIndTabla >= 0 Then
            INDICE = iIndTabla
        Else
            INDICE = aVectorDeIndicesDeTablasDelDiagrama(INDICE)
        End If
        With mshTabla(INDICE)
            iConRow = 0
            For F = .FixedRows To .Rows - 1
                iConRow = iConRow + 1
                If F > Ki_TABLA_ROWS_MAX Then Exit For
            Next F
            .Height = Ki_TABLA_ROWHEIGHT_FX1 _
                    + Ki_TABLA_ROWHEIGHT_FX2 _
                    + Ki_TABLA_ROWHEIGHT_FX1 _
                    + (iConRow * Ki_TABLA_ROWHEIGHT)
        End With
    Next i
End Sub

Private Sub DiseñarRelaciones()
    'On Error Resume Next
    '=============================================================================================================================
    '=> DECLARACION DE VARIABLES
    '=============================================================================================================================
    Dim IPK As Integer
    Dim IFK As Integer
    Dim IRL As Integer
    
    Dim iIndPK As Integer
    Dim iIndFK As Integer
    Dim iIndRL As Integer
    
    Dim sCodTabla_PK As String
    Dim oRsTablaPKFK As New ADODB.Recordset
        
    Dim iIndPK_FND As Integer
    Dim iIndFK_FND As Integer
        
    Dim sPKFK_Relacion As String
    
    Dim IndTablasFKxPK As Integer
    Dim aTablasFKporPK() As Integer
    Dim blSW_TieneFK_Pendiente As Boolean
    '=============================================================================================================================
    '=> LEO TODO EL ARREGLO DE CONTROLES PARA BUSCAR LAS RELACIONES, TOMANDO CADA TABLA DEL VECTOR COMO TABLA PK
    '=============================================================================================================================
    For IPK = LBound(aVectorDeIndicesDeTablasDelDiagrama()) To _
              UBound(aVectorDeIndicesDeTablasDelDiagrama())
        iIndPK = aVectorDeIndicesDeTablasDelDiagrama(IPK)
        If iIndPK <> Ki_Vector_ValorNULL Then
            '---------------------------------------------------------------------------------------------------------------------
            '=> CONSULTA DE TODAS LAS LLAVES FORANEAS QUE TIENE LA TABLA PK
            '---------------------------------------------------------------------------------------------------------------------
            sCodTabla_PK = pctTabla(iIndPK).Tag
            Set oRsTablaPKFK = PU_002_ObtenerRelacionesPKFK(sCodTabla_PK)
            '---------------------------------------------------------------------------------------------------------------------
            '=> SI EXISTEN RELACIONES ENTONCES PROCEDO A DIBUJARLAS EN EL DIAGRAMA, PARA ELLO LEO EL RS, CADA REGISTRO ES UNA
            '   RELACION PK CON FK
            '---------------------------------------------------------------------------------------------------------------------
            If GO_ECNLIB02_VALIDA.RsEsValidoParaLectura(oRsTablaPKFK) Then
                With oRsTablaPKFK
                    .MoveFirst
                    Do While Not .EOF
                        '--------------------------------------------------------------------------------------------------------
                        '=> LEO TODO EL ARREGLO DE CONTROLES DE LOS MSH, ESTA VEZ TOMO CADA TABLA COMO UNA TABLA FK, SI CADA
                        '   TABLA QUE LEO COINCIDE CON EL VALOR FK (CAMPO "FK_TABLA_DES") DEL REGISTRO ACTUAL DEL RECORDSET
                        '   ENTONCES CREO Y/O ALIMENTO UN VECTOR DONDE GUARDO SOLO LOS INDICES DE LAS TABLAS QUE SON FK
                        '--------------------------------------------------------------------------------------------------------
                        IndTablasFKxPK = 0
                        Erase aTablasFKporPK
                        For IFK = LBound(aVectorDeIndicesDeTablasDelDiagrama()) To _
                                  UBound(aVectorDeIndicesDeTablasDelDiagrama())
                            iIndFK = aVectorDeIndicesDeTablasDelDiagrama(IFK)
                            If iIndFK <> Ki_Vector_ValorNULL Then
                                If UCase(Trim(mshTabla(iIndFK).Tag)) = UCase(Trim(.Fields("FK_TABLA_DES").Value)) Then
                                    IndTablasFKxPK = IndTablasFKxPK + 1
                                    ReDim Preserve aTablasFKporPK(1 To IndTablasFKxPK) As Integer
                                    aTablasFKporPK(IndTablasFKxPK) = iIndFK
                                End If
                            End If
                        Next IFK
                        '--------------------------------------------------------------------------------------------------------
                        '=> TENIENDO TODAS LAS TABLAS FK DE LA TABLA PK () EN UN VECTOR, PROCEDO A LEER DICHO VECTOR Y BUSCO EN
                        '   EL DIAGRAMA SI SUS RELACIONES ESTAN CREADAS, SI NO FUERA ASI ENTONCES PROCEDO A DIBUJARLAS DENTRO
                        '   DEL DIAGRAMA
                        '--------------------------------------------------------------------------------------------------------
                        If IndTablasFKxPK > 0 Then
                            For IFK = LBound(aTablasFKporPK()) To _
                                      UBound(aTablasFKporPK())
                                iIndFK = aTablasFKporPK(IFK)
                                blSW_TieneFK_Pendiente = True
                                If IND_VECTOR_DE_RELACIONES > 0 Then
                                    For IRL = LBound(aVectorDeIndicesRelacionesDelDiagrama()) To _
                                              UBound(aVectorDeIndicesRelacionesDelDiagrama())
                                        iIndRL = aVectorDeIndicesRelacionesDelDiagrama(IRL)
                                        If iIndRL <> Ki_Vector_ValorNULL Then
                                            If CInt(Val(imgRelacion_PK(iIndRL).Tag)) = iIndPK And _
                                               CInt(Val(imgRelacion_FK(iIndRL).Tag)) = iIndFK Then
                                                    blSW_TieneFK_Pendiente = False
                                                    Exit For
                                            End If
                                        End If
                                    Next IRL
                                End If
                                
                                If blSW_TieneFK_Pendiente Then
                                    If IND_VECTOR_DE_RELACIONES <> 0 Then
                                        Load imgRelacion_PK(IND_VECTOR_DE_RELACIONES)
                                        Load linRelacionPFK(IND_VECTOR_DE_RELACIONES)
                                        Load imgRelacion_FK(IND_VECTOR_DE_RELACIONES)
                                    End If
                                    imgRelacion_PK(IND_VECTOR_DE_RELACIONES).Tag = CStr(iIndPK) '=> INDICE DE LA TABLA QUE LLEVA LA LLAVE PK
                                    imgRelacion_FK(IND_VECTOR_DE_RELACIONES).Tag = CStr(iIndFK) '=> INDICE DE LA TABLA QUE LLEVA LA LLAVE FK
                                    
                                    sPKFK_Relacion = Ks_CarPK & CStr(iIndPK) _
                                                   & Ks_CarSepPKFK _
                                                   & Ks_CarFK & CStr(iIndFK)
                                 
                                    linRelacionPFK(IND_VECTOR_DE_RELACIONES).Tag = Trim(sPKFK_Relacion)
                                    Call UbicaRelacion(iIndPK, iIndFK, IND_VECTOR_DE_RELACIONES)
                                    
                                    ReDim Preserve aVectorDeIndicesRelacionesDelDiagrama(1 To (IND_VECTOR_DE_RELACIONES + 1)) As Integer
                                    aVectorDeIndicesRelacionesDelDiagrama(IND_VECTOR_DE_RELACIONES + 1) = IND_VECTOR_DE_RELACIONES
                                    
                                    IND_VECTOR_DE_RELACIONES = IND_VECTOR_DE_RELACIONES + 1
                                    PU_SW_DISEÑO_CON_RELACIONES = True
                                End If
                            Next IFK
                        End If
                        
                        .MoveNext
                    Loop
                End With
            End If
        End If
    Next IPK
End Sub

Private Sub UbicaRelacion(ByVal iIndTablaPK As Integer, _
                          ByVal iIndTablaFK As Integer, _
                          ByVal iIndiceRelac As Integer)
    On Error Resume Next
    If iIndTablaPK < 0 Then Exit Sub
    If iIndTablaFK < 0 Then Exit Sub
    
    With imgRelacion_PK(iIndiceRelac)
        mshTabla(iIndTablaPK).Row = 1
        .Left = mshTabla(iIndTablaPK).Left + mshTabla(iIndTablaPK).Width
        .Top = mshTabla(iIndTablaPK).Top + mshTabla(iIndTablaPK).CellTop
        .Visible = True
        If .Left + .Width + Kdbl_Espacio_Entre_Tablas > pctDiseño.Width Then
            pctDiseño.Width = .Left + .Width + Kdbl_Espacio_Entre_Tablas
        End If
    End With
    
    With imgRelacion_FK(iIndiceRelac)
        mshTabla(iIndTablaFK).Row = 1
        .Left = mshTabla(iIndTablaFK).Left - .Width
        .Top = mshTabla(iIndTablaFK).Top + mshTabla(iIndTablaFK).CellTop
        .Visible = True
        If .Left + .Width + Kdbl_Espacio_Entre_Tablas > pctDiseño.Width Then
            pctDiseño.Width = .Left + .Width + Kdbl_Espacio_Entre_Tablas
        End If
    End With
    
    With linRelacionPFK(iIndiceRelac)
        .X1 = imgRelacion_PK(iIndiceRelac).Left + imgRelacion_PK(iIndiceRelac).Width
        .Y1 = imgRelacion_PK(iIndiceRelac).Top + (imgRelacion_PK(iIndiceRelac).Height / 2)
        
        .X2 = imgRelacion_FK(iIndiceRelac).Left
        .Y2 = imgRelacion_FK(iIndiceRelac).Top + (imgRelacion_FK(iIndiceRelac).Height / 2)
        .Visible = True
    End With
End Sub

Private Sub ConfiguraGrillaTabla(Optional ByVal iIndTabla As Integer = -1)
    Dim i As Integer
    Dim F As Integer
    Dim C As Integer
        
    Dim dblBackColorHeader As Double
    Dim dblForecolorHeader As Double
    
    Dim blSW_Key As Boolean
    Dim opcTipCampo As GE_TIPO_TABLA_CLAVE
    
    Dim iIni As Integer
    Dim iFin As Integer
    
    Dim INDICE As Integer
    
    If iIndTabla >= 0 Then
        iIni = 1
        iFin = 1
    Else
        iIni = LBound(aVectorDeIndicesDeTablasDelDiagrama())
        iFin = UBound(aVectorDeIndicesDeTablasDelDiagrama())
    End If
    
    For i = iIni To iFin
        If iIndTabla >= 0 Then
            INDICE = iIndTabla
        Else
            INDICE = aVectorDeIndicesDeTablasDelDiagrama(INDICE)
        End If
        
        Select Case CInt(Val(chkTabla(INDICE).Tag))
            Case GE_TIPO_TABLA.TT_TABLA
                dblBackColorHeader = Kdbl_COLOR_TABLA
                dblForecolorHeader = vbWhite
            Case GE_TIPO_TABLA.TT_VISTA
                dblBackColorHeader = Kdbl_COLOR_VISTA
                dblForecolorHeader = vbWhite
        End Select
        
        With mshTabla(INDICE)
            .Cols = 7
                                                
            .ColWidth(Ki_TABLA_Col_Codigo) = 300
            .ColWidth(Ki_TABLA_Col_Descripcion) = 1350
            .ColWidth(Ki_TABLA_Col_IsPK) = 0
            .ColWidth(Ki_TABLA_Col_IsUK) = 300
            .ColWidth(Ki_TABLA_Col_IsFK) = 300
            .ColWidth(Ki_TABLA_Col_CheckVal) = 0
            .ColWidth(Ki_TABLA_Col_CheckIco) = 300
            
            .RowHeight(0) = Ki_TABLA_ROWHEIGHT_FX1
            .RowHeight(1) = Ki_TABLA_ROWHEIGHT_FX2
            .RowHeight(2) = Ki_TABLA_ROWHEIGHT_FX3
            .MergeCells = flexMergeFree
            
            For C = 0 To .Cols - 1
                
                .TextMatrix(0, C) = Space(5) & lblTabla(INDICE).Caption
                .TextMatrix(1, C) = "."
                .TextMatrix(2, C) = "."
                
                .Col = C
                
                .Row = 0
                .CellAlignment = flexAlignLeftCenter
                .CellBackColor = dblBackColorHeader
                .CellForeColor = dblForecolorHeader
                
                .Row = 1
                .CellBackColor = &H8000000F
                                
                .Row = 2
                .CellBackColor = &H404040
                .CellForeColor = .CellBackColor
            Next C

            Call UbicaControlesDeTabla(INDICE)

            .MergeRow(0) = True
            .MergeRow(1) = True
            .MergeRow(2) = True
            
            For F = .FixedRows To .Rows - 1
                .Row = F
                .RowHeight(F) = Ki_TABLA_ROWHEIGHT
                
                .Col = Ki_TABLA_Col_Codigo
                blSW_Key = CBool(CInt(Val(Trim(.TextMatrix(F, Ki_TABLA_Col_IsPK)))))
                If blSW_Key = True Then
                    Set .CellPicture = imgL.ListImages(Ki_Ico_ColPK).Picture
                End If
                .CellPictureAlignment = flexAlignCenterCenter
                .CellForeColor = .CellBackColor
                .CellAlignment = flexAlignRightCenter
                .CellFontSize = "1"
                
                .Col = Ki_TABLA_Col_IsUK
                blSW_Key = CBool(CInt(Val(Trim(.TextMatrix(F, Ki_TABLA_Col_IsUK)))))
                If blSW_Key = True Then
                    Set .CellPicture = imgL.ListImages(Ki_Ico_ColUK).Picture
                End If
                .CellPictureAlignment = flexAlignCenterCenter
                .CellForeColor = .CellBackColor
                .CellAlignment = flexAlignRightCenter
                .CellFontSize = "1"
                
                .Col = Ki_TABLA_Col_IsFK
                blSW_Key = CBool(CInt(Val(Trim(.TextMatrix(F, Ki_TABLA_Col_IsFK)))))
                If blSW_Key = True Then
                    Set .CellPicture = imgL.ListImages(Ki_Ico_ColFK).Picture
                End If
                .CellPictureAlignment = flexAlignCenterCenter
                .CellForeColor = .CellBackColor
                .CellAlignment = flexAlignRightCenter
                .CellFontSize = "1"
                                
                
                .Col = Ki_TABLA_Col_Descripcion
                .CellTextStyle = flexTextInsetLight
                
                .TextMatrix(F, Ki_TABLA_Col_CheckVal) = GO_ECNLIB00_CONST.VAL_UNCHK
                
                .Col = Ki_TABLA_Col_CheckIco
                Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
                .CellPictureAlignment = flexAlignCenterCenter
            Next F
            .Refresh
        End With
    Next i
    pctDiseño.Refresh
End Sub

Private Sub MostrarControlesDeTabla(ByVal iIndice As Integer, ByVal blSW As Boolean)
    mshTabla(iIndice).Visible = blSW
    cmdTabla(iIndice).Visible = blSW
    pctTabla(iIndice).Visible = blSW
    chkTabla(iIndice).Visible = blSW
    
    If blSW Then
        cmdTabla(iIndice).ZOrder 0
        pctTabla(iIndice).ZOrder 0
        chkTabla(iIndice).ZOrder 0
    End If
End Sub

Private Sub UbicaControlesDeTabla(ByVal iIndice As Integer)
    On Error Resume Next
    
    With mshTabla(iIndice)
        .Row = 0
        .Col = Ki_TABLA_Col_CheckIco
        cmdTabla(iIndice).Left = .Left + .Width - 600
        cmdTabla(iIndice).Top = .Top + .CellTop + 30
                
        .Row = 1
        .Col = Ki_TABLA_Col_Codigo
        chkTabla(iIndice).Left = .Left + .CellLeft + 50
        chkTabla(iIndice).Top = .Top + .CellTop + 10
        
        pctTabla(iIndice).Left = chkTabla(iIndice).Left + 20
        pctTabla(iIndice).Top = .Top + 90
        Select Case CInt(Val(chkTabla(iIndice).Tag))
            Case GE_TIPO_TABLA.TT_TABLA
'ECN : LINEAS COMENTADAS PORQUE ESTA CONFIGURACION YA SE REALIZA AL CREAR EL CONTROL DINAMICAMENTE
'      YA QUE EL CONTROL PRINCIPAL (INDICE 0) TIENE DICHA CONFIGURACION
'                pctTabla(iIndice).Picture = imgL.ListImages(Ki_Ico_Tabla).Picture
'                pctTabla(iIndice).BackColor = Kdbl_COLOR_TABLA
            Case GE_TIPO_TABLA.TT_VISTA
                pctTabla(iIndice).Picture = imgL.ListImages(Ki_Ico_Vista).Picture
                pctTabla(iIndice).BackColor = Kdbl_COLOR_VISTA
        End Select
        .ZOrder 0
    End With
    Call MostrarControlesDeTabla(iIndice, True)
End Sub



Private Sub ReUbicaRelaciones(ByVal iIndTabla As Integer)
    If iIndTabla < 0 Then Exit Sub
    
    Dim i As Integer
    Dim INDICE As Integer
    Dim iIndPK As Integer
    Dim iIndFK As Integer
    
    For i = LBound(aVectorDeIndicesRelacionesDelDiagrama()) To _
            UBound(aVectorDeIndicesRelacionesDelDiagrama())
        INDICE = aVectorDeIndicesRelacionesDelDiagrama(i)
        Call EncuentraIndiceTablaPKyFKdesdeUnaLineaRelacion(iIndPK, iIndFK, INDICE)
        If iIndPK = iIndTabla Or iIndFK = iIndTabla Then
            Call UbicaRelacion(iIndPK, iIndFK, INDICE)
        End If
    Next i
End Sub

Private Sub EncuentraIndiceTablaPKyFKdesdeUnaLineaRelacion(ByRef iIndTabla_PK As Integer, _
                                                           ByRef iIndTabla_FK As Integer, _
                                                           ByVal iIndRelacion As Integer)
    On Error Resume Next
    
    iIndTabla_PK = -1
    iIndTabla_FK = -1
    If iIndRelacion < 0 Then Exit Sub
   
    Dim iPosPK As Integer
    Dim iPosFK As Integer
    Dim iPosSP As Integer
   
    With linRelacionPFK(iIndRelacion)
        iPosPK = InStr(1, .Tag, Ks_CarPK, vbTextCompare)
        iPosFK = InStr(1, .Tag, Ks_CarFK, vbTextCompare)
        iPosSP = InStr(1, .Tag, Ks_CarSepPKFK, vbTextCompare)
        
        If iPosPK > 0 Then
            iIndTabla_PK = Mid(.Tag, iPosPK + 1, iPosSP - iPosPK - 1)
        End If
        If iPosFK > 0 Then
            iIndTabla_FK = Mid(.Tag, iPosFK + 1, Len(.Tag) - iPosFK)
        End If
    End With
End Sub

Private Function EncuentraIndiceDeTablaPorRelacionFK(ByVal iIndImgFK As Integer) As Integer
    EncuentraIndiceDeTablaPorRelacionFK = -1
    If iIndImgFK < 0 Then Exit Function
    
    Dim iPosFK As Integer
    iPosFK = InStr(1, linRelacionPFK(iIndImgFK).Tag, Ks_CarFK, vbTextCompare)
        
    If iPosFK > 0 Then
        EncuentraIndiceDeTablaPorRelacionFK = Mid(linRelacionPFK(iIndImgFK).Tag, _
                                                  iPosFK + 1, _
                                                  Len(linRelacionPFK(iIndImgFK).Tag) - iPosFK)
    End If
End Function

Private Sub ConfiguraMSHCampos(Optional ByVal iFila As Integer = -1, _
                               Optional ByVal blWait As Boolean = False)
    On Error Resume Next
    
    If blWait Then
        pctDiseñoDatos.Visible = True
        ecnPbrCir_Wait.Interval = 5
        Me.Refresh
    End If
    
    Dim F As Integer
    Dim C As Integer
    
    Dim iFilaINI As Integer
    Dim iFilaFIN As Integer
    Dim blSWFullDesign As Boolean
    
    Dim iAnchoDeBarra As Integer
    iAnchoDeBarra = 50
    
    With mshDiseñoDatos
        If iFila < .FixedRows Then
            iFilaINI = .FixedRows
            iFilaFIN = .Rows - 1
            blSWFullDesign = True
        Else
            iFilaINI = iFila
            iFilaFIN = iFila
            blSWFullDesign = False
        End If
        
        If blSWFullDesign = True Then
            Call ConfiguraMSHCampos_Cabecera(mshDiseñoDatos)
            If blSW_Load = True Then
                Call ConfiguraMSHCampos_Cabecera(mshFinta)
                Call CargarDatosTipoSQLJoin(0)
                Call CargarDatosTipoDeCampo(0)
                Call CargarDatosTipoDeOrdenamiento(0)
            End If
        End If
        
        Dim iColumnAux As Integer
        
        For F = iFilaINI To iFilaFIN
            .RowHeight(F) = 300
            .Row = F
                        
            For C = 0 To .Cols - 1
                .Col = C
                Select Case C
                    Case Ki_MSH_CAMPOS_01_FIXED__COL
                        .CellFontName = "Webdings"
                        .CellBackColor = .GridColor
                        .CellFontSize = 12
                        .CellTextStyle = flexTextInset
                        .CellForeColor = vbBlue
                    Case Ki_MSH_CAMPOS_03_TABLA_JOIN, _
                         Ki_MSH_CAMPOS_05_SELEC_TIPO, _
                         Ki_MSH_CAMPOS_32_ORDER_TIPO
                        .CellFontName = "Tahoma"
                        .CellFontSize = 1
                        .CellPictureAlignment = flexAlignCenterCenter
                        .CellForeColor = .CellBackColor
                        If blSW_Load = False Then
                            Select Case C
                                Case Ki_MSH_CAMPOS_03_TABLA_JOIN
                                    Select Case .TextMatrix(F, C)
                                        Case GO_002_Ks_TIPO_DE_JOIN_FR: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_From).Picture
                                        Case GO_002_Ks_TIPO_DE_JOIN_IN: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_InnerJoin).Picture
                                        Case GO_002_Ks_TIPO_DE_JOIN_LF: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_LeftJoin).Picture
                                        Case GO_002_Ks_TIPO_DE_JOIN_RI: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_RightJoin).Picture
                                        Case GO_002_Ks_TIPO_DE_JOIN_UN: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_Union).Picture
                                        Case Else
                                            Set .CellPicture = Nothing
                                    End Select
                                Case Ki_MSH_CAMPOS_05_SELEC_TIPO
                                    Select Case .TextMatrix(F, C)
                                        Case GO_002_Ks_TIPO_DE_CAMPO_FD: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_Column).Picture
                                        Case GO_002_Ks_TIPO_DE_CAMPO_TX: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_Text).Picture
                                        Case GO_002_Ks_TIPO_DE_CAMPO_FX: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_Fx).Picture
                                        Case GO_002_Ks_TIPO_DE_CAMPO_AD: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_Zigma).Picture
                                        Case Else
                                            Set .CellPicture = Nothing
                                    End Select
                                Case Ki_MSH_CAMPOS_32_ORDER_TIPO
                                    Select Case .TextMatrix(F, C)
                                        Case GO_002_Ks_TIPO_DE_ORDENAMIENTO_ASC: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_SortAsc).Picture
                                        Case GO_002_Ks_TIPO_DE_ORDENAMIENTO_DES: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_SortDes).Picture
                                        Case Else
                                            Set .CellPicture = Nothing
                                    End Select
                            End Select
                        End If
                    Case Ki_MSH_CAMPOS_04_BARRA_0001, Ki_MSH_CAMPOS_10_BARRA_0002, _
                         Ki_MSH_CAMPOS_17_BARRA_0003, Ki_MSH_CAMPOS_21_BARRA_0004, _
                         Ki_MSH_CAMPOS_28_BARRA_0005, Ki_MSH_CAMPOS_33_BARRA_0006
                        .CellBackColor = .GridColor
                    Case Ki_MSH_CAMPOS_11_WHERE_CHKI, Ki_MSH_CAMPOS_18_GROUP_CHKI, _
                         Ki_MSH_CAMPOS_22_HAVIN_CHKI, Ki_MSH_CAMPOS_29_ORDER_CHKI
                         If blSW_Load = True Then
                            Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
                         Else
                            Select Case C
                                Case Ki_MSH_CAMPOS_11_WHERE_CHKI: iColumnAux = Ki_MSH_CAMPOS_12_WHERE_CHKV
                                Case Ki_MSH_CAMPOS_18_GROUP_CHKI: iColumnAux = Ki_MSH_CAMPOS_19_GROUP_CHKV
                                Case Ki_MSH_CAMPOS_22_HAVIN_CHKI: iColumnAux = Ki_MSH_CAMPOS_23_HAVIN_CHKV
                                Case Ki_MSH_CAMPOS_29_ORDER_CHKI: iColumnAux = Ki_MSH_CAMPOS_30_ORDER_CHKV
                            End Select
                            Select Case .TextMatrix(F, iColumnAux)
                                Case GO_ECNLIB00_CONST.VAL_CHECK: Set .CellPicture = imgL.ListImages(Ki_Ico_Check).Picture
                                Case GO_ECNLIB00_CONST.VAL_UNCHK: Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
                            End Select
                         End If
                        .CellPictureAlignment = flexAlignCenterCenter
                    Case Ki_MSH_CAMPOS_08_SELEC_ACTI
                            .CellPictureAlignment = flexAlignCenterCenter
                    Case Ki_MSH_CAMPOS_09_SELEC_ACTV, Ki_MSH_CAMPOS_12_WHERE_CHKV, Ki_MSH_CAMPOS_19_GROUP_CHKV, _
                         Ki_MSH_CAMPOS_23_HAVIN_CHKV, Ki_MSH_CAMPOS_30_ORDER_CHKV
                         If blSW_Load = True Then
                            .TextMatrix(F, C) = GO_ECNLIB00_CONST.VAL_UNCHK
                         End If
                    Case Ki_MSH_CAMPOS_13_WHERE_TIPO
                        .CellForeColor = cboTipoWHERE(0).ForeColor
                        .CellFontName = "verdana"
                        .CellFontSize = 7
                        .CellFontBold = True
                    Case Ki_MSH_CAMPOS_24_HAVIN_TIPO
                        .CellForeColor = cboTipoHAVING(0).ForeColor
                        .CellFontName = "verdana"
                        .CellFontSize = 7
                        .CellFontBold = True
                    Case Ki_MSH_CAMPOS_20_GROUP_NUME, _
                         Ki_MSH_CAMPOS_31_ORDER_NUME
                        .CellFontBold = True
                        .CellTextStyle = flexTextInsetLight
                End Select
            Next C
            DoEvents
        Next F
        .Row = 0
        .Col = 0
    End With
    
    If blSWFullDesign = True Then
        Call ReEnumeraCamposEnDgMatricial
        Call UbicaImgDeMSHCab
    End If
    
    If blWait Then
        ecnPbrCir_Wait.Interval = 0
        pctDiseñoDatos.Visible = False
        Me.Refresh
    End If
End Sub


Private Sub ConfiguraMSHCampos_Cabecera(ByRef ctrlMSH As MSHFlexGrid)
    On Error Resume Next
    Dim F As Integer
    Dim C As Integer
    
    Dim iAnchoDeBarra As Integer
    iAnchoDeBarra = 50
    
    With ctrlMSH 'mshDiseñoDatos
        .Cols = Ki_MSH_CAMPOS_33_BARRA_0006 + 1
        .BackColorFixed = RGB(200, 214, 228)
                    
        .ColWidth(Ki_MSH_CAMPOS_00_FIXED__COL) = 280
        .ColWidth(Ki_MSH_CAMPOS_01_FIXED__COL) = 200
        .ColWidth(Ki_MSH_CAMPOS_02_TABLA_NOMB) = 1100
        .ColWidth(Ki_MSH_CAMPOS_03_TABLA_JOIN) = 310
        .ColWidth(Ki_MSH_CAMPOS_04_BARRA_0001) = iAnchoDeBarra
        .ColWidth(Ki_MSH_CAMPOS_05_SELEC_TIPO) = 310
        .ColWidth(Ki_MSH_CAMPOS_06_SELEC_NOMB) = 1050
        .ColWidth(Ki_MSH_CAMPOS_07_SELEC_ALEA) = 1000
        .ColWidth(Ki_MSH_CAMPOS_08_SELEC_ACTI) = 350
        .ColWidth(Ki_MSH_CAMPOS_09_SELEC_ACTV) = 0
        .ColWidth(Ki_MSH_CAMPOS_10_BARRA_0002) = iAnchoDeBarra
        .ColWidth(Ki_MSH_CAMPOS_11_WHERE_CHKI) = Ki_MSH_CAMPOS_11_WHERE_CHKI_WIDTH
        .ColWidth(Ki_MSH_CAMPOS_12_WHERE_CHKV) = Ki_MSH_CAMPOS_12_WHERE_CHKV_WIDTH
        .ColWidth(Ki_MSH_CAMPOS_13_WHERE_TIPO) = Ki_MSH_CAMPOS_13_WHERE_TIPO_WIDTH
        .ColWidth(Ki_MSH_CAMPOS_14_WHERE_OPER) = Ki_MSH_CAMPOS_14_WHERE_OPER_WIDTH
        .ColWidth(Ki_MSH_CAMPOS_15_WHERE_CRI1) = Ki_MSH_CAMPOS_15_WHERE_CRI1_WIDTH
        .ColWidth(Ki_MSH_CAMPOS_16_WHERE_CRI2) = Ki_MSH_CAMPOS_16_WHERE_CRI2_WIDTH
        .ColWidth(Ki_MSH_CAMPOS_17_BARRA_0003) = iAnchoDeBarra
        .ColWidth(Ki_MSH_CAMPOS_18_GROUP_CHKI) = Ki_MSH_CAMPOS_18_GROUP_CHKI_WIDTH
        .ColWidth(Ki_MSH_CAMPOS_19_GROUP_CHKV) = Ki_MSH_CAMPOS_19_GROUP_CHKV_WIDTH
        .ColWidth(Ki_MSH_CAMPOS_20_GROUP_NUME) = Ki_MSH_CAMPOS_20_GROUP_NUME_WIDTH
        .ColWidth(Ki_MSH_CAMPOS_21_BARRA_0004) = iAnchoDeBarra
        .ColWidth(Ki_MSH_CAMPOS_22_HAVIN_CHKI) = Ki_MSH_CAMPOS_22_HAVIN_CHKI_WIDTH
        .ColWidth(Ki_MSH_CAMPOS_23_HAVIN_CHKV) = Ki_MSH_CAMPOS_23_HAVIN_CHKV_WIDTH
        .ColWidth(Ki_MSH_CAMPOS_24_HAVIN_TIPO) = Ki_MSH_CAMPOS_24_HAVIN_TIPO_WIDTH
        .ColWidth(Ki_MSH_CAMPOS_25_HAVIN_OPER) = Ki_MSH_CAMPOS_25_HAVIN_OPER_WIDTH
        .ColWidth(Ki_MSH_CAMPOS_26_HAVIN_CRI1) = Ki_MSH_CAMPOS_26_HAVIN_CRI1_WIDTH
        .ColWidth(Ki_MSH_CAMPOS_27_HAVIN_CRI2) = Ki_MSH_CAMPOS_27_HAVIN_CRI2_WIDTH
        .ColWidth(Ki_MSH_CAMPOS_28_BARRA_0005) = iAnchoDeBarra
        .ColWidth(Ki_MSH_CAMPOS_29_ORDER_CHKI) = 300
        .ColWidth(Ki_MSH_CAMPOS_30_ORDER_CHKV) = 0
        .ColWidth(Ki_MSH_CAMPOS_31_ORDER_NUME) = 300
        .ColWidth(Ki_MSH_CAMPOS_32_ORDER_TIPO) = 310
        .ColWidth(Ki_MSH_CAMPOS_33_BARRA_0006) = iAnchoDeBarra
        
         
        .ColAlignment(Ki_MSH_CAMPOS_00_FIXED__COL) = flexAlignRightCenter
        .ColAlignment(Ki_MSH_CAMPOS_01_FIXED__COL) = flexAlignCenterCenter
        .ColAlignment(Ki_MSH_CAMPOS_02_TABLA_NOMB) = flexAlignLeftCenter
        .ColAlignment(Ki_MSH_CAMPOS_03_TABLA_JOIN) = flexAlignRightCenter
        .ColAlignment(Ki_MSH_CAMPOS_04_BARRA_0001) = flexAlignRightCenter
        .ColAlignment(Ki_MSH_CAMPOS_05_SELEC_TIPO) = flexAlignCenterCenter
        .ColAlignment(Ki_MSH_CAMPOS_06_SELEC_NOMB) = flexAlignLeftCenter
        .ColAlignment(Ki_MSH_CAMPOS_07_SELEC_ALEA) = flexAlignLeftCenter
        .ColAlignment(Ki_MSH_CAMPOS_08_SELEC_ACTI) = flexAlignCenterCenter
        .ColAlignment(Ki_MSH_CAMPOS_09_SELEC_ACTV) = flexAlignCenterCenter
        .ColAlignment(Ki_MSH_CAMPOS_10_BARRA_0002) = flexAlignRightCenter
        .ColAlignment(Ki_MSH_CAMPOS_11_WHERE_CHKI) = flexAlignCenterCenter
        .ColAlignment(Ki_MSH_CAMPOS_12_WHERE_CHKV) = flexAlignRightCenter
        .ColAlignment(Ki_MSH_CAMPOS_13_WHERE_TIPO) = flexAlignCenterCenter
        .ColAlignment(Ki_MSH_CAMPOS_14_WHERE_OPER) = flexAlignCenterCenter
        .ColAlignment(Ki_MSH_CAMPOS_15_WHERE_CRI1) = flexAlignLeftCenter
        .ColAlignment(Ki_MSH_CAMPOS_16_WHERE_CRI2) = flexAlignLeftCenter
        .ColAlignment(Ki_MSH_CAMPOS_17_BARRA_0003) = flexAlignRightCenter
        .ColAlignment(Ki_MSH_CAMPOS_18_GROUP_CHKI) = flexAlignCenterCenter
        .ColAlignment(Ki_MSH_CAMPOS_19_GROUP_CHKV) = flexAlignRightCenter
        .ColAlignment(Ki_MSH_CAMPOS_20_GROUP_NUME) = flexAlignRightCenter
        .ColAlignment(Ki_MSH_CAMPOS_21_BARRA_0004) = flexAlignRightCenter
        .ColAlignment(Ki_MSH_CAMPOS_22_HAVIN_CHKI) = flexAlignCenterCenter
        .ColAlignment(Ki_MSH_CAMPOS_23_HAVIN_CHKV) = flexAlignRightCenter
        .ColAlignment(Ki_MSH_CAMPOS_24_HAVIN_TIPO) = flexAlignCenterCenter
        .ColAlignment(Ki_MSH_CAMPOS_25_HAVIN_OPER) = flexAlignCenterCenter
        .ColAlignment(Ki_MSH_CAMPOS_26_HAVIN_CRI1) = flexAlignLeftCenter
        .ColAlignment(Ki_MSH_CAMPOS_27_HAVIN_CRI2) = flexAlignLeftCenter
        .ColAlignment(Ki_MSH_CAMPOS_28_BARRA_0005) = flexAlignRightCenter
        .ColAlignment(Ki_MSH_CAMPOS_29_ORDER_CHKI) = flexAlignCenterCenter
        .ColAlignment(Ki_MSH_CAMPOS_30_ORDER_CHKV) = flexAlignRightCenter
        .ColAlignment(Ki_MSH_CAMPOS_31_ORDER_NUME) = flexAlignRightCenter
        .ColAlignment(Ki_MSH_CAMPOS_32_ORDER_TIPO) = flexAlignCenterCenter
        .ColAlignment(Ki_MSH_CAMPOS_33_BARRA_0006) = flexAlignRightCenter
        

        .TextMatrix(0, Ki_MSH_CAMPOS_00_FIXED__COL) = Space(30) & "."
        .TextMatrix(0, Ki_MSH_CAMPOS_01_FIXED__COL) = Space(30) & GO_ECNLIB00_CONST.CARESP_WEB_DERECHA
        .TextMatrix(0, Ki_MSH_CAMPOS_02_TABLA_NOMB) = "    TABLA"
        .TextMatrix(0, Ki_MSH_CAMPOS_03_TABLA_JOIN) = "    TABLA"
        .TextMatrix(0, Ki_MSH_CAMPOS_04_BARRA_0001) = "BARRA"
        .TextMatrix(0, Ki_MSH_CAMPOS_05_SELEC_TIPO) = "CAMPO"
        .TextMatrix(0, Ki_MSH_CAMPOS_06_SELEC_NOMB) = "CAMPO"
        .TextMatrix(0, Ki_MSH_CAMPOS_07_SELEC_ALEA) = "CAMPO"
        .TextMatrix(0, Ki_MSH_CAMPOS_08_SELEC_ACTI) = GO_ECNLIB00_CONST.CARESP_CHK_CHECKED_01
        .TextMatrix(0, Ki_MSH_CAMPOS_10_BARRA_0002) = "BARRA"
        .TextMatrix(0, Ki_MSH_CAMPOS_11_WHERE_CHKI) = "FILTRO"
        .TextMatrix(0, Ki_MSH_CAMPOS_12_WHERE_CHKV) = "FILTRO"
        .TextMatrix(0, Ki_MSH_CAMPOS_13_WHERE_TIPO) = "FILTRO"
        .TextMatrix(0, Ki_MSH_CAMPOS_14_WHERE_OPER) = "FILTRO"
        .TextMatrix(0, Ki_MSH_CAMPOS_15_WHERE_CRI1) = "FILTRO"
        .TextMatrix(0, Ki_MSH_CAMPOS_16_WHERE_CRI2) = "FILTRO"
        .TextMatrix(0, Ki_MSH_CAMPOS_17_BARRA_0003) = "BARRA"
        .TextMatrix(0, Ki_MSH_CAMPOS_18_GROUP_CHKI) = "." '"GRUP."
        .TextMatrix(0, Ki_MSH_CAMPOS_19_GROUP_CHKV) = "." 'GRUP."
        .TextMatrix(0, Ki_MSH_CAMPOS_20_GROUP_NUME) = "." 'GRUP."
        .TextMatrix(0, Ki_MSH_CAMPOS_21_BARRA_0004) = "BARRA"
        .TextMatrix(0, Ki_MSH_CAMPOS_22_HAVIN_CHKI) = "FILTRO DE AGRUPAMIENTO"
        .TextMatrix(0, Ki_MSH_CAMPOS_23_HAVIN_CHKV) = "FILTRO DE AGRUPAMIENTO"
        .TextMatrix(0, Ki_MSH_CAMPOS_24_HAVIN_TIPO) = "FILTRO DE AGRUPAMIENTO"
        .TextMatrix(0, Ki_MSH_CAMPOS_25_HAVIN_OPER) = "FILTRO DE AGRUPAMIENTO"
        .TextMatrix(0, Ki_MSH_CAMPOS_26_HAVIN_CRI1) = "FILTRO DE AGRUPAMIENTO"
        .TextMatrix(0, Ki_MSH_CAMPOS_27_HAVIN_CRI2) = "FILTRO DE AGRUPAMIENTO"
        .TextMatrix(0, Ki_MSH_CAMPOS_28_BARRA_0005) = "BARRA"
        .TextMatrix(0, Ki_MSH_CAMPOS_29_ORDER_CHKI) = "     ORDEN"
        .TextMatrix(0, Ki_MSH_CAMPOS_30_ORDER_CHKV) = "     ORDEN"
        .TextMatrix(0, Ki_MSH_CAMPOS_31_ORDER_NUME) = "     ORDEN"
        .TextMatrix(0, Ki_MSH_CAMPOS_32_ORDER_TIPO) = "     ORDEN"
        .TextMatrix(0, Ki_MSH_CAMPOS_33_BARRA_0006) = "BARRA"
        
        
        .TextMatrix(1, Ki_MSH_CAMPOS_00_FIXED__COL) = Space(30) & "."
        .TextMatrix(1, Ki_MSH_CAMPOS_01_FIXED__COL) = Space(30) & GO_ECNLIB00_CONST.CARESP_WEB_DERECHA
        .TextMatrix(1, Ki_MSH_CAMPOS_02_TABLA_NOMB) = "NOMBRE"
        .TextMatrix(1, Ki_MSH_CAMPOS_03_TABLA_JOIN) = "TP"
        .TextMatrix(1, Ki_MSH_CAMPOS_04_BARRA_0001) = "BARRA"
        .TextMatrix(1, Ki_MSH_CAMPOS_05_SELEC_TIPO) = "TP"
        .TextMatrix(1, Ki_MSH_CAMPOS_06_SELEC_NOMB) = "NOMBRE"
        .TextMatrix(1, Ki_MSH_CAMPOS_07_SELEC_ALEA) = "ALEAS"
        .TextMatrix(1, Ki_MSH_CAMPOS_08_SELEC_ACTI) = GO_ECNLIB00_CONST.CARESP_CHK_CHECKED_01
        .TextMatrix(1, Ki_MSH_CAMPOS_10_BARRA_0002) = "BARRA"
        .TextMatrix(1, Ki_MSH_CAMPOS_11_WHERE_CHKI) = ".?"
        .TextMatrix(1, Ki_MSH_CAMPOS_12_WHERE_CHKV) = Empty
        .TextMatrix(1, Ki_MSH_CAMPOS_13_WHERE_TIPO) = "TP"
        .TextMatrix(1, Ki_MSH_CAMPOS_14_WHERE_OPER) = "OPE."
        .TextMatrix(1, Ki_MSH_CAMPOS_15_WHERE_CRI1) = "CRI.01"
        .TextMatrix(1, Ki_MSH_CAMPOS_16_WHERE_CRI2) = "CRI.02"
        .TextMatrix(1, Ki_MSH_CAMPOS_17_BARRA_0003) = "BARRA"
        .TextMatrix(1, Ki_MSH_CAMPOS_18_GROUP_CHKI) = ".?"
        .TextMatrix(1, Ki_MSH_CAMPOS_19_GROUP_CHKV) = Empty
        .TextMatrix(1, Ki_MSH_CAMPOS_20_GROUP_NUME) = "N°"
        .TextMatrix(1, Ki_MSH_CAMPOS_21_BARRA_0004) = "BARRA"
        .TextMatrix(1, Ki_MSH_CAMPOS_22_HAVIN_CHKI) = ".?"
        .TextMatrix(1, Ki_MSH_CAMPOS_23_HAVIN_CHKV) = Empty
        .TextMatrix(1, Ki_MSH_CAMPOS_24_HAVIN_TIPO) = "TP"
        .TextMatrix(1, Ki_MSH_CAMPOS_25_HAVIN_OPER) = "OPE."
        .TextMatrix(1, Ki_MSH_CAMPOS_26_HAVIN_CRI1) = "CRI.01"
        .TextMatrix(1, Ki_MSH_CAMPOS_27_HAVIN_CRI2) = "CRI.02"
        .TextMatrix(1, Ki_MSH_CAMPOS_28_BARRA_0005) = "BARRA"
        .TextMatrix(1, Ki_MSH_CAMPOS_29_ORDER_CHKI) = ".?"
        .TextMatrix(1, Ki_MSH_CAMPOS_30_ORDER_CHKV) = Empty
        .TextMatrix(1, Ki_MSH_CAMPOS_31_ORDER_NUME) = "N°"
        .TextMatrix(1, Ki_MSH_CAMPOS_32_ORDER_TIPO) = "TP"
        .TextMatrix(1, Ki_MSH_CAMPOS_33_BARRA_0006) = "BARRA"
        
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        
        .MergeCol(Ki_MSH_CAMPOS_02_TABLA_NOMB) = True
        .MergeCol(Ki_MSH_CAMPOS_03_TABLA_JOIN) = True
        For C = 0 To .Cols - 1
            .Col = C
            .ColAlignmentHeader(C) = flexAlignCenterCenter
            Select Case C
                Case Ki_MSH_CAMPOS_00_FIXED__COL
                    .MergeCol(C) = True
                Case Ki_MSH_CAMPOS_01_FIXED__COL
                    .MergeCol(C) = True
                    .Row = 1
                    .CellFontName = "Webdings"
                    .CellFontSize = 1
                    .CellBackColor = .GridColor
                    .CellForeColor = .CellBackColor
                Case Ki_MSH_CAMPOS_02_TABLA_NOMB, _
                     Ki_MSH_CAMPOS_03_TABLA_JOIN
                    .MergeCol(C) = True
                    GoTo SALTO_FIXED_ROW_1
                Case Ki_MSH_CAMPOS_04_BARRA_0001, Ki_MSH_CAMPOS_10_BARRA_0002, _
                     Ki_MSH_CAMPOS_17_BARRA_0003, Ki_MSH_CAMPOS_21_BARRA_0004, _
                     Ki_MSH_CAMPOS_28_BARRA_0005, Ki_MSH_CAMPOS_33_BARRA_0006
                    .MergeCol(C) = True
                    .CellBackColor = .GridColor
                    .CellForeColor = .CellBackColor
                    .CellTextStyle = flexTextFlat
                Case Ki_MSH_CAMPOS_11_WHERE_CHKI, Ki_MSH_CAMPOS_18_GROUP_CHKI, _
                     Ki_MSH_CAMPOS_22_HAVIN_CHKI, Ki_MSH_CAMPOS_29_ORDER_CHKI
                     .Row = 1
                     .CellForeColor = vbBlue
                     .CellTextStyle = flexTextRaised
                     .CellFontSize = 10
                     .CellBackColor = RGB(248, 247, 246)
                Case Ki_MSH_CAMPOS_08_SELEC_ACTI
                    .MergeCol(C) = True
                    .Row = 0: .CellForeColor = RGB(0, 204, 0): .CellBackColor = RGB(248, 247, 246): .CellFontName = "Wingdings": .CellFontSize = "18"
                    .Row = 1: .CellForeColor = RGB(0, 204, 0): .CellBackColor = RGB(248, 247, 246): .CellFontName = "Wingdings": .CellFontSize = "18"
                Case Else
SALTO_FIXED_ROW_1:
                    .Row = 1
                    .CellFontBold = False
                    .CellBackColor = RGB(248, 247, 246)
                    .CellTextStyle = flexTextFlat
                    .CellForeColor = &H808080
            End Select
            .Row = 0: .CellAlignment = flexAlignCenterCenter
            .Row = 1: .CellAlignment = flexAlignCenterCenter
        Next C

        .RowHeight(0) = 450
        .RowHeight(1) = 300
    End With
    Call UbicaImgDeMSHCab
End Sub


Private Sub UbicaImgDeMSHCab()
    With mshDiseñoDatos
        .Row = 0
        
        .Col = Ki_MSH_CAMPOS_02_TABLA_NOMB
        pctMSHCab(0).Left = .Left + .CellLeft + 250
        pctMSHCab(0).Top = .Top + (.CellHeight / 2) - (pctMSHCab(0).Height / 2) + 30
        pctMSHCab(0).BackColor = .BackColorFixed
        
        .Col = Ki_MSH_CAMPOS_07_SELEC_ALEA
        pctMSHCab(1).Left = .Left + .CellLeft + 600
        pctMSHCab(1).Top = .Top + (.CellHeight / 2) - 80
        pctMSHCab(1).BackColor = .BackColorFixed
        
        .Col = Ki_MSH_CAMPOS_14_WHERE_OPER
        pctMSHCab(2).Left = .Left + .CellLeft + 1080
        pctMSHCab(2).Top = pctMSHCab(1).Top
        pctMSHCab(2).BackColor = .BackColorFixed
        
        .Col = Ki_MSH_CAMPOS_20_GROUP_NUME
        pctMSHCab(3).Left = .Left + .CellLeft + (.CellWidth / 2) - (pctMSHCab(3).Width / 2)
        pctMSHCab(3).Top = pctMSHCab(1).Top
        pctMSHCab(3).BackColor = .BackColorFixed
        
        .Col = Ki_MSH_CAMPOS_25_HAVIN_OPER
        pctMSHCab(4).Left = .Left + .CellLeft + 240
        pctMSHCab(4).Top = pctMSHCab(1).Top
        pctMSHCab(4).BackColor = .BackColorFixed
        
        .Col = Ki_MSH_CAMPOS_32_ORDER_TIPO
        pctMSHCab(5).Left = .Left + .CellLeft + 50
        pctMSHCab(5).Top = pctMSHCab(1).Top
        pctMSHCab(5).BackColor = .BackColorFixed
    End With
End Sub

Private Sub MNMAIN_Click(ByVal ID As Long)
    If EjecucionMenuDgGrafico(MNMAIN, ID) = True Then
    ElseIf EjecucionMenuDgMatricial(MNMAIN, ID) = True Then
    ElseIf MNMAIN.MenuItems.Key(ID) = Ks_MNMAIN_00_04___ Then
        Call MENU_OPC_00_04
    End If
End Sub
  
Private Sub MENU_OPC_00_01_01()
    Call AgregarTabla
End Sub

Private Sub MENU_OPC_00_01_02()
    Call BorraDiseño
End Sub

Private Sub MENU_OPC_00_01_03()
    Call MiniMaximizarTodasLasTablas(True)
End Sub

Private Sub MENU_OPC_00_01_04()
    Call MiniMaximizarTodasLasTablas(False)
End Sub

Private Sub MENU_OPC_00_02_01(ByVal blSW As Boolean)
    Dim F As Integer
    
    With mshDiseñoDatos
        If blSW = True Then
            .ColWidth(Ki_MSH_CAMPOS_11_WHERE_CHKI) = Ki_MSH_CAMPOS_11_WHERE_CHKI_WIDTH
            .ColWidth(Ki_MSH_CAMPOS_12_WHERE_CHKV) = Ki_MSH_CAMPOS_12_WHERE_CHKV_WIDTH
            .ColWidth(Ki_MSH_CAMPOS_13_WHERE_TIPO) = Ki_MSH_CAMPOS_13_WHERE_TIPO_WIDTH
            .ColWidth(Ki_MSH_CAMPOS_14_WHERE_OPER) = Ki_MSH_CAMPOS_14_WHERE_OPER_WIDTH
            .ColWidth(Ki_MSH_CAMPOS_15_WHERE_CRI1) = Ki_MSH_CAMPOS_15_WHERE_CRI1_WIDTH
            .ColWidth(Ki_MSH_CAMPOS_16_WHERE_CRI2) = Ki_MSH_CAMPOS_16_WHERE_CRI2_WIDTH
            
            .Col = Ki_MSH_CAMPOS_11_WHERE_CHKI
            For F = .FixedRows To .Rows - 1
                .Row = F
                Select Case .TextMatrix(F, Ki_MSH_CAMPOS_12_WHERE_CHKV)
                    Case GO_ECNLIB00_CONST.VAL_CHECK: Set .CellPicture = imgL.ListImages(Ki_Ico_Check).Picture
                    Case GO_ECNLIB00_CONST.VAL_UNCHK: Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
                End Select
            Next F
        Else
            .ColWidth(Ki_MSH_CAMPOS_11_WHERE_CHKI) = 0
            .ColWidth(Ki_MSH_CAMPOS_12_WHERE_CHKV) = 0
            .ColWidth(Ki_MSH_CAMPOS_13_WHERE_TIPO) = 0
            .ColWidth(Ki_MSH_CAMPOS_14_WHERE_OPER) = 0
            .ColWidth(Ki_MSH_CAMPOS_15_WHERE_CRI1) = 0
            .ColWidth(Ki_MSH_CAMPOS_16_WHERE_CRI2) = 0
            
            .Col = Ki_MSH_CAMPOS_11_WHERE_CHKI
            For F = .FixedRows To .Rows - 1
                .Row = F
                Set .CellPicture = Nothing
            Next F
        End If
        .Refresh
    End With
End Sub

Private Sub MENU_OPC_00_02_02(ByVal blSW As Boolean)
    Dim F As Integer
    
    With mshDiseñoDatos
        If blSW = True Then
            .ColWidth(Ki_MSH_CAMPOS_18_GROUP_CHKI) = Ki_MSH_CAMPOS_18_GROUP_CHKI_WIDTH
            .ColWidth(Ki_MSH_CAMPOS_19_GROUP_CHKV) = Ki_MSH_CAMPOS_19_GROUP_CHKV_WIDTH
            .ColWidth(Ki_MSH_CAMPOS_20_GROUP_NUME) = Ki_MSH_CAMPOS_20_GROUP_NUME_WIDTH
            
            .Col = Ki_MSH_CAMPOS_18_GROUP_CHKI
            For F = .FixedRows To .Rows - 1
                .Row = F
                Select Case .TextMatrix(F, Ki_MSH_CAMPOS_19_GROUP_CHKV)
                    Case GO_ECNLIB00_CONST.VAL_CHECK: Set .CellPicture = imgL.ListImages(Ki_Ico_Check).Picture
                    Case GO_ECNLIB00_CONST.VAL_UNCHK: Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
                End Select
            Next F
        Else
            .ColWidth(Ki_MSH_CAMPOS_18_GROUP_CHKI) = 0
            .ColWidth(Ki_MSH_CAMPOS_19_GROUP_CHKV) = 0
            .ColWidth(Ki_MSH_CAMPOS_20_GROUP_NUME) = 0
            
            .Col = Ki_MSH_CAMPOS_18_GROUP_CHKI
            For F = .FixedRows To .Rows - 1
                .Row = F
                Set .CellPicture = Nothing
            Next F
        End If
        .Refresh
    End With
End Sub

Private Sub MENU_OPC_00_02_03(ByVal blSW As Boolean)
    Dim F As Integer
    
    With mshDiseñoDatos
        If blSW = True Then
            .ColWidth(Ki_MSH_CAMPOS_22_HAVIN_CHKI) = Ki_MSH_CAMPOS_22_HAVIN_CHKI_WIDTH
            .ColWidth(Ki_MSH_CAMPOS_23_HAVIN_CHKV) = Ki_MSH_CAMPOS_23_HAVIN_CHKV_WIDTH
            .ColWidth(Ki_MSH_CAMPOS_24_HAVIN_TIPO) = Ki_MSH_CAMPOS_24_HAVIN_TIPO_WIDTH
            .ColWidth(Ki_MSH_CAMPOS_25_HAVIN_OPER) = Ki_MSH_CAMPOS_25_HAVIN_OPER_WIDTH
            .ColWidth(Ki_MSH_CAMPOS_26_HAVIN_CRI1) = Ki_MSH_CAMPOS_26_HAVIN_CRI1_WIDTH
            .ColWidth(Ki_MSH_CAMPOS_27_HAVIN_CRI2) = Ki_MSH_CAMPOS_27_HAVIN_CRI2_WIDTH
            
            .Col = Ki_MSH_CAMPOS_22_HAVIN_CHKI
            For F = .FixedRows To .Rows - 1
                .Row = F
                Select Case .TextMatrix(F, Ki_MSH_CAMPOS_23_HAVIN_CHKV)
                    Case GO_ECNLIB00_CONST.VAL_CHECK: Set .CellPicture = imgL.ListImages(Ki_Ico_Check).Picture
                    Case GO_ECNLIB00_CONST.VAL_UNCHK: Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
                End Select
            Next F
        Else
            .ColWidth(Ki_MSH_CAMPOS_22_HAVIN_CHKI) = 0
            .ColWidth(Ki_MSH_CAMPOS_23_HAVIN_CHKV) = 0
            .ColWidth(Ki_MSH_CAMPOS_24_HAVIN_TIPO) = 0
            .ColWidth(Ki_MSH_CAMPOS_25_HAVIN_OPER) = 0
            .ColWidth(Ki_MSH_CAMPOS_26_HAVIN_CRI1) = 0
            .ColWidth(Ki_MSH_CAMPOS_27_HAVIN_CRI2) = 0
            
            .Col = Ki_MSH_CAMPOS_22_HAVIN_CHKI
            For F = .FixedRows To .Rows - 1
                .Row = F
                Set .CellPicture = Nothing
            Next F
        End If
        .Refresh
    End With
End Sub

Private Sub MENU_OPC_00_02_05_01()
    On Error Resume Next
    With mshDiseñoDatos
        .AddItem "", .FixedRows
        Call ReEnumeraCamposEnDgMatricial
        Call ConfiguraMSHCampos(.FixedRows)
        .Refresh
    End With
End Sub

Private Sub MENU_OPC_00_02_05_02()
    On Error Resume Next
    With mshDiseñoDatos
        .AddItem "", .Row - 1
        Call ReEnumeraCamposEnDgMatricial
        Call ConfiguraMSHCampos(.Row - 1)
        .Refresh
    End With
End Sub

Private Sub MENU_OPC_00_02_05_03()
    On Error Resume Next
    With mshDiseñoDatos
        .AddItem "", .Row + 1
        Call ReEnumeraCamposEnDgMatricial
        Call ConfiguraMSHCampos(.Row + 1)
        .Refresh
    End With
End Sub

Private Sub MENU_OPC_00_02_05_04()
    On Error Resume Next
    With mshDiseñoDatos
        .AddItem ""
        Call ReEnumeraCamposEnDgMatricial
        Call ConfiguraMSHCampos(.Rows - 1)
        .Refresh
    End With
End Sub

Private Sub MENU_OPC_00_02_06()
    On Error Resume Next
    With mshDiseñoDatos
        .RemoveItem .Row
        Call ReEnumeraCamposEnDgMatricial
        .Refresh
    End With
End Sub

Private Sub MENU_OPC_00_02_07()
    On Error Resume Next
    Call mshDiseñoDatos_RowColChange
    With mshDiseñoDatos
        Dim iFilaDondeReemplazar As Integer
        iFilaDondeReemplazar = InputBox("Ingrese N° de Fila donde se reemplazarán" & vbNewLine & _
                                        "los datos de la fila actual (" & .TextMatrix(.Row, Ki_MSH_CAMPOS_00_FIXED__COL) & ")", _
                                        "Reemplazar Datos de Fila", "0")
        If iFilaDondeReemplazar <= 0 Then
            MsgBox "El valor ingresado no es válido para realizar la operación solicitada, no se ha realizado ningún reemplazo", vbCritical, Me.Caption
            .Refresh
            Exit Sub
        End If
        
        Dim i As Integer
        Dim iFilaOrigen As Integer
        Dim iFilaDestin As Integer
     
        iFilaOrigen = .Row
     
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, Ki_MSH_CAMPOS_00_FIXED__COL)) = iFilaDondeReemplazar Then
                iFilaDestin = i
                Exit For
            End If
        Next i
        
        For i = .FixedCols To .Cols - 1
            .TextMatrix(iFilaDestin, i) = .TextMatrix(iFilaOrigen, i)
        Next i
        
        Call ConfigurarFilaSegunValores(iFilaDestin)
        .Refresh
    End With
End Sub

Private Sub MENU_OPC_00_02_08()
    On Error Resume Next
    
    Dim oRs As New ADODB.Recordset
    Dim F As Integer
    Dim C As Integer
    
    With mshDiseñoDatos
        oRs.CursorLocation = adUseClient
        For C = .FixedCols To .Cols - 1
            oRs.Fields.Append CStr(C), adVarChar, 200
        Next C
        oRs.Open
        
        For F = .FixedRows To .Rows - 1
            If Len(.TextMatrix(F, Ki_MSH_CAMPOS_02_TABLA_NOMB)) > 0 Then
                oRs.AddNew
                For C = .FixedCols To .Cols - 1
                    oRs.Fields(CStr(C)) = .TextMatrix(F, C)
                Next C
                oRs.Update
                oRs.MoveLast
            End If
        Next F
        
        oRs.Sort = CStr(Ki_MSH_CAMPOS_02_TABLA_NOMB)
        Set .DataSource = oRs
        
        Call ConfiguraMSHCampos(, True)
    End With
End Sub

Private Sub MENU_OPC_00_02_09()
    On Error Resume Next
    
    Dim C As Integer
    
    Call MuestraControlesDGMatricial(False)
    
    With mshDiseñoDatos
        For C = .FixedCols To .Cols - 1
            .TextMatrix(.Row, C) = Empty
            .Col = C
            Select Case C
                Case Ki_MSH_CAMPOS_03_TABLA_JOIN, _
                     Ki_MSH_CAMPOS_05_SELEC_TIPO, _
                     Ki_MSH_CAMPOS_32_ORDER_TIPO
                    Set .CellPicture = Nothing
                Case Ki_MSH_CAMPOS_08_SELEC_ACTI
                    Set .CellPicture = Nothing
                Case Ki_MSH_CAMPOS_09_SELEC_ACTV
                    .TextMatrix(.Row, C) = GO_ECNLIB00_CONST.VAL_UNCHK
                Case Ki_MSH_CAMPOS_11_WHERE_CHKI, _
                     Ki_MSH_CAMPOS_18_GROUP_CHKI, _
                     Ki_MSH_CAMPOS_22_HAVIN_CHKI, _
                     Ki_MSH_CAMPOS_29_ORDER_CHKI
                     Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
                Case Ki_MSH_CAMPOS_12_WHERE_CHKV, _
                     Ki_MSH_CAMPOS_19_GROUP_CHKV, _
                     Ki_MSH_CAMPOS_23_HAVIN_CHKV, _
                     Ki_MSH_CAMPOS_30_ORDER_CHKV
                     .TextMatrix(.Row, C) = GO_ECNLIB00_CONST.VAL_UNCHK
            End Select
        Next C
    End With
End Sub

Private Sub MENU_OPC_00_04()
    Unload Me
End Sub

Private Sub CrearMenusDeLaApp()
    On Error Resume Next
    '--+------------------------------------------------------------------------------------------------------------------------------------------------+--
    '=> MENU PRINCIPAL
    '--+------------------------------------------------------------------------------------------------------------------------------------------------+--
    Call CrearMenuDgGrafico(MNMAIN)
    Call CrearMenuDgMatricial(MNMAIN)
    With MNMAIN.MenuItems
        .Add 0, Ks_MNMAIN_00_04___, smiNone, "&Salir"
    End With
    '--+------------------------------------------------------------------------------------------------------------------------------------------------+--
    '=> MENU POPUD DIAGRAMA GRAFICO
    '--+------------------------------------------------------------------------------------------------------------------------------------------------+--
    Call CrearMenuDgGrafico(mnPopPup_DgGrafico)
    '--+------------------------------------------------------------------------------------------------------------------------------------------------+--
    '=> MENU POPUD DIAGRAMA MATRICIAL
    '--+------------------------------------------------------------------------------------------------------------------------------------------------+--
    Call CrearMenuDgMatricial(mnPopPup_DgMatricial)
    '--+------------------------------------------------------------------------------------------------------------------------------------------------+--
    '=> MENU POPUD TABLA : DENTRO DEL DIAGRAMA GRAFICO
    '--+------------------------------------------------------------------------------------------------------------------------------------------------+--
    Call CrearMenuDgGraficoTabla(mnPopPup_Tabla)
End Sub
    
Private Sub CrearMenuDgGrafico(ByRef xMenuXP As SmartMenuXP)
    On Error Resume Next
    With xMenuXP.MenuItems
        .Add 0, Ks_MNMAIN_00_01___, smiNone, "Diagrama Gráfico"
            .Add Ks_MNMAIN_00_01___, Ks_MNMAIN_00_01_01, smiPicture, "Agregar Tabla", imgL.ListImages(Ki_Ico_TableAdd).Picture, , vbKeyF4, , True, True
            .Add Ks_MNMAIN_00_01___, "BARRA", smiSeparator
            .Add Ks_MNMAIN_00_01___, Ks_MNMAIN_00_01_02, smiPicture, "Borrar Diseño", imgL.ListImages(Ki_Ico_DiagramDel).Picture, , vbKeyF7, , True, True
            .Add Ks_MNMAIN_00_01___, Ks_MNMAIN_00_01_03, smiPicture, "Minimizar todo", imgL.ListImages(Ki_Ico_TableMin).Picture, , vbKeyF9, , True, True
            .Add Ks_MNMAIN_00_01___, Ks_MNMAIN_00_01_04, smiPicture, "Maximizar todo", imgL.ListImages(Ki_Ico_TableMax).Picture, , vbKeyF10, , False, True
    End With
End Sub
    
Private Sub CrearMenuDgMatricial(ByRef xMenuXP As SmartMenuXP)
    On Error Resume Next
    With xMenuXP.MenuItems
        .Add 0, Ks_MNMAIN_00_02___, smiNone, "Diagrama matricial"
            .Add Ks_MNMAIN_00_02___, Ks_MNMAIN_00_02_01, smiCheckBox, "Sección FILTROS", , , , smiChecked, True, True
            .Add Ks_MNMAIN_00_02___, Ks_MNMAIN_00_02_02, smiCheckBox, "Sección AGRUPAMIENTO", , , , smiChecked, True, True
            .Add Ks_MNMAIN_00_02___, Ks_MNMAIN_00_02_03, smiCheckBox, "Sección FILTROS PARA EL GRUPO", , , , smiChecked, True, True
            .Add Ks_MNMAIN_00_02___, Ks_MNMAIN_00_02___ & "BARRA_00", smiSeparator
            .Add Ks_MNMAIN_00_02___, Ks_MNMAIN_00_02_04, smiPicture, "Guardar datos de la matriz", imgL.ListImages(Ki_Ico_TableSave).Picture, , , , True, True
            .Add Ks_MNMAIN_00_02___, Ks_MNMAIN_00_02___ & "BARRA_01", smiSeparator
            .Add Ks_MNMAIN_00_02___, Ks_MNMAIN_00_02_05, smiPicture, "Insertar Nueva fila", imgL.ListImages(Ki_Ico_TableInsert).Picture, , , , True, True
                .Add Ks_MNMAIN_00_02_05, Ks_MNMAIN_00_02_05_01, smiPicture, "Al Inicio", imgL.ListImages(Ki_Ico_BorderTop).Picture, , , , True, True
                .Add Ks_MNMAIN_00_02_05, Ks_MNMAIN_00_02_05_02, smiPicture, "Arriba de fila actual", imgL.ListImages(Ki_Ico_AlignTop).Picture, , , , True, True
                .Add Ks_MNMAIN_00_02_05, Ks_MNMAIN_00_02_05_03, smiPicture, "Debajo de Fila Actual", imgL.ListImages(Ki_Ico_AlignBottom).Picture, , , , True, True
                .Add Ks_MNMAIN_00_02_05, Ks_MNMAIN_00_02_05_04, smiPicture, "Al Final", imgL.ListImages(Ki_Ico_BorderTop).Picture, , , , True, True
            .Add Ks_MNMAIN_00_02___, Ks_MNMAIN_00_02_06, smiPicture, "Eliminar Fila Actual", imgL.ListImages(Ki_Ico_TableDelete).Picture, , , , True, True
            .Add Ks_MNMAIN_00_02___, Ks_MNMAIN_00_02_07, smiPicture, "Reemplazar Datos", imgL.ListImages(Ki_Ico_TableReplace).Picture, , , , True, True
            .Add Ks_MNMAIN_00_02___, Ks_MNMAIN_00_02_08, smiPicture, "Agrupar y Ordenar Campos configurados", imgL.ListImages(Ki_Ico_TableSort).Picture, , , , True, True
            .Add Ks_MNMAIN_00_02___, Ks_MNMAIN_00_02_09, smiPicture, "Borrar Datos de fila actual", imgL.ListImages(Ki_Ico_TablePaint).Picture, , , , True, True
    End With
End Sub

Private Sub CrearMenuDgGraficoTabla(ByRef xMenuXP As SmartMenuXP)
    On Error Resume Next
    With xMenuXP.MenuItems
        .Add 0, Ks_MNMAIN_00_03___, smiNone, "Tabla"
            .Add Ks_MNMAIN_00_03___, Ks_MNMAIN_00_03_01, smiPicture, "Quitar Tabla", imgL.ListImages(Ki_Ico_TableDel).Picture, , vbKeyF2, , True, True
    End With
End Sub

Private Function EjecucionMenuDgGrafico(ByRef xMenuXP As SmartMenuXP, _
                                          ByVal ID As Integer) As Boolean
    On Error Resume Next
    Dim blSW_Ejecutado As Boolean
        
    Select Case xMenuXP.MenuItems.Key(ID)
        Case Ks_MNMAIN_00_01_01
            Call MENU_OPC_00_01_01
            blSW_Ejecutado = True
        Case Ks_MNMAIN_00_01_02
            Call MENU_OPC_00_01_02
            blSW_Ejecutado = True
        Case Ks_MNMAIN_00_01_03
            Call MENU_OPC_00_01_03
            blSW_Ejecutado = True
        Case Ks_MNMAIN_00_01_04
            Call MENU_OPC_00_01_04
            blSW_Ejecutado = True
    End Select
End Function

Private Function EjecucionMenuDgMatricial(ByRef xMenuXP As SmartMenuXP, _
                                          ByVal ID As Integer) As Boolean
    On Error Resume Next
    Dim blSW_Ejecutado As Boolean
        
    Select Case xMenuXP.MenuItems.Key(ID)
        Case Ks_MNMAIN_00_02_01
            Call MENU_OPC_00_02_01(xMenuXP.MenuItems.Value(ID))
            blSW_Ejecutado = True
        Case Ks_MNMAIN_00_02_02
            Call MENU_OPC_00_02_02(xMenuXP.MenuItems.Value(ID))
            blSW_Ejecutado = True
        Case Ks_MNMAIN_00_02_03
            Call MENU_OPC_00_02_03(xMenuXP.MenuItems.Value(ID))
            blSW_Ejecutado = True
        Case Ks_MNMAIN_00_02_05_01
            Call MENU_OPC_00_02_05_01
            blSW_Ejecutado = True
        Case Ks_MNMAIN_00_02_05_02
            Call MENU_OPC_00_02_05_02
            blSW_Ejecutado = True
        Case Ks_MNMAIN_00_02_05_03
            Call MENU_OPC_00_02_05_03
            blSW_Ejecutado = True
        Case Ks_MNMAIN_00_02_05_04
            Call MENU_OPC_00_02_05_04
            blSW_Ejecutado = True
        Case Ks_MNMAIN_00_02_06
            Call MENU_OPC_00_02_06
            blSW_Ejecutado = True
        Case Ks_MNMAIN_00_02_07
            Call MENU_OPC_00_02_07
            blSW_Ejecutado = True
        Case Ks_MNMAIN_00_02_08
            Call MENU_OPC_00_02_08
            blSW_Ejecutado = True
        Case Ks_MNMAIN_00_02_09
            Call MENU_OPC_00_02_09
            blSW_Ejecutado = True
    End Select
End Function

Private Sub AgregarTabla()
    Load frm002_ECNSQLQRYDESIGN_TV
    With frm002_ECNSQLQRYDESIGN_TV
        Call .PU_CargarInfo
        .Show 1
    End With
    Set frm002_ECNSQLQRYDESIGN_TV = Nothing
    Call DiseñarTablas
End Sub

Private Sub MiniMaximizarTodasLasTablas(ByVal blSW_Minimizar As Boolean)
    Dim i As Integer
    Dim INDICE As Integer
    
    For i = LBound(aVectorDeIndicesDeTablasDelDiagrama()) To _
            UBound(aVectorDeIndicesDeTablasDelDiagrama())
        INDICE = aVectorDeIndicesDeTablasDelDiagrama(i)
        Select Case blSW_Minimizar
            Case True: cmdTabla(INDICE).Tag = GO_ECNLIB00_CONST.VAL_UNCHK
            Case False: cmdTabla(INDICE).Tag = GO_ECNLIB00_CONST.VAL_CHECK
        End Select
        Call cmdTabla_Click(INDICE)
        Call ReUbicaRelaciones(INDICE)
    Next i
    
    Select Case blSW_Minimizar
        Case True
            mnPopPup_DgGrafico.MenuItems.Enabled(Ki_MNPOPU_PCTDISEÑO_ID5_MINTABLE) = False
            mnPopPup_DgGrafico.MenuItems.Enabled(Ki_MNPOPU_PCTDISEÑO_ID6_MAXTABLE) = True
        Case False
            mnPopPup_DgGrafico.MenuItems.Enabled(Ki_MNPOPU_PCTDISEÑO_ID5_MINTABLE) = True
            mnPopPup_DgGrafico.MenuItems.Enabled(Ki_MNPOPU_PCTDISEÑO_ID6_MAXTABLE) = False
    End Select
End Sub

Private Sub BorraDiseño()
    On Error Resume Next
    Dim i As Integer
    Dim INDICE As Integer
    
    '----------------------------------------------------------------------------------------
    '=> ELIMINO LAS RELACIONES CREADAS,PERO DEJO LA PRIMERA (CONTROLES BASE)
    '----------------------------------------------------------------------------------------
    For i = LBound(aVectorDeIndicesRelacionesDelDiagrama()) To _
            UBound(aVectorDeIndicesRelacionesDelDiagrama())
        
        INDICE = aVectorDeIndicesRelacionesDelDiagrama(i)
        If INDICE <> 0 Then
            Unload linRelacionPFK(INDICE)
            Unload imgRelacion_PK(INDICE)
            Unload imgRelacion_FK(INDICE)
            
            aVectorDeIndicesRelacionesDelDiagrama(i) = Ki_Vector_ValorNULL
        End If
    Next i
    '----------------------------------------------------------------------------------------
    '=> ELIMINO LAS TABLAS Y SUS RESPECTIVOS CONTROLES, PERO DEJO LA PRIMERA (CONTROLES BASE)
    '----------------------------------------------------------------------------------------
    For i = LBound(aVectorDeIndicesDeTablasDelDiagrama()) To _
            UBound(aVectorDeIndicesDeTablasDelDiagrama())
        
        INDICE = aVectorDeIndicesDeTablasDelDiagrama(i)
        If INDICE <> 0 Then
            Unload mshTabla(INDICE)
            Unload pctTabla(INDICE)
            Unload chkTabla(INDICE)
            Unload cmdTabla(INDICE)
            Unload lblTabla(INDICE)
            
            aVectorDeIndicesDeTablasDelDiagrama(i) = Ki_Vector_ValorNULL
        End If
    Next i
    '----------------------------------------------------------------------------------------
    '=> REALIZO LA CONFIGURACION BASE PARA EL FUNCIONAMIENTO DEL DIAGRAMA
    '----------------------------------------------------------------------------------------
    Call ConfiguracionBase
    '----------------------------------------------------------------------------------------
    '=> REINICIO EL VECTOR QUE CONTROLA LOS INDICES DE LAS TABLAS DEL ARREGLO DE CONTROLES
    '   Y EL VECTOR QUE CONTROLA LOS INDICES DE LAS RELACIONES
    '----------------------------------------------------------------------------------------
    Erase aVectorDeIndicesDeTablasDelDiagrama()
    Erase aVectorDeIndicesRelacionesDelDiagrama()
    IND_VECTOR_DE_TABLAS = 0
    IND_VECTOR_DE_RELACIONES = 0
    '----------------------------------------------------------------------------------------
    '=> DESACIVO EL SW DE DISEÑO
    '----------------------------------------------------------------------------------------
    PU_SW_DISEÑO_CON_TABLAS = False
    PU_SW_DISEÑO_CON_RELACIONES = False
End Sub

Private Sub ConfiguracionBase()
    Call ConfiguracionBase_Tablas
    Call ConfiguracionBase_Relaciones
End Sub

Private Sub ConfiguracionBase_Relaciones()
    On Error Resume Next
    imgRelacion_PK(0).Tag = Empty
    imgRelacion_PK(0).Visible = False
    
    imgRelacion_FK(0).Tag = Empty
    imgRelacion_FK(0).Visible = False
    
    linRelacionPFK(0).Tag = Empty
    linRelacionPFK(0).Visible = False
End Sub

Private Sub ConfiguracionBase_Tablas()
    On Error Resume Next
    With mshTabla(0)
        .Clear
        .Tag = Empty
        .Rows = 4
        .FixedRows = 3
        .FixedCols = 0
        .BackColorFixed = Kdbl_COLOR_TABLA
        .Visible = False
    End With
    
    lblTabla(0).Caption = Empty
    lblTabla(0).Tag = GO_ECNLIB00_CONST.VAL_UNCHK
    lblTabla(0).Visible = False
    
    chkTabla(0).Tag = Empty
    chkTabla(0).Visible = False
    
    cmdTabla(0).Caption = GO_ECNLIB00_CONST.CARESP_WEB_RESTAURADO
    cmdTabla(0).Tag = GO_ECNLIB00_CONST.VAL_UNCHK
    cmdTabla(0).Visible = False
    
    pctTabla(0).Tag = Empty
    pctTabla(0).Visible = False
    pctTabla(0).BackColor = Kdbl_COLOR_TABLA
End Sub

Private Function IndiceEsParteDelArregloDeControlesDeTablas(ByVal iIndiceTabla As Integer) As Boolean
    IndiceEsParteDelArregloDeControlesDeTablas = False
    If iIndiceTabla = Ki_Vector_ValorNULL Then Exit Function
    
    Dim i As Integer
    For i = LBound(aVectorDeIndicesDeTablasDelDiagrama) To _
            UBound(aVectorDeIndicesDeTablasDelDiagrama)
        If aVectorDeIndicesDeTablasDelDiagrama(i) = iIndiceTabla Then
            IndiceEsParteDelArregloDeControlesDeTablas = True
            Exit For
        End If
    Next i
End Function

Private Function IndiceEsParteDelArregloDeControlesDeRelaciones(ByVal iIndiceRelacion As Integer) As Boolean
    IndiceEsParteDelArregloDeControlesDeRelaciones = False
    
    Dim i As Integer
    For i = LBound(aVectorDeIndicesDeTablasDelDiagrama) To _
            UBound(aVectorDeIndicesDeTablasDelDiagrama)
        If aVectorDeIndicesDeTablasDelDiagrama(i) = iIndiceRelacion Then
            IndiceEsParteDelArregloDeControlesDeRelaciones = True
            Exit For
        End If
    Next i
End Function

Private Function SeteaNullPorValorEnVectorDeIndicesDeTablasEnDiagrama(ByVal iIndiceTabla As Integer) As Boolean
    SeteaNullPorValorEnVectorDeIndicesDeTablasEnDiagrama = False
    Dim i As Integer
    Dim INDICE As Integer
    
    For i = LBound(aVectorDeIndicesDeTablasDelDiagrama) To _
            UBound(aVectorDeIndicesDeTablasDelDiagrama)
        INDICE = aVectorDeIndicesDeTablasDelDiagrama(i)
        If INDICE = iIndiceTabla Then
            aVectorDeIndicesDeTablasDelDiagrama(i) = Ki_Vector_ValorNULL
            SeteaNullPorValorEnVectorDeIndicesDeTablasEnDiagrama = True
            Exit For
        End If
    Next i
End Function

Private Function SeteaNullPorValorEnVectorDeIndicesDeRelacionesEnDiagrama(ByVal iIndiceRelacion As Integer) As Boolean
    SeteaNullPorValorEnVectorDeIndicesDeRelacionesEnDiagrama = False
    Dim i As Integer
    Dim INDICE As Integer
    
    For i = LBound(aVectorDeIndicesRelacionesDelDiagrama) To _
            UBound(aVectorDeIndicesRelacionesDelDiagrama)
        INDICE = aVectorDeIndicesRelacionesDelDiagrama(i)
        If INDICE = iIndiceRelacion Then
            aVectorDeIndicesRelacionesDelDiagrama(i) = Ki_Vector_ValorNULL
            SeteaNullPorValorEnVectorDeIndicesDeRelacionesEnDiagrama = True
            Exit For
        End If
    Next i
End Function

Private Function FindIndMasCercanoEnVectorDeTablas(ByVal iIndiceTbla As Integer) As Integer
    On Error GoTo SALTO_ERROR
    '--------------------------------------------------------------------------------------------------------------------
    '=> DECLARACION DE VARIABLES
    '--------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim J As Integer
    
    Dim iVALOR_I As Integer
    Dim iVALOR_J As Integer
    '--------------------------------------------------------------------------------------------------------------------
    '=> REALIZO UNA COPIA DEL VECTOR DE TABLAS
    '--------------------------------------------------------------------------------------------------------------------
    Dim aVECTOR_TEMPORAL() As Integer
    
    For i = LBound(aVectorDeIndicesDeTablasDelDiagrama) To _
            UBound(aVectorDeIndicesDeTablasDelDiagrama)
        ReDim Preserve aVECTOR_TEMPORAL(1 To i) As Integer
        aVECTOR_TEMPORAL(i) = aVectorDeIndicesDeTablasDelDiagrama(i)
        If IndiceEsParteDelArregloDeControlesDeTablas(aVECTOR_TEMPORAL(i)) = False Then _
            aVECTOR_TEMPORAL(i) = Ki_Vector_ValorNULL
    Next i
    '--------------------------------------------------------------------------------------------------------------------
    '=> ORDENO EL VECTOR TEMPORAL
    '--------------------------------------------------------------------------------------------------------------------
    Dim iAuxilio As Integer
    
    For i = LBound(aVECTOR_TEMPORAL) To _
            UBound(aVECTOR_TEMPORAL) - 1
        iVALOR_I = aVECTOR_TEMPORAL(i)
        For J = LBound(aVECTOR_TEMPORAL) + 1 To _
                UBound(aVECTOR_TEMPORAL)
            iVALOR_J = aVECTOR_TEMPORAL(J)
            If iVALOR_J < iVALOR_I Then
                aVECTOR_TEMPORAL(i) = iVALOR_J
                aVECTOR_TEMPORAL(J) = iVALOR_I
                
                iAuxilio = iVALOR_I
                iVALOR_I = iVALOR_J
                iVALOR_J = iAuxilio
            End If
        Next J
    Next i
    '--------------------------------------------------------------------------------------------------------------------
    '=> BUSCO EL VALOR MAS CERCANO AL VALOR DEL INDICE DE LA TABLA
    '--------------------------------------------------------------------------------------------------------------------
    If IndiceEsParteDelArregloDeControlesDeTablas(UBound(aVECTOR_TEMPORAL) - 1) Then
        FindIndMasCercanoEnVectorDeTablas = UBound(aVECTOR_TEMPORAL) - 1
    Else
        FindIndMasCercanoEnVectorDeTablas = 0
    End If
    Exit Function
SALTO_ERROR:
    FindIndMasCercanoEnVectorDeTablas = 0
End Function

Private Sub CargarDatosTipoDeOrdenamiento(ByVal iIndice As Integer)
    Dim objCboItem As ComboItem
    
    With cboTipOrden(iIndice)
        .ComboItems.Clear
        Set .ImageList = imgL2
        Set objCboItem = .ComboItems.Add(1, GO_002_Ks_TIPO_DE_ORDENAMIENTO_ASC, "Asc", Ki_Ico_ImgL2_SortAsc)
        Set objCboItem = .ComboItems.Add(2, GO_002_Ks_TIPO_DE_ORDENAMIENTO_DES, "Desc", Ki_Ico_ImgL2_SortDes)
    End With
    Call GO_ECNLIB01_FUNSUB.CambiarComboListaAncho(cboTipOrden(iIndice), 10)
End Sub

Private Sub CargarDatosTipoSQLJoin(ByVal iIndice As Integer)
    Dim objCboItem As ComboItem
    
    With cboTipSQLJoin(iIndice)
        .ComboItems.Clear
        Set .ImageList = imgL2
        Set objCboItem = .ComboItems.Add(1, GO_002_Ks_TIPO_DE_JOIN_FR, "From", Ki_Ico_ImgL2_From)
        Set objCboItem = .ComboItems.Add(2, GO_002_Ks_TIPO_DE_JOIN_LF, "Izquierda", Ki_Ico_ImgL2_LeftJoin)
        Set objCboItem = .ComboItems.Add(3, GO_002_Ks_TIPO_DE_JOIN_IN, "Interseccion", Ki_Ico_ImgL2_InnerJoin)
        Set objCboItem = .ComboItems.Add(4, GO_002_Ks_TIPO_DE_JOIN_RI, "Derecha", Ki_Ico_ImgL2_RightJoin)
        Set objCboItem = .ComboItems.Add(5, GO_002_Ks_TIPO_DE_JOIN_UN, "Union All", Ki_Ico_ImgL2_Union)
    End With
    Call GO_ECNLIB01_FUNSUB.CambiarComboListaAncho(cboTipSQLJoin(iIndice), 95)
End Sub

Private Sub CargarDatosTipoDeCampo(ByVal iIndice As Integer)
    Dim objCboItem As ComboItem
    
    With cboTipCampo(iIndice)
        .ComboItems.Clear
        Set .ImageList = imgL2
        Set objCboItem = .ComboItems.Add(1, GO_002_Ks_TIPO_DE_CAMPO_FD, "Campo", Ki_Ico_ImgL2_Column)
        Set objCboItem = .ComboItems.Add(2, GO_002_Ks_TIPO_DE_CAMPO_TX, "TxT", Ki_Ico_ImgL2_Text)
        Set objCboItem = .ComboItems.Add(3, GO_002_Ks_TIPO_DE_CAMPO_FX, "Fx", Ki_Ico_ImgL2_Fx)
        Set objCboItem = .ComboItems.Add(4, GO_002_Ks_TIPO_DE_CAMPO_AD, "Agregado", Ki_Ico_ImgL2_Zigma)
    End With
    Call GO_ECNLIB01_FUNSUB.CambiarComboListaAncho(cboTipCampo(iIndice), 80)
End Sub

Private Sub CargarDatosOperadoresWHERE(ByVal iIndice As Integer)
    Call GO_ECNLIB01_FUNSUB.CambiarComboListaAncho(cboOperWHERE(iIndice), 80)
    Call GO_ECNLIB01_FUNSUB.CambiarComboListaLargo(cboOperWHERE(iIndice), 300)
    If iIndice = 0 Then Exit Sub
    
    cboOperWHERE(iIndice).Clear
    
    Dim i As Integer
    For i = 0 To cboOperWHERE(0).ListCount - 1
        cboOperWHERE(iIndice).AddItem cboOperWHERE(0).List(i)
    Next i
End Sub

Private Sub CargarDatosOperadoresHAVING(ByVal iIndice As Integer)
    Call GO_ECNLIB01_FUNSUB.CambiarComboListaAncho(cboOperHAVING(iIndice), 80)
    Call GO_ECNLIB01_FUNSUB.CambiarComboListaLargo(cboOperHAVING(iIndice), 300)
    If iIndice = 0 Then Exit Sub
    
    cboOperHAVING(iIndice).Clear
    
    Dim i As Integer
    For i = 0 To cboOperHAVING(0).ListCount - 1
        cboOperHAVING(iIndice).AddItem cboOperHAVING(0).List(i)
    Next i
End Sub

Private Sub CargarDatosTipoWHERE(ByVal iIndice As Integer)
    Call GO_ECNLIB01_FUNSUB.CambiarComboListaAncho(cboTipoWHERE(iIndice), 20)
    If iIndice = 0 Then Exit Sub
    
    cboTipoWHERE(iIndice).Clear
    
    Dim i As Integer
    For i = 0 To cboTipoWHERE(0).ListCount - 1
        cboTipoWHERE(iIndice).AddItem cboTipoWHERE(0).List(i)
    Next i
End Sub

Private Sub CargarDatosTipoHAVING(ByVal iIndice As Integer)
    Call GO_ECNLIB01_FUNSUB.CambiarComboListaAncho(cboTipoHAVING(iIndice), 20)
    If iIndice = 0 Then Exit Sub
    
    cboTipoHAVING(iIndice).Clear
    
    Dim i As Integer
    For i = 0 To cboTipoHAVING(0).ListCount - 1
        cboTipoHAVING(iIndice).AddItem cboTipoHAVING(0).List(i)
    Next i
End Sub

Private Sub UbicaControlesMSH()
    On Error GoTo SALTO_ERROR
    Dim F As Integer
    Dim iIndice As Integer
    
    With mshDiseñoDatos
        For F = .FixedRows To .Rows - 1
                .RowHeight(F) = 280
            iIndice = CInt(Val(.TextMatrix(F, Ki_MSH_CAMPOS_00_FIXED__COL)))
            If iIndice > 0 Then
                .Row = F
                                
                .Col = Ki_MSH_CAMPOS_05_SELEC_TIPO
                cboTipCampo(iIndice).Left = .Left + .CellLeft
                cboTipCampo(iIndice).Top = .Top + .CellTop
                cboTipCampo(iIndice).Width = .CellWidth
                cboTipCampo(iIndice).Height = .CellHeight
                cboTipCampo(iIndice).Visible = True
                cboTipCampo(iIndice).ZOrder 0
                
                .Col = Ki_MSH_CAMPOS_32_ORDER_TIPO
                cboTipOrden(iIndice).Left = .Left + .CellLeft
                cboTipOrden(iIndice).Top = .Top + .CellTop
                cboTipOrden(iIndice).Width = .CellWidth
                cboTipOrden(iIndice).Height = .CellHeight
                cboTipOrden(iIndice).Visible = True
                cboTipOrden(iIndice).ZOrder 0
                
                .Col = Ki_MSH_CAMPOS_13_WHERE_TIPO
                cboTipoWHERE(iIndice).Left = .Left + .CellLeft
                cboTipoWHERE(iIndice).Top = .Top + .CellTop
                cboTipoWHERE(iIndice).Width = .CellWidth
                cboTipoWHERE(iIndice).Visible = True
                cboTipoWHERE(iIndice).ZOrder 0
                
                .Col = Ki_MSH_CAMPOS_14_WHERE_OPER
                cboOperWHERE(iIndice).Left = .Left + .CellLeft
                cboOperWHERE(iIndice).Top = .Top + .CellTop
                cboOperWHERE(iIndice).Width = .CellWidth
                cboOperWHERE(iIndice).Visible = True
                cboOperWHERE(iIndice).ZOrder 0
                
                .Col = Ki_MSH_CAMPOS_24_HAVIN_TIPO
                cboTipoHAVING(iIndice).Left = .Left + .CellLeft
                cboTipoHAVING(iIndice).Top = .Top + .CellTop
                cboTipoHAVING(iIndice).Width = .CellWidth
                cboTipoHAVING(iIndice).Visible = True
                cboTipoHAVING(iIndice).ZOrder 0
                
                .Col = Ki_MSH_CAMPOS_25_HAVIN_OPER
                cboOperHAVING(iIndice).Left = .Left + .CellLeft
                cboOperHAVING(iIndice).Top = .Top + .CellTop
                cboOperHAVING(iIndice).Width = .CellWidth
                cboOperHAVING(iIndice).Visible = True
                cboOperHAVING(iIndice).ZOrder 0
            End If
        Next F
    End With
SALTO_ERROR:
End Sub

Private Sub ReEnumeraCamposEnDgMatricial()
    On Error Resume Next
    Dim F As Integer
    Dim iFila As Integer
    
    With mshDiseñoDatos
        For F = .FixedRows To .Rows - 1
            iFila = iFila + 1
            .TextMatrix(F, Ki_MSH_CAMPOS_00_FIXED__COL) = CStr(iFila)
        Next F
    End With
End Sub

Private Sub SetearValoresPorTablaNull(ByVal iFILA_SETEAR As Integer)
    Dim iFilaActual As Integer
    Dim iColuActual As Integer
    
    With mshDiseñoDatos
        If .TextMatrix(iFILA_SETEAR, Ki_MSH_CAMPOS_04_TABLA_CODI) = Ks_CBO_TABLA_NULL_COD Then
            iFilaActual = .Row
            iColuActual = .Col
            
            .Row = iFILA_SETEAR
            
            .TextMatrix(.Row, Ki_MSH_CAMPOS_03_TABLA_JOIN) = Empty
            .Col = Ki_MSH_CAMPOS_03_TABLA_JOIN: Set .CellPicture = Nothing
            
            If .TextMatrix(.Row, Ki_MSH_CAMPOS_05_SELEC_TIPO) = GO_002_Ks_TIPO_DE_CAMPO_FD Then
                .TextMatrix(.Row, Ki_MSH_CAMPOS_05_SELEC_TIPO) = Empty
                .Col = Ki_MSH_CAMPOS_05_SELEC_TIPO: Set .CellPicture = Nothing
                .TextMatrix(.Row, Ki_MSH_CAMPOS_06_SELEC_NOMB) = Empty
            End If
            
            .Col = Ki_MSH_CAMPOS_11_WHERE_CHKI: Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
            .Col = Ki_MSH_CAMPOS_18_GROUP_CHKI: Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
            .Col = Ki_MSH_CAMPOS_22_HAVIN_CHKI: Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
            .Col = Ki_MSH_CAMPOS_29_ORDER_CHKI: Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
            
            .Col = Ki_MSH_CAMPOS_32_ORDER_TIPO: Set .CellPicture = Nothing
            
            .TextMatrix(.Row, Ki_MSH_CAMPOS_12_WHERE_CHKV) = GO_ECNLIB00_CONST.VAL_UNCHK
            .TextMatrix(.Row, Ki_MSH_CAMPOS_13_WHERE_TIPO) = Empty
            .TextMatrix(.Row, Ki_MSH_CAMPOS_14_WHERE_OPER) = Empty
            .TextMatrix(.Row, Ki_MSH_CAMPOS_15_WHERE_CRI1) = Empty
            .TextMatrix(.Row, Ki_MSH_CAMPOS_16_WHERE_CRI2) = Empty
            
            .TextMatrix(.Row, Ki_MSH_CAMPOS_19_GROUP_CHKV) = GO_ECNLIB00_CONST.VAL_UNCHK
            .TextMatrix(.Row, Ki_MSH_CAMPOS_20_GROUP_NUME) = Empty
            
            .TextMatrix(.Row, Ki_MSH_CAMPOS_23_HAVIN_CHKV) = GO_ECNLIB00_CONST.VAL_UNCHK
            .TextMatrix(.Row, Ki_MSH_CAMPOS_24_HAVIN_TIPO) = Empty
            .TextMatrix(.Row, Ki_MSH_CAMPOS_25_HAVIN_OPER) = Empty
            .TextMatrix(.Row, Ki_MSH_CAMPOS_26_HAVIN_CRI1) = Empty
            .TextMatrix(.Row, Ki_MSH_CAMPOS_27_HAVIN_CRI2) = Empty
            
            .TextMatrix(.Row, Ki_MSH_CAMPOS_30_ORDER_CHKV) = GO_ECNLIB00_CONST.VAL_UNCHK
            .TextMatrix(.Row, Ki_MSH_CAMPOS_31_ORDER_NUME) = Empty
            
            .Row = iFilaActual
            .Col = iColuActual
        End If
    End With
End Sub

Private Sub ConfigurarFilaSegunValores(ByVal iFILA_SETEAR As Integer)
    Dim iFilaActual As Integer
    Dim iColuActual As Integer
    
    With mshDiseñoDatos
        iFilaActual = .Row
        iColuActual = .Col
        
        .Row = iFILA_SETEAR
        
        .Col = Ki_MSH_CAMPOS_03_TABLA_JOIN
        Select Case .TextMatrix(.Row, Ki_MSH_CAMPOS_03_TABLA_JOIN)
            Case GO_002_Ks_TIPO_DE_JOIN_FR: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_From).Picture
            Case GO_002_Ks_TIPO_DE_JOIN_IN: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_InnerJoin).Picture
            Case GO_002_Ks_TIPO_DE_JOIN_LF: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_LeftJoin).Picture
            Case GO_002_Ks_TIPO_DE_JOIN_RI: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_RightJoin).Picture
            Case GO_002_Ks_TIPO_DE_JOIN_UN: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_Union).Picture
            Case Else
                Set .CellPicture = Nothing
        End Select
        
        .Col = Ki_MSH_CAMPOS_05_SELEC_TIPO
        Select Case .TextMatrix(.Row, Ki_MSH_CAMPOS_05_SELEC_TIPO)
            Case GO_002_Ks_TIPO_DE_CAMPO_FD: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_Column).Picture
            Case GO_002_Ks_TIPO_DE_CAMPO_TX: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_Text).Picture
            Case GO_002_Ks_TIPO_DE_CAMPO_FX: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_Fx).Picture
            Case GO_002_Ks_TIPO_DE_CAMPO_AD: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_Zigma).Picture
            Case Else
                Set .CellPicture = Nothing
        End Select
        
        .Col = Ki_MSH_CAMPOS_08_SELEC_ACTI
        Select Case .TextMatrix(.Row, Ki_MSH_CAMPOS_09_SELEC_ACTV)
            Case GO_ECNLIB00_CONST.VAL_CHECK: Set .CellPicture = imgL.ListImages(Ki_Ico_Tick).Picture
            Case GO_ECNLIB00_CONST.VAL_UNCHK: Set .CellPicture = Nothing
        End Select
        
        .Col = Ki_MSH_CAMPOS_11_WHERE_CHKI
        Select Case .TextMatrix(.Row, Ki_MSH_CAMPOS_12_WHERE_CHKV)
            Case GO_ECNLIB00_CONST.VAL_CHECK: Set .CellPicture = imgL.ListImages(Ki_Ico_Check).Picture
            Case GO_ECNLIB00_CONST.VAL_UNCHK: Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
        End Select
        
        .Col = Ki_MSH_CAMPOS_18_GROUP_CHKI
        Select Case .TextMatrix(.Row, Ki_MSH_CAMPOS_19_GROUP_CHKV)
            Case GO_ECNLIB00_CONST.VAL_CHECK: Set .CellPicture = imgL.ListImages(Ki_Ico_Check).Picture
            Case GO_ECNLIB00_CONST.VAL_UNCHK: Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
        End Select
        
        .Col = Ki_MSH_CAMPOS_22_HAVIN_CHKI
        Select Case .TextMatrix(.Row, Ki_MSH_CAMPOS_23_HAVIN_CHKV)
            Case GO_ECNLIB00_CONST.VAL_CHECK: Set .CellPicture = imgL.ListImages(Ki_Ico_Check).Picture
            Case GO_ECNLIB00_CONST.VAL_UNCHK: Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
        End Select
        
        .Col = Ki_MSH_CAMPOS_29_ORDER_CHKI
        Select Case .TextMatrix(.Row, Ki_MSH_CAMPOS_30_ORDER_CHKV)
            Case GO_ECNLIB00_CONST.VAL_CHECK: Set .CellPicture = imgL.ListImages(Ki_Ico_Check).Picture
            Case GO_ECNLIB00_CONST.VAL_UNCHK: Set .CellPicture = imgL.ListImages(Ki_Ico_UnChk).Picture
        End Select
               
        .Col = Ki_MSH_CAMPOS_32_ORDER_TIPO
        Select Case .TextMatrix(.Row, Ki_MSH_CAMPOS_32_ORDER_TIPO)
            Case GO_002_Ks_TIPO_DE_ORDENAMIENTO_ASC: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_SortAsc).Picture
            Case GO_002_Ks_TIPO_DE_ORDENAMIENTO_DES: Set .CellPicture = imgL2.ListImages(Ki_Ico_ImgL2_SortDes).Picture
            Case Else
                Set .CellPicture = Nothing
        End Select
               
        .Row = iFilaActual
        .Col = iColuActual
    End With
End Sub
