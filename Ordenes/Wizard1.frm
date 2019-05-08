VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWizard 
   Appearance      =   0  'Flat
   Caption         =   "Generación de Purchase Order"
   ClientHeight    =   8535
   ClientLeft      =   1980
   ClientTop       =   1830
   ClientWidth     =   9855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Purchase Order Generation"
   Begin VB.PictureBox picNav 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   9855
      TabIndex        =   38
      Top             =   7965
      Width           =   9855
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "OK"
         Height          =   312
         Left            =   7650
         MaskColor       =   &H00000000&
         TabIndex        =   88
         Tag             =   "Ok"
         Top             =   120
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   312
         Left            =   8745
         MaskColor       =   &H00000000&
         TabIndex        =   87
         Tag             =   "Cancel"
         Top             =   120
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         Height          =   312
         Index           =   4
         Left            =   5640
         MaskColor       =   &H00000000&
         TabIndex        =   37
         Tag             =   "Finish"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   312
         Index           =   3
         Left            =   4530
         MaskColor       =   &H00000000&
         TabIndex        =   42
         Tag             =   "Next"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Height          =   312
         Index           =   2
         Left            =   3435
         MaskColor       =   &H00000000&
         TabIndex        =   41
         Tag             =   "Back"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Cancel"
         Height          =   312
         Index           =   1
         Left            =   2250
         MaskColor       =   &H00000000&
         TabIndex        =   40
         Tag             =   "Cancel"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Help"
         Height          =   312
         Index           =   0
         Left            =   108
         MaskColor       =   &H00000000&
         TabIndex        =   39
         Tag             =   "Help"
         Top             =   120
         Width           =   1092
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   0
         X2              =   11350
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   255
         X2              =   9735
         Y1              =   -30
         Y2              =   -45
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Datos Generales"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7905
      Index           =   0
      Left            =   30
      TabIndex        =   43
      Tag             =   "1000"
      Top             =   60
      Width           =   9945
      Begin VB.CommandButton cmdGrupoPro 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6390
         TabIndex        =   111
         Top             =   5085
         Width           =   390
      End
      Begin VB.TextBox txtDes_GrupoPro 
         Height          =   285
         Left            =   2340
         MaxLength       =   50
         TabIndex        =   66
         Top             =   5085
         Width           =   4035
      End
      Begin VB.TextBox txtCod_GrupoPro 
         Height          =   285
         Left            =   1710
         MaxLength       =   8
         TabIndex        =   13
         Top             =   5085
         Width           =   615
      End
      Begin VB.Frame Frame3 
         Caption         =   "Prendas a Producir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1965
         Left            =   7320
         TabIndex        =   112
         Tag             =   "Production Garments"
         Top             =   3420
         Width           =   2355
         Begin VB.TextBox TxtAd 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   18
            Top             =   1440
            Width           =   1185
         End
         Begin VB.TextBox TxtPorc 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   17
            Top             =   840
            Width           =   1185
         End
         Begin VB.TextBox TxtCritico 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   16
            Top             =   300
            Width           =   1185
         End
         Begin VB.Label Label3 
            Caption         =   "Adicionales:"
            Height          =   345
            Left            =   120
            TabIndex        =   115
            Tag             =   "Add"
            Top             =   1470
            Width           =   1305
         End
         Begin VB.Label Label2 
            Caption         =   "Porcentaje:"
            Height          =   345
            Left            =   120
            TabIndex        =   114
            Tag             =   "%"
            Top             =   870
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "Criticas:"
            Height          =   255
            Left            =   120
            TabIndex        =   113
            Tag             =   "Critict"
            Top             =   330
            Width           =   1035
         End
      End
      Begin VB.CommandButton cmdCod_Banco 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6390
         TabIndex        =   110
         Top             =   4725
         Width           =   405
      End
      Begin VB.CommandButton cmdCod_Embarque 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6390
         TabIndex        =   109
         Top             =   3510
         Width           =   405
      End
      Begin VB.CommandButton cmdCod_PagEmb 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6390
         TabIndex        =   108
         Top             =   3135
         Width           =   405
      End
      Begin VB.CommandButton cmdCod_TemCli 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6390
         TabIndex        =   107
         Top             =   2745
         Width           =   405
      End
      Begin VB.CommandButton cmdCod_DivCli 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6390
         TabIndex        =   106
         Top             =   2370
         Width           =   405
      End
      Begin VB.CommandButton cmdCod_Destino 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6390
         TabIndex        =   105
         Top             =   825
         Width           =   405
      End
      Begin VB.Frame Frame2 
         Height          =   795
         Left            =   7320
         TabIndex        =   102
         Top             =   1650
         Visible         =   0   'False
         Width           =   2355
         Begin VB.OptionButton optNoRegular 
            Caption         =   "No Regular"
            Height          =   195
            Left            =   360
            TabIndex        =   104
            Tag             =   "Not Regular"
            Top             =   450
            Width           =   1815
         End
         Begin VB.OptionButton optRegular 
            Caption         =   "Regular"
            Height          =   195
            Left            =   360
            TabIndex        =   103
            Tag             =   "Regular"
            Top             =   210
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         Height          =   795
         Left            =   7320
         TabIndex        =   99
         Top             =   840
         Width           =   2355
         Begin VB.OptionButton optOrden 
            Caption         =   "Orden"
            Height          =   195
            Left            =   360
            TabIndex        =   101
            Tag             =   "Order"
            Top             =   210
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton optReorden 
            Caption         =   "Reorden"
            Height          =   195
            Left            =   360
            TabIndex        =   100
            Tag             =   "Reorder"
            Top             =   450
            Width           =   1815
         End
      End
      Begin VB.OptionButton optFlg_CartaNoAprobada 
         Caption         =   "No Aprobada"
         Height          =   195
         Left            =   2880
         TabIndex        =   11
         Tag             =   "Not Approved"
         Top             =   4365
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.OptionButton optFlg_CartaAprobada 
         Caption         =   "Aprobada"
         Height          =   195
         Left            =   1710
         TabIndex        =   72
         Tag             =   "Approved"
         Top             =   4350
         Width           =   1365
      End
      Begin VB.CommandButton cmdDes_Despacho 
         BackColor       =   &H00808080&
         Caption         =   "Comentario para Despachos"
         Enabled         =   0   'False
         Height          =   225
         Index           =   1
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   70
         Tag             =   "Shipment Comments"
         Top             =   6720
         Width           =   9495
      End
      Begin VB.CommandButton cmdDes_General 
         BackColor       =   &H00808080&
         Caption         =   "Comentario General"
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   69
         Tag             =   "General Comments"
         Top             =   5595
         Width           =   9495
      End
      Begin VB.TextBox txtDes_Despacho 
         BackColor       =   &H80000018&
         Height          =   810
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   6945
         Width           =   9525
      End
      Begin VB.TextBox txtDes_General 
         BackColor       =   &H80000018&
         Height          =   645
         Left            =   165
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   5835
         Width           =   9525
      End
      Begin VB.TextBox txtPor_Slush 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   8880
         TabIndex        =   15
         Text            =   "0"
         Top             =   465
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox txtCod_Banco 
         Height          =   285
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   12
         Top             =   4725
         Width           =   615
      End
      Begin VB.TextBox txtNom_Banco 
         Height          =   285
         Left            =   2340
         MaxLength       =   30
         TabIndex        =   65
         Top             =   4725
         Width           =   4035
      End
      Begin VB.TextBox txtCod_Moneda 
         Height          =   285
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   10
         Top             =   3915
         Width           =   615
      End
      Begin VB.TextBox txtNom_Moneda 
         Height          =   285
         Left            =   2340
         MaxLength       =   30
         TabIndex        =   63
         Top             =   3915
         Width           =   4035
      End
      Begin VB.TextBox txtCod_Embarque 
         Height          =   285
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   9
         Top             =   3525
         Width           =   615
      End
      Begin VB.TextBox txtDes_Embarque 
         Height          =   285
         Left            =   2340
         MaxLength       =   30
         TabIndex        =   61
         Top             =   3525
         Width           =   4035
      End
      Begin VB.TextBox txtCod_PagEmb 
         Height          =   285
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   8
         Top             =   3135
         Width           =   615
      End
      Begin VB.TextBox txtDes_PagEmb 
         Height          =   285
         Left            =   2340
         MaxLength       =   30
         TabIndex        =   59
         Top             =   3135
         Width           =   4035
      End
      Begin VB.TextBox txtPor_Comision 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   8880
         TabIndex        =   14
         Text            =   "0"
         Top             =   105
         Width           =   750
      End
      Begin VB.TextBox txtCod_TemCli 
         Height          =   285
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   7
         Top             =   2730
         Width           =   615
      End
      Begin VB.TextBox txtNom_TemCli 
         Height          =   285
         Left            =   2340
         MaxLength       =   30
         TabIndex        =   56
         Top             =   2730
         Width           =   4035
      End
      Begin VB.TextBox txtNom_DivCli 
         Height          =   285
         Left            =   2340
         MaxLength       =   50
         TabIndex        =   55
         Top             =   2370
         Width           =   4035
      End
      Begin VB.TextBox txtCod_DivCli 
         Height          =   285
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   6
         Top             =   2370
         Width           =   615
      End
      Begin VB.TextBox txtDes_Destino 
         Height          =   285
         Left            =   2355
         MaxLength       =   30
         TabIndex        =   53
         Top             =   855
         Width           =   4050
      End
      Begin VB.TextBox txtCod_Destino 
         Height          =   285
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   4
         Top             =   855
         Width           =   615
      End
      Begin VB.TextBox txtNom_Fabrica 
         Height          =   285
         Left            =   2355
         MaxLength       =   50
         TabIndex        =   52
         Top             =   555
         Width           =   4050
      End
      Begin VB.TextBox txtAbr_Fabrica 
         Height          =   285
         Left            =   1710
         MaxLength       =   4
         TabIndex        =   3
         Top             =   555
         Width           =   630
      End
      Begin MSComCtl2.DTPicker dtpFec_DespachoAct 
         Height          =   315
         Left            =   1710
         TabIndex        =   5
         Top             =   1740
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DateIsNull      =   -1  'True
         Format          =   110166016
         CurrentDate     =   37159
      End
      Begin SSDataWidgets_B.SSDBCombo cboCod_ClaPurOrd 
         Height          =   315
         Left            =   1710
         TabIndex        =   0
         Top             =   90
         Width           =   2235
         DataFieldList   =   "Column 0"
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   953
         Columns(0).Caption=   "Clase"
         Columns(0).Name =   "Cod_ClaPurOrd"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3175
         Columns(1).Caption=   "Descripción"
         Columns(1).Name =   "Des_ClaPurOrd"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3200
         Columns(2).Visible=   0   'False
         Columns(2).Caption=   "Num_NivPurOrd"
         Columns(2).Name =   "Num_NivPurOrd"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         _ExtentX        =   3942
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483624
         DataFieldToDisplay=   "Column 1"
      End
      Begin MSComCtl2.DTPicker dtpFec_Emision 
         Height          =   315
         Left            =   5235
         TabIndex        =   128
         Top             =   1410
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         _Version        =   393216
         DateIsNull      =   -1  'True
         Format          =   110166017
         CurrentDate     =   37159
      End
      Begin MSComCtl2.DTPicker dtpFec_LlegadaPO 
         Height          =   315
         Left            =   1710
         TabIndex        =   2
         Top             =   1380
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   110166019
         CurrentDate     =   40263.5095949074
      End
      Begin VB.Label labels 
         Caption         =   "Fecha y Hora Llegada PO"
         Height          =   435
         Index           =   23
         Left            =   180
         TabIndex        =   130
         Tag             =   "Class"
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label labels 
         Caption         =   "Fec.Emisión"
         Height          =   255
         Index           =   22
         Left            =   4380
         TabIndex        =   129
         Tag             =   "Emision Date"
         Top             =   1470
         Width           =   1080
      End
      Begin VB.Label Label8 
         Caption         =   "Grupo"
         Height          =   195
         Left            =   210
         TabIndex        =   126
         Tag             =   "Group"
         Top             =   5100
         Width           =   585
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         FillStyle       =   7  'Diagonal Cross
         Height          =   5310
         Left            =   6990
         Top             =   60
         Width           =   90
      End
      Begin VB.Label labels 
         Caption         =   "Estado de Carta Credito"
         Height          =   360
         Index           =   12
         Left            =   195
         TabIndex        =   71
         Tag             =   "Status L/C"
         Top             =   4290
         Width           =   1365
      End
      Begin VB.Label labels 
         Caption         =   "Slush"
         Height          =   255
         Index           =   11
         Left            =   7350
         TabIndex        =   68
         Tag             =   "Slush"
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label labels 
         Caption         =   "Banco"
         Height          =   255
         Index           =   10
         Left            =   210
         TabIndex        =   67
         Tag             =   "Bank"
         Top             =   4755
         Width           =   1335
      End
      Begin VB.Label labels 
         Caption         =   "Moneda"
         Height          =   255
         Index           =   9
         Left            =   195
         TabIndex        =   64
         Tag             =   "Currency"
         Top             =   3945
         Width           =   1335
      End
      Begin VB.Label labels 
         Caption         =   "Tipo de Embarque"
         Height          =   255
         Index           =   8
         Left            =   180
         TabIndex        =   62
         Tag             =   "Shipment Type"
         Top             =   3555
         Width           =   1335
      End
      Begin VB.Label labels 
         Caption         =   "Pago de  Embarque"
         Height          =   255
         Index           =   7
         Left            =   180
         TabIndex        =   60
         Tag             =   "Shipment Terms"
         Top             =   3150
         Width           =   1440
      End
      Begin VB.Label labels 
         Caption         =   "Comisión"
         Height          =   255
         Index           =   6
         Left            =   7335
         TabIndex        =   58
         Tag             =   "Commision"
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label labels 
         Caption         =   "Temporada"
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   57
         Tag             =   "Season"
         Top             =   2790
         Width           =   1335
      End
      Begin VB.Label labels 
         Caption         =   "División del Cliente"
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   54
         Tag             =   "Client Division"
         Top             =   2445
         Width           =   1335
      End
      Begin VB.Label labels 
         Caption         =   "Destino"
         Height          =   255
         Index           =   3
         Left            =   165
         TabIndex        =   51
         Tag             =   "Destination"
         Top             =   930
         Width           =   1200
      End
      Begin VB.Label labels 
         Caption         =   "Fabrica"
         Height          =   255
         Index           =   2
         Left            =   165
         TabIndex        =   50
         Tag             =   "Fabric"
         Top             =   540
         Width           =   1200
      End
      Begin VB.Label labels 
         Caption         =   "Fecha Despacho (Ex-Factory)"
         Height          =   405
         Index           =   1
         Left            =   180
         TabIndex        =   49
         Tag             =   "Ex-Factory DAte"
         Top             =   1740
         Width           =   1380
      End
      Begin VB.Label labels 
         AutoSize        =   -1  'True
         Caption         =   "Clase de P.O"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Tag             =   "Class"
         Top             =   120
         Width           =   930
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Elegir Estilo Cliente"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7950
      Index           =   1
      Left            =   -15
      TabIndex        =   44
      Tag             =   "2000"
      Top             =   10
      Width           =   9885
      Begin VB.Frame Frame4 
         Caption         =   "Modo de Comisión"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   60
         TabIndex        =   131
         Top             =   2610
         Width           =   9765
         Begin VB.OptionButton optComisionEnPorcentaje 
            Caption         =   "En Porcentaje"
            Height          =   240
            Left            =   210
            TabIndex        =   137
            Tag             =   "%"
            Top             =   285
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optComisionEnImporte 
            Caption         =   "En Importe"
            Height          =   240
            Left            =   1650
            TabIndex        =   136
            Tag             =   "In Importe"
            Top             =   285
            Width           =   1335
         End
         Begin VB.TextBox txtPor_ComisionLOT 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   6450
            TabIndex        =   133
            Text            =   "0"
            Top             =   225
            Width           =   750
         End
         Begin VB.TextBox txtImp_Comision 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   8910
            TabIndex        =   132
            Text            =   "0"
            Top             =   210
            Width           =   750
         End
         Begin VB.Label labels 
            Caption         =   "% de Comisión"
            Height          =   240
            Index           =   18
            Left            =   5340
            TabIndex        =   135
            Tag             =   "Commision %"
            Top             =   270
            Width           =   1155
         End
         Begin VB.Label labels 
            Caption         =   "Importe Comisión"
            Height          =   240
            Index           =   21
            Left            =   7635
            TabIndex        =   134
            Tag             =   "Commision Import"
            Top             =   255
            Width           =   1440
         End
      End
      Begin VB.CommandButton cmdCod_DivPre 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   98
         Top             =   900
         Width           =   405
      End
      Begin VB.CheckBox chkDivPreIgual 
         Caption         =   "Considerar la misma División de Prenda para todos los Colores / Talla"
         ForeColor       =   &H00E97834&
         Height          =   285
         Left            =   2850
         TabIndex        =   93
         Tag             =   "Same Garments Division for All Colors/Size"
         Top             =   960
         Value           =   1  'Checked
         Width           =   5310
      End
      Begin VB.TextBox txtCod_DivPreLOT 
         Height          =   300
         Left            =   1635
         MaxLength       =   3
         TabIndex        =   23
         Top             =   915
         Width           =   630
      End
      Begin VB.CommandButton cmdCod_EstCli 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4470
         TabIndex        =   89
         Top             =   150
         Width           =   405
      End
      Begin VB.CheckBox chkPrecioIgual 
         Caption         =   "Considerar el mismo Precio para todos los Colores / Talla"
         ForeColor       =   &H00E97834&
         Height          =   285
         Left            =   2850
         TabIndex        =   86
         Tag             =   "Same Price for All Colors/Size"
         Top             =   570
         Value           =   1  'Checked
         Width           =   4485
      End
      Begin VB.Frame fraTallas 
         Caption         =   "Tallas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4560
         Left            =   4980
         TabIndex        =   81
         Tag             =   "Size"
         Top             =   3240
         Width           =   4860
         Begin VB.ListBox lstTallasSELEC 
            Appearance      =   0  'Flat
            Columns         =   1
            ForeColor       =   &H00E97834&
            Height          =   3540
            ItemData        =   "Wizard1.frx":0000
            Left            =   2700
            List            =   "Wizard1.frx":0002
            TabIndex        =   97
            Top             =   450
            Width           =   1995
         End
         Begin VB.ListBox lstTallas 
            Appearance      =   0  'Flat
            Columns         =   1
            Height          =   3540
            ItemData        =   "Wizard1.frx":0004
            Left            =   120
            List            =   "Wizard1.frx":0006
            TabIndex        =   96
            Top             =   450
            Width           =   2025
         End
         Begin VB.CommandButton cmdTG_Talla 
            Caption         =   "Tallas ..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   150
            TabIndex        =   91
            Tag             =   "Sizes..."
            Top             =   4170
            Width           =   1995
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Seleccionados"
            Height          =   210
            Index           =   1
            Left            =   2700
            TabIndex        =   85
            Tag             =   "Selected"
            Top             =   225
            Width           =   1995
         End
         Begin VB.CommandButton cmdColores 
            Caption         =   "Para Seleccionar"
            Height          =   210
            Index           =   1
            Left            =   135
            TabIndex        =   84
            Tag             =   "For Selection"
            Top             =   225
            Width           =   1995
         End
         Begin VB.CommandButton cmdTallasToRight 
            Caption         =   ">"
            Height          =   315
            Left            =   2220
            TabIndex        =   31
            Top             =   1470
            Width           =   360
         End
         Begin VB.CommandButton cmdTallasToLeft 
            Caption         =   "<"
            Height          =   315
            Left            =   2220
            TabIndex        =   32
            Top             =   1830
            Width           =   360
         End
         Begin VB.CommandButton cmdAllTallasToRight 
            Caption         =   ">>"
            Height          =   315
            Left            =   2250
            TabIndex        =   33
            Top             =   2370
            Width           =   360
         End
         Begin VB.CommandButton cmdAllTallasToLeft 
            Caption         =   "<<"
            Height          =   315
            Left            =   2235
            TabIndex        =   34
            Top             =   2730
            Width           =   360
         End
      End
      Begin VB.Frame fraColores 
         Caption         =   "Colores"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4560
         Left            =   60
         TabIndex        =   80
         Tag             =   "Colors"
         Top             =   3240
         Width           =   4860
         Begin VB.ListBox lstColores 
            Appearance      =   0  'Flat
            Columns         =   1
            Height          =   3540
            ItemData        =   "Wizard1.frx":0008
            Left            =   120
            List            =   "Wizard1.frx":000A
            TabIndex        =   95
            Top             =   480
            Width           =   1995
         End
         Begin VB.ListBox lstColoresSELEC 
            Appearance      =   0  'Flat
            Columns         =   1
            ForeColor       =   &H00E97834&
            Height          =   3540
            ItemData        =   "Wizard1.frx":000C
            Left            =   2700
            List            =   "Wizard1.frx":000E
            TabIndex        =   94
            Top             =   480
            Width           =   1995
         End
         Begin VB.CommandButton cmdTG_ColCli 
            Caption         =   "Colores ..."
            Height          =   315
            Left            =   135
            TabIndex        =   90
            Tag             =   "Colors...."
            Top             =   4170
            Width           =   1995
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Seleccionados"
            Height          =   210
            Index           =   0
            Left            =   2700
            TabIndex        =   83
            Tag             =   "Selected"
            Top             =   255
            Width           =   1995
         End
         Begin VB.CommandButton cmdColores 
            Caption         =   "Para Seleccionar"
            Height          =   210
            Index           =   0
            Left            =   135
            TabIndex        =   82
            Tag             =   "For Selection"
            Top             =   255
            Width           =   1995
         End
         Begin VB.CommandButton cmdAllColoresToLeft 
            Caption         =   "<<"
            Height          =   315
            Left            =   2220
            TabIndex        =   30
            Top             =   2730
            Width           =   360
         End
         Begin VB.CommandButton cmdAllColoresToRight 
            Caption         =   ">>"
            Height          =   315
            Left            =   2220
            TabIndex        =   29
            Top             =   2370
            Width           =   360
         End
         Begin VB.CommandButton cmdColoresToLeft 
            Caption         =   "<"
            Height          =   315
            Left            =   2220
            TabIndex        =   28
            Top             =   1830
            Width           =   360
         End
         Begin VB.CommandButton cmdColoresToRight 
            Caption         =   ">"
            Height          =   315
            Left            =   2220
            TabIndex        =   27
            Top             =   1500
            Width           =   360
         End
      End
      Begin VB.TextBox txtAbr_FabricaLOT 
         Height          =   300
         Left            =   1635
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   24
         Top             =   1320
         Width           =   630
      End
      Begin VB.TextBox txtNom_FabricaLOT 
         Height          =   300
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   76
         Top             =   1320
         Width           =   4050
      End
      Begin VB.TextBox txtCod_DestinoLOT 
         Height          =   285
         Left            =   1635
         MaxLength       =   3
         TabIndex        =   25
         Top             =   1710
         Width           =   615
      End
      Begin VB.TextBox txtDes_DestinoLOT 
         Height          =   285
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   75
         Top             =   1710
         Width           =   4050
      End
      Begin VB.TextBox txtPrecioLOT 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1635
         TabIndex        =   22
         Text            =   "0"
         Top             =   540
         Width           =   630
      End
      Begin VB.TextBox txtCod_EstCliLOT 
         Height          =   285
         Left            =   1635
         TabIndex        =   21
         Top             =   135
         Width           =   2790
      End
      Begin MSComCtl2.DTPicker dtpFec_DespachoActLOT 
         Height          =   315
         Left            =   1635
         TabIndex        =   26
         Top             =   2100
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   110166016
         CurrentDate     =   37159
      End
      Begin VB.Label labels 
         Caption         =   "Division de Prenda"
         Height          =   255
         Index           =   19
         Left            =   150
         TabIndex        =   92
         Tag             =   "Garment Division"
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label labels 
         Caption         =   "Fecha Fin Producción"
         Height          =   360
         Index           =   17
         Left            =   150
         TabIndex        =   79
         Tag             =   "End Production Date"
         Top             =   2050
         Width           =   1300
      End
      Begin VB.Label labels 
         Caption         =   "Fabrica"
         Height          =   240
         Index           =   16
         Left            =   150
         TabIndex        =   78
         Tag             =   "Fabric"
         Top             =   1365
         Width           =   1200
      End
      Begin VB.Label labels 
         Caption         =   "Destino"
         Height          =   255
         Index           =   15
         Left            =   150
         TabIndex        =   77
         Tag             =   "Destination"
         Top             =   1725
         Width           =   1200
      End
      Begin VB.Label labels 
         Caption         =   "Precio de Vta."
         Height          =   255
         Index           =   14
         Left            =   150
         TabIndex        =   74
         Tag             =   "Price"
         Top             =   585
         Width           =   1335
      End
      Begin VB.Label labels 
         Caption         =   "Estilo del Cliente"
         Height          =   255
         Index           =   13
         Left            =   150
         TabIndex        =   73
         Tag             =   "Client Style"
         Top             =   195
         Width           =   1200
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Matriz de Cantidades"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Index           =   2
      Left            =   -15
      TabIndex        =   45
      Top             =   10
      Width           =   9855
      Begin VB.CommandButton cmdMatrizDestinoEmpaque 
         Caption         =   "Ingresar Color/Talla a Nivel P.O-Destino / Empaque"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7020
         TabIndex        =   127
         Tag             =   "Color/Size Entry - Nivel PO /Destinity/Package"
         Top             =   105
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Frame FraDatos 
         Caption         =   "Datos Generales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   105
         TabIndex        =   116
         Tag             =   "Generales"
         Top             =   30
         Width           =   6855
         Begin VB.TextBox txtEstilo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   3915
            TabIndex        =   120
            Top             =   255
            Width           =   2565
         End
         Begin VB.TextBox txtPO 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   690
            TabIndex        =   118
            Top             =   255
            Width           =   1785
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estilo del Cliente"
            Height          =   195
            Left            =   2670
            TabIndex        =   119
            Tag             =   "Client Style"
            Top             =   300
            Width           =   1170
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nro PO"
            Height          =   195
            Left            =   150
            TabIndex        =   117
            Tag             =   "PO Number"
            Top             =   315
            Width           =   510
         End
      End
      Begin SSDataWidgets_B.SSDBGrid ssgrdDatosCantid 
         Height          =   7020
         Left            =   105
         TabIndex        =   35
         Tag             =   "Quantuty Required Matrix"
         Top             =   840
         Width           =   9705
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeadLines       =   3
         Col.Count       =   0
         HeadFont3D      =   3
         BackColorOdd    =   10354687
         RowHeight       =   423
         ExtraHeight     =   212
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   17119
         _ExtentY        =   12382
         _StockProps     =   79
         Caption         =   "Matriz de Cantidades Requeridas"
         BackColor       =   16777215
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Matriz de Cantidades y Precios"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Index           =   3
      Left            =   -15
      TabIndex        =   46
      Tag             =   "2004"
      Top             =   -10
      Width           =   9855
      Begin SSDataWidgets_B.SSDBGrid SSgrdDatosPrec 
         Height          =   7035
         Left            =   105
         TabIndex        =   36
         Top             =   795
         Width           =   9705
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeadLines       =   3
         Col.Count       =   0
         HeadFont3D      =   3
         BackColorOdd    =   10354687
         RowHeight       =   423
         ExtraHeight     =   212
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   17119
         _ExtentY        =   12409
         _StockProps     =   79
         Caption         =   "Matriz de Precios"
         BackColor       =   16777215
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame fraDatos2 
         Caption         =   "Datos Generales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   105
         TabIndex        =   121
         Top             =   90
         Width           =   9705
         Begin VB.TextBox txtPO2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   870
            TabIndex        =   123
            Top             =   255
            Width           =   1785
         End
         Begin VB.TextBox txtEstilo2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   4275
            TabIndex        =   122
            Top             =   255
            Width           =   2565
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Nro PO"
            Height          =   195
            Left            =   240
            TabIndex        =   125
            Top             =   315
            Width           =   510
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Estilo del Cliente"
            Height          =   195
            Left            =   3030
            TabIndex        =   124
            Top             =   300
            Width           =   1170
         End
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Generar Lote Estilo"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7845
      Index           =   4
      Left            =   -15
      TabIndex        =   47
      Tag             =   "3000"
      Top             =   10
      Width           =   9795
      Begin VB.Label lblStepFinal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "El Sistema generará información de acuerdo a los datos proporcionados por Ud."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1470
         Left            =   3240
         TabIndex        =   48
         Tag             =   "3001"
         Top             =   2370
         Width           =   3960
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   1065
         Index           =   5
         Left            =   0
         Picture         =   "Wizard1.frx":0010
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public oParent As Object
Public sCod_FabricaLot As String
Public sCod_Fabrica As String
Public sLote As String
Public sCod_EStCli As String
Public bGrabar As Boolean
Public bGrabarySalir As Boolean

Public sAccionName As String
Public sModoWizard As String
Public sCod_DestinoLOT As String
Public sCod_Destino As String
Public bHastaNivelDetalle As Boolean
Public dPor_ComisionCliente As Double
Public bSoloUnNum_EstProRea  As Boolean
Public sCod_GruTal As String

Dim sValueCombo As String
Dim bNotFirstActivate As Boolean
Public sCod_EstPro As String
Public bInhabilita As Boolean

Dim sFlag As String

Const NUM_STEPS = 5

Dim aCarga(NUM_STEPS) As Boolean

Const RES_ERROR_MSG = 30000

'BASE VALUE FOR HELP FILE FOR THIS WIZARD:
Const HELP_BASE = 1000
Const HELP_FILE = "MYWIZARD.HLP"

Const BTN_HELP = 0
Const BTN_CANCEL = 1
Const BTN_BACK = 2
Const BTN_NEXT = 3
Const BTN_FINISH = 4

Const STEP_INTRO = 0
Const STEP_1 = 1
Const STEP_2 = 2
Const STEP_3 = 3
Const STEP_4 = 4
Const STEP_FINISH = 4

Const DIR_NONE = 0
Const DIR_BACK = 1
Const DIR_NEXT = 2

Const FRM_TITLE = ""
Const INTRO_KEY = ""
Const SHOW_INTRO = ""
Const TOPIC_TEXT = "<TOPIC_TEXT>"

'module level vars
Dim mnCurStep       As Integer
Dim mbHelpStarted   As Boolean

'Public VBInst       As VBIDE.VBE
Dim mbFinishOK      As Boolean

Dim bGeneraMatrizenTemporal As Boolean

Public sCod_Cliente As String
Public sCod_PurOrd As String

Dim varValorAntiguo As Double

Sub CargaValores()
Me.TxtAd = DevuelveCampo("select Pre_AdicProd from tg_control", cCONNECT)
Me.TxtCritico = DevuelveCampo("select Num_PreCri from tg_control", cCONNECT)
Me.TxtPorc = DevuelveCampo("select Por_AdicProd from tg_control", cCONNECT)
End Sub


Private Sub cmdAceptar_Click()
    Select Case sAccionName
    Case "MODIFICAR"
        If ValidStep(STEP_INTRO, DIR_NEXT, True) Then
            UpdatePurOrd
            oParent.BUSCAR
            Unload Me
        End If
    Case "ELIMINAR"
        DeletePurOrd
        oParent.BUSCAR
        Unload Me
    Case "DETALLEXTALLA"
        Unload Me
    End Select
End Sub

Private Sub cmdAllColoresToLeft_Click()
    ComboBoxToComboBox lstColoresSELEC, lstColores, 1
End Sub

Private Sub cmdAllColoresToRight_Click()
    ComboBoxToComboBox lstColores, lstColoresSELEC, 1
End Sub

Private Sub cmdAllTallasToLeft_Click()
    ComboBoxToComboBox lstTallasSELEC, lstTallas, 1
End Sub

Private Sub cmdAllTallasToRight_Click()
    ComboBoxToComboBox lstTallas, lstTallasSELEC, 1
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub EditBanco(Optional bEnabledCodigo As Boolean)
    Load frmAddRecord
    frmAddRecord.bEnabledCodigo = bEnabledCodigo
    frmAddRecord.LoadFrame frmAddRecord.fraBanco
    frmAddRecord.txtCod_Banco.Text = Me.txtCod_Banco.Text
    frmAddRecord.Show vbModal
    If frmAddRecord.bOk Then
        Me.txtCod_Banco.Text = frmAddRecord.sDato
    End If
    Set frmAddRecord = Nothing
End Sub
Private Sub cmdCod_Banco_Click()
    EditBanco True
End Sub

Sub EditDestino(Optional bEnabledCodigo As Boolean)
    Load frmAddRecord
    frmAddRecord.bEnabledCodigo = bEnabledCodigo
    frmAddRecord.LoadFrame frmAddRecord.fraDestino
    frmAddRecord.txtCod_Destino.Text = Me.txtCod_Destino.Text
    frmAddRecord.Show vbModal
    If frmAddRecord.bOk Then
        Me.txtCod_Destino.Text = frmAddRecord.sDato
    End If
    Set frmAddRecord = Nothing
End Sub
Private Sub cmdCod_Destino_Click()
    EditDestino True
End Sub

Private Sub EditDivCli(Optional bEnabledCodigo As Boolean)
    Load frmAddRecord
    frmAddRecord.bEnabledCodigo = bEnabledCodigo
    frmAddRecord.LoadFrame frmAddRecord.fraDivCli
    frmAddRecord.txtCod_DivCli.Text = Me.txtCod_DivCli.Text
    frmAddRecord.sCod_Cliente = sCod_Cliente
    frmAddRecord.Show vbModal
    If frmAddRecord.bOk Then
        Me.txtCod_DivCli.Text = frmAddRecord.sDato
    End If
    Set frmAddRecord = Nothing
End Sub
Private Sub cmdCod_DivCli_Click()
    EditDivCli True
End Sub

Private Sub cmdCod_DivPre_Click()
    Load frmDivPre
    Set frmDivPre.oParent = Me
    frmDivPre.sCod_DivPRe = Me.txtCod_DivPreLOT.Text
    frmDivPre.Show vbModal
    If frmDivPre.bOk Then
        Me.txtCod_DivPreLOT.Text = frmDivPre.sCod_DivPRe
        If Me.txtAbr_FabricaLOT.Enabled Then
            Me.txtAbr_FabricaLOT.SetFocus
        End If
    End If
    Set frmDivPre = Nothing
End Sub

Sub EditTipEmb(Optional bEnabledCodigo As Boolean)
    Load frmAddRecord
    frmAddRecord.bEnabledCodigo = bEnabledCodigo
    frmAddRecord.LoadFrame frmAddRecord.fraTipEmb
    frmAddRecord.txtCod_Embarque.Text = Me.txtCod_Embarque.Text
    frmAddRecord.Show vbModal
    If frmAddRecord.bOk Then
        Me.txtCod_Embarque.Text = frmAddRecord.sDato
    End If
    Set frmAddRecord = Nothing


End Sub

Private Sub cmdCod_Embarque_Click()
    EditTipEmb True
End Sub

Private Sub cmdCod_EstCli_Click()
    Load frmAddTG_EstCli
    Set frmAddTG_EstCli.oParent = Me
    frmAddTG_EstCli.sCod_Cliente = sCod_Cliente
    frmAddTG_EstCli.sCod_EStCli = Me.txtCod_EstCliLOT
    frmAddTG_EstCli.sCod_TemCli = Me.txtCod_TemCli.Text
    frmAddTG_EstCli.Show vbModal
    If frmAddTG_EstCli.bOk Then
        Me.txtCod_EstCliLOT.Text = frmAddTG_EstCli.sCod_EStCli
        If Me.txtPrecioLOT.Enabled Then
            Me.txtPrecioLOT.SetFocus
        End If
    End If
    Set frmAddTG_EstCli = Nothing
End Sub

Sub EditPagEmb(Optional bEnabledCodigo As Boolean)
    Load frmAddRecord
    frmAddRecord.bEnabledCodigo = bEnabledCodigo
    frmAddRecord.LoadFrame frmAddRecord.fraPagEmb
    frmAddRecord.txtCod_PagEmb.Text = Me.txtCod_PagEmb.Text
    frmAddRecord.Show vbModal
    If frmAddRecord.bOk Then
        Me.txtCod_PagEmb.Text = frmAddRecord.sDato
    End If
    Set frmAddRecord = Nothing
End Sub

Private Sub cmdCod_PagEmb_Click()
    EditPagEmb True
End Sub


Sub EditTemCli(Optional bEnabledCodigo As Boolean)
    Load frmAddRecord
    frmAddRecord.bEnabledCodigo = bEnabledCodigo
    frmAddRecord.LoadFrame frmAddRecord.fraTemCli
    frmAddRecord.txtCod_TemCli.Text = Me.txtCod_TemCli.Text
    frmAddRecord.sCod_Cliente = sCod_Cliente
    frmAddRecord.Show vbModal
    If frmAddRecord.bOk Then
        Me.txtCod_TemCli.Text = frmAddRecord.sDato
    End If
    Set frmAddRecord = Nothing

End Sub

Private Sub cmdCod_TemCli_Click()
    EditTemCli True
End Sub

Private Sub cmdColoresToLeft_Click()
    ComboBoxToComboBox lstColoresSELEC, lstColores, 0
End Sub

Private Sub cmdColoresToRight_Click()
    If bHastaNivelDetalle And RTrim(lstColores.List(lstColores.ListIndex)) = "" Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_INVALID_SELECC
        Exit Sub
    End If
    ComboBoxToComboBox lstColores, lstColoresSELEC, 0
End Sub

Private Sub cmdGrupoPro_Click()
    Dim Strsql As String
    Load frmAddGrupoPro
    Set frmAddGrupoPro.oParent = Me
    Strsql = "SELECT Abr_Cliente FROM TG_CLIENTE WHERE Cod_Cliente = '" & Me.sCod_Cliente & "'"
    frmAddGrupoPro.txtAbr_Cliente = DevuelveCampo(Strsql, cCONNECT)
    Strsql = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE Cod_Cliente = '" & Me.sCod_Cliente & "'"
    frmAddGrupoPro.txtNom_Cliente = DevuelveCampo(Strsql, cCONNECT)
    Strsql = "SELECT Ult_Grupo FROM TG_CLIENTE WHERE Cod_Cliente = '" & Me.sCod_Cliente & "'"
    frmAddGrupoPro.txtCod_GrupoPro = Trim(frmAddGrupoPro.txtAbr_Cliente) & "-" & Format(CStr(DevuelveCampo(Strsql, cCONNECT) + 1), "00#")
    frmAddGrupoPro.Show 1
End Sub

Private Sub cmdMatrizDestinoEmpaque_Click()
    If ssgrdDatosCantid.Rows > 0 Then
        If Not VerificaGruposOperativos Then
            Aviso "Orden de Producción se encuentra en Grupo (Logístico/Textil) Explosionado", 1
        End If
        
        Load frmTG_PurOrdDestinos
        Set frmTG_PurOrdDestinos.oParent = Me
        frmTG_PurOrdDestinos.sFlgOpcion_Nueva = "N"
        frmTG_PurOrdDestinos.sAccionName = sAccionName
        frmTG_PurOrdDestinos.sModoWizard = sModoWizard
        frmTG_PurOrdDestinos.sCod_Cliente = sCod_Cliente
        frmTG_PurOrdDestinos.sCod_PurOrd = sCod_PurOrd
        frmTG_PurOrdDestinos.scod_LotPurOrd = sLote
        frmTG_PurOrdDestinos.sCod_EStCli = Me.txtCod_EstCliLOT.Text
        frmTG_PurOrdDestinos.sCod_TemCli = Me.txtCod_TemCli.Text
        frmTG_PurOrdDestinos.BUSCAR
        frmTG_PurOrdDestinos.Show vbModal
        Set frmTG_PurOrdDestinos = Nothing
    End If
End Sub

Private Sub cmdNav_Click(Index As Integer)
On Error GoTo hand
    Dim nAltStep As Integer
    Dim lHelpTopic As Long
    Dim varCancelar As Integer
    Dim rc As Long
    
    Select Case Index
        Case BTN_HELP
            mbHelpStarted = True
            lHelpTopic = HELP_BASE + 10 * (1 + mnCurStep)
            rc = WinHelp(Me.hwnd, HELP_FILE, HELP_CONTEXT, lHelpTopic)
        
        Case BTN_CANCEL
            'Mensaje MESSAGECODE.kMESSAGE_ASK_CANCEL_PURORD
            'varCancelar = MsgBox("Esta usted seguro de cancelar?.", vbInformation + vbYesNo, "Pregunta")
            varCancelar = MsgBox("Esta usted seguro de cancelar?", vbInformation + vbYesNo, "Pregunta")
            If varCancelar = vbYes Then
                oParent.bChangedPODetalleDestino = False
                Unload Me
            End If
          
        Case BTN_BACK
            'place special cases here to jump
            'to alternate steps
            nAltStep = mnCurStep - 1
            SetStep nAltStep, DIR_BACK
          
        Case BTN_NEXT
            'place special cases here to jump
            'to alternate steps
            If ValidStep(mnCurStep, DIR_NEXT, True) Then
                nAltStep = mnCurStep + 1
                SetStep nAltStep, DIR_NEXT
            End If
            If nAltStep = 1 Then
'                txtCod_EstCliLOT.SetFocus
            End If
        Case BTN_FINISH
            'wizard creation code goes here
            If sAccionName = "ADICIONAR" Or sAccionName = "INCORPORAR" Then
                GenerarInformacion sModoWizard
            End If
            If sAccionName = "MODIFICAR" Then
                UpdateInformacion sModoWizard
            End If
            Unload Me
                    
    End Select
Exit Sub
hand:
ErrorHandler Err, "cmdNav_Click"
End Sub

Private Sub cmdToLeft_Click()
    ComboBoxToComboBox Me.lstColoresSELEC, Me.lstColores, 0
End Sub

Private Sub cmdToRight_Click()
    ComboBoxToComboBox Me.lstColores, Me.lstColoresSELEC, 0
End Sub

Private Sub Command4_Click()
    
End Sub

Private Sub cmdTallasToLeft_Click()
    ComboBoxToComboBox lstTallasSELEC, lstTallas, 0
End Sub

Private Sub cmdTallasToRight_Click()
    If bHastaNivelDetalle And RTrim(lstTallas.List(lstTallas.ListIndex)) = "" Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_INVALID_SELECC
        Exit Sub
    End If

    ComboBoxToComboBox lstTallas, lstTallasSELEC, 0
End Sub

Private Sub cmdTG_ColCli_Click()
    If bSoloUnNum_EstProRea Then
        Load frmTG_ColCli
        frmTG_ColCli.sCod_Cliente = sCod_Cliente
        frmTG_ColCli.sCod_TemCli = Me.txtCod_TemCli.Text
        frmTG_ColCli.sCod_EStCli = Me.txtCod_EstCliLOT.Text
        frmTG_ColCli.sCod_EstPro = Me.sCod_EstPro
        frmTG_ColCli.CargarPresentaciones frmTG_ColCli.sCod_EstPro
        frmTG_ColCli.Show vbModal
        If frmTG_ColCli.bOk Then
            Me.lstColoresSELEC.AddItem frmTG_ColCli.sCod_ColCli
        End If
        Set frmTG_ColCli = Nothing
    Else
        Mensaje kMESSAGE_ERR_STYLE_HAVE_MORE_ESTPRO
    End If
End Sub

Private Sub cmdTG_Talla_Click()
    Load frmTG_Talla
    frmTG_Talla.sCod_GruTal = sCod_GruTal
    frmTG_Talla.Show vbModal
    If frmTG_Talla.bOk Then
        Me.lstTallasSELEC.AddItem frmTG_Talla.sCod_Talla
    End If
    Set frmTG_Talla = Nothing

End Sub

Private Sub Command5_Click(Index As Integer)
    'Con este boton ordenaremos los colores
    If Index = 0 Then
        Call ORDENA_LISTOX(Me.lstColoresSELEC)
    Else
        Call ORDENA_LISTOX(Me.lstTallasSELEC)
    End If
End Sub

Private Sub dtpFec_DespachoAct_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpFec_LlegadaPO.SetFocus
    End If
End Sub

Private Sub dtpFec_Emision_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtCod_DivCli.Enabled Then
            txtCod_DivCli.SetFocus
        End If
    End If
End Sub

Private Sub Form_Activate()
  If Not bNotFirstActivate Then
    If sAccionName = "ADICIONAR" Or sAccionName = "INCORPORAR" Then
        dtpFec_DespachoAct.MinDate = Date
        dtpFec_Emision.value = Date
        If sAccionName = "INCORPORAR" Then
            dtpFec_DespachoActLOT.MinDate = Date
        End If
        Me.txtPor_Comision.Text = dPor_ComisionCliente
        cboCod_ClaPurOrd.value = sValueCombo
        CargaValores
    End If
    bNotFirstActivate = True
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        cmdNav_Click BTN_HELP
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    'init all vars
    Call FormSet(Me)
    'Me.Caption = sCaptionForm
    
    If sAccionName <> "MODIFICAR" Then
        dtpFec_LlegadaPO.value = Now
        dtpFec_LlegadaPO.value = ""
    End If
    
    mbFinishOK = False
    cmdTG_ColCli.Enabled = False
    For i = 0 To NUM_STEPS - 1
      fraStep(i).Left = -10000
    Next
    dtpFec_DespachoAct.value = Date
    
    cmdCod_EstCli.Enabled = False
    SetStep 0, DIR_NONE
    
End Sub

Public Sub SetStep(nStep As Integer, nDirection As Integer)
Dim varStep3 As Boolean
        
    varStep3 = False
    
    Select Case nStep
        Case STEP_INTRO
            LoadDataGeneral
        Case STEP_1
            If sAccionName = "ADICIONAR" Or sAccionName = "INCORPORAR" Then
                Me.dtpFec_DespachoActLOT.value = Me.dtpFec_DespachoAct.value
            End If
            
            LoadDataColores
            LoadDataTallas
            If sAccionName = "MODIFICAR" Or sAccionName = "DETALLEXTALLA" Or sAccionName = "VIEWLOTE" Then
                LoadLOTEST sCod_Cliente, sCod_PurOrd, sLote, sCod_EStCli
                LoadDataColoresSELEC
                LoadDataTallasSELEC
                AddTallaBlanco
            End If
        Case STEP_2
            Me.txtPO.Text = Me.sCod_PurOrd
            Me.txtPO2.Text = Me.sCod_PurOrd
            Me.txtEstilo.Text = Me.txtCod_EstCliLOT
            Me.txtEstilo2.Text = Me.txtCod_EstCliLOT
            
            If sAccionName = "ADICIONAR" Or sAccionName = "INCORPORAR" Then
                GenerarMatrizEnTemporal
                LoadMatrizCantidades
            End If
            If sAccionName = "MODIFICAR" Or sAccionName = "DETALLEXTALLA" Then
                LimpiaMatrizKeyEnTemporal  ' BORRA TM_LOTCOLTAL Y CARGA DESDE TG_LOTCOLTAL
                GenerarMatrizEnTemporalWithKey  'GRABA EN TM_LOTCOLTAL COLOREST/TALLA NUEVOSÇ
                EliminaNoSeleccionadosWithKey
                bInhabilita = InhabilitaModifCantidades ' DEBIDO A LA MATRIZ DEL PO DESTINOS /EMPAQUES
                LoadDataMatrizCantidadesWithKey  'CARGA MATRIZ
                
            End If
            
        Case STEP_3
            If sAccionName = "ADICIONAR" Or sAccionName = "INCORPORAR" Then
                GrabarCantidadesEnTemporal ssgrdDatosCantid, "QR1"
                LoadMatrizPrecios
                
                If chkPrecioIgual.value = 0 Then
                    varStep3 = True
                End If
                
            End If
            
            If sAccionName = "DETALLEXTALLA" Then
                LoadMatrizPrecios
            End If
            If sAccionName = "MODIFICAR" Or sAccionName = "DETALLEXTALLA" Then
                GrabarCantidadesEnTemporal ssgrdDatosCantid, "QR1"
                LoadMatrizPreciosWithKey
            End If
            
'            Call COLOCA_NOMBRECOLOR(Me.SSgrdDatosPrec)
            
            mbFinishOK = False
        Case STEP_FINISH
            GrabarPreciosEnTemporal SSgrdDatosPrec, "PR1"
            mbFinishOK = True
        
    End Select
    
    'move to new step
    fraStep(mnCurStep).Enabled = False
    fraStep(nStep).Left = 0
    fraStep(nStep).Top = fraStep(mnCurStep).Top
    If nStep <> mnCurStep Then
        fraStep(mnCurStep).Left = -10000
    End If
    fraStep(nStep).Enabled = True
  
    SetCaption nStep
    SetNavBtns nStep
    
    If STEP_2 = nStep Then
        Call COLOCA_NOMBRECOLOR(Me.ssgrdDatosCantid)
    End If
    If STEP_3 = nStep Then
        Call COLOCA_NOMBRECOLOR(Me.SSgrdDatosPrec)
    End If
    
    If varStep3 = True Then
    
        'Aqui llamaremos al formulario
        Load frmEleccionPrecios
        Set frmEleccionPrecios.oParent = Me
        Call frmEleccionPrecios.GENERA_GRILLA(lstTallasSELEC)
        frmEleccionPrecios.Show 1
        Set frmEleccionPrecios = Nothing
    
    End If
    
End Sub

Private Sub SetNavBtns(nStep As Integer)
    mnCurStep = nStep
    
    If mnCurStep = 0 Then
        cmdNav(BTN_BACK).Enabled = False
        cmdNav(BTN_NEXT).Enabled = True
    ElseIf mnCurStep = NUM_STEPS - 1 Then
        cmdNav(BTN_NEXT).Enabled = False
        cmdNav(BTN_BACK).Enabled = True
    Else
        cmdNav(BTN_BACK).Enabled = True
        cmdNav(BTN_NEXT).Enabled = True
    End If
    
    If mbFinishOK Then
        cmdNav(BTN_FINISH).Enabled = True
    Else
        cmdNav(BTN_FINISH).Enabled = False
    End If
End Sub

Private Sub SetCaption(nStep As Integer)
    On Error Resume Next

    

End Sub

'=========================================================
'this sub displays an error message when the user has
'not entered enough data to continue
'=========================================================
Sub IncompleteData(nIndex As Integer)
    On Error Resume Next
    Dim sTmp As String
      
    'get the base error message
    sTmp = LoadResString(RES_ERROR_MSG)
    'get the specific message
    sTmp = sTmp & vbCrLf & LoadResString(RES_ERROR_MSG + nIndex)
    Beep
    MsgBox sTmp, vbInformation
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        oParent.bChangedPODetalleDestino = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim rc As Long

    If mbHelpStarted Then
        rc = WinHelp(Me.hwnd, HELP_FILE, HELP_QUIT, 0)
    End If
    
End Sub

Private Sub LoadDataGeneral()
On Error GoTo errores
    Dim vbuff
    Dim obj As New clsTG_PurOrd
    Dim i As Long
    
    If aCarga(STEP_INTRO) = False Then
        'dtpFec_DespachoAct.value = Date
        
        'dtpFec_DespachoActLOT.value = Date
        
        Me.cboCod_ClaPurOrd.TagVariant = cboCod_ClaPurOrd.Cols
        
        Set obj = New clsTG_PurOrd
        obj.ConexionString = cCONNECT
        vbuff = obj.ViewAllClaPurOrd
        
        LibraryVBToSSDBCombo obj, vbuff, cboCod_ClaPurOrd
        Set obj = Nothing
                
'        If Not IsEmpty(vbuff) Then
'            For i = 0 To UBound(vbuff, 2)
'                Me.cboCod_ClaPurOrd.AddItem
'                Me.cboCod_ClaPurOrd.Column(0, i) = vbuff(0, i)
'                Me.cboCod_ClaPurOrd.Column(1, i) = vbuff(1, i)
'                Me.cboCod_ClaPurOrd.Column(2, i) = vbuff(2, i)
'            Next
'            'Me.cboCod_ClaPurOrd.ColumnWidths = "30;30"
'            BuscarComboD cboCod_ClaPurOrd, vbuff(0, 0)
'
'            aCarga(STEP_INTRO) = True
'        End If
    End If
    
Exit Sub
errores:
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    ErrorHandler Err, Err.Description
End Sub

Public Sub LoadDataColores()
On Error GoTo errores
    Dim vbuff
    Dim obj As New clsTG_ColCli
    
    Dim i As Long
    
    If aCarga(STEP_1) = False Then
        Set obj = New clsTG_ColCli
        obj.ConexionString = cCONNECT
        If sAccionName = "ADICIONAR" Or sAccionName = "INCORPORAR" Then
            vbuff = obj.ViewxClie(sCod_Cliente, Me.txtCod_TemCli.Text, Me.txtCod_EstCliLOT.Text)
        Else
            vbuff = obj.ViewxClie(sCod_Cliente, Me.txtCod_TemCli.Text, sCod_EStCli)
        End If
        Set obj = Nothing
        
        lstColores.Clear

        If Not IsEmpty(vbuff) Then
            For i = 0 To UBound(vbuff, 2)
                Me.lstColores.AddItem vbuff(0, i)
            Next
        End If
        
    End If

Exit Sub
errores:
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    ErrorHandler Err, Err.Description

End Sub

Public Sub LoadDataTallas()
On Error GoTo errores
    Dim vbuff
    Dim obj As New clsTG_Talla
    Dim objPO As New clsTG_PurOrdTal
    Dim i As Long
    
    If aCarga(STEP_1) = False Then
        Set obj = New clsTG_Talla
        obj.ConexionString = cCONNECT
        If sAccionName = "ADICIONAR" Or sAccionName = "INCORPORAR" Then
            vbuff = obj.ViewxTalla(sCod_Cliente, Me.txtCod_TemCli.Text, txtCod_EstCliLOT.Text)
        Else
            vbuff = obj.ViewxTalla(sCod_Cliente, Me.txtCod_TemCli.Text, sCod_EStCli)
        End If
        Set obj = Nothing
        
        lstTallas.Clear

        If Not IsEmpty(vbuff) Then
            For i = 0 To UBound(vbuff, 2)
                Me.lstTallas.AddItem vbuff(0, i)
                
            Next
            aCarga(STEP_1) = True
        End If

        
    End If
Exit Sub
errores:
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    ErrorHandler Err, Err.Description

End Sub

Public Function LoadMatrizCantidades() As Boolean
On Error GoTo errores
    Dim vbuff
    Dim obj As New clsTG_LotColTal
    Dim i As Long
    Dim irow As Variant
    
    If aCarga(STEP_3) = False Then
            
        LoadMatrizCantidades = False
        
        irow = Me.ssgrdDatosCantid.Bookmark
        Me.ssgrdDatosCantid.Redraw = False
        
        SSDBGridREmove Me.ssgrdDatosCantid
        
        Set obj = New clsTG_LotColTal
        obj.ConexionString = cCONNECT
        vbuff = obj.ViewMatriz(vusu, sCod_Cliente, sCod_PurOrd, Me.txtCod_EstCliLOT.Text, sCod_FabricaLot, sCod_DestinoLOT, Me.dtpFec_DespachoActLOT.value, txtCod_TemCli.Text)
                
        
        LibraryVBToMatriz obj, vbuff, Me.ssgrdDatosCantid, True, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, True
        'Call Me.SUB_TOTALES
        ssgrdDatosCantid.ActiveRowStyleSet = "RowActive"
        ssgrdDatosCantid.SelectTypeRow = ssSelectionTypeMultiSelectRange
        Set obj = Nothing
        
        If Not IsEmpty(vbuff) Then
            LoadMatrizCantidades = True
        End If
        
        Me.ssgrdDatosCantid.Redraw = True
        
        'aCarga(STEP_3) = True
        Exit Function

    End If
Exit Function
errores:
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    ErrorHandler Err, Err.Description

End Function


Public Function LoadMatrizPrecios() As Boolean
On Error GoTo errores
    Dim vbuff
    Dim obj As New clsTG_LotColTal
    Dim bLoocked As Boolean
    Dim bLockedDivPre      As Boolean
    
    If aCarga(STEP_4) = False Then
        Set obj = New clsTG_LotColTal
        obj.ConexionString = cCONNECT
        vbuff = obj.ViewMatriz(vusu, sCod_Cliente, sCod_PurOrd, Me.txtCod_EstCliLOT.Text, sCod_FabricaLot, sCod_DestinoLOT, Me.dtpFec_DespachoActLOT.value, txtCod_TemCli.Text)
        
        SSDBGridREmove Me.SSgrdDatosPrec
        
        
        If chkPrecioIgual.value = 1 Then
            bLoocked = True
        Else
            bLoocked = False
        End If
        
        If chkDivPreIgual.value = 1 Then
            bLockedDivPre = True
        Else
            bLockedDivPre = False
        End If
        
        LibraryVBToMatriz obj, vbuff, SSgrdDatosPrec, False, False, False, False, True, False, True, False, False, False, False, False, False, False, bLoocked, True, bLockedDivPre
        SSgrdDatosPrec.ActiveRowStyleSet = "RowActive"
        SSgrdDatosPrec.SelectTypeRow = ssSelectionTypeMultiSelectRange
        Set obj = Nothing
        
        If Not IsEmpty(vbuff) Then
            LoadMatrizPrecios = True
        End If
        
        Me.SSgrdDatosPrec.Redraw = True
        'aCarga(STEP_4) = True
        
    End If
Exit Function
errores:
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    ErrorHandler Err, Err.Description

End Function

Public Function LoadMatrizPreciosWithKey() As Boolean
On Error GoTo errores
    Dim vbuff
    Dim obj As New clsTG_LotColTal
    Dim bLoocked As Boolean
    Dim bLockedDivPre      As Boolean
    
    
    
    Set obj = New clsTG_LotColTal
    obj.ConexionString = cCONNECT
    vbuff = obj.ViewMatrizKeyUpdate(vusu, sCod_Cliente, sCod_PurOrd, sLote, Me.txtCod_EstCliLOT.Text)
    
    SSDBGridREmove Me.SSgrdDatosPrec
    
    
    If chkPrecioIgual.value = 1 Then
        bLoocked = True
    Else
        bLoocked = False
    End If
    
    If chkDivPreIgual.value = 1 Then
        bLockedDivPre = True
    Else
        bLockedDivPre = False
    End If
    
    LibraryVBToMatriz obj, vbuff, SSgrdDatosPrec, False, False, False, False, True, False, True, False, False, False, False, False, False, False, bLoocked, True, bLockedDivPre
    SSgrdDatosPrec.ActiveRowStyleSet = "RowActive"
    SSgrdDatosPrec.SelectTypeRow = ssSelectionTypeMultiSelectRange
    Set obj = Nothing
    
    If Not IsEmpty(vbuff) Then
        LoadMatrizPreciosWithKey = True
    End If
    
    Me.SSgrdDatosPrec.Redraw = True
    Exit Function
    
errores:
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    ErrorHandler Err, Err.Description

End Function

Private Sub lstColores_DblClick()
    If bHastaNivelDetalle And RTrim(lstColores.ItemData(lstColores.ListIndex)) = "" Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_INVALID_SELECC
        Exit Sub
    End If

    ComboBoxToComboBox lstColores, lstColoresSELEC, 0
End Sub

Private Sub lstColores_KeyPress(KeyAscii As Integer)
    Dim varIndex As Integer
    If KeyAscii = 13 Then
        If lstColores.ListCount < 1 Or lstColores.ListIndex = -1 Then
            Exit Sub
        End If
        If bHastaNivelDetalle And RTrim(lstColores.ItemData(lstColores.ListIndex)) = "" Then
            Mensaje MESSAGECODE.kMESSAGE_ERR_INVALID_SELECC
            Exit Sub
        End If
        
        varIndex = lstColores.ListIndex
        ComboBoxToComboBox lstColores, lstColoresSELEC, 0
        If varIndex > 0 Then
            varIndex = varIndex - 1
        Else
            varIndex = 0
        End If
        
        If lstColores.ListCount > 0 Then
            lstColores.ListIndex = varIndex
        End If

    End If
End Sub

Private Sub lstColoresSELEC_DBLClick()
    ComboBoxToComboBox lstColoresSELEC, lstColores, 0
End Sub

Private Sub lstColoresSELEC1_DblClick()
    ComboBoxToComboBox lstColoresSELEC, lstColores, 0
End Sub


Private Sub lstTallas_DblClick()
    If bHastaNivelDetalle And RTrim(lstTallas.ItemData(lstTallas.ListIndex)) = "" Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_INVALID_SELECC
        Exit Sub
    End If

    ComboBoxToComboBox lstTallas, lstTallasSELEC, 0
End Sub

Private Sub lstTallas_KeyPress(KeyAscii As Integer)
    Dim varIndex As Integer
    If KeyAscii = 13 Then
        If lstTallas.ListCount < 1 Or lstTallas.ListIndex = -1 Then
            Exit Sub
        End If
        If bHastaNivelDetalle And RTrim(lstTallas.ItemData(lstTallas.ListIndex)) = "" Then
            Mensaje MESSAGECODE.kMESSAGE_ERR_INVALID_SELECC
            Exit Sub
        End If
        
        varIndex = lstTallas.ListIndex
        ComboBoxToComboBox lstTallas, lstTallasSELEC, 0
        If varIndex > 0 Then
            varIndex = varIndex - 1
        Else
            varIndex = 0
        End If
        
        If lstTallas.ListCount > 0 Then
            lstTallas.ListIndex = varIndex
        End If

    End If

End Sub

Private Sub lstTallasSELEC_DblClick()
    ComboBoxToComboBox lstTallasSELEC, lstTallas, 0
End Sub

Private Sub optComisionEnImporte_Click()
    txtImp_Comision.Enabled = True
    txtPor_ComisionLOT.Enabled = False
    txtPor_ComisionLOT.Text = 0

    If Me.txtImp_Comision.Enabled Then
        txtImp_Comision.SetFocus
    End If
End Sub

Private Sub optComisionEnPorcentaje_Click()
    txtPor_ComisionLOT.Enabled = True
    txtImp_Comision.Enabled = False
    txtImp_Comision.Text = 0

    If Me.txtPor_ComisionLOT.Enabled And Me.Visible Then
        txtPor_ComisionLOT.SetFocus
    End If
End Sub

Private Sub ssgrdDatosCantid_AfterUpdate(RtnDispErrMsg As Integer)
    Dim oColumn As Object
'    Dim NroFila As Variant          'ahsp
'    Dim NroColumna As Integer       'ahsp
'    Dim varValorNuevo As Double     'ahsp
'    Dim varProvValorAntiguo As Double 'ahsp

    For Each oColumn In SSgrdDatosPrec.Columns
        If Mid(oColumn.Name, 1, 3) = "DP1" Then
            oColumn.DataType = 8 ' SSgrdDatosPrec.Columns("COD_DIVPRE").DataType = 8
        End If
    Next
  
'    'ahsp
'    NroColumna = ssgrdDatosCantid.Col
'    NroFila = ssgrdDatosCantid.Bookmark
'
'    If NroColumna < 0 Or CStr(NroFila) = CStr(ssgrdDatosCantid.Rows - 1) Then
'        Exit Sub
'    End If
'
'
'    If Mid(ssgrdDatosCantid.Columns(NroColumna).Name, 1, 3) = "QR1" Then
'        varValorNuevo = ssgrdDatosCantid.Columns(NroColumna).value
'        ssgrdDatosCantid.Bookmark = ssgrdDatosCantid.Rows - 1
'        varProvValorAntiguo = ssgrdDatosCantid.Columns(NroColumna).value
'        ssgrdDatosCantid.Columns(NroColumna).value = ssgrdDatosCantid.Columns(NroColumna).value - varValorAntiguo + varValorNuevo
'        varValorAntiguo = varProvValorAntiguo
'        ssgrdDatosCantid.Bookmark = NroFila
'        ssgrdDatosCantid.Col = NroColumna
'    End If
'    ssgrdDatosCantid.Col = NroColumna
  
End Sub

Private Sub ssgrdDatosCantid_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
   If Mid(ssgrdDatosCantid.Columns(ColIndex).Name, 1, 3) = "QR1" Then
        ssgrdDatosCantid.Columns("TOTAL").Text = ssgrdDatosCantid.Columns("TOTAL").value - varValorAntiguo + ssgrdDatosCantid.Columns(ColIndex).value
        ssgrdDatosCantid.Columns("IMPORTE").Text = ssgrdDatosCantid.Columns("TOTAL").value * Me.txtPrecioLOT.Text
        ssgrdDatosCantid.Bookmark = ssgrdDatosCantid.Bookmark
    End If
End Sub

Private Sub ssgrdDatosCantid_BeforeUpdate(Cancel As Integer)
Dim oColumn As Object

    For Each oColumn In SSgrdDatosPrec.Columns
        If Mid(oColumn.Name, 1, 3) = "DP1" Then
            oColumn.DataType = 8 ' SSgrdDatosPrec.Columns("COD_DIVPRE").DataType = 8
        End If
    Next

End Sub


Private Sub ssgrdDatosCantid_Error(ByVal DataError As Integer, Response As Integer)
DataError = 0
Response = 0
End Sub

Private Sub ssgrdDatosCantid_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Trim(ssgrdDatosCantid.Columns(0).Text = "Totales") Then
'        If KeyCode = 46 Then
'            KeyCode = 0
'        End If
'    End If
'    If KeyCode = 39 Or KeyCode = 37 Then
'        'ssgrdDatosCantid.Bookmark = ssgrdDatosCantid.Bookmark
'        'Call ssgrdDatosCantid_AfterUpdate(0)
'    End If
End Sub

Private Sub ssgrdDatosCantid_KeyPress(KeyAscii As Integer)
'    If Trim(ssgrdDatosCantid.Columns(0).Text = "Totales") Then
'        KeyAscii = 0
'    End If
End Sub

Private Sub ssgrdDatosCantid_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    SendKeys "+{TAB}"
    SendKeys "{TAB}"
End Sub

Private Sub ssgrdDatosCantid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    Dim varColumn As Integer
    varColumn = ssgrdDatosCantid.col
    If varColumn < 0 Then
        Exit Sub
    End If
    If Mid(ssgrdDatosCantid.Columns(varColumn).Name, 1, 3) = "QR1" Then
        varValorAntiguo = Val(ssgrdDatosCantid.Columns(varColumn).Text)
    End If
End Sub

Private Sub SSgrdDatosPrec_AfterColUpdate(ByVal ColIndex As Integer)
On Error Resume Next
Dim sColCAlculate As String
Dim sColCantidad As String
Dim sColPrecio As String

    If Mid(SSgrdDatosPrec.Columns(ColIndex).Name, 1, 3) = "PR1" Then
        sColCantidad = "QR1_" + Mid(SSgrdDatosPrec.Columns(ColIndex).Name, 5)
        sColPrecio = "PR1_" + Mid(SSgrdDatosPrec.Columns(ColIndex).Name, 5)
        sColCAlculate = "TR1_" + Mid(SSgrdDatosPrec.Columns(ColIndex).Name, 5)
        SSgrdDatosPrec.Columns(sColCAlculate).Text = SSgrdDatosPrec.Columns(sColCantidad).value * SSgrdDatosPrec.Columns(sColPrecio).value
    End If
End Sub

Private Sub SSgrdDatosPrec_AfterUpdate(RtnDispErrMsg As Integer)
On Error GoTo ERROR1
Dim i As Long
Dim oColumn As Object

    For Each oColumn In SSgrdDatosPrec.Columns
        If Mid(oColumn.Name, 1, 3) = "DP1" Then
            oColumn.DataType = 8 ' SSgrdDatosPrec.Columns("COD_DIVPRE").DataType = 8
        End If
    Next
Exit Sub
ERROR1:
    'MsgBox "1"

End Sub

Private Sub SSgrdDatosPrec_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
       If Mid(SSgrdDatosPrec.Columns(ColIndex).Name, 1, 3) = "PR1" Then
        'SSgrdDatosPrec.Columns("TOTAL").Text = SSgrdDatosPrec.Columns("TOTAL").value - varValorAntiguo + ssgrdDatosCantid.Columns(ColIndex).value
        SSgrdDatosPrec.Columns("IMPORTE").Text = SSgrdDatosPrec.Columns("IMPORTE").value - varValorAntiguo + SSgrdDatosPrec.Columns(ColIndex).value * SSgrdDatosPrec.Columns(ColIndex - 4).value
    End If
End Sub

Private Sub SSgrdDatosPrec_BeforeUpdate(Cancel As Integer)
On Error GoTo ERROR1
Dim i As Long
Dim oColumn As Object

    For Each oColumn In SSgrdDatosPrec.Columns
        If Mid(oColumn.Name, 1, 3) = "DP1" Then
            oColumn.DataType = 8 ' SSgrdDatosPrec.Columns("COD_DIVPRE").DataType = 8
        End If
    Next
Exit Sub
ERROR1:
    'MsgBox "1"
End Sub

Private Sub SSgrdDatosPrec_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    SendKeys "+{TAB}"
    SendKeys "{TAB}"
End Sub

Private Sub SSgrdDatosPrec_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    Dim varColumn As Integer
    varColumn = SSgrdDatosPrec.col
    If varColumn < 0 Then
        Exit Sub
    End If
    If Mid(SSgrdDatosPrec.Columns(varColumn).Name, 1, 3) = "PR1" Then
        varValorAntiguo = CDbl(SSgrdDatosPrec.Columns(varColumn).Text) * CDbl(SSgrdDatosPrec.Columns(varColumn - 4).Text)
    End If
End Sub

Private Sub txtAbr_Fabrica_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sCod_Fabrica = ""
        sFlag = "ABR_FABRICA"
        If Filtrar(sFlag, Me, txtAbr_Fabrica, txtNom_Fabrica) Then
            If txtCod_Destino.Enabled And txtCod_Destino.Visible Then
                txtCod_Destino.SetFocus
            Else
                Me.dtpFec_Emision.SetFocus
            End If
            Me.txtAbr_FabricaLOT.Text = Me.txtAbr_Fabrica.Text
            Me.txtNom_FabricaLOT.Text = Me.txtNom_Fabrica.Text
            Me.sCod_FabricaLot = Me.sCod_Fabrica
        End If
    End If

End Sub

Private Sub txtAbr_FabricaLOT_GotFocus()
    If Not VAlidDivPre(txtCod_DivPreLOT.Text) Then
        If txtCod_DivPreLOT.Enabled Then
            txtCod_DivPreLOT.SetFocus
        End If
    End If
End Sub

Private Sub TxtAd_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
Else
    SoloNumeros TxtAd, KeyAscii, False, 0, 4
End If
End Sub


Private Sub txtCod_Banco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sFlag = "COD_BANCO"
        If Filtrar(sFlag, Me, txtCod_Banco, txtNom_Banco) Then
            If txtCod_GrupoPro.Enabled Then
                txtCod_GrupoPro.SetFocus
            Else
                If txtPor_Comision.Enabled Then
                    txtPor_Comision.SetFocus
                End If
            End If
            'Me.txtPor_Comision.SetFocus
        Else
            EditBanco False
        End If
    End If
End Sub

Private Sub txtCod_GrupoPro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sFlag = "COD_GRUPO"
        If Filtrar(sFlag, Me, txtCod_GrupoPro, txtDes_GrupoPro) Then
            'txtCod_GrupoPro.SetFocus
            Me.txtPor_Comision.SetFocus
        Else
            'Me.cmdCod_Banco.value = True
        End If
    End If

End Sub


Private Sub txtCod_Destino_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sFlag = "COD_DESTINO"
        sCod_Destino = ""
        If Filtrar(sFlag, Me, txtCod_Destino, txtDes_Destino) Then
            Me.dtpFec_LlegadaPO.SetFocus
            'Me.dtpFec_DespachoAct.SetFocus
            Me.txtCod_DestinoLOT.Text = Me.txtCod_Destino.Text
            Me.txtDes_DestinoLOT.Text = Me.txtDes_Destino.Text
            Me.sCod_DestinoLOT = Me.txtCod_Destino.Text
        Else
            EditDestino False
            
        End If
    End If
End Sub

Private Sub txtAbr_FabricaLOT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sFlag = "ABR_FABRICALOT"
        If Filtrar(sFlag, Me, txtAbr_FabricaLOT, txtNom_FabricaLOT) Then
            Me.txtCod_DestinoLOT.SetFocus
        End If
    End If
End Sub

Private Sub txtCod_DestinoLOT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sFlag = "COD_DESTINOLOT"
        If Filtrar(sFlag, Me, txtCod_DestinoLOT, txtDes_DestinoLOT) Then
            Me.dtpFec_DespachoActLOT.SetFocus
        End If
    End If
End Sub

Private Sub txtCod_EstCliLOT_KeyDown(KeyCode As Integer, Shift As Integer)
    bSoloUnNum_EstProRea = False
    sCod_EstPro = ""
    sCod_GruTal = ""
    If KeyCode = vbKeyReturn Then
        sFlag = "COD_ESTCLI"
        If RTrim(txtCod_EstCliLOT.Text) = "" Then
            Filtrar sFlag, Me, txtCod_EstCliLOT, Nothing, True
            txtCod_EstCliLOT_KeyPress (13)
        Else
            'txtPrecioLOT.SetFocus
            'Avanza KeyCode
        End If
    End If
End Sub



Public Sub LoadDataColoresSELEC()
'Este es el codigo antiguo
'On Error GoTo errores
'    Dim vbuff
'
'    Dim objPO As New clsTG_LotColTal
'    Dim i As Long
'    Dim j As Long
'
'    Set objPO = New clsTG_LotColTal
'    objPO.ConexionString= cCONNECT
'    'vbuff = objPO.ViewxClieEstilo(sCod_Cliente, sCod_PurOrd, Me.txtCod_EstCliLOT)
'    vbuff = objPO.ViewColoresSELEC_Matriz(sCod_Cliente, sCod_PurOrd, sLote, Me.txtCod_EstCliLOT)
'    Set objPO = Nothing
'
'    lstColoresSELEC.Clear
'
'    If Not IsEmpty(vbuff) Then
'        For i = 0 To UBound(vbuff, 2)
'            Me.lstColoresSELEC.AddItem vbuff(0, i)
'            For j = lstColores.ListCount - 1 To 0 Step -1
'                If UCase(RTrim(lstColores.List(j))) = UCase(RTrim(vbuff(0, i))) Then
'                    lstColores.RemoveItem (j)
'                    Exit For
'                End If
'            Next
'        Next
'    End If

On Error GoTo errores
    Dim vbuff
    
    Dim objPO As New clsTG_LotColTal
    Dim i As Long
    Dim j As Long
    
    Dim varEncontro As Boolean      'ahsp
    
    Set objPO = New clsTG_LotColTal
    objPO.ConexionString = cCONNECT
    'vbuff = objPO.ViewxClieEstilo(sCod_Cliente, sCod_PurOrd, Me.txtCod_EstCliLOT)
    vbuff = objPO.ViewColoresSELEC_Matriz(sCod_Cliente, sCod_PurOrd, sLote, Me.txtCod_EstCliLOT)
    Set objPO = Nothing
    
    lstColoresSELEC.Clear

    If Not IsEmpty(vbuff) Then
        For i = 0 To UBound(vbuff, 2)
            varEncontro = False
            'Me.lstColoresSELEC.AddItem vbuff(0, i)
            For j = lstColores.ListCount - 1 To 0 Step -1
                If Mid(UCase(RTrim(lstColores.List(j))), 1, Len(vbuff(0, i))) = UCase(vbuff(0, i)) Then
                    varEncontro = True
                    'Aqui anado a los seleccionados el original
                    Me.lstColoresSELEC.AddItem RTrim(lstColores.List(j))
                    lstColores.RemoveItem (j)
                    Exit For
                End If
            Next
            If varEncontro = False Then
                Me.lstColoresSELEC.AddItem vbuff(0, i)
            End If
        Next
    End If


Exit Sub
errores:
    If Not objPO Is Nothing Then
        Set objPO = Nothing
    End If
    
    ErrorHandler Err, Err.Description

End Sub

Public Sub LoadDataTallasSELEC()
On Error GoTo errores
    Dim vbuff
    Dim objPO As New clsTG_LotColTal
    Dim i As Long
    Dim j As Long
    
    Set objPO = New clsTG_LotColTal
    objPO.ConexionString = cCONNECT
    vbuff = objPO.ViewTAllasSELEC_Matriz(sCod_Cliente, sCod_PurOrd, sLote, Me.txtCod_EstCliLOT)
    Set objPO = Nothing
    
    lstTallasSELEC.Clear

    If Not IsEmpty(vbuff) Then
        For i = 0 To UBound(vbuff, 2)
            Me.lstTallasSELEC.AddItem vbuff(0, i)
            For j = lstTallas.ListCount - 1 To 0 Step -1
                If UCase(RTrim(lstTallas.List(j))) = UCase(RTrim(vbuff(0, i))) Then
                    lstTallas.RemoveItem (j)
                    Exit For
                End If
            Next

        Next
    End If

Exit Sub
errores:
    If Not objPO Is Nothing Then
        Set objPO = Nothing
    End If
    
    ErrorHandler Err, Err.Description

End Sub


Public Sub GenerarMatrizEnTemporal()
On Error GoTo errores
    Dim vbuff
    Dim objPO As New clsTG_LotColTal
    Dim iColores As Long
    Dim iTallas As Long
    Dim sCod_ColCli As String
    Dim sCod_Talla As String
    Dim sCod_DivPRe  As String
    Dim dPrecio As Double
    
    Set objPO = New clsTG_LotColTal
    objPO.ConexionString = cCONNECT
    objPO.LimpiaTodoMatrizKeyEnTemporal vusu, sCod_Cliente, sCod_PurOrd, sLote, Me.txtCod_EstCliLOT.Text
    
    For iColores = 0 To lstColoresSELEC.ListCount - 1
        For iTallas = 0 To lstTallasSELEC.ListCount - 1
            sCod_ColCli = lstColoresSELEC.List(iColores)
            sCod_Talla = lstTallasSELEC.List(iTallas)
            
            If chkPrecioIgual.value = "1" Then
                dPrecio = CDbl(Me.txtPrecioLOT.Text)
            Else
                dPrecio = 0
            End If
            
            If chkDivPreIgual.value = "1" Then
                sCod_DivPRe = Me.txtCod_DivPreLOT.Text
            Else
                sCod_DivPRe = ""
            End If
                    
            
            objPO.SaveToTemporal vusu, sCod_Cliente, sCod_PurOrd, "", Me.txtCod_EstCliLOT.Text, sCod_ColCli, sCod_Talla, sCod_FabricaLot, sCod_DestinoLOT, Me.dtpFec_DespachoActLOT, dPrecio, 0, sCod_DivPRe
        Next
    Next
    Set objPO = Nothing
           
Exit Sub
errores:
    If Not objPO Is Nothing Then
        Set objPO = Nothing
    End If
    
    ErrorHandler Err, "GenerarMatrizEnTemporal"
    
End Sub


Public Sub GrabarCantidadesEnTemporal(ByRef ssgrdDatos As SSDataWidgets_B.ssDBGrid, ByVal sColumnUpdate As String)
On Error GoTo errores
    Dim vbuff
    Dim objPO As New clsTG_LotColTal
    Dim iColores As Long
    Dim iTallas As Long
    Dim sCod_ColCli As String
    Dim sCod_Talla As String
    Dim dNum_PreReq As Long
    
    
    
    
    
    Set objPO = New clsTG_LotColTal
    objPO.ConexionString = cCONNECT
    For iColores = 0 To ssgrdDatos.Rows - 1
        ssgrdDatos.Bookmark = iColores
        sCod_ColCli = ssgrdDatos.Columns("Cod_ColCli").Text
        If sCod_ColCli <> "Totales" Then        'AHSP
            For iTallas = 0 To ssgrdDatos.Cols - 1
                If Mid(ssgrdDatos.Columns(iTallas).Name, 1, 3) = UCase(sColumnUpdate) Then
                    sCod_Talla = Mid(ssgrdDatos.Columns(iTallas).Name, 5)
                    dNum_PreReq = ssgrdDatos.Columns(iTallas).value
                    objPO.SaveCantidadesToTemporal vusu, sCod_Cliente, sCod_PurOrd, sLote, Me.txtCod_EstCliLOT.Text, sCod_ColCli, sCod_Talla, sCod_FabricaLot, sCod_DestinoLOT, Me.dtpFec_DespachoActLOT, dNum_PreReq
                End If
            Next
        End If
    Next
    Set objPO = Nothing
           
Exit Sub
errores:
    If Not objPO Is Nothing Then
        Set objPO = Nothing
    End If
    
    ErrorHandler Err, Err.Description
End Sub


Public Sub GrabarPreciosEnTemporal(ByRef ssgrdDatos As SSDataWidgets_B.ssDBGrid, ByVal sColumnUpdate As String)
On Error GoTo errores
    Dim vbuff
    Dim objPO As New clsTG_LotColTal
    Dim iColores As Long
    Dim iTallas As Long
    Dim sCod_ColCli As String
    Dim sCod_Talla As String
    Dim dPrecio As Double
    Dim sCod_DivPRe As String
    
    
    
    
    Set objPO = New clsTG_LotColTal
    objPO.ConexionString = cCONNECT
    For iColores = 0 To ssgrdDatos.Rows - 1
        ssgrdDatos.Bookmark = iColores
        sCod_ColCli = ssgrdDatos.Columns("Cod_ColCli").Text
        
        For iTallas = 0 To ssgrdDatos.Cols - 1
            If Mid(ssgrdDatos.Columns(iTallas).Name, 1, 3) = UCase(sColumnUpdate) Then
                sCod_Talla = Mid(ssgrdDatos.Columns(iTallas).Name, 5)
                dPrecio = ssgrdDatos.Columns(iTallas).value
                sCod_DivPRe = ssgrdDatos.Columns("DP1_" & sCod_Talla).Text
                objPO.SavePreciosToTemporal vusu, sCod_Cliente, sCod_PurOrd, sLote, Me.txtCod_EstCliLOT.Text, sCod_ColCli, sCod_Talla, sCod_FabricaLot, sCod_DestinoLOT, Me.dtpFec_DespachoActLOT, dPrecio, sCod_DivPRe
            End If
        Next
    Next
    Set objPO = Nothing
           
Exit Sub
errores:
    If Not objPO Is Nothing Then
        Set objPO = Nothing
    End If
    
    ErrorHandler Err, "GrabarPreciosEnTemporal"
End Sub


Private Function GenerarInformacion(ByVal sModoWizard As String) As Boolean
On Error GoTo errores
    Dim vbuff
    Dim objPO As clsTG_LotColTal
    Dim iColores As Long
    Dim iTallas As Long
    Dim sCod_ColCli As String
    Dim sCod_Talla As String
    Dim dPrecio As Double
    Dim sFlg_Carta As String
    Dim sFlg_OrdenREorden As String
    Dim sFlg_Regular As String
    Dim sFlg_ComisionEnPorcentaje  As String
    Dim sCod_PurOrd_Retorno  As String
    
    If Me.optOrden.value = True Then
        sFlg_OrdenREorden = "O"
    Else
        sFlg_OrdenREorden = "R"
    End If
    
    If Me.optRegular.value = True Then
        sFlg_Regular = "S"
    Else
        sFlg_Regular = "N"
    End If
    
    
    If Me.optFlg_CartaAprobada.value = True Then
        sFlg_Carta = "S"
    Else
        sFlg_Carta = ""
    End If
    
    If Me.optComisionEnPorcentaje = True Then
        sFlg_ComisionEnPorcentaje = "S"
    Else
        sFlg_ComisionEnPorcentaje = "N"
    End If
    
    Set objPO = New clsTG_LotColTal
    objPO.ConexionString = cCONNECT
        
    sCod_PurOrd_Retorno = objPO.GenerarInformacion(sModoWizard, vusu, sCod_Cliente, sCod_PurOrd, Me.cboCod_ClaPurOrd.value, _
                                                   CStr(Me.dtpFec_DespachoAct.value), Me.txtCod_PagEmb.Text, Me.sCod_Fabrica, Me.sCod_Destino, Me.txtCod_Embarque.Text, _
                                                    Me.txtCod_Moneda.Text, Me.txtCod_DivCli.Text, Me.txtCod_TemCli.Text, sFlg_Carta, Me.txtCod_Banco.Text, _
                                                    CDbl(Me.txtPor_Slush.Text), Me.txtDes_General.Text, Me.txtDes_Despacho.Text, CDbl(Me.txtPor_Comision.Text), _
                                                    "", Me.sCod_FabricaLot, Me.sCod_DestinoLOT, CStr(Me.dtpFec_DespachoActLOT.value), Me.txtCod_EstCliLOT.Text, _
                                                    CDbl(Me.txtPor_ComisionLOT.Text), CDbl(Me.txtPrecioLOT.Text), vusu, CStr(ComputerName()), sFlg_OrdenREorden, sFlg_Regular, _
                                                    Me.TxtPorc, Me.TxtAd, Me.TxtCritico, Me.txtCod_GrupoPro.Text, sFlg_ComisionEnPorcentaje, CDbl(txtImp_Comision.Text), CStr(dtpFec_Emision.value), CStr(Format(dtpFec_LlegadaPO.value, "dd/MM/yyyy HH:mm")))
    
    If Me.cboCod_ClaPurOrd.value = "MU" Then
        oParent.txtCod_PurOrd.Text = sCod_PurOrd_Retorno
    End If
    oParent.BUSCAR
    oParent.BuscarEStilos
    Set objPO = Nothing
    
Exit Function
errores:
    If Not objPO Is Nothing Then
        Set objPO = Nothing
    End If
    
    ErrorHandler Err, Err.Description

End Function

Public Sub SSDBGridREmove(ByRef ssDBGrid As Object)
    Dim i As Long
    Dim n As Long
    
    ssDBGrid.col = 0
    ssDBGrid.SplitterPos = 0
    ssDBGrid.SplitterVisible = False
    ssDBGrid.RemoveAll
    ssDBGrid.Refresh
    ssDBGrid.Redraw = False
    n = ssDBGrid.Cols
    For i = n To ssDBGrid.TagVariant + 1 Step -1
        ssDBGrid.Columns.Remove ssDBGrid.Cols - 1
    Next
    ssDBGrid.Redraw = True
    ssDBGrid.Refresh
End Sub




Public Sub LibraryVBToMatriz(ByRef oData As Object, ByRef pBuff As Variant, ByRef ssDBGrid As SSDataWidgets_B.ssDBGrid, _
ByVal bShowCantidRequOrig As Boolean, bShowCantidRequMod As Boolean, _
ByVal bShowCantidDespOrig As Boolean, bShowCantidDespMod As Boolean, _
ByVal bShowPrecioRequOrig As Boolean, bShowPrecioRequMod As Boolean, _
ByVal bShowImportRequOrig As Boolean, bShowImportRequMod As Boolean, _
ByVal bShowImportDespOrig As Boolean, bShowImportDespMod As Boolean, _
ByVal bShowTitulRequ As Boolean, bShowTitulDesp As Boolean, bShowTitulOrig As Boolean, bShowTitulModif As Boolean, _
ByVal bLookedPRecio As Boolean, ByVal bShowDivPre As Boolean, ByVal bLockedDivPre As Boolean)
On Error Resume Next
Dim rsBuff As LibraryVB.clsRecords
Dim iContador As Long
Dim nCols As Integer
Dim iVerif As Integer
Dim temp As String
Dim NVEZ As Boolean
Dim x%
Dim total1 As Long
Dim Y%
Dim i As Long
Dim ic As Long
Dim iLenCol As Long
Dim sTalla As String
Dim sOrig As String
Dim sModi As String
Dim sRequ As String
Dim sDesp As String

If iLanguage = "1" Then
    If bShowTitulRequ Then
        sRequ = "Requ"
    End If
    If bShowTitulDesp Then
        sDesp = "Desp"
    End If
    If bShowTitulOrig Then
        sOrig = "Orig"
    End If
    If bShowTitulModif Then
        sModi = "Modif"
    End If
Else
    If bShowTitulRequ Then
        sRequ = "Requ"
    End If
    If bShowTitulDesp Then
        sDesp = "Desp"
    End If
    If bShowTitulOrig Then
        sOrig = "Orig"
    End If
    If bShowTitulModif Then
        sModi = "Modif"
    End If
End If

 ssDBGrid.FieldSeparator = "~"
 Set rsBuff = New LibraryVB.clsRecords
 Set rsBuff.RefObject = oData

 rsBuff.Buffer = pBuff
 ssDBGrid.Redraw = False
 nCols = rsBuff.count

 ic = ssDBGrid.Cols
 If ssDBGrid.Cols < nCols Then
    For i = nCols To ic + 1 Step -1
       ssDBGrid.Columns.Add ssDBGrid.Cols    ' "Column" & i, 500, False, Nothing, "Column" & i
       ssDBGrid.Columns(ssDBGrid.Cols - 1).Name = rsBuff(ssDBGrid.Cols).Name
       ssDBGrid.Columns(ssDBGrid.Cols - 1).Caption = rsBuff(ssDBGrid.Cols).Name
    Next i
 End If

 For Y = 0 To ssDBGrid.Cols - 1
   If ssDBGrid.Columns(Y).DataType = 5 Or ssDBGrid.Columns(Y).DataType = 6 Or ssDBGrid.Columns(Y).DataType = 9 Then
      ssDBGrid.Columns(Y).TagVariant = 0
   End If
 Next

 NVEZ = True


 x = 0
 Do While Not rsBuff.EOF
   temp = ""
   For iContador = 0 To nCols - 1
      If NVEZ Then
      End If
      iLenCol = 900
      sTalla = Mid(ssDBGrid.Columns(iContador).Name, 5)
      Select Case Mid(ssDBGrid.Columns(iContador).Name, 1, 3)
        Case "QR1"
            ssDBGrid.Columns(iContador).Visible = bShowCantidRequOrig
            ssDBGrid.Columns(iContador).Caption = "Cantid " + sRequ + " " + sOrig + " " + sTalla
            ssDBGrid.Columns(iContador).Width = iLenCol
            ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
            ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
            ssDBGrid.Columns(iContador).DataType = 5
            If bInhabilita Then
                ssDBGrid.Columns(iContador).Style = 4 'ssStyleButton
            Else
                ssDBGrid.Columns(iContador).Style = ssStyleEdit
            End If
        Case "QR2"
            ssDBGrid.Columns(iContador).Visible = bShowCantidRequMod
            ssDBGrid.Columns(iContador).Caption = "Cantid " + sRequ + " " + sOrig + " " + sTalla
            ssDBGrid.Columns(iContador).Width = iLenCol
            ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
            ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
            ssDBGrid.Columns(iContador).DataType = 5
            If bInhabilita Then
                ssDBGrid.Columns(iContador).Style = 4 'ssStyleButton
            Else
                ssDBGrid.Columns(iContador).Style = ssStyleEdit
            End If
        Case "QD1"
            ssDBGrid.Columns(iContador).Visible = bShowCantidDespOrig
            ssDBGrid.Columns(iContador).Caption = "Cantid " + sDesp + " " + sOrig + " " + sTalla
            ssDBGrid.Columns(iContador).Width = iLenCol
            ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
            ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
            ssDBGrid.Columns(iContador).DataType = 5
            ssDBGrid.Columns(iContador).Style = ssStyleEdit
        Case "QD2"
            ssDBGrid.Columns(iContador).Visible = bShowCantidDespMod
            ssDBGrid.Columns(iContador).Caption = "Cantid " + sDesp + " " + sOrig + " " + sTalla
            ssDBGrid.Columns(iContador).Width = iLenCol
            ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
            ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
            ssDBGrid.Columns(iContador).DataType = 5
            ssDBGrid.Columns(iContador).Style = ssStyleEdit
        Case "PR1"
            ssDBGrid.Columns(iContador).Visible = bShowPrecioRequOrig
            ssDBGrid.Columns(iContador).Caption = "Precio " + sRequ + " " + sOrig + " " + sTalla
            ssDBGrid.Columns(iContador).Width = iLenCol
            ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
            ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
            ssDBGrid.Columns(iContador).DataType = 5
            ssDBGrid.Columns(iContador).Style = ssStyleEdit
            ssDBGrid.Columns(iContador).Locked = bLookedPRecio
        Case "PR2"
            ssDBGrid.Columns(iContador).Visible = bShowPrecioRequMod
            ssDBGrid.Columns(iContador).Caption = "Precio " + sRequ + " " + sOrig + " " + sTalla
            ssDBGrid.Columns(iContador).Width = iLenCol
            ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
            ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
            ssDBGrid.Columns(iContador).DataType = 5
            ssDBGrid.Columns(iContador).Style = ssStyleEdit
            ssDBGrid.Columns(iContador).Locked = bLookedPRecio
        Case "TR1"
            ssDBGrid.Columns(iContador).Locked = True
            ssDBGrid.Columns(iContador).Visible = bShowImportRequOrig
            ssDBGrid.Columns(iContador).Caption = "Importe " + sRequ + " " + sOrig + " " + sTalla
            ssDBGrid.Columns(iContador).Width = iLenCol
            ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
            ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
            ssDBGrid.Columns(iContador).DataType = 5
            ssDBGrid.Columns(iContador).Style = ssStyleEdit
        Case "TR2"
            ssDBGrid.Columns(iContador).Locked = True
            ssDBGrid.Columns(iContador).Visible = bShowImportRequMod
            ssDBGrid.Columns(iContador).Caption = "Importe " + sRequ + " " + sOrig + " " + sTalla
            ssDBGrid.Columns(iContador).Width = iLenCol
            ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
            ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
            ssDBGrid.Columns(iContador).DataType = 5
            ssDBGrid.Columns(iContador).Style = ssStyleEdit
        Case "TD1"
            ssDBGrid.Columns(iContador).Locked = True
            ssDBGrid.Columns(iContador).Visible = bShowImportDespOrig
            ssDBGrid.Columns(iContador).Caption = "Importe " + sDesp + " " + sOrig + " " + sTalla
            ssDBGrid.Columns(iContador).Width = iLenCol
            ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
            ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
            ssDBGrid.Columns(iContador).DataType = 5
            ssDBGrid.Columns(iContador).Style = ssStyleEdit
        Case "TD2"
            ssDBGrid.Columns(iContador).Locked = True
            ssDBGrid.Columns(iContador).Visible = bShowImportDespMod
            ssDBGrid.Columns(iContador).Caption = "Importe " + sDesp + " " + sOrig + " " + sTalla
            ssDBGrid.Columns(iContador).Width = iLenCol
            ssDBGrid.Columns(iContador).Locked = False
            ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
            ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
            ssDBGrid.Columns(iContador).DataType = 5
            ssDBGrid.Columns(iContador).Style = ssStyleEdit
        Case "DP1"
            ssDBGrid.Columns(iContador).Visible = bShowDivPre
            ssDBGrid.Columns(iContador).Caption = "Division " + sDesp + " " + sOrig + " " + sTalla
            ssDBGrid.Columns(iContador).Width = 600
            ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
            ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
            ssDBGrid.Columns(iContador).Locked = bLockedDivPre
            ssDBGrid.Columns(iContador).mask = "AAAA"
            
            ssDBGrid.Columns(iContador).Style = ssStyleEdit
            ssDBGrid.Columns(iContador).DataType = 8
        Case "TOT"
            ssDBGrid.Columns(iContador).Visible = bShowCantidRequOrig
            ssDBGrid.Columns(iContador).Caption = "Totales"
            ssDBGrid.Columns(iContador).Width = 800
            ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
            ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
            ssDBGrid.Columns(iContador).Locked = True
            ssDBGrid.Columns(iContador).mask = "####"
            ssDBGrid.Columns(iContador).Style = 4

        Case "IMP"
            'ssDBGrid.Columns(iContador).Visible = bShowCantidRequOrig      ERA ASI 17/06/2002
            ssDBGrid.Columns(iContador).Visible = bShowPrecioRequOrig
            ssDBGrid.Columns(iContador).Caption = "Importe Total"
            ssDBGrid.Columns(iContador).Width = 800
            ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
            ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
            ssDBGrid.Columns(iContador).Locked = True
            ssDBGrid.Columns(iContador).mask = "####"
            ssDBGrid.Columns(iContador).Style = 4
            
        Case Else
            ssDBGrid.Columns(iContador).Locked = True
            ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
            ssDBGrid.Columns(iContador).Style = 4 'ssStyleButton
      End Select
      
      If UCase(ssDBGrid.Columns(iContador).Name) = "COD_COLCLI" Then
        ssDBGrid.Columns(iContador).Caption = "Color Cliente"
      End If
      
      temp = temp & FixNulos(rsBuff(iContador + 1), vbstring)
      If iContador < nCols - 1 Then
         temp = temp & "~"
      End If

'      If iContador >= FixNulos(ssDBGrid.TagVariant, vbLong) Then
'            ssDBGrid.Columns(iContador).DataType = 5
'            ssDBGrid.Columns(iContador).Alignment = 1
'      End If
      'ssDbgrid.Columns(iContador).DataType = 5
      If ssDBGrid.Columns(iContador).DataType = 5 Or ssDBGrid.Columns(iContador).DataType = 6 Or ssDBGrid.Columns(iContador).DataType = 9 Then
        If Val(FixNulos(rsBuff(iContador + 1), vbDouble)) > 0 Then
            ssDBGrid.Columns(iContador).TagVariant = Val(ssDBGrid.Columns(iContador).TagVariant) + FixNulos(rsBuff(iContador + 1), vbDouble)
        End If
      End If
   Next
   NVEZ = False
   ssDBGrid.AddItem temp
  rsBuff.MoveNext
  x = x + 1
 Loop
 ssDBGrid.AllowDragDrop = True
 ssDBGrid.RowHeight = 300 ' SSDBGrid.RowHeight * 1.25
 ssDBGrid.Refresh

 ssDBGrid.Redraw = True
 Set rsBuff.RefObject = Nothing
 Set rsBuff = Nothing

End Sub

Private Sub txtCod_EstCliLOT_KeyPress(KeyAscii As Integer)
    Dim Strsql As String
    
    If KeyAscii = 13 Then
        If Trim(txtCod_EstCliLOT.Text) = "" Then
            txtCod_EstCliLOT.Text = ""
            'txtCod_EstCliLOT.SetFocus
            Exit Sub
        End If
        
        If Not VAlidEStCli(sCod_Cliente, Me.txtCod_EstCliLOT.Text, Me.txtCod_TemCli.Text) Then
            If txtCod_EstCliLOT.Enabled Then
                txtCod_EstCliLOT.SetFocus
            End If
        Else
            
        
            aCarga(STEP_1) = False
            LoadDataColores
            LoadDataTallas
            LoadDataColoresSELEC
            LoadDataTallasSELEC
            AddTallaBlanco
            
           Strsql = "SELECT Precio FROM TG_LOTEST WHERE " & _
                    "Cod_Cliente ='" & sCod_Cliente & "' AND " & _
                    "Cod_EstCli  ='" & txtCod_EstCliLOT.Text & "' AND " & _
                    "Fec_DespachoAct = (SELECT MAX(Fec_DespachoAct) FROM TG_LOTEST WHERE " & _
                    "Cod_Cliente ='" & sCod_Cliente & "' AND " & _
                    "Cod_EstCli  ='" & txtCod_EstCliLOT.Text & "')"
                    
'Esto todavia se vera, si es correcto debemos colocarlo en c/u de las partes del query
'"Cod_PurOrd  ='" & sCod_PurOrd & "' AND " & _
'"Cod_PurOrd  ='" & sCod_PurOrd & "' AND " & _

            If DevuelveCampo(Strsql, cCONNECT) = "" Then
                txtPrecioLOT.Text = "0"
            Else
                txtPrecioLOT.Text = DevuelveCampo(Strsql, cCONNECT)
            End If
            txtPrecioLOT.SetFocus
            SelectionText txtPrecioLOT
            
            Strsql = "select Flg_Asigna_Version_CosteoPO from tg_Control"
            If DevuelveCampo(Strsql, cCONNECT) = "S" Then
                Load frmHelpEstPro
                frmHelpEstPro.sCod_Cliente = Me.sCod_Cliente
                frmHelpEstPro.sCod_TemCli = Me.txtCod_TemCli.Text
                frmHelpEstPro.sCod_PurOrd = Me.sCod_PurOrd
                frmHelpEstPro.TxtEstCli.Text = Me.txtCod_EstCliLOT.Text
                frmHelpEstPro.CARGA_GRID
                frmHelpEstPro.Show 1
            End If
            
        End If
    End If
    
End Sub

Private Sub txtCod_EstCliLOT_LostFocus()
'    txtCod_EstCliLOT_KeyPress (13)
End Sub

Private Sub txtCod_PagEmb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sFlag = "COD_PAGEMB"
        If Filtrar(sFlag, Me, txtCod_PagEmb, txtDes_PagEmb) Then
            Me.txtCod_Embarque.SetFocus
        Else
            EditPagEmb False
        End If
    End If
End Sub


Private Sub txtCod_embarque_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sFlag = "COD_TIPEMB"
        If Filtrar(sFlag, Me, txtCod_Embarque, txtDes_Embarque) Then
            Me.txtCod_Moneda.SetFocus
        Else
            EditTipEmb False
        End If
    End If
End Sub


Private Sub txtCod_Moneda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sFlag = "COD_MONEDA"
        If Filtrar(sFlag, Me, txtCod_Moneda, txtNom_Moneda) Then
            Me.txtCod_Banco.SetFocus
        End If
    End If
End Sub

Private Sub txtCod_TemCli_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        sFlag = "COD_TEMCLI"
        If Filtrar(sFlag, Me, txtCod_TemCli, txtNom_TemCli) Then
            txtCod_PagEmb.SetFocus
        Else
            EditTemCli False
        End If
    End If
End Sub

Private Sub txtCod_DivCli_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        sFlag = "COD_DIVCLI"
        If Filtrar(sFlag, Me, txtCod_DivCli, txtNom_DivCli) Then
            If txtCod_TemCli.Enabled Then
                txtCod_TemCli.SetFocus
            Else
                If txtCod_PagEmb.Enabled Then
                    txtCod_PagEmb.SetFocus
                End If
            End If
        Else
            EditDivCli False
        End If
    End If
End Sub



Public Function ValidStep(nStep As Integer, nDirection As Integer, ByVal bValidaFabrica As Boolean) As Boolean
Dim aMess(4)
Dim amensaje As clsMessages
Set amensaje = New clsMessages
  
    Select Case nStep
        Case STEP_INTRO
            If cboCod_ClaPurOrd.value = "" Then
                Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
                Exit Function
            End If
            If sCod_Fabrica = "" Then
                Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
                If txtAbr_Fabrica.Enabled Then
                    Me.txtAbr_Fabrica.SetFocus
                End If
                Exit Function
            End If
            If sCod_Destino = "" Then
                Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
                If txtCod_Destino.Enabled Then
                    Me.txtCod_Destino.SetFocus
                End If
                Exit Function
            End If
            If dtpFec_DespachoAct.value = "" Then
                Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
                If dtpFec_DespachoAct.Enabled Then
                    Me.dtpFec_DespachoAct.SetFocus
                End If
                Exit Function
            End If
            
'            If FixNulos(dtpFec_LlegadaPO.value, vbstring) = "" Then
'                MsgBox "Falta ingresar Fecha/Hora Llegada P.O. Revisar", vbCritical, "Validación"
'                If dtpFec_LlegadaPO.Enabled Then
'                    Me.dtpFec_LlegadaPO.SetFocus
'                End If
'                Exit Function
'            End If
                        
            If Not VAlidFechaDespacho(CStr(dtpFec_DespachoAct.value)) Then
                Mensaje MESSAGECODE.kMESSAGE_ERR_INVALID_SELECC
                If dtpFec_DespachoAct.Enabled Then
                    'dtpFec_DespachoAct.SetFocus
                End If
                Exit Function
            End If
            
            If Not ValidCod_DivCli() Then
                Exit Function
            End If
                        
            If Me.txtCod_Moneda.Text = "" Then
                Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
                If txtCod_Moneda.Enabled Then
                    Me.txtCod_Moneda.SetFocus
                End If
                Exit Function
            End If
            
            If cboCod_ClaPurOrd.Columns("Num_NivPurOrd").Text = "1" Then
                If txtCod_TemCli = "" Then
                    Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
                    If txtCod_TemCli.Enabled Then
                        Me.txtCod_TemCli.SetFocus
                    End If
                    Exit Function
                End If
            End If
            
            If txtAbr_Fabrica.Text <> "" Then
                If bValidaFabrica Then
                    If Not ValidCod_Fabrica() Then
                        Exit Function
                    End If
                End If
            End If
            
            If txtCod_Destino.Text <> "" Then
                If Not ValidCod_Destino() Then
                    Exit Function
                End If
            End If
            
            If txtCod_TemCli.Text <> "" Then
                If Not ValidCod_TemCli() Then
                    Exit Function
                End If
            End If
            
            If txtCod_PagEmb.Text <> "" Then
                If Not ValidCod_PagEmb() Then
                    Exit Function
                End If
            End If
            
            If txtCod_Embarque.Text <> "" Then
                If Not ValidCod_Embarque() Then
                    Exit Function
                End If
            End If
            
            If txtCod_Moneda.Text <> "" Then
                If Not ValidCod_Moneda() Then
                    Exit Function
                End If
            End If
            
            If txtCod_Banco.Text <> "" Then
                If Not ValidCod_Banco() Then
                    Exit Function
                End If
            End If
            
            If cboCod_ClaPurOrd.value = sValueCombo Then
                bHastaNivelDetalle = True
            Else
                bHastaNivelDetalle = True       'Por ahora le ponemos esto para considerarlo = a PO
                'bHastaNivelDetalle = False      'Esto era cuando era diferente a PO
            End If
            If sModoWizard = "POCESTDAT" And txtPor_ComisionLOT.Text = 0 Then
                txtPor_ComisionLOT.Text = txtPor_Comision.Text
            End If
            
            If sModoWizard = "   ESTDAT" And txtPor_ComisionLOT.Text = 0 Then
                txtPor_ComisionLOT.Text = txtPor_Comision.Text
            End If
            
            If Not bHastaNivelDetalle Then
                AddTallaBlanco
            Else
                fraColores.Enabled = True
                fraTallas.Enabled = True
            End If
                        
        Case STEP_1
            If RTrim(txtCod_EstCliLOT.Text) = "" Then
                Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
                If Me.txtCod_EstCliLOT.Enabled Then
                    Me.txtCod_EstCliLOT.SetFocus
                End If
                Exit Function
            End If
            If Not VAlidEStCli(sCod_Cliente, Me.txtCod_EstCliLOT.Text, Me.txtCod_TemCli.Text) Then
                If txtCod_EstCliLOT.Enabled Then
                    txtCod_EstCliLOT.SetFocus
                End If
                Exit Function
            End If
            
            
            If txtPrecioLOT.Text = 0 Then
                Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
                If txtPrecioLOT.Enabled Then
                    txtPrecioLOT.SetFocus
                End If
                Exit Function
            End If
        
'            If optComisionEnPorcentaje And CDbl(txtPor_ComisionLOT.Text) <= 0 Then
'                Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
'                If txtPor_ComisionLOT.Enabled Then
'                    txtPor_ComisionLOT.SetFocus
'                    Exit Function
'                End If
'            End If
'
'            If optComisionEnImporte And CDbl(txtImp_Comision.Text) <= 0 Then
'                Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
'                If txtImp_Comision.Enabled Then
'                    txtImp_Comision.SetFocus
'                    Exit Function
'                End If
'            End If
            
            If RTrim(txtCod_DivPreLOT.Text) <> "" Then
                If Not VAlidDivPre(Me.txtCod_DivPreLOT.Text) Then
                    If txtCod_DivPreLOT.Enabled Then
                        txtCod_DivPreLOT.SetFocus
                    End If
                    Exit Function
                End If
            End If
        
            If sCod_FabricaLot = "" Then
                Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
                If txtAbr_FabricaLOT.Enabled Then
                    Me.txtAbr_FabricaLOT.SetFocus
                End If
                Exit Function
            End If
            
            If sCod_DestinoLOT = "" Then
                Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
                If txtCod_DestinoLOT.Enabled Then
                    Me.txtCod_DestinoLOT.SetFocus
                End If
                Exit Function
            End If
        
            If txtAbr_FabricaLOT.Text <> "" Then
                If bValidaFabrica Then
                    If Not ValidCod_FabricaLot() Then
                        Exit Function
                    End If
                End If
            End If
            
            If txtCod_DestinoLOT.Text <> "" Then
                If Not ValidCod_DestinoLot() Then
                    Exit Function
                End If
            End If
        
            If dtpFec_DespachoActLOT.value = "" Then
                Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
                If dtpFec_DespachoActLOT.Enabled Then
                    Me.dtpFec_DespachoActLOT.SetFocus
                End If
                Exit Function
            End If
            If Not VAlidFechaDespacho(CStr(dtpFec_DespachoActLOT.value)) Then
                Mensaje MESSAGECODE.kMESSAGE_ERR_INVALID_SELECC
                If dtpFec_DespachoActLOT.Enabled Then
                    dtpFec_DespachoActLOT.SetFocus
                End If
                Exit Function
            End If
            
            If lstColoresSELEC.ListCount > 0 Then
                If lstTallasSELEC.ListCount = 0 Then
                    Mensaje MESSAGECODE.kMESSAGE_ERR_INVALID_SELECC
                    If lstTallasSELEC.Enabled Then
                        lstTallasSELEC.SetFocus
                    End If
                    Exit Function
                End If
            End If
            If lstTallasSELEC.ListCount > 0 Then
                If lstColoresSELEC.ListCount = 0 Then
                    Mensaje MESSAGECODE.kMESSAGE_ERR_INVALID_SELECC
                    If lstColoresSELEC.Enabled Then
                        lstColoresSELEC.SetFocus
                    End If
                    Exit Function
                End If
            End If
            
'            If lstColoresSELEC.ListCount = 0 And lstTallasSELEC.ListCount = 0 Then
'                If lstColores.Enabled Then
'                    lstColores.SetFocus
'                End If
'                Mensaje MESSAGECODE.kMESSAGE_ERR_INVALID_SELECC
'                Exit Function
'            End If

        Case STEP_2
        Case STEP_3
        Case STEP_FINISH
    End Select
    ValidStep = True
End Function

Private Function ValidCod_DivCli() As Boolean

    sFlag = "COD_DIVCLI"
    If Not Filtrar(sFlag, Me, Me.txtCod_DivCli, Me.txtNom_DivCli, False) Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTFOUND
        If Me.txtCod_DivCli.Enabled Then
            Me.txtCod_DivCli.SetFocus
        End If
        Exit Function
    End If
    ValidCod_DivCli = True
End Function

Private Function ValidCod_TemCli() As Boolean

    sFlag = "COD_TEMCLI"
    If Not Filtrar(sFlag, Me, Me.txtCod_TemCli, Me.txtNom_TemCli, False) Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTFOUND
        If Me.txtCod_TemCli.Enabled Then
            Me.txtCod_TemCli.SetFocus
        End If
        Exit Function
    End If

    ValidCod_TemCli = True
End Function

Private Function ValidCod_PagEmb() As Boolean

    sFlag = "COD_PAGEMB"
    If Not Filtrar(sFlag, Me, Me.txtCod_PagEmb, Me.txtDes_PagEmb, False) Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTFOUND
        If Me.txtCod_PagEmb.Enabled Then
            Me.txtCod_PagEmb.SetFocus
        End If
        Exit Function
    End If

    ValidCod_PagEmb = True
End Function

Private Function ValidCod_Embarque() As Boolean

    sFlag = "COD_TIPEMB"
    If Not Filtrar(sFlag, Me, Me.txtCod_Embarque, Me.txtDes_Embarque, False) Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTFOUND
        If Me.txtCod_Embarque.Enabled Then
            Me.txtCod_Embarque.SetFocus
        End If
        Exit Function
    End If

    ValidCod_Embarque = True
End Function


Private Function ValidCod_Moneda() As Boolean

    sFlag = "COD_MONEDA"
    If Not Filtrar(sFlag, Me, Me.txtCod_Moneda, Me.txtNom_Moneda, False) Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTFOUND
        If Me.txtCod_Moneda.Enabled Then
            Me.txtCod_Moneda.SetFocus
        End If
        Exit Function
    End If

    ValidCod_Moneda = True
End Function

Private Function ValidCod_Banco() As Boolean

    sFlag = "COD_BANCO"
    If Not Filtrar(sFlag, Me, Me.txtCod_Banco, Me.txtNom_Banco, False) Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTFOUND
        If Me.txtCod_Banco.Enabled Then
            Me.txtCod_Banco.SetFocus
        End If
        Exit Function
    End If

    ValidCod_Banco = True
End Function



Private Function ValidCod_Fabrica() As Boolean

    sFlag = "ABR_FABRICA"
    If Not Filtrar(sFlag, Me, Me.txtAbr_Fabrica, Me.txtNom_Fabrica, False) Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTFOUND
        If Me.txtAbr_Fabrica.Enabled Then
            Me.txtAbr_Fabrica.SetFocus
        End If
        Exit Function
    End If

    ValidCod_Fabrica = True
End Function



Private Function ValidCod_FabricaLot() As Boolean

    sFlag = "ABR_FABRICA"
    If Not Filtrar(sFlag, Me, Me.txtAbr_FabricaLOT, Me.txtNom_FabricaLOT, False) Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTFOUND
        If Me.txtAbr_FabricaLOT.Enabled Then
            Me.txtAbr_FabricaLOT.SetFocus
        End If
        Exit Function
    End If

    ValidCod_FabricaLot = True
End Function


Private Function ValidCod_Destino() As Boolean

    sFlag = "COD_DESTINO"
    If Not Filtrar(sFlag, Me, Me.txtCod_Destino, Me.txtDes_Destino, False) Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTFOUND
        If Me.txtCod_Destino.Enabled Then
            Me.txtCod_Destino.SetFocus
        End If
        Exit Function
    End If

    ValidCod_Destino = True
End Function


Private Function ValidCod_DestinoLot() As Boolean

    sFlag = "COD_DESTINO"
    If Not Filtrar(sFlag, Me, Me.txtCod_DestinoLOT, Me.txtDes_DestinoLOT, False) Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTFOUND
        If Me.txtCod_DestinoLOT.Enabled Then
            Me.txtCod_DestinoLOT.SetFocus
        End If
        Exit Function
    End If

    ValidCod_DestinoLot = True
End Function

Private Function UpdatePurOrd() As Boolean
On Error GoTo errores
    Dim vbuff
    Dim objPO As clsTG_LotColTal
    Dim iColores As Long
    Dim iTallas As Long
    Dim sCod_ColCli As String
    Dim sCod_Talla As String
    Dim dPrecio As Double
    Dim sFlg_Carta As String
    Dim sFlg_OrdenREorden As String
    Dim sFlg_Regular As String


    If Me.optOrden.value = True Then
        sFlg_OrdenREorden = "O"
    Else
        sFlg_OrdenREorden = "R"
    End If
    
    If Me.optRegular.value = True Then
        sFlg_Regular = "S"
    Else
        sFlg_Regular = "N"
    End If
    
    
    If Me.optFlg_CartaAprobada.value = True Then
        sFlg_Carta = "A"
    Else
        sFlg_Carta = "S"
    End If
    
    Set objPO = New clsTG_LotColTal
    objPO.ConexionString = cCONNECT
        
    objPO.UpdatePOC sCod_Cliente, sCod_PurOrd, Me.cboCod_ClaPurOrd.value, _
    CStr(Me.dtpFec_DespachoAct.value), Me.txtCod_PagEmb.Text, Me.sCod_Fabrica, Me.sCod_Destino, Me.txtCod_Embarque.Text, _
    Me.txtCod_Moneda.Text, Me.txtCod_DivCli.Text, Me.txtCod_TemCli.Text, sFlg_Carta, Me.txtCod_Banco.Text, _
    CDbl(Me.txtPor_Slush.Text), Me.txtDes_General.Text, Me.txtDes_Despacho.Text, CDbl(Me.txtPor_Comision.Text), sFlg_OrdenREorden, sFlg_Regular, Me.TxtPorc, Me.TxtAd, Me.TxtCritico, Me.txtCod_GrupoPro, vusu, CStr(Me.dtpFec_Emision.value), CStr(Format(Me.dtpFec_LlegadaPO.value, "dd/MM/yyyy HH:mm"))
    
    Set objPO = Nothing
    
Exit Function
errores:
    If Not objPO Is Nothing Then
        Set objPO = Nothing
    End If
    
    ErrorHandler Err, Err.Description

End Function


Public Sub LoadPOC(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String)
On Error Resume Next
'On Error GoTo errores
    Dim vbuff
    Dim objPO As clsTG_LotColTal
    Dim rsBuff As LibraryVB.clsRecords
    Dim varStrsql As String
    Dim i As Integer
    
    Set objPO = New clsTG_LotColTal
    objPO.ConexionString = cCONNECT
        
    Set rsBuff = New LibraryVB.clsRecords
    Set rsBuff.RefObject = objPO
           
    rsBuff.Buffer = objPO.LoadDataPOC(sCod_Cliente, sCod_PurOrd, vusu)
        
    Do While Not rsBuff.EOF
           Me.sCod_PurOrd = FixNulos(rsBuff!cod_purord, vbstring)
           'BuscarComboD cboCod_ClaPurOrd, FixNulos(rsBuff!Cod_ClaPurOrd, vbString)
           cboCod_ClaPurOrd.value = FixNulos(rsBuff!Cod_ClaPurOrd, vbstring)
           cboCod_ClaPurOrd.Enabled = False
           
           Me.dtpFec_DespachoAct.value = FixNulos(rsBuff!Fec_DespachoAct, vbstring)
           Me.txtCod_PagEmb.Text = FixNulos(rsBuff!Cod_PagEmb, vbstring)
           Me.txtDes_PagEmb.Text = FixNulos(rsBuff!Des_PagEmb, vbstring)
           Me.sCod_Fabrica = FixNulos(rsBuff!Cod_Fabrica, vbstring)
           Me.txtNom_Fabrica.Text = FixNulos(rsBuff!Nom_Fabrica, vbstring)
           Me.txtAbr_Fabrica.Text = FixNulos(rsBuff!Abr_Fabrica, vbstring)
           
           Me.txtCod_Destino.Text = FixNulos(rsBuff!Cod_Destino, vbstring)
           Me.sCod_Destino = Me.txtCod_Destino.Text
           Me.txtDes_Destino.Text = FixNulos(rsBuff!Des_Destino, vbstring)
           Me.txtCod_Embarque.Text = FixNulos(rsBuff!Cod_Embarque, vbstring)
           Me.txtDes_Embarque.Text = FixNulos(rsBuff!Des_Embarque, vbstring)
           Me.txtCod_Moneda.Text = FixNulos(rsBuff!Cod_Moneda, vbstring)
           Me.txtNom_Moneda.Text = FixNulos(rsBuff!Nom_Moneda, vbstring)
           Me.txtCod_DivCli.Text = FixNulos(rsBuff!Cod_DivCli, vbstring)
           Me.txtNom_DivCli.Text = FixNulos(rsBuff!Nom_DivCli, vbstring)
           Me.txtCod_TemCli.Text = FixNulos(rsBuff!Cod_TemCli, vbstring)
           Me.txtNom_TemCli.Text = FixNulos(rsBuff!Nom_TemCli, vbstring)
                   
           Me.dtpFec_Emision.value = FixNulos(rsBuff!Fec_Emision, vbstring)

           If FixNulos(rsBuff!Fec_Hora_Llegada_PO, vbstring) <> "" Then
              Me.dtpFec_LlegadaPO.value = FixNulos(rsBuff!Fec_Hora_Llegada_PO, vbstring)
           End If
           If FixNulos(rsBuff!LotPurOrd, vbLong) > 0 Then
                txtCod_TemCli.Enabled = False
                txtNom_TemCli.Enabled = False
                cmdCod_TemCli.Enabled = False
           End If
           
           
           If FixNulos(rsBuff!Cod_TipPurOrd, vbstring) = "O" Then
                Me.optOrden.value = True
                Me.optReorden.value = False
           Else
                Me.optOrden.value = False
                Me.optReorden.value = True
           End If
           
           If FixNulos(rsBuff!Flg_Regular, vbstring) = "S" Then
                Me.optRegular.value = True
                Me.optNoRegular.value = False
           Else
                Me.optRegular.value = False
                Me.optNoRegular.value = True
           End If
                  
           If FixNulos(rsBuff!Flg_Carta, vbstring) = "S" Or FixNulos(rsBuff!Flg_Carta, vbstring) = "" Then
              Me.optFlg_CartaAprobada.value = True
              Me.optFlg_CartaNoAprobada.value = False
           Else
              Me.optFlg_CartaAprobada.value = False
              Me.optFlg_CartaNoAprobada.value = True
           End If
           
           Me.txtCod_Banco.Text = FixNulos(rsBuff!Cod_Banco, vbstring)
           Me.txtNom_Banco.Text = FixNulos(rsBuff!Nom_Banco, vbstring)
           Me.txtPor_Slush.Text = FixNulos(rsBuff!Por_Slush, vbstring)
           Me.txtDes_General.Text = FixNulos(rsBuff!Des_General, vbstring)
           Me.txtDes_Despacho.Text = FixNulos(rsBuff!Des_Despacho, vbstring)
           Me.txtPor_Comision.Text = FixNulos(rsBuff!Por_Comision, vbDouble)
    
            'daniel franco
            Me.TxtAd = FixNulos(rsBuff!Pre_AdicProd, vbInteger)
            Me.TxtPorc = FixNulos(rsBuff!Por_AdicProd, vbDouble)
            Me.TxtCritico = FixNulos(rsBuff!Num_PreCri, vbInteger)
    
    
            Me.txtCod_GrupoPro.Text = FixNulos(rsBuff!Cod_GrupoPro, vbstring)
            Me.txtDes_GrupoPro.Text = FixNulos(rsBuff!Des_GrupoPro, vbstring)
    
            'AHSP
            varStrsql = "SELECT  count(*)  From TG_LOTESTPRO " & _
                        "WHERE   Cod_Cliente  = '" & sCod_Cliente & "'   AND " & _
                                "Cod_PurOrd   = '" & sCod_PurOrd & "'    AND " & _
                                "Cod_OrdPro  <> ''"
    
            'Esto implica que ya tiene un OP asignada
            If DevuelveCampo(varStrsql, cCONNECT) > 0 Then
                Me.txtCod_GrupoPro.Enabled = False
                Me.txtDes_GrupoPro.Enabled = False
                Me.cmdGrupoPro.Enabled = False
            Else
                Me.txtCod_GrupoPro.Enabled = True
                Me.txtDes_GrupoPro.Enabled = True
                Me.cmdGrupoPro.Enabled = True
            End If
            
        rsBuff.MoveNext
    Loop
    
    
    
    Set rsBuff.RefObject = Nothing
    Set rsBuff = Nothing
    Set objPO = Nothing
    
    
'Exit Sub
'errores:
'    If Not objPO Is Nothing Then
'        Set objPO = Nothing
'    End If
'
'    If Not rsBuff.RefObject Is Nothing Then
'        Set rsBuff.RefObject = Nothing
'    End If
'    ErrorHandler Err, Err.Description
End Sub

Private Sub TxtCritico_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
Else
    SoloNumeros TxtCritico, KeyAscii, False, 0, 4
End If
End Sub

Private Sub txtDes_Despacho_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtDes_General_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtImp_Comision_GotFocus()
    SelectionText txtImp_Comision
End Sub

Private Sub txtImp_Comision_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And optComisionEnImporte.value Then
         txtImp_Comision.Text = FixNulos(CDbl(txtImp_Comision.Text), vbDouble)
         txtPor_ComisionLOT.Text = 0
    End If
End Sub

Private Sub txtPor_Comision_GotFocus()
    SelectionText txtPor_Comision
End Sub

Private Sub txtPor_Comision_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtCritico.SetFocus
    End If
End Sub

Private Sub txtPor_ComisionLOT_GotFocus()
    SelectionText txtPor_ComisionLOT
End Sub

Private Sub txtPor_ComisionLOT_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If optComisionEnPorcentaje Then
         txtImp_Comision.Text = 0
    End If
End If
End Sub

Private Sub txtPor_Slush_GotFocus()
    SelectionText txtPor_Slush
End Sub

Private Sub TxtPorc_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
Else
    SoloNumeros TxtPorc, KeyAscii, True, 2, 2
End If
End Sub


Private Sub txtPrecioLOT_GotFocus()
'    If Not VAlidEStCli(sCod_Cliente, Me.txtCod_EstCliLOT.Text, Me.txtCod_TemCli.Text) Then
'        If txtCod_EstCliLOT.Enabled Then
'            txtCod_EstCliLOT.SetFocus
'        End If
'    Else
'        aCarga(STEP_1) = False
'        LoadDataColores
'        LoadDataTallas
'        LoadDataColoresSELEC
'        LoadDataTallasSELEC
'        AddTallaBlanco
'        SelectionText txtPrecioLOT
'
'    End If
'
End Sub

Private Function VAlidEStCli(sCod_Cliente As String, ByVal sCod_EStCli As String, ByVal sCod_TemCli As String) As Boolean
On Error GoTo errores
    Dim vbuff
    Dim obj As clsTG_Cliente
    Dim sValid  As String
    Dim bValid As Boolean
    Dim oMess As clsMessages
    
    Dim sModoAddEstCli  As String
    
    sCod_EstPro = ""
    sCod_GruTal = ""
    
    Set obj = New clsTG_Cliente
    obj.ConexionString = cCONNECT
    sValid = obj.ValidEstilos(sCod_Cliente, sCod_EStCli, sCod_TemCli, sCod_EstPro, sCod_GruTal)
    Set obj = Nothing
    
    If sValid = "EXISTE" Then
        bValid = True
    End If
    
    bSoloUnNum_EstProRea = True
    
    If sValid = "TIENE MAS DE 1 ESTILO PROPIO" Then
        'Mensaje kMESSAGE_ERR_STYLE_HAVE_MORE_ESTPRO
        bSoloUnNum_EstProRea = False
        bValid = True
        VAlidEStCli = bValid
        Exit Function
    End If
    
    If sValid = "NO EXISTE" Then
        bValid = False
    End If
    
    sModoAddEstCli = "ADICIONAR Y ASIGNAR"
    
    If sValid = "ESTILO EXISTE Y NO ASIG A TEMCLI" Then
        Set oMess = New clsMessages
        oMess.Codigo = MESSAGECODE.kMESSAGE_ERR_ASIGN_STYLE_TEMCLI
        oMess.ShowMesage iLanguage
        bValid = False
        VAlidEStCli = bValid
        Exit Function
    End If
               
    If Not bValid Then
        Load frmAddTG_EstCli
        frmAddTG_EstCli.sModoAddEstCli = sModoAddEstCli
        
        Set frmAddTG_EstCli.oParent = Me
        frmAddTG_EstCli.sCod_Cliente = sCod_Cliente
        frmAddTG_EstCli.sCod_TemCli = Me.txtCod_TemCli.Text
        frmAddTG_EstCli.sCod_EStCli = txtCod_EstCliLOT.Text
        frmAddTG_EstCli.txtIdestilo = frmAddTG_EstCli.sCod_EStCli
        If sModoAddEstCli = "SOLO ASIGNACION" Then
            frmAddTG_EstCli.txtIdestilo.Enabled = False
            frmAddTG_EstCli.txtNomestilo.Enabled = False
            frmAddTG_EstCli.txttelaestilo.Enabled = False
        End If
        
        frmAddTG_EstCli.Show vbModal
        VAlidEStCli = frmAddTG_EstCli.bOk
        If VAlidEStCli Then
            VAlidEStCli = VAlidEStCli(sCod_Cliente, txtCod_EstCliLOT.Text, Me.txtCod_TemCli.Text)
        End If
        Set frmAddTG_EstCli = Nothing
    Else
        VAlidEStCli = True
    End If
Exit Function
errores:
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    ErrorHandler Err, Err.Description
End Function

Private Sub AddTallaBlanco()
    Dim vbuff
    Dim obj As Object

    If Not bHastaNivelDetalle Then
        lstColoresSELEC.Clear
        lstColoresSELEC.AddItem ""
        
        lstTallasSELEC.Clear
        lstTallasSELEC.AddItem ""
        
        fraColores.Enabled = False
        fraTallas.Enabled = False
        
        Set obj = New clsTG_ColCli
        obj.ConexionString = cCONNECT
        obj.Add sCod_Cliente, "", "", ""
        Set obj = Nothing
        
        
        
        Set obj = New clsTG_Talla
        obj.ConexionString = cCONNECT
        'daniel franco 26-02-2002 obj.Add ""
        obj.Add "", ""
        Set obj = Nothing
        
    End If
End Sub



Private Function DeletePurOrd() As Boolean
On Error GoTo errores
    Dim vbuff
    Dim objPO As clsTG_LotColTal
    Dim iColores As Long
    Dim iTallas As Long
    Dim sCod_ColCli As String
    Dim sCod_Talla As String
    Dim dPrecio As Double
    Dim sFlg_Carta As String
    

'    Dim oMensaje As clsMessages
    
    Dim oWizard As frmWizard

'    Set oMensaje = New clsMessages
'    oMensaje.Codigo = MESSAGECODE.kMESSAGE_ASK_DELETE_PURORD
'
'
'
'    If Not oMensaje.ShowMesage(iLanguage) Then
'        Exit Function
'    End If
                
    Set objPO = New clsTG_LotColTal
    objPO.ConexionString = cCONNECT
        
    objPO.DeletePOC sCod_Cliente, sCod_PurOrd
    Set objPO = Nothing
    
Exit Function
errores:
    If Not objPO Is Nothing Then
        Set objPO = Nothing
    End If
    ErrorHandler Err, "DeletePurOrd"
'    ErrorHandler Err, Err.Description

End Function



Public Function LoadDataMatrizCantidadesWithKey() As Boolean
On Error GoTo errores
    Dim vbuff
    Dim obj As New clsTG_LotColTal
    Dim i As Long
    Dim irow As Variant
    
    
    
            
    LoadDataMatrizCantidadesWithKey = False
    
    irow = Me.ssgrdDatosCantid.Bookmark
    Me.ssgrdDatosCantid.Redraw = False
    
    SSDBGridREmove Me.ssgrdDatosCantid
    
    Set obj = New clsTG_LotColTal
    obj.ConexionString = cCONNECT
    vbuff = obj.ViewMatrizKeyUpdate(vusu, sCod_Cliente, sCod_PurOrd, sLote, Me.txtCod_EstCliLOT.Text)
                
    LibraryVBToMatriz obj, vbuff, Me.ssgrdDatosCantid, True, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, True
    'Call Me.SUB_TOTALES
    ssgrdDatosCantid.ActiveRowStyleSet = "RowActive"
    ssgrdDatosCantid.SelectTypeRow = ssSelectionTypeMultiSelectRange
    Set obj = Nothing
    
    If Not IsEmpty(vbuff) Then
        LoadDataMatrizCantidadesWithKey = True
    End If
    
    Me.ssgrdDatosCantid.Redraw = True
    
    
    Exit Function

Exit Function
errores:
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    ErrorHandler Err, Err.Description

End Function



Public Sub GenerarMatrizEnTemporalWithKey()
On Error GoTo errores
    Dim vbuff
    Dim objPO As New clsTG_LotColTal
    Dim iColores As Long
    Dim iTallas As Long
    Dim sCod_ColCli As String
    Dim sCod_Talla As String
    Dim sCod_DivPRe  As String
    Dim dPrecio As Double
    
    
    Set objPO = New clsTG_LotColTal
    objPO.ConexionString = cCONNECT
    For iColores = 0 To lstColoresSELEC.ListCount - 1
        For iTallas = 0 To lstTallasSELEC.ListCount - 1
            sCod_ColCli = lstColoresSELEC.List(iColores)
            sCod_Talla = lstTallasSELEC.List(iTallas)
            
            dPrecio = CDbl(Me.txtPrecioLOT.Text)
            
            sCod_DivPRe = Me.txtCod_DivPreLOT.Text
            
            objPO.SaveToTemporal vusu, sCod_Cliente, sCod_PurOrd, sLote, Me.txtCod_EstCliLOT.Text, sCod_ColCli, sCod_Talla, sCod_FabricaLot, sCod_DestinoLOT, Me.dtpFec_DespachoActLOT, dPrecio, 1, sCod_DivPRe
        Next
    Next
    Set objPO = Nothing
           
Exit Sub
errores:
    If Not objPO Is Nothing Then
        Set objPO = Nothing
    End If
    
    ErrorHandler Err, "GenerarMatrizEnTemporalWithKey"
End Sub


Private Function UpdateInformacion(ByVal sModoWizard As String) As Boolean
On Error GoTo errores
    Dim vbuff
    Dim objPO As clsTG_LotColTal
    Dim iColores As Long
    Dim iTallas As Long
    Dim sCod_ColCli As String
    Dim sCod_Talla As String
    Dim dPrecio As Double
    Dim sFlg_ComisionEnPorcentaje  As String
    
    If Me.optComisionEnPorcentaje = True Then
        sFlg_ComisionEnPorcentaje = "S"
    Else
        sFlg_ComisionEnPorcentaje = "N"
    End If
    
    
        
    Set objPO = New clsTG_LotColTal
    objPO.ConexionString = cCONNECT
        
    objPO.UpdateInformacion sModoWizard, vusu, sCod_Cliente, sCod_PurOrd, sLote, Me.txtCod_EstCliLOT.Text, _
    Me.cboCod_ClaPurOrd.value, Me.txtCod_Moneda.Text, _
    Me.sCod_DestinoLOT, CStr(Me.dtpFec_DespachoActLOT.value), _
    CDbl(Me.txtPor_ComisionLOT.Text), vusu, CStr(ComputerName()), sFlg_ComisionEnPorcentaje, CDbl(txtImp_Comision)
    
    oParent.BUSCAR
    oParent.BuscarEStilos
    
    Set objPO = Nothing
    
Exit Function
errores:
    If Not objPO Is Nothing Then
        Set objPO = Nothing
    End If
    
    ErrorHandler Err, Err.Description

End Function



Public Sub LoadLOTEST(ByVal sCod_Cliente As String, ByVal sCod_PurOrd As String, ByVal scod_LotPurOrd As String, ByVal sCod_EStCli As String)
On Error Resume Next
'On Error GoTo errores
    Dim vbuff
    Dim objPO As clsTG_LotColTal
    Dim rsBuff As LibraryVB.clsRecords
    
    Set objPO = New clsTG_LotColTal
    objPO.ConexionString = cCONNECT
        
    Set rsBuff = New LibraryVB.clsRecords
    Set rsBuff.RefObject = objPO
           
    rsBuff.Buffer = objPO.LoadDataLOTEST(sCod_Cliente, sCod_PurOrd, scod_LotPurOrd, sCod_EStCli)
        
    Do While Not rsBuff.EOF
       Me.txtCod_EstCliLOT.Text = sCod_EStCli
       Me.txtPrecioLOT.Text = FixNulos(rsBuff!precio, vbDouble)
       Me.txtAbr_FabricaLOT.Text = FixNulos(rsBuff!Abr_Fabrica, vbstring)
       Me.txtNom_FabricaLOT.Text = FixNulos(rsBuff!Nom_Fabrica, vbstring)
       Me.sCod_FabricaLot = FixNulos(rsBuff!Cod_Fabrica, vbstring)
       
       Me.txtCod_DestinoLOT.Text = FixNulos(rsBuff!Cod_Destino, vbstring)
       Me.sCod_DestinoLOT = Me.txtCod_DestinoLOT.Text
       Me.txtDes_DestinoLOT.Text = FixNulos(rsBuff!Des_Destino, vbstring)
       
       If sAccionName = "MODIFICAR" Then
        Me.txtCod_Destino.Text = FixNulos(rsBuff!Cod_Destino, vbstring)
        Me.sCod_Destino = Me.txtCod_DestinoLOT.Text
        Me.txtDes_Destino.Text = FixNulos(rsBuff!Des_Destino, vbstring)
       
        Me.txtCod_Destino.Locked = True
        Me.txtDes_Destino.Locked = True
       End If
       
       Me.dtpFec_DespachoActLOT.value = FixNulos(rsBuff!Fec_DespachoAct, vbstring)
                             
       Me.txtPor_ComisionLOT.Text = FixNulos(rsBuff!Por_Comision, vbDouble)
       Me.txtCod_DivPreLOT.Text = FixNulos(rsBuff!Cod_DivPre, vbstring)
       
       If FixNulos(rsBuff!Flg_PreDif, vbstring) = "*" Then
            Me.chkPrecioIgual.value = "0"
       Else
            Me.chkPrecioIgual.value = "1"
       End If
       
       If FixNulos(rsBuff!Flg_DivPreDif, vbstring) = "*" Then
            Me.chkDivPreIgual.value = "0"
       Else
            Me.chkDivPreIgual.value = "1"
       End If
       
       If FixNulos(rsBuff!Flg_ComisionEnPorcentaje, vbstring) = "S" Then
            optComisionEnPorcentaje.value = True
       Else
            optComisionEnImporte.value = True
            'txtImp_Comision.Text = FixNulos(rsBuff!Imp_Comision, vbDouble)
            'If FixNulos(rsBuff!precio, vbDouble) = 0 Then
            '    txtPor_ComisionLOT.Text = 0
            'Else
            '    txtPor_ComisionLOT.Text = (FixNulos(rsBuff!Imp_Comision, vbDouble) * 100) / FixNulos(rsBuff!precio, vbDouble)
            'End If
       End If
       txtImp_Comision.Text = FixNulos(rsBuff!IMP_COMISION, vbDouble)
       
       rsBuff.MoveNext
    Loop
    
    Set rsBuff.RefObject = Nothing
    Set rsBuff = Nothing
    Set objPO = Nothing
    
    
'Exit Sub
'errores:
'    If Not objPO Is Nothing Then
'        Set objPO = Nothing
'    End If
'
'    If Not rsBuff.RefObject Is Nothing Then
'        Set rsBuff.RefObject = Nothing
'    End If
'    ErrorHandler Err, Err.Description
End Sub


Private Function LimpiaMatrizKeyEnTemporal() As Boolean
On Error GoTo errores
    Dim vbuff
    Dim objPO As clsTG_LotColTal
           
    Set objPO = New clsTG_LotColTal
    objPO.ConexionString = cCONNECT
        
    objPO.LimpiaMatrizKeyEnTemporal vusu, sCod_Cliente, sCod_PurOrd, sLote, Me.txtCod_EstCliLOT.Text
       
    Set objPO = Nothing
    
Exit Function
errores:
    If Not objPO Is Nothing Then
        Set objPO = Nothing
    End If
    
    ErrorHandler Err, Err.Description
End Function




Public Sub LibraryVBToMatrizResumen(ByRef ssDBGrid As SSDataWidgets_B.ssDBGrid, _
ByVal bShowCantidRequOrig As Boolean, bShowCantidRequMod As Boolean, _
ByVal bShowCantidDespOrig As Boolean, bShowCantidDespMod As Boolean, _
ByVal bShowPrecioRequOrig As Boolean, bShowPrecioRequMod As Boolean, _
ByVal bShowImportRequOrig As Boolean, bShowImportRequMod As Boolean, _
ByVal bShowImportDespOrig As Boolean, bShowImportDespMod As Boolean, _
ByVal bShowTitulRequ As Boolean, bShowTitulDesp As Boolean, bShowTitulOrig As Boolean, bShowTitulModif As Boolean, _
ByVal bLookedPRecio As Boolean, ByVal bShowDivPre As Boolean)
On Error Resume Next
Dim rsBuff As LibraryVB.clsRecords
Dim iContador As Long
Dim nCols As Integer
Dim iVerif As Integer
Dim temp As String
Dim NVEZ As Boolean
Dim x%
Dim total1 As Long
Dim Y%
Dim i As Long
Dim ic As Long
Dim iLenCol As Long
Dim sTalla As String
Dim sOrig As String
Dim sModi As String
Dim sRequ As String
Dim sDesp As String

If iLanguage = "1" Then
    If bShowTitulRequ Then
        sRequ = "Requ"
    End If
    If bShowTitulDesp Then
        sDesp = "Desp"
    End If
    If bShowTitulOrig Then
        sOrig = "Orig"
    End If
    If bShowTitulModif Then
        sModi = "Modif"
    End If
Else
    If bShowTitulRequ Then
        sRequ = "Requ"
    End If
    If bShowTitulDesp Then
        sDesp = "Desp"
    End If
    If bShowTitulOrig Then
        sOrig = "Orig"
    End If
    If bShowTitulModif Then
        sModi = "Modif"
    End If
End If

 iLenCol = 900
 For iContador = 0 To ssDBGrid.Cols - 1
    sTalla = Mid(ssDBGrid.Columns(iContador).Name, 5)
    Select Case Mid(ssDBGrid.Columns(iContador).Name, 1, 3)
      Case "QR1"
          ssDBGrid.Columns(iContador).Visible = bShowCantidRequOrig
          ssDBGrid.Columns(iContador).Caption = "Cantid " + sRequ + " " + sOrig + " " + sTalla
          ssDBGrid.Columns(iContador).Width = iLenCol
          ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
          ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
          ssDBGrid.Columns(iContador).DataType = 5
          ssDBGrid.Columns(iContador).Style = ssStyleEdit
      Case "QR2"
          ssDBGrid.Columns(iContador).Visible = bShowCantidRequMod
          ssDBGrid.Columns(iContador).Caption = "Cantid " + sRequ + " " + sOrig + " " + sTalla
          ssDBGrid.Columns(iContador).Width = iLenCol
          ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
          ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
          ssDBGrid.Columns(iContador).DataType = 5
          ssDBGrid.Columns(iContador).Style = ssStyleEdit
      Case "QD1"
          ssDBGrid.Columns(iContador).Visible = bShowCantidDespOrig
          ssDBGrid.Columns(iContador).Caption = "Cantid " + sDesp + " " + sOrig + " " + sTalla
          ssDBGrid.Columns(iContador).Width = iLenCol
          ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
          ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
          ssDBGrid.Columns(iContador).DataType = 5
          ssDBGrid.Columns(iContador).Style = ssStyleEdit
      Case "QD2"
          ssDBGrid.Columns(iContador).Visible = bShowCantidDespMod
          ssDBGrid.Columns(iContador).Caption = "Cantid " + sDesp + " " + sOrig + " " + sTalla
          ssDBGrid.Columns(iContador).Width = iLenCol
          ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
          ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
          ssDBGrid.Columns(iContador).DataType = 5
          ssDBGrid.Columns(iContador).Style = ssStyleEdit
      Case "PR1"
          ssDBGrid.Columns(iContador).Visible = bShowPrecioRequOrig
          ssDBGrid.Columns(iContador).Caption = "Precio " + sRequ + " " + sOrig + " " + sTalla
          ssDBGrid.Columns(iContador).Width = iLenCol
          ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
          ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
          ssDBGrid.Columns(iContador).DataType = 5
          ssDBGrid.Columns(iContador).Style = ssStyleEdit
          ssDBGrid.Columns(iContador).Locked = bLookedPRecio
      Case "PR2"
          ssDBGrid.Columns(iContador).Visible = bShowPrecioRequMod
          ssDBGrid.Columns(iContador).Caption = "Precio " + sRequ + " " + sOrig + " " + sTalla
          ssDBGrid.Columns(iContador).Width = iLenCol
          ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
          ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
          ssDBGrid.Columns(iContador).DataType = 5
          ssDBGrid.Columns(iContador).Style = ssStyleEdit
          ssDBGrid.Columns(iContador).Locked = bLookedPRecio
      Case "TR1"
          ssDBGrid.Columns(iContador).Locked = True
          ssDBGrid.Columns(iContador).Visible = bShowImportRequOrig
          ssDBGrid.Columns(iContador).Caption = "Importe " + sRequ + " " + sOrig + " " + sTalla
          ssDBGrid.Columns(iContador).Width = iLenCol
          ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
          ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
          ssDBGrid.Columns(iContador).DataType = 5
          ssDBGrid.Columns(iContador).Style = ssStyleEdit
      Case "TR2"
          ssDBGrid.Columns(iContador).Locked = True
          ssDBGrid.Columns(iContador).Visible = bShowImportRequMod
          ssDBGrid.Columns(iContador).Caption = "Importe " + sRequ + " " + sOrig + " " + sTalla
          ssDBGrid.Columns(iContador).Width = iLenCol
          ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
          ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
          ssDBGrid.Columns(iContador).DataType = 5
          ssDBGrid.Columns(iContador).Style = ssStyleEdit
      Case "TD1"
          ssDBGrid.Columns(iContador).Locked = True
          ssDBGrid.Columns(iContador).Visible = bShowImportDespOrig
          ssDBGrid.Columns(iContador).Caption = "Importe " + sDesp + " " + sOrig + " " + sTalla
          ssDBGrid.Columns(iContador).Width = iLenCol
          ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
          ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
          ssDBGrid.Columns(iContador).DataType = 5
          ssDBGrid.Columns(iContador).Style = ssStyleEdit
      Case "TD2"
          ssDBGrid.Columns(iContador).Locked = True
          ssDBGrid.Columns(iContador).Visible = bShowImportDespMod
          ssDBGrid.Columns(iContador).Caption = "Importe " + sDesp + " " + sOrig + " " + sTalla
          ssDBGrid.Columns(iContador).Width = iLenCol
          ssDBGrid.Columns(iContador).Locked = False
          ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
          ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
          ssDBGrid.Columns(iContador).DataType = 5
          ssDBGrid.Columns(iContador).Style = ssStyleEdit
      Case "DP1"
          ssDBGrid.Columns(iContador).Visible = bShowDivPre
          ssDBGrid.Columns(iContador).Caption = "Division " + sDesp + " " + sOrig + " " + sTalla
          ssDBGrid.Columns(iContador).Width = 600
          ssDBGrid.Columns(iContador).Alignment = ssCaptionAlignmentRight
          ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
          ssDBGrid.Columns(iContador).DataType = 8
          ssDBGrid.Columns(iContador).Locked = False
          ssDBGrid.Columns(iContador).mask = "####"
          
          ssDBGrid.Columns(iContador).Style = ssStyleEdit
          
      Case Else
          ssDBGrid.Columns(iContador).Locked = True
          ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
          ssDBGrid.Columns(iContador).Style = 4 'ssStyleButton
    End Select
    
    ssDBGrid.Columns(iContador).Style = 4 'ssStyleButton
 Next

End Sub



Public Sub LibraryVBToSSDBCombo(ByRef oData As Object, ByRef pBuff As Variant, ByRef ssDBCombo As SSDataWidgets_B.ssDBCombo)
On Error Resume Next
Dim rsBuff As LibraryVB.clsRecords
Dim iContador As Long
Dim nCols As Integer
Dim iVerif As Integer
Dim temp As String
Dim NVEZ As Boolean
Dim x%
Dim total1 As Long
Dim Y%
Dim i As Long
Dim ic As Long
Dim bPrimero As Boolean

 ssDBCombo.FieldSeparator = "~"
 Set rsBuff = New LibraryVB.clsRecords
 Set rsBuff.RefObject = oData

 rsBuff.Buffer = pBuff
 ssDBCombo.Redraw = False
 nCols = rsBuff.count

 ic = ssDBCombo.Cols
 If ssDBCombo.Cols < nCols Then
    For i = nCols To ic + 1 Step -1
       ssDBCombo.Columns.Add ssDBCombo.Cols    ' "Column" & i, 500, False, Nothing, "Column" & i
       ssDBCombo.Columns(ssDBCombo.Cols - 1).Name = rsBuff(ssDBCombo.Cols).Name
       ssDBCombo.Columns(ssDBCombo.Cols - 1).Caption = rsBuff(ssDBCombo.Cols).Name
    Next i
 End If

 For Y = 0 To ssDBCombo.Cols - 1
   If ssDBCombo.Columns(Y).DataType = 5 Or ssDBCombo.Columns(Y).DataType = 6 Or ssDBCombo.Columns(Y).DataType = 9 Then
      ssDBCombo.Columns(Y).TagVariant = 0
   End If
 Next

 NVEZ = True

 bPrimero = True
 x = 0
 ssDBCombo.RemoveAll
 Do While Not rsBuff.EOF
   temp = ""
   For iContador = 0 To nCols - 1
      ssDBCombo.Columns(iContador).Locked = True
      ssDBCombo.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
      ssDBCombo.Columns(iContador).Style = 4 'ssStyleButton
      ssDBCombo.Columns(iContador).Width = 2000
      temp = temp & FixNulos(rsBuff(iContador + 1), vbstring)
      
      If rsBuff(iContador + 1).value = "1" And iContador = 2 Then
        sValueCombo = rsBuff("Cod_ClaPurORd").value
        bPrimero = False
      End If
      
      If iContador < nCols - 1 Then
         temp = temp & "~"
      End If

      If iContador >= FixNulos(ssDBCombo.TagVariant, vbLong) Then
            ssDBCombo.Columns(iContador).DataType = 5
            ssDBCombo.Columns(iContador).Alignment = 1
      End If

      'ssdbCombo.Columns(iContador).DataType = 5
      If ssDBCombo.Columns(iContador).DataType = 5 Or ssDBCombo.Columns(iContador).DataType = 6 Or ssDBCombo.Columns(iContador).DataType = 9 Or iContador > FixNulos(ssDBCombo.TagVariant, vbLong) Then
        If Val(FixNulos(rsBuff(iContador + 1), vbDouble)) > 0 Then
            ssDBCombo.Columns(iContador).TagVariant = Val(ssDBCombo.Columns(iContador).TagVariant) + FixNulos(rsBuff(iContador + 1), vbDouble)
        End If
      End If
   Next
   NVEZ = False
   ssDBCombo.AddItem temp
  rsBuff.MoveNext
  x = x + 1
 Loop
 
 ssDBCombo.RowHeight = 300 ' ssdbCombo.RowHeight * 1.25
 ssDBCombo.Refresh

 ssDBCombo.Redraw = True
 Set rsBuff.RefObject = Nothing
 Set rsBuff = Nothing

End Sub


Private Function VAlidDivPre(sCod_DivPRe As String) As Boolean
On Error GoTo errores
    Dim vbuff
    Dim obj As clsTG_LotColTal
    Dim bValid  As Boolean
    
    Set obj = New clsTG_LotColTal
    obj.ConexionString = cCONNECT
    bValid = obj.VAlidDivPre(sCod_DivPRe)
    Set obj = Nothing
    
    If Not bValid Then
        Load frmDivPre
        Set frmDivPre.oParent = Me
        frmDivPre.sCod_DivPRe = Me.txtCod_DivPreLOT.Text
        frmDivPre.txtCod_DivPre.Text = frmDivPre.sCod_DivPRe
        frmDivPre.Show vbModal
        If frmDivPre.bOk Then
            Me.txtCod_DivPreLOT.Text = frmDivPre.sCod_DivPRe
            If Me.txtAbr_FabricaLOT.Enabled Then
                Me.txtAbr_FabricaLOT.SetFocus
            End If
            VAlidDivPre = frmDivPre.bOk
        End If
        
        Set frmDivPre = Nothing
        
    Else
        VAlidDivPre = True
    End If
Exit Function
errores:
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    ErrorHandler Err, Err.Description
End Function


Private Sub txtCod_DivPreLOT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sFlag = "COD_DIVPRE"
        If RTrim(txtCod_DivPreLOT.Text) = "" Then
            Filtrar sFlag, Me, txtCod_DivPreLOT, Nothing, True
        Else
            'Avanza KeyCode
        End If
        'dtpFec_DespachoOriLOT.SetFocus
        dtpFec_DespachoActLOT.SetFocus
    End If
End Sub




Private Function VAlidFechaDespacho(dFecha As String) As Boolean
On Error GoTo errores
    Dim vbuff
    Dim obj As clsTG_LotColTal
    Dim iRet As Integer
    
    Set obj = New clsTG_LotColTal
    obj.ConexionString = cCONNECT
    iRet = obj.VAlidFechaDespacho(dFecha)
    Set obj = Nothing
    
    If iRet = 0 Then
        VAlidFechaDespacho = True
    Else
        VAlidFechaDespacho = False
    End If
Exit Function
errores:
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    ErrorHandler Err, Err.Description
End Function


Private Function EliminaNoSeleccionadosWithKey() As Boolean
On Error GoTo errores
    Dim vbuff
    Dim objPO As clsTG_LotColTal
           
    Set objPO = New clsTG_LotColTal
    objPO.ConexionString = cCONNECT
        
    objPO.EliminaNoSeleccionadosWithKey vusu, sCod_Cliente, sCod_PurOrd, sLote, Me.txtCod_EstCliLOT.Text
       
    Set objPO = Nothing
    
Exit Function
errores:
    If Not objPO Is Nothing Then
        Set objPO = Nothing
    End If
    
    ErrorHandler Err, Err.Description
End Function

Private Sub txtPrecioLOT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If optComisionEnPorcentaje Then
             txtImp_Comision.Text = FixNulos(CDbl(txtPrecioLOT.Text), vbDouble) * (CDbl(txtPor_ComisionLOT.Text) / 100)
        End If
        txtCod_DivPreLOT.SetFocus
    End If
End Sub


Public Sub SUB_TOTALES()
    Dim ItemIngresar As String
    Dim varSubTotal As Double
    Dim NroColumnas As Integer
    Dim NroFilas As Integer
    NroColumnas = ssgrdDatosCantid.Cols
    If ssgrdDatosCantid.Rows < 1 Then
        Exit Sub
    End If
    ItemIngresar = ""
    'Aqui seleccionamos el primer registro para que nos sirva como base
    ssgrdDatosCantid.Bookmark = 0
    
    ItemIngresar = "Totales~"
    For NroColumnas = 1 To ssgrdDatosCantid.Cols - 1
        If Mid(ssgrdDatosCantid.Columns(NroColumnas).Name, 1, 3) = "QR1" Or Mid(ssgrdDatosCantid.Columns(NroColumnas).Name, 1, 3) = "TOT" Then
            'ssgrdDatosCantid.Columns("TOTAL").Text = ssgrdDatosCantid.Columns("TOTAL").value - varValorAntiguo + ssgrdDatosCantid.Columns(ColIndex).value
            'ssgrdDatosCantid.Columns("IMPORTE").Text = ssgrdDatosCantid.Columns("TOTAL").value * Me.txtPrecioLOT.Text
            varSubTotal = 0
            For NroFilas = 0 To ssgrdDatosCantid.Rows - 1
                ssgrdDatosCantid.Bookmark = NroFilas
                varSubTotal = varSubTotal + ssgrdDatosCantid.Columns(NroColumnas).value
            Next
            ItemIngresar = ItemIngresar & CStr(varSubTotal) & "~"
        Else
            ItemIngresar = ItemIngresar & "" & "~"
        End If
        
    Next
    ssgrdDatosCantid.AddItem ItemIngresar
    ssgrdDatosCantid.Bookmark = 0
End Sub

Public Sub ORDENA_LISTOX(lstprov As ListBox)
    'Dim lstprov As ListBox
    Dim Contador1 As Integer
    Dim Contador2 As Integer
    'Variables temporales
    Dim ItemTemp1 As String
    Dim ItemTemp2 As String
    Dim IndiceTemp As Integer
        
    'For Contador1 = 0 To lstBox.ListCount - 1
    '    lstprov.AddItem "lstBox.List(Contador1)", Contador1
    'Next
    'Hasta el paso anterior ya generamos una copia del listbox origen
    'Usaremos la burbuja para ordenar
    'Inicializamos los contadores
    Contador1 = 0
    Contador2 = 0
    
    For Contador1 = 0 To lstprov.ListCount - 1 - Contador2
        For Contador2 = 0 To lstprov.ListCount - 2
            If lstprov.List(Contador2) > lstprov.List(Contador2 + 1) Then
                ItemTemp1 = lstprov.List(Contador2)
                ItemTemp2 = lstprov.List(Contador2 + 1)
                IndiceTemp = Contador2 + 1
                'Efectuamos el cambio
                lstprov.RemoveItem (Contador2 + 1)
                lstprov.RemoveItem (Contador2)
                Call lstprov.AddItem(ItemTemp2, Contador2)
                Call lstprov.AddItem(ItemTemp1, Contador2 + 1)
            End If
        Next
    Next
End Sub

'Esta funcion fue creada por AHSP
Public Sub LoadMatrizPreciosGENERAGRILLA(ByVal varGridPrecios As ssDBGrid)
    Dim i As Integer
    Dim j As Integer
    'Dim k As Integer
    Dim varPrecio As String
    Dim varExiste As Boolean
    Dim varIndice As Integer
    
    If varGridPrecios.Rows > 0 Then
        'Aqui nos paseamos en la grilla de precios
        
        For i = 0 To varGridPrecios.Rows - 1
            varGridPrecios.Bookmark = i
            
            varExiste = False
            
            SSgrdDatosPrec.Bookmark = 0
            'Aqui nos paseamos por las columnas
            For j = 0 To SSgrdDatosPrec.Cols - 1
                If Trim(SSgrdDatosPrec.Columns(j).Name) = "PR1_" & varGridPrecios.Columns(0).value Then
                    'SSgrdDatosPrec.Columns(j).value = varGridPrecios.Columns(1).value
                    varIndice = j
                    varExiste = True
                    Exit For
                End If
            Next
            
            If varExiste = True Then
                'SSgrdDatosPrec.Bookmark = varIndice
                SSgrdDatosPrec.Bookmark = 0
                For j = 0 To SSgrdDatosPrec.Rows - 1
                    SSgrdDatosPrec.Bookmark = j
                    'If Mid(SSgrdDatosPrec.Columns(j).Name, 1, 4) = "PR1_" Then
                        SSgrdDatosPrec.Columns(varIndice).value = varGridPrecios.Columns(1).value
                    'End If
                Next
            End If
            
        Next
        
        SSgrdDatosPrec.Bookmark = 0

        Dim varTotal As Double
        Dim varSubTotal As Double
        'Aqui llenaremos la data de los precios correspondientes
        For i = 0 To SSgrdDatosPrec.Rows - 1
            SSgrdDatosPrec.Bookmark = i
            varTotal = 0
            For j = 0 To SSgrdDatosPrec.Cols - 1
                If Mid(Trim(SSgrdDatosPrec.Columns(j).Name), 1, 4) = "QR1_" Then
                    SSgrdDatosPrec.Columns(j + 6).value = CStr(CDbl(SSgrdDatosPrec.Columns(j).value) * CDbl(SSgrdDatosPrec.Columns(j + 4).value))
                    varTotal = varTotal + CDbl(SSgrdDatosPrec.Columns(j).value * SSgrdDatosPrec.Columns(j + 4).value)
                End If
            Next
            SSgrdDatosPrec.Columns("IMPORTE").value = CStr(varTotal)
        Next
        SSgrdDatosPrec.Bookmark = 0
    End If
End Sub

Public Sub COLOCA_NOMBRECOLOR(ssgrdDatos As ssDBGrid)
    Dim j As Integer
    Dim i As Integer
    
    For j = 0 To ssgrdDatos.Rows - 1
        ssgrdDatos.Bookmark = j
    
        For i = 0 To Me.lstColoresSELEC.ListCount - 1
            If ssgrdDatos.Columns(0).value = Mid(Me.lstColoresSELEC.List(i), 1, 20) Then
                ssgrdDatos.Columns(0).value = Me.lstColoresSELEC.List(i)
                Exit For
            End If
        Next
    Next
End Sub

Public Function InhabilitaModifCantidades() As Boolean
On Error GoTo errores
    Dim vbuff
    Dim obj As New clsTG_LotColTal
    Dim i As Long
            
    InhabilitaModifCantidades = False
    
    Set obj = New clsTG_LotColTal
    obj.ConexionString = cCONNECT
    vbuff = obj.InhabilitaModifCantidades(sCod_Cliente, sCod_PurOrd, sLote, Me.txtCod_EstCliLOT.Text)
                    
    If Not IsEmpty(vbuff) Then
        If vbuff(0, 0) = "1" Then
            InhabilitaModifCantidades = True
        End If
    End If

Exit Function
errores:
    If Not obj Is Nothing Then
        Set obj = Nothing
    End If
    
    ErrorHandler Err, Err.Description

End Function



Private Function VerificaGruposOperativos() As Boolean
On Error GoTo errx
Dim ssql As String
Dim mRs As ADODB.Recordset


ssql = "UP_VERIFICA_GRUPOS_OPERATIVOS '$' , '$' , '$' ,'$' "
ssql = VBsprintf(ssql, sCod_Cliente, sCod_PurOrd, sLote, Me.txtCod_EstCliLOT.Text)
Set mRs = GetDataSet(cCONNECT, ssql)

If Not mRs Then
    If mRs!CountOperativos = 0 Then
        VerificaGruposOperativos = True
    Else
        VerificaGruposOperativos = False
    End If
End If

Exit Function
errx:
    errores Err.Number
End Function



