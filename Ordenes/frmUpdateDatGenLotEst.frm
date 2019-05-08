VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUpdateDatGenLotEst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualización Datos Generales a Est.Cli."
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Update General Data"
   Begin VB.Frame Frame1 
      Caption         =   "DETALLE DEL PRECIO DE COSTO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E97834&
      Height          =   1155
      Left            =   60
      TabIndex        =   35
      Top             =   2970
      Width           =   5895
      Begin VB.TextBox txtCostoFijo 
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
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4770
         TabIndex        =   46
         Text            =   "0"
         Top             =   660
         Width           =   750
      End
      Begin VB.TextBox txtCostoFinanciero 
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
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4770
         TabIndex        =   44
         Text            =   "0"
         Top             =   300
         Width           =   750
      End
      Begin VB.TextBox txtCostoFOB 
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
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2550
         TabIndex        =   42
         Text            =   "0"
         Top             =   660
         Width           =   750
      End
      Begin VB.TextBox txtCostoMOD 
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
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2550
         TabIndex        =   40
         Text            =   "0"
         Top             =   300
         Width           =   750
      End
      Begin VB.TextBox txtCostoAvios 
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
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   870
         TabIndex        =   38
         Text            =   "0"
         Top             =   660
         Width           =   750
      End
      Begin VB.TextBox txtCostoTela 
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
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   870
         TabIndex        =   36
         Text            =   "0"
         Top             =   300
         Width           =   750
      End
      Begin VB.Label labels 
         AutoSize        =   -1  'True
         Caption         =   "FIJOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   13
         Left            =   3630
         TabIndex        =   47
         Top             =   720
         Width           =   420
      End
      Begin VB.Label labels 
         AutoSize        =   -1  'True
         Caption         =   "FINANCIEROS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   12
         Left            =   3630
         TabIndex        =   45
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label labels 
         AutoSize        =   -1  'True
         Caption         =   "FOB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   11
         Left            =   1950
         TabIndex        =   43
         Top             =   720
         Width           =   300
      End
      Begin VB.Label labels 
         AutoSize        =   -1  'True
         Caption         =   "M.O.D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   10
         Left            =   1950
         TabIndex        =   41
         Top             =   360
         Width           =   435
      End
      Begin VB.Label labels 
         AutoSize        =   -1  'True
         Caption         =   "AVIOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   9
         Left            =   270
         TabIndex        =   39
         Top             =   720
         Width           =   480
      End
      Begin VB.Label labels 
         AutoSize        =   -1  'True
         Caption         =   "TELA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   8
         Left            =   270
         TabIndex        =   37
         Top             =   360
         Width           =   345
      End
   End
   Begin VB.TextBox txtPrecioCostoLOT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   5190
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "0"
      Top             =   4110
      Width           =   765
   End
   Begin VB.OptionButton optPrePackNo 
      Caption         =   "NO"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E97834&
      Height          =   195
      Left            =   2520
      TabIndex        =   26
      Top             =   7080
      Width           =   615
   End
   Begin VB.OptionButton optPrePackSi 
      Caption         =   "SI"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E97834&
      Height          =   195
      Left            =   1830
      TabIndex        =   25
      Top             =   7080
      Width           =   615
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
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   22
      Text            =   "0"
      Top             =   1725
      Width           =   750
   End
   Begin VB.OptionButton optComisionEnPorcentaje 
      Caption         =   "EN PORCENTAJE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E97834&
      Height          =   180
      Left            =   1785
      TabIndex        =   19
      Top             =   1470
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton optComisionEnImporte 
      Caption         =   "EN IMPORTE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E97834&
      Height          =   180
      Left            =   3495
      TabIndex        =   18
      Top             =   1470
      Width           =   1575
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
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   17
      Text            =   "0"
      Top             =   2040
      Width           =   750
   End
   Begin VB.TextBox txtDes_General 
      Height          =   615
      Left            =   1770
      MaxLength       =   1000
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   4995
      Width           =   4125
   End
   Begin VB.TextBox txtCod_DivPreLOT 
      Height          =   300
      Left            =   1770
      MaxLength       =   3
      TabIndex        =   13
      Top             =   4665
      Width           =   615
   End
   Begin VB.TextBox txtPrecioLOT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E97834&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Text            =   "0"
      Top             =   2520
      Width           =   750
   End
   Begin VB.Frame fraNORegular 
      Appearance      =   0  'Flat
      Caption         =   "DATOS  P.O. NO REGULAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E97834&
      Height          =   1170
      Left            =   60
      TabIndex        =   6
      Top             =   5790
      Width           =   5880
      Begin VB.TextBox txtPrecio_RecCliLOT 
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
         Left            =   2430
         TabIndex        =   7
         Text            =   "0"
         Top             =   315
         Width           =   750
      End
      Begin MSComCtl2.DTPicker dtpFec_RecCliLOT 
         Height          =   315
         Left            =   2430
         TabIndex        =   8
         Top             =   630
         Width           =   2130
         _ExtentX        =   3757
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
         Format          =   16711681
         CurrentDate     =   37159
      End
      Begin VB.Label labels 
         AutoSize        =   -1  'True
         Caption         =   "PRECIO DEL CLIENTE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   21
         Left            =   120
         TabIndex        =   10
         Tag             =   "Client Price"
         Top             =   345
         Width           =   1590
      End
      Begin VB.Label labels 
         AutoSize        =   -1  'True
         Caption         =   "FECHA INGRESO A ALMACEN"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   22
         Left            =   135
         TabIndex        =   9
         Tag             =   "Reception Warehouse Date"
         Top             =   735
         Width           =   2175
      End
   End
   Begin VB.TextBox txtDes_DestinoLOT 
      Height          =   285
      Left            =   2430
      MaxLength       =   30
      TabIndex        =   3
      Top             =   60
      Width           =   3480
   End
   Begin VB.TextBox txtCod_DestinoLOT 
      Height          =   285
      Left            =   1770
      MaxLength       =   3
      TabIndex        =   0
      Top             =   60
      Width           =   615
   End
   Begin MSComCtl2.DTPicker dtpFec_DespachoOriLOT 
      Height          =   315
      Left            =   3570
      TabIndex        =   1
      Top             =   525
      Width           =   2340
      _ExtentX        =   4128
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
      Format          =   16711681
      CurrentDate     =   37159
   End
   Begin FunctionsButtons.FunctButt acbForm 
      Height          =   405
      Left            =   4080
      TabIndex        =   2
      Top             =   7500
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   661
      Custom          =   $"frmUpdateDatGenLotEst.frx":0000
      Orientacion     =   0
      Style           =   1
      Language        =   1
      TypeImageList   =   0
      ControlWidth    =   900
      ControlHeigth   =   350
      ControlSeparator=   10
   End
   Begin MSComCtl2.DTPicker dtpFecAceptacionDelCliente 
      Height          =   345
      Left            =   3540
      TabIndex        =   50
      Top             =   870
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   609
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
      CheckBox        =   -1  'True
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   16711683
      CurrentDate     =   40263.5095949074
   End
   Begin VB.Label labels 
      AutoSize        =   -1  'True
      Caption         =   "FECHA DE ACEPTACION DEL CLIENTE.............."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   23
      Left            =   90
      TabIndex        =   49
      Tag             =   "Delivery Date"
      Top             =   1020
      Width           =   3420
   End
   Begin VB.Label labels 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Index           =   16
      Left            =   60
      TabIndex        =   48
      Tag             =   "Delivery Date"
      Top             =   420
      Width           =   5865
   End
   Begin VB.Label labels 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Index           =   7
      Left            =   60
      TabIndex        =   34
      Tag             =   "Delivery Date"
      Top             =   2850
      Width           =   5865
   End
   Begin VB.Label labels 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Index           =   6
      Left            =   90
      TabIndex        =   33
      Tag             =   "Delivery Date"
      Top             =   7350
      Width           =   5865
   End
   Begin VB.Label labels 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Index           =   5
      Left            =   90
      TabIndex        =   32
      Tag             =   "Delivery Date"
      Top             =   5670
      Width           =   5865
   End
   Begin VB.Label labels 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Index           =   4
      Left            =   90
      TabIndex        =   31
      Tag             =   "Delivery Date"
      Top             =   4530
      Width           =   5865
   End
   Begin VB.Label labels 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Index           =   3
      Left            =   90
      TabIndex        =   30
      Tag             =   "Delivery Date"
      Top             =   2430
      Width           =   5865
   End
   Begin VB.Label labels 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Index           =   2
      Left            =   90
      TabIndex        =   29
      Tag             =   "Delivery Date"
      Top             =   1290
      Width           =   5865
   End
   Begin VB.Label labels 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL PRECIO DE COSTO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   3180
      TabIndex        =   28
      Tag             =   "Price"
      Top             =   4290
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "PRE PACK"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   90
      TabIndex        =   24
      Top             =   7110
      Width           =   765
   End
   Begin VB.Label labels 
      AutoSize        =   -1  'True
      Caption         =   "% COMISION"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   18
      Left            =   90
      TabIndex        =   23
      Tag             =   "Commision"
      Top             =   1800
      Width           =   1005
   End
   Begin VB.Label labels 
      AutoSize        =   -1  'True
      Caption         =   "MODO DE COMISION"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   20
      Left            =   90
      TabIndex        =   21
      Top             =   1455
      Width           =   1545
   End
   Begin VB.Label labels 
      AutoSize        =   -1  'True
      Caption         =   "IMPORTE COMISION"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   90
      TabIndex        =   20
      Top             =   2100
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "LETTER CREDIT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   90
      TabIndex        =   16
      Tag             =   "Letter Credit :"
      Top             =   5085
      Width           =   1155
   End
   Begin VB.Label labels 
      AutoSize        =   -1  'True
      Caption         =   "DIVISION DE PRENDA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   19
      Left            =   90
      TabIndex        =   14
      Tag             =   "Garment Division"
      Top             =   4740
      Width           =   1665
   End
   Begin VB.Label labels 
      AutoSize        =   -1  'True
      Caption         =   "PRECIO DE VTA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   14
      Left            =   90
      TabIndex        =   12
      Tag             =   "Price"
      Top             =   2580
      Width           =   1185
   End
   Begin VB.Label labels 
      AutoSize        =   -1  'True
      Caption         =   "DESTINO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   15
      Left            =   90
      TabIndex        =   5
      Tag             =   "Destination"
      Top             =   75
      Width           =   675
   End
   Begin VB.Label labels 
      AutoSize        =   -1  'True
      Caption         =   "FECHA DE DPCHO (EX-FACTORY)....................."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   17
      Left            =   90
      TabIndex        =   4
      Tag             =   "Delivery Date"
      Top             =   645
      Width           =   3420
   End
End
Attribute VB_Name = "frmUpdateDatGenLotEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public sCod_PurOrd As String
Public sCod_LotPurOrd As String
Public sCod_Cliente As String
Public sCod_EstCli As String
Public sFlag As String
Public sCod_Destino As String
Public sCod_DestinoLOT As String
Public sFlg_Regular As String
Dim Rs_cargaanexo As ADODB.Recordset

Public Sub CARGA_DATOSANEXOS()
    Dim strSql As String
    
    Set Rs_cargaanexo = New ADODB.Recordset
    Rs_cargaanexo.ActiveConnection = cCONNECT
    Rs_cargaanexo.CursorType = adOpenStatic
    Rs_cargaanexo.CursorLocation = adUseClient
    Rs_cargaanexo.LockType = adLockReadOnly
        
    'Esta cadena es la que nos devolvera los grupos de produccion
    strSql = "SELECT Cod_DivPre, Precio_RecCli, Fec_RecCli FROM TG_LOTEST WHERE " & _
    "Cod_Cliente ='" & sCod_Cliente & "' AND " & _
    "Cod_PurOrd ='" & sCod_PurOrd & "' AND " & _
    "Cod_LotPurOrd='" & sCod_LotPurOrd & "' AND " & _
    "Cod_EstCli='" & sCod_EstCli & "'"

    Rs_cargaanexo.Open strSql
    If Rs_cargaanexo.RecordCount > 0 Then
        If Not IsNull(Rs_cargaanexo("Cod_DivPre").value) Then
            txtCod_DivPreLOT.Text = Rs_cargaanexo("Cod_DivPre").value
        Else
            txtCod_DivPreLOT.Text = ""
        End If
        
        If Not IsNull(Rs_cargaanexo("Precio_RecCli").value) Then
            txtPrecio_RecCliLOT.Text = Rs_cargaanexo("Precio_RecCli").value
        Else
            txtPrecio_RecCliLOT.Text = ""
        End If

        If Not IsNull(Rs_cargaanexo("Fec_RecCli")) Then
            dtpFec_RecCliLOT.value = Rs_cargaanexo("Fec_RecCli").value
        Else
            dtpFec_RecCliLOT.value = Date
        End If
        
    End If
    
    Rs_cargaanexo.Close
    Set Rs_cargaanexo = Nothing
    
    
End Sub

Private Sub ObtenerTotalCosto()
    Dim dblTotalCosto As Double
        
    dblTotalCosto = Val(txtCostoTela.Text) + Val(txtCostoAvios.Text) + Val(txtCostoMOD.Text) + Val(txtCostoFOB.Text) + Val(txtCostoFinanciero.Text) + Val(txtCostoFijo.Text)
    txtPrecioCostoLOT.Text = Format(dblTotalCosto, "###,###.00")
End Sub

Private Sub txtCostoTela_Change()
    ObtenerTotalCosto
End Sub
Private Sub txtCostoAvios_Change()
    ObtenerTotalCosto
End Sub
Private Sub txtCostoMOD_Change()
    ObtenerTotalCosto
End Sub
Private Sub txtCostoFOB_Change()
    ObtenerTotalCosto
End Sub
Private Sub txtCostoFinanciero_Change()
    ObtenerTotalCosto
End Sub
Private Sub txtCostoFijo_Change()
    ObtenerTotalCosto
End Sub

Private Sub acbForm_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            If Not ValidStep Then
                Exit Sub
            End If
            ObtenerTotalCosto
            UpdateDatGen
        Case "CANCELAR"
            Unload Me
    End Select
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Call FormSet(Me)
End Sub

Private Function UpdateDatGen() As Boolean
On Error GoTo errores
    Dim vbuff
    Dim objPO As clsTG_LotColTal
    Dim sFlg_NoRegular
    Dim sComisionEnPorcentaje As String
    
    Set objPO = New clsTG_LotColTal
    objPO.ConexionString = cCONNECT
    
    If optComisionEnPorcentaje Then
        sComisionEnPorcentaje = "S"
    Else
        sComisionEnPorcentaje = "N"
    End If
        
    'objPO.UpdateDatGenPurORd sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, txtCod_DestinoLOT.Text, FechaOK(dtpFec_DespachoActLOT.value), CDbl(txtPor_ComisionLOT.Text), vusu, ComputerName, FechaOK(dtpFec_DespachoOriLOT.value), FixNulos(txtPrecioLOT.Text, vbDouble), sFlg_Regular, FixNulos(txtPrecio_RecCliLOT.Text, vbDouble), FechaOK(dtpFec_RecCliLOT.value), txtCod_DivPreLOT.Text, txtDes_General.Text, sComisionEnPorcentaje, CDbl(txtImp_Comision.Text)
    objPO.UpdateDatGenPurORd sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, txtCod_DestinoLOT.Text, CDbl(txtPor_ComisionLOT.Text), vusu, ComputerName, _
                             FechaOK(dtpFec_DespachoOriLOT.value), FixNulos(txtPrecioLOT.Text, vbDouble), sFlg_Regular, FixNulos(txtPrecio_RecCliLOT.Text, vbDouble), _
                             FechaOK(dtpFec_RecCliLOT.value), txtCod_DivPreLOT.Text, txtDes_General.Text, sComisionEnPorcentaje, _
                             CDbl(Val(txtImp_Comision.Text)), _
                             CDbl(Val(txtPrecioCostoLOT.Text)), _
                             CDbl(Val(txtCostoTela.Text)), _
                             CDbl(Val(txtCostoAvios.Text)), _
                             CDbl(Val(txtCostoMOD.Text)), _
                             CDbl(Val(txtCostoFOB.Text)), _
                             CDbl(Val(txtCostoFinanciero.Text)), _
                             CDbl(Val(txtCostoFijo.Text)), _
                             dtpFecAceptacionDelCliente.value
'    oParent.Buscar
'    oParent.BuscarEStilos
    
    Set objPO = Nothing
    Unload Me
Exit Function
errores:
    If Not objPO Is Nothing Then
        Set objPO = Nothing
    End If
    
    ErrorHandler Err, Err.Description

End Function



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

Private Sub txtCod_DestinoLOT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sFlag = "COD_DESTINOLOT"
        If Filtrar(sFlag, Me, txtCod_DestinoLOT, txtDes_DestinoLOT) Then
            Me.dtpFec_DespachoOriLOT.SetFocus
        End If
    End If
End Sub
Public Function ValidStep() As Boolean
Dim aMess(4)
Dim amensaje As clsMessages
Set amensaje = New clsMessages
  
    If txtCod_DestinoLOT.Text = "" Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
        If txtCod_DestinoLOT.Enabled Then
            Me.txtCod_DestinoLOT.SetFocus
        End If
        Exit Function
    End If

    If txtCod_DestinoLOT.Text <> "" Then
        If Not ValidCod_DestinoLot() Then
            Exit Function
        End If
    End If

    If dtpFec_DespachoOriLOT.value = "" Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
        If dtpFec_DespachoOriLOT.Enabled Then
            Me.dtpFec_DespachoOriLOT.SetFocus
        End If
        Exit Function
    End If
    
    If Not VAlidFechaDespacho(FechaOK(dtpFec_DespachoOriLOT.value)) Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_INVALID_SELECC
        If dtpFec_DespachoOriLOT.Enabled Then
            dtpFec_DespachoOriLOT.SetFocus
        End If
        Exit Function
    End If
    
'    If Not VAlidFechaDespacho(FechaOK(dtpFec_DespachoOriLOT.Value)) Then
'        Mensaje MESSAGECODE.kMESSAGE_ERR_INVALID_SELECC
'        If dtpFec_DespachoOriLOT.Enabled Then
'            dtpFec_DespachoOriLOT.SetFocus
'        End If
'        Exit Function
'    End If
    
'    If optComisionEnPorcentaje And CDbl(txtPor_ComisionLOT.Text) <= 0 Then
'        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
'        If txtPor_ComisionLOT.Enabled Then
'            txtPor_ComisionLOT.SetFocus
'            Exit Function
'        End If
'    End If
'
'    If optComisionEnImporte And CDbl(txtImp_Comision.Text) <= 0 Then
'        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
'        If txtImp_Comision.Enabled Then
'            txtImp_Comision.SetFocus
'            Exit Function
'        End If
'    End If
'
    
    If FixNulos(txtPrecioLOT.Text, vbDouble) = 0 Then
        Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
        If txtPrecioLOT.Enabled Then
            Me.txtPrecioLOT.SetFocus
        End If
        Exit Function
    End If
    
    If RTrim(txtCod_DivPreLOT.Text) <> "" Then
        If Not VAlidDivPre(Me.txtCod_DivPreLOT.Text) Then
            If txtCod_DivPreLOT.Enabled Then
                txtCod_DivPreLOT.SetFocus
            End If
            Exit Function
        End If
    End If
    
    If sFlg_Regular = "N" Then
        If FixNulos(txtPrecio_RecCliLOT.Text, vbDouble) = 0 Then
            Mensaje MESSAGECODE.kMESSAGE_ERR_NOTEMPTY
            If txtPrecio_RecCliLOT.Enabled Then
                Me.txtPrecio_RecCliLOT.SetFocus
            End If
            Exit Function
        End If
    End If
    
    ValidStep = True
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

Private Sub txtCod_DivPreLOT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sFlag = "COD_DIVPRE"
            If Filtrar(sFlag, Me, txtCod_DivPreLOT, Nothing, True) Then
                acbForm.SetFocus
            Else
                If Not VAlidDivPre(Me.txtCod_DivPreLOT.Text) Then
                    Exit Sub
                Else
                    acbForm.SetFocus
                End If
            End If

    End If
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
        Set frmDivPre = Nothing
        VAlidDivPre = True
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


Private Sub optComisionEnImporte_Click()
    txtPor_ComisionLOT.Enabled = False
    txtImp_Comision.Enabled = True
    txtPor_ComisionLOT.Text = 0
    If Me.txtImp_Comision.Enabled Then
        If Me.Visible Then
            txtImp_Comision.SetFocus
        End If
    End If
End Sub

Private Sub optComisionEnPorcentaje_Click()
    txtPor_ComisionLOT.Enabled = True
    
    txtImp_Comision.Text = 0
    txtImp_Comision.Enabled = False
    If Me.txtPor_ComisionLOT.Enabled Then
        txtPor_ComisionLOT.SetFocus

    End If
End Sub



Private Sub txtPor_ComisionLOT_GotFocus()
    SelectionText txtPor_ComisionLOT
End Sub

Private Sub txtPor_ComisionLOT_KeyPress(keyascii As Integer)
If keyascii = vbKeyReturn Then
    If optComisionEnPorcentaje Then
         txtImp_Comision.Text = 0
    End If
End If
End Sub

Private Sub txtImp_Comision_GotFocus()
    SelectionText txtImp_Comision
End Sub

Private Sub txtImp_Comision_KeyPress(keyascii As Integer)
    If keyascii = vbKeyReturn And optComisionEnImporte.value Then
         txtImp_Comision.Text = FixNulos(CDbl(txtImp_Comision.Text), vbDouble)
         txtPor_ComisionLOT.Text = 0
    End If
End Sub


Private Sub txtPrecioLOT_KeyPress(keyascii As Integer)
    If keyascii = vbKeyReturn Then
        txtCod_DivPreLOT.SetFocus
    End If
End Sub
