VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form frmShowTG_PurOrd 
   Caption         =   "Purchase Order"
   ClientHeight    =   8550
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "frmShowTG_PurOrd4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Elija"
      Height          =   1035
      Left            =   45
      TabIndex        =   5
      Tag             =   "Selection"
      Top             =   -15
      Width           =   15285
      Begin VB.OptionButton optCod_EstCli 
         Caption         =   "Estilo del Cliente"
         Height          =   195
         Left            =   7350
         TabIndex        =   10
         Tag             =   "Client Style"
         Top             =   195
         Width           =   1470
      End
      Begin VB.OptionButton optCod_TemCli 
         Caption         =   "Temporada"
         Height          =   195
         Left            =   4500
         TabIndex        =   9
         Tag             =   "Season"
         Top             =   195
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton optCod_PurOrd 
         Caption         =   "Purchase Order"
         Height          =   195
         Left            =   5775
         TabIndex        =   8
         Tag             =   "Purchase Order"
         Top             =   195
         Width           =   1470
      End
      Begin VB.OptionButton optCod_OrdPro 
         Caption         =   "O/P"
         Height          =   195
         Left            =   9030
         TabIndex        =   7
         Tag             =   "O/P"
         Top             =   210
         Width           =   735
      End
      Begin VB.TextBox txtAbr_Cliente 
         Height          =   285
         Left            =   810
         TabIndex        =   0
         Top             =   375
         Width           =   690
      End
      Begin VB.TextBox txtNom_Cliente 
         Height          =   285
         Left            =   1545
         TabIndex        =   1
         Top             =   375
         Width           =   2400
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   420
         Left            =   10260
         TabIndex        =   2
         Tag             =   "Find"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Frame fraTemporada 
         Height          =   525
         Left            =   4200
         TabIndex        =   11
         Top             =   375
         Width           =   6015
         Begin VB.TextBox txtNom_TemCli 
            Height          =   285
            Left            =   2040
            TabIndex        =   13
            Top             =   165
            Width           =   3900
         End
         Begin VB.TextBox txtCod_TemCli 
            Height          =   285
            Left            =   1380
            MaxLength       =   3
            TabIndex        =   12
            Top             =   165
            Width           =   600
         End
         Begin VB.Label Label1 
            Caption         =   "Temporada"
            Height          =   180
            Left            =   150
            TabIndex        =   14
            Tag             =   "Season"
            Top             =   225
            Width           =   855
         End
      End
      Begin VB.Frame fraOP 
         Height          =   525
         Left            =   4200
         TabIndex        =   21
         Top             =   375
         Width           =   6015
         Begin VB.TextBox txtDes_estpro 
            Height          =   285
            Left            =   2220
            TabIndex        =   23
            Top             =   180
            Width           =   3765
         End
         Begin VB.TextBox txtCod_Ordpro 
            Height          =   285
            Left            =   1350
            MaxLength       =   5
            TabIndex        =   22
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "O/P"
            Height          =   195
            Left            =   150
            TabIndex        =   24
            Top             =   225
            Width           =   300
         End
      End
      Begin VB.Frame fraPurOrd 
         Height          =   525
         Left            =   4200
         TabIndex        =   18
         Top             =   375
         Width           =   6015
         Begin VB.TextBox txtCod_PurOrd 
            Height          =   285
            Left            =   1380
            TabIndex        =   19
            Top             =   165
            Width           =   4530
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Purchase Order"
            Height          =   195
            Left            =   135
            TabIndex        =   20
            Top             =   225
            Width           =   1110
         End
      End
      Begin VB.Frame fraEstCli 
         Height          =   525
         Left            =   4200
         TabIndex        =   15
         Top             =   375
         Width           =   6015
         Begin VB.TextBox txtCod_EstCli 
            Height          =   285
            Left            =   1395
            TabIndex        =   16
            Top             =   165
            Width           =   4530
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Estilo Cliente"
            Height          =   195
            Left            =   210
            TabIndex        =   17
            Top             =   210
            Width           =   900
         End
      End
      Begin VB.Label lblCod_Cliente 
         Caption         =   "Cliente"
         Height          =   285
         Left            =   135
         TabIndex        =   6
         Tag             =   "Client"
         Top             =   390
         Width           =   765
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5400
      Left            =   45
      TabIndex        =   25
      Top             =   960
      Width           =   13815
      Begin SSDataWidgets_B.SSDBGrid ssgrdDatos 
         Height          =   2520
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   13665
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         HeadLines       =   2
         Col.Count       =   46
         DividerType     =   1
         ForeColorEven   =   0
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   46
         Columns(0).Width=   2752
         Columns(0).Caption=   "Purchase Order"
         Columns(0).Name =   "Cod_PurOrd"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "Fecha Despacho Actual"
         Columns(1).Name =   "Fec_DespachoAct"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   2011
         Columns(2).Caption=   "Total Prendas"
         Columns(2).Name =   "Num_PreReq"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   1826
         Columns(3).Caption=   "Importe Prendas"
         Columns(3).Name =   "Imp_TotalPre"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3200
         Columns(4).Visible=   0   'False
         Columns(4).Caption=   "Destino"
         Columns(4).Name =   "Cod_Destino"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Visible=   0   'False
         Columns(5).Caption=   "Des_Destino"
         Columns(5).Name =   "Des_Destino"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   1588
         Columns(6).Caption=   "Grupo Pro"
         Columns(6).Name =   "Cod_GrupoPro"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   2143
         Columns(7).Caption=   "Desc. Grupo"
         Columns(7).Name =   "Des_GrupoPro"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(8).Width=   1693
         Columns(8).Caption=   "Temporada Cliente"
         Columns(8).Name =   "Cod_TemCli"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   256
         Columns(9).Width=   3122
         Columns(9).Caption=   "Nombre Temporada Cliente"
         Columns(9).Name =   "Nom_TemCli"
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   8
         Columns(9).FieldLen=   256
         Columns(10).Width=   1879
         Columns(10).Caption=   "Codigo de Embarque"
         Columns(10).Name=   "Cod_Embarque"
         Columns(10).DataField=   "Column 10"
         Columns(10).DataType=   8
         Columns(10).FieldLen=   256
         Columns(11).Width=   2619
         Columns(11).Caption=   "Descripción de Embarque"
         Columns(11).Name=   "Des_Embarque"
         Columns(11).DataField=   "Column 11"
         Columns(11).DataType=   8
         Columns(11).FieldLen=   256
         Columns(12).Width=   1852
         Columns(12).Caption=   "Codigo Pago Embarque"
         Columns(12).Name=   "Cod_PagEmb"
         Columns(12).DataField=   "Column 12"
         Columns(12).DataType=   8
         Columns(12).FieldLen=   256
         Columns(13).Width=   3200
         Columns(13).Caption=   "Descripción Pago Embarque"
         Columns(13).Name=   "Des_PagEmb"
         Columns(13).DataField=   "Column 13"
         Columns(13).DataType=   8
         Columns(13).FieldLen=   256
         Columns(14).Width=   1085
         Columns(14).Caption=   "Class"
         Columns(14).Name=   "Cod_ClaPurOrd"
         Columns(14).DataField=   "Column 14"
         Columns(14).DataType=   8
         Columns(14).FieldLen=   256
         Columns(15).Width=   1588
         Columns(15).Caption=   "% de Comisión"
         Columns(15).Name=   "Por_Comision"
         Columns(15).DataField=   "Column 15"
         Columns(15).DataType=   8
         Columns(15).NumberFormat=   "##0.00"
         Columns(15).FieldLen=   256
         Columns(16).Width=   1535
         Columns(16).Caption=   "Order/ Reorder"
         Columns(16).Name=   "Cod_TipPurOrd"
         Columns(16).DataField=   "Column 16"
         Columns(16).DataType=   8
         Columns(16).FieldLen=   256
         Columns(17).Width=   3200
         Columns(17).Visible=   0   'False
         Columns(17).Caption=   "Description"
         Columns(17).Name=   "Des_ClaPurOrd"
         Columns(17).DataField=   "Column 17"
         Columns(17).DataType=   8
         Columns(17).FieldLen=   256
         Columns(18).Width=   3200
         Columns(18).Visible=   0   'False
         Columns(18).Caption=   "Proforma"
         Columns(18).Name=   "Cod_Proforma"
         Columns(18).DataField=   "Column 18"
         Columns(18).DataType=   8
         Columns(18).FieldLen=   256
         Columns(19).Width=   3200
         Columns(19).Visible=   0   'False
         Columns(19).Caption=   "Nivel P.O."
         Columns(19).Name=   "Num_NivPurOrd"
         Columns(19).DataField=   "Column 19"
         Columns(19).DataType=   8
         Columns(19).FieldLen=   256
         Columns(20).Width=   3200
         Columns(20).Visible=   0   'False
         Columns(20).Caption=   "Fecha Despacho Original"
         Columns(20).Name=   "Fec_DespachoOri"
         Columns(20).DataField=   "Column 20"
         Columns(20).DataType=   8
         Columns(20).FieldLen=   256
         Columns(21).Width=   3200
         Columns(21).Visible=   0   'False
         Columns(21).Caption=   "Fabrica"
         Columns(21).Name=   "Cod_Fabrica"
         Columns(21).DataField=   "Column 21"
         Columns(21).DataType=   8
         Columns(21).FieldLen=   256
         Columns(22).Width=   3200
         Columns(22).Visible=   0   'False
         Columns(22).Caption=   "Nombre de Fabrica"
         Columns(22).Name=   "Nom_Fabrica"
         Columns(22).DataField=   "Column 22"
         Columns(22).DataType=   8
         Columns(22).FieldLen=   256
         Columns(23).Width=   3200
         Columns(23).Visible=   0   'False
         Columns(23).Caption=   "Abreviatura Fabrica"
         Columns(23).Name=   "Abr_Fabrica"
         Columns(23).DataField=   "Column 23"
         Columns(23).DataType=   8
         Columns(23).FieldLen=   256
         Columns(24).Width=   3200
         Columns(24).Visible=   0   'False
         Columns(24).Caption=   "Fecha Despacho Real"
         Columns(24).Name=   "Fec_DespachoReal"
         Columns(24).DataField=   "Column 24"
         Columns(24).DataType=   8
         Columns(24).FieldLen=   256
         Columns(25).Width=   3200
         Columns(25).Visible=   0   'False
         Columns(25).Caption=   "Moneda"
         Columns(25).Name=   "Cod_Moneda"
         Columns(25).DataField=   "Column 25"
         Columns(25).DataType=   8
         Columns(25).FieldLen=   256
         Columns(26).Width=   3200
         Columns(26).Visible=   0   'False
         Columns(26).Caption=   "Nombre de Moneda"
         Columns(26).Name=   "Nom_Moneda"
         Columns(26).DataField=   "Column 26"
         Columns(26).DataType=   8
         Columns(26).FieldLen=   256
         Columns(27).Width=   3200
         Columns(27).Visible=   0   'False
         Columns(27).Caption=   "División de Cliente"
         Columns(27).Name=   "Cod_DivCli"
         Columns(27).DataField=   "Column 27"
         Columns(27).DataType=   8
         Columns(27).FieldLen=   256
         Columns(28).Width=   3200
         Columns(28).Visible=   0   'False
         Columns(28).Caption=   "Nombre de Division de Cliente"
         Columns(28).Name=   "Nom_DivCli"
         Columns(28).DataField=   "Column 28"
         Columns(28).DataType=   8
         Columns(28).FieldLen=   256
         Columns(29).Width=   3200
         Columns(29).Visible=   0   'False
         Columns(29).Caption=   "P.O. Original"
         Columns(29).Name=   "Cod_PurOrdOri"
         Columns(29).DataField=   "Column 29"
         Columns(29).DataType=   8
         Columns(29).FieldLen=   256
         Columns(30).Width=   3200
         Columns(30).Visible=   0   'False
         Columns(30).Caption=   "Flag Carta"
         Columns(30).Name=   "Flg_Carta"
         Columns(30).DataField=   "Column 30"
         Columns(30).DataType=   8
         Columns(30).FieldLen=   256
         Columns(31).Width=   3200
         Columns(31).Visible=   0   'False
         Columns(31).Caption=   "Banco"
         Columns(31).Name=   "Cod_Banco"
         Columns(31).DataField=   "Column 31"
         Columns(31).DataType=   8
         Columns(31).FieldLen=   256
         Columns(32).Width=   3200
         Columns(32).Visible=   0   'False
         Columns(32).Caption=   "Nombre de Banco"
         Columns(32).Name=   "Nom_Banco"
         Columns(32).DataField=   "Column 32"
         Columns(32).DataType=   8
         Columns(32).FieldLen=   256
         Columns(33).Width=   3200
         Columns(33).Visible=   0   'False
         Columns(33).Caption=   "% Slush"
         Columns(33).Name=   "Por_Slush"
         Columns(33).DataField=   "Column 33"
         Columns(33).DataType=   8
         Columns(33).NumberFormat=   "##0.00"
         Columns(33).FieldLen=   256
         Columns(34).Width=   3200
         Columns(34).Visible=   0   'False
         Columns(34).Caption=   "Descripción General"
         Columns(34).Name=   "Des_General"
         Columns(34).DataField=   "Column 34"
         Columns(34).DataType=   8
         Columns(34).FieldLen=   256
         Columns(35).Width=   3200
         Columns(35).Visible=   0   'False
         Columns(35).Caption=   "Descripción del Despacho"
         Columns(35).Name=   "Des_Despacho"
         Columns(35).DataField=   "Column 35"
         Columns(35).DataType=   8
         Columns(35).FieldLen=   256
         Columns(36).Width=   3200
         Columns(36).Visible=   0   'False
         Columns(36).Caption=   "Regular/ No Regular"
         Columns(36).Name=   "Flg_Regular"
         Columns(36).DataField=   "Column 36"
         Columns(36).DataType=   8
         Columns(36).FieldLen=   256
         Columns(37).Width=   3200
         Columns(37).Visible=   0   'False
         Columns(37).Caption=   "Access Level"
         Columns(37).Name=   "NivAcc"
         Columns(37).DataField=   "Column 37"
         Columns(37).DataType=   8
         Columns(37).FieldLen=   256
         Columns(38).Width=   3200
         Columns(38).Visible=   0   'False
         Columns(38).Caption=   "Nro Lotes"
         Columns(38).Name=   "LotPurOrd"
         Columns(38).DataField=   "Column 38"
         Columns(38).DataType=   8
         Columns(38).FieldLen=   256
         Columns(39).Width=   3200
         Columns(39).Visible=   0   'False
         Columns(39).Caption=   "% Adic. Prod."
         Columns(39).Name=   "Por_AdicProd"
         Columns(39).DataField=   "Column 39"
         Columns(39).DataType=   8
         Columns(39).FieldLen=   256
         Columns(40).Width=   3200
         Columns(40).Visible=   0   'False
         Columns(40).Caption=   "Pre. Adic.Prod."
         Columns(40).Name=   "Pre_AdicProd"
         Columns(40).DataField=   "Column 40"
         Columns(40).DataType=   8
         Columns(40).FieldLen=   256
         Columns(41).Width=   3200
         Columns(41).Visible=   0   'False
         Columns(41).Caption=   "Num. Pre. Cri"
         Columns(41).Name=   "Num_PreCri"
         Columns(41).DataField=   "Column 41"
         Columns(41).DataType=   8
         Columns(41).FieldLen=   256
         Columns(42).Width=   3200
         Columns(42).Caption=   "FLG_IFINANCIERA_PO"
         Columns(42).Name=   "FLG_IFINANCIERA_PO"
         Columns(42).DataField=   "Column 42"
         Columns(42).DataType=   8
         Columns(42).FieldLen=   256
         Columns(43).Width=   3200
         Columns(43).Caption=   "FLG_IFINANCIERA_COMMIT"
         Columns(43).Name=   "FLG_IFINANCIERA_COMMIT"
         Columns(43).DataField=   "Column 43"
         Columns(43).DataType=   8
         Columns(43).FieldLen=   256
         Columns(44).Width=   3200
         Columns(44).Caption=   "Cod_Cliente"
         Columns(44).Name=   "Cod_Cliente"
         Columns(44).DataField=   "Column 44"
         Columns(44).DataType=   8
         Columns(44).FieldLen=   256
         Columns(45).Width=   3200
         Columns(45).Caption=   "Fecha Emision"
         Columns(45).Name=   "Fec_Emision"
         Columns(45).CaptionAlignment=   0
         Columns(45).DataField=   "Column 45"
         Columns(45).DataType=   8
         Columns(45).FieldLen=   256
         _ExtentX        =   24104
         _ExtentY        =   4445
         _StockProps     =   79
         Caption         =   "Purchase Orders"
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
      Begin SSDataWidgets_B.SSDBGrid ssgrdDatos2 
         Height          =   2640
         Left            =   120
         TabIndex        =   27
         Top             =   2745
         Width           =   13575
         _Version        =   196617
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Col.Count       =   33
         UseGroups       =   -1  'True
         DividerType     =   1
         ForeColorEven   =   0
         BackColorOdd    =   16777215
         RowHeight       =   423
         Groups.Count    =   10
         Groups(0).Width =   10134
         Groups(0).Caption=   "Estilos"
         Groups(0).Columns.Count=   7
         Groups(0).Columns(0).Width=   1032
         Groups(0).Columns(0).Caption=   "Lote"
         Groups(0).Columns(0).Name=   "Cod_LotPurOrd"
         Groups(0).Columns(0).DataField=   "Column 0"
         Groups(0).Columns(0).DataType=   8
         Groups(0).Columns(0).FieldLen=   256
         Groups(0).Columns(1).Width=   1455
         Groups(0).Columns(1).Caption=   "Numero"
         Groups(0).Columns(1).Name=   "Cod_EstCli"
         Groups(0).Columns(1).DataField=   "Column 1"
         Groups(0).Columns(1).DataType=   8
         Groups(0).Columns(1).FieldLen=   256
         Groups(0).Columns(2).Width=   1164
         Groups(0).Columns(2).Caption=   "E. Propios"
         Groups(0).Columns(2).Name=   "EstPropio"
         Groups(0).Columns(2).DataField=   "Column 2"
         Groups(0).Columns(2).DataType=   8
         Groups(0).Columns(2).FieldLen=   256
         Groups(0).Columns(3).Width=   1508
         Groups(0).Columns(3).Caption=   "O/PS"
         Groups(0).Columns(3).Name=   "OrdPro"
         Groups(0).Columns(3).DataField=   "Column 3"
         Groups(0).Columns(3).DataType=   8
         Groups(0).Columns(3).FieldLen=   256
         Groups(0).Columns(4).Width=   2223
         Groups(0).Columns(4).Caption=   "Descripcion"
         Groups(0).Columns(4).Name=   "Des_EstCli"
         Groups(0).Columns(4).DataField=   "Column 4"
         Groups(0).Columns(4).DataType=   8
         Groups(0).Columns(4).FieldLen=   256
         Groups(0).Columns(5).Width=   847
         Groups(0).Columns(5).Caption=   "Destino"
         Groups(0).Columns(5).Name=   "Cod_Destino"
         Groups(0).Columns(5).DataField=   "Column 5"
         Groups(0).Columns(5).DataType=   8
         Groups(0).Columns(5).FieldLen=   256
         Groups(0).Columns(6).Width=   1905
         Groups(0).Columns(6).Caption=   "Descrip Destino"
         Groups(0).Columns(6).Name=   "Des_Destino"
         Groups(0).Columns(6).DataField=   "Column 6"
         Groups(0).Columns(6).DataType=   8
         Groups(0).Columns(6).FieldLen=   256
         Groups(1).Width =   6826
         Groups(1).Caption=   "Despacho"
         Groups(1).Columns.Count=   4
         Groups(1).Columns(0).Width=   2064
         Groups(1).Columns(0).Caption=   "Ex-Factory"
         Groups(1).Columns(0).Name=   "Fec_DespachoOri"
         Groups(1).Columns(0).DataField=   "Column 7"
         Groups(1).Columns(0).DataType=   8
         Groups(1).Columns(0).FieldLen=   256
         Groups(1).Columns(0).Mask=   "##/##"
         Groups(1).Columns(1).Width=   1905
         Groups(1).Columns(1).Caption=   "Fin Produc"
         Groups(1).Columns(1).Name=   "Fec_DespachoAct"
         Groups(1).Columns(1).DataField=   "Column 8"
         Groups(1).Columns(1).DataType=   8
         Groups(1).Columns(1).FieldLen=   256
         Groups(1).Columns(1).Mask=   "##/##"
         Groups(1).Columns(2).Width=   1931
         Groups(1).Columns(2).Caption=   "Real"
         Groups(1).Columns(2).Name=   "Fec_DespachoReal"
         Groups(1).Columns(2).DataField=   "Column 9"
         Groups(1).Columns(2).DataType=   8
         Groups(1).Columns(2).FieldLen=   256
         Groups(1).Columns(2).Mask=   "##/##"
         Groups(1).Columns(3).Width=   926
         Groups(1).Columns(3).Caption=   "R/V"
         Groups(1).Columns(3).Name=   "Flg_CuadraDetalle"
         Groups(1).Columns(3).DataField=   "Column 10"
         Groups(1).Columns(3).DataType=   8
         Groups(1).Columns(3).FieldLen=   256
         Groups(2).Width =   2487
         Groups(2).Caption=   "Precio"
         Groups(2).Columns.Count=   2
         Groups(2).Columns(0).Width=   1402
         Groups(2).Columns(0).Caption=   "Precio"
         Groups(2).Columns(0).Name=   "Precio"
         Groups(2).Columns(0).Alignment=   1
         Groups(2).Columns(0).CaptionAlignment=   0
         Groups(2).Columns(0).DataField=   "Column 11"
         Groups(2).Columns(0).DataType=   8
         Groups(2).Columns(0).FieldLen=   256
         Groups(2).Columns(1).Width=   1085
         Groups(2).Columns(1).Caption=   "Flag  "
         Groups(2).Columns(1).Name=   "Flg_PreDif"
         Groups(2).Columns(1).Alignment=   1
         Groups(2).Columns(1).CaptionAlignment=   0
         Groups(2).Columns(1).DataField=   "Column 12"
         Groups(2).Columns(1).DataType=   8
         Groups(2).Columns(1).FieldLen=   256
         Groups(3).Width =   2328
         Groups(3).Caption=   "Prendas"
         Groups(3).Columns.Count=   2
         Groups(3).Columns(0).Width=   1111
         Groups(3).Columns(0).Caption=   "Request"
         Groups(3).Columns(0).Name=   "Num_PreReq"
         Groups(3).Columns(0).Alignment=   1
         Groups(3).Columns(0).CaptionAlignment=   0
         Groups(3).Columns(0).DataField=   "Column 13"
         Groups(3).Columns(0).DataType=   8
         Groups(3).Columns(0).FieldLen=   256
         Groups(3).Columns(1).Width=   1217
         Groups(3).Columns(1).Caption=   "Desp"
         Groups(3).Columns(1).Name=   "Num_PreDes"
         Groups(3).Columns(1).Alignment=   1
         Groups(3).Columns(1).CaptionAlignment=   0
         Groups(3).Columns(1).DataField=   "Column 14"
         Groups(3).Columns(1).DataType=   8
         Groups(3).Columns(1).FieldLen=   256
         Groups(4).Width =   3572
         Groups(4).Caption=   "Importes"
         Groups(4).Columns.Count=   2
         Groups(4).Columns(0).Width=   1826
         Groups(4).Columns(0).Caption=   "Request"
         Groups(4).Columns(0).Name=   "Imp_TotalPRe"
         Groups(4).Columns(0).Alignment=   1
         Groups(4).Columns(0).CaptionAlignment=   0
         Groups(4).Columns(0).DataField=   "Column 15"
         Groups(4).Columns(0).DataType=   8
         Groups(4).Columns(0).FieldLen=   256
         Groups(4).Columns(1).Width=   1746
         Groups(4).Columns(1).Caption=   "Desp"
         Groups(4).Columns(1).Name=   "Imp_TotalDes"
         Groups(4).Columns(1).Alignment=   1
         Groups(4).Columns(1).CaptionAlignment=   0
         Groups(4).Columns(1).DataField=   "Column 16"
         Groups(4).Columns(1).DataType=   8
         Groups(4).Columns(1).FieldLen=   256
         Groups(5).Width =   3200
         Groups(5).Caption=   "Nro Factura"
         Groups(5).Columns.Count=   2
         Groups(5).Columns(0).Width=   1588
         Groups(5).Columns(0).Caption=   "Serie"
         Groups(5).Columns(0).Name=   "Cod_SerFac"
         Groups(5).Columns(0).DataField=   "Column 17"
         Groups(5).Columns(0).DataType=   8
         Groups(5).Columns(0).FieldLen=   256
         Groups(5).Columns(1).Width=   1614
         Groups(5).Columns(1).Caption=   "Number"
         Groups(5).Columns(1).Name=   "Cod_Factura"
         Groups(5).Columns(1).DataField=   "Column 18"
         Groups(5).Columns(1).DataType=   8
         Groups(5).Columns(1).FieldLen=   256
         Groups(6).Width =   1508
         Groups(6).Caption=   "TipoLote"
         Groups(6).Columns(0).Width=   1508
         Groups(6).Columns(0).Caption=   "Lote Type"
         Groups(6).Columns(0).Name=   "Tip_LotEst"
         Groups(6).Columns(0).DataField=   "Column 19"
         Groups(6).Columns(0).DataType=   8
         Groups(6).Columns(0).FieldLen=   256
         Groups(7).Width =   1984
         Groups(7).Caption=   "Tipo"
         Groups(7).Columns(0).Width=   1984
         Groups(7).Columns(0).Caption=   "Open/ Closed"
         Groups(7).Columns(0).Name=   "Flg_Status"
         Groups(7).Columns(0).DataField=   "Column 20"
         Groups(7).Columns(0).DataType=   8
         Groups(7).Columns(0).FieldLen=   256
         Groups(8).Width =   6853
         Groups(8).Caption=   "Motivo Atraso"
         Groups(8).Columns.Count=   2
         Groups(8).Columns(0).Width=   1852
         Groups(8).Columns(0).Caption=   "Motivo"
         Groups(8).Columns(0).Name=   "Cod_MotAtr"
         Groups(8).Columns(0).DataField=   "Column 21"
         Groups(8).Columns(0).DataType=   8
         Groups(8).Columns(0).FieldLen=   256
         Groups(8).Columns(1).Width=   5001
         Groups(8).Columns(1).Caption=   "Nombre"
         Groups(8).Columns(1).Name=   "Des_MotAtr"
         Groups(8).Columns(1).DataField=   "Column 22"
         Groups(8).Columns(1).DataType=   8
         Groups(8).Columns(1).FieldLen=   256
         Groups(9).Width =   14605
         Groups(9).Caption=   "Datos Generales"
         Groups(9).Columns.Count=   10
         Groups(9).Columns(0).Width=   741
         Groups(9).Columns(0).Caption=   "% Comision"
         Groups(9).Columns(0).Name=   "Por_Comision"
         Groups(9).Columns(0).DataField=   "Column 23"
         Groups(9).Columns(0).DataType=   8
         Groups(9).Columns(0).FieldLen=   256
         Groups(9).Columns(1).Width=   741
         Groups(9).Columns(1).Caption=   "Div.Pre"
         Groups(9).Columns(1).Name=   "Cod_DivPre"
         Groups(9).Columns(1).DataField=   "Column 24"
         Groups(9).Columns(1).DataType=   8
         Groups(9).Columns(1).FieldLen=   256
         Groups(9).Columns(2).Width=   688
         Groups(9).Columns(2).Caption=   "Flag Div.Pre"
         Groups(9).Columns(2).Name=   "Flg_DivPreDif"
         Groups(9).Columns(2).DataField=   "Column 25"
         Groups(9).Columns(2).DataType=   8
         Groups(9).Columns(2).FieldLen=   256
         Groups(9).Columns(3).Width=   714
         Groups(9).Columns(3).Caption=   "Flg_ComisionEnPorcentaje"
         Groups(9).Columns(3).Name=   "Flg_ComisionEnPorcentaje"
         Groups(9).Columns(3).DataField=   "Column 26"
         Groups(9).Columns(3).DataType=   8
         Groups(9).Columns(3).FieldLen=   256
         Groups(9).Columns(4).Width=   714
         Groups(9).Columns(4).Caption=   "Imp_Comision"
         Groups(9).Columns(4).Name=   "Imp_Comision"
         Groups(9).Columns(4).DataField=   "Column 27"
         Groups(9).Columns(4).DataType=   8
         Groups(9).Columns(4).FieldLen=   256
         Groups(9).Columns(5).Width=   714
         Groups(9).Columns(5).Caption=   "flg_PrePack"
         Groups(9).Columns(5).Name=   "flg_PrePack"
         Groups(9).Columns(5).DataField=   "Column 28"
         Groups(9).Columns(5).DataType=   8
         Groups(9).Columns(5).FieldLen=   256
         Groups(9).Columns(6).Width=   2805
         Groups(9).Columns(6).Caption=   "Ex-Fact Reprogr"
         Groups(9).Columns(6).Name=   "Fec_DespachoOri_Reprogramada"
         Groups(9).Columns(6).DataField=   "Column 29"
         Groups(9).Columns(6).DataType=   8
         Groups(9).Columns(6).FieldLen=   256
         Groups(9).Columns(7).Width=   3281
         Groups(9).Columns(7).Caption=   "Fec_ExFactory_Ajustada_Facturacion"
         Groups(9).Columns(7).Name=   "Fec_ExFactory_Ajustada_Facturacion"
         Groups(9).Columns(7).DataField=   "Column 30"
         Groups(9).Columns(7).DataType=   8
         Groups(9).Columns(7).FieldLen=   256
         Groups(9).Columns(8).Width=   979
         Groups(9).Columns(8).Caption=   "Tip_Lead_time"
         Groups(9).Columns(8).Name=   "Tip_Lead_time"
         Groups(9).Columns(8).DataField=   "Column 31"
         Groups(9).Columns(8).DataType=   8
         Groups(9).Columns(8).FieldLen=   256
         Groups(9).Columns(9).Width=   3228
         Groups(9).Columns(9).Caption=   "Des_Tip_Lead_Time"
         Groups(9).Columns(9).Name=   "Des_Tip_Lead_Time"
         Groups(9).Columns(9).DataField=   "Column 32"
         Groups(9).Columns(9).DataType=   8
         Groups(9).Columns(9).FieldLen=   256
         _ExtentX        =   23945
         _ExtentY        =   4657
         _StockProps     =   79
         Caption         =   "Styles "
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
   Begin FunctionsButtons.FunctButt FunctButt3 
      Height          =   495
      Left            =   13920
      TabIndex        =   28
      Top             =   5880
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      Custom          =   "0~0~AVANENCAJADO~Verdadero~Verdadero~&Avance Encajado~0~0~1~~0~Falso~Falso~&Avance Encajado~"
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1350
      ControlHeigth   =   490
      ControlSeparator=   0
   End
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   705
      Left            =   0
      TabIndex        =   4
      Top             =   6360
      Width           =   15345
      _ExtentX        =   27093
      _ExtentY        =   1244
      Custom          =   $"frmShowTG_PurOrd4.frx":0442
      Orientacion     =   0
      Style           =   0
      Language        =   2
      TypeImageList   =   0
      ControlWidth    =   1150
      ControlHeigth   =   680
      ControlSeparator=   30
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   5280
      Left            =   13920
      TabIndex        =   3
      Top             =   600
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   9313
      Custom          =   $"frmShowTG_PurOrd4.frx":0A1F
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1350
      ControlHeigth   =   450
      ControlSeparator=   30
   End
   Begin FunctionsButtons.FunctButt FunctButt4 
      Height          =   705
      Left            =   -15
      TabIndex        =   29
      Top             =   7080
      Width           =   15330
      _ExtentX        =   27040
      _ExtentY        =   1244
      Custom          =   $"frmShowTG_PurOrd4.frx":0F25
      Orientacion     =   0
      Style           =   0
      Language        =   2
      TypeImageList   =   0
      ControlWidth    =   1500
      ControlHeigth   =   680
      ControlSeparator=   32
   End
   Begin FunctionsButtons.FunctButt FunctButt5 
      Height          =   705
      Left            =   0
      TabIndex        =   30
      Top             =   7800
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   1244
      Custom          =   $"frmShowTG_PurOrd4.frx":1330
      Orientacion     =   0
      Style           =   0
      Language        =   2
      TypeImageList   =   0
      ControlWidth    =   1500
      ControlHeigth   =   680
      ControlSeparator=   20
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   3480
      Top             =   8040
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmShowTG_PurOrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declaración de Variables Nivel Formulario
Option Explicit

Public oParent      As Object

Public sCaptionForm As String

Public PrinterHeight

Public iLin                     As Integer

Public iMante                   As Integer

Public sCod_Cliente             As String

Public sCod_PurOrd              As String

Public dPor_ComisionCliente     As Double

Public sNivAccUsuario           As String

Public Tipo_Rep                 As String

Public Tipo_RepAcum             As String

Public sPONew                   As String

Public sCod_TemCli              As String

Public sCod_Fabrica             As String

Public sCod_EstPro              As String

Public bChangedPODetalleDestino As Boolean

Public vFilaActual              As Variant

Public nColumnaActual           As Integer

Dim sFlag                       As String

Public Sub Reporte()

    On Error GoTo hand

    Dim oo   As Object

    Dim Ruta As String

    Dim Usu  As String

    Ruta = App.Path & "\Proforma.XLT"
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
        
    'oo.Run "reporte", CStr(DevuelveCampo("select cod_cliente from tg_cliente where abr_cliente='" & txtAbr_Cliente & "'", cCONNECT)), CStr(ssgrdDatos.Columns("Purchase Order").Text), CStr(cCONNECT), CStr(vemp)
    oo.Run "reporte", CStr(DevuelveCampo("select cod_cliente from tg_cliente where abr_cliente='" & txtAbr_Cliente & "'", cCONNECT)), CStr(cCONNECT), CStr(vemp), CStr(ssgrdDatos.Columns("Purchase Order").Text)
    Set oo = Nothing

    Exit Sub

hand:
    ErrorHandler Err, "GeneraReportes"
    Set oo = Nothing

End Sub

Public Sub Reporte_Control()

    Dim Ruta, Usu As String

    Dim strSql     As String

    Dim sRuta_Logo As String

    Dim oo         As Object

    On Error GoTo ImprimirErr

    strSql = "SELECT Ruta_Logo=ISNULL(Ruta_Logo,'') FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA='" & vemp & "'"
    sRuta_Logo = DevuelveCampo(strSql, cCONNECT)
    
    Ruta = App.Path & "\RptControl.xlt"
    Usu = "Usuario : " & vusu
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "Reporte", Me.sCod_Cliente, Me.ssgrdDatos.Columns("Cod_PurOrd").Text, cCONNECT, sRuta_Logo
    Tipo_Rep = ""
    Set oo = Nothing

    Exit Sub
    
ImprimirErr:
    ErrorHandler Err, "Reporte_Control"
    Set oo = Nothing
End Sub

Private Sub acbOperaciones2_ActionClick(ByVal Index As Integer, _
                                        ByVal ActionType As Integer, _
                                        ByVal ActionName As String)
End Sub

Private Sub Form_Load()

    Dim X As Variant

    IdiomaEtiquetas Me
    'Call FormSet(Me)
    Me.Caption = sCaptionForm
    SSDBGridSetGrid0 Me.ssgrdDatos
    SSDBGridSetGrid0 Me.ssgrdDatos2
    Me.FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    Me.FunctButt2.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    'Me.FunctButt3.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    Me.FunctButt4.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    Me.FunctButt5.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me

    If Not oParent Is Nothing Then
        oParent.DropWindowList Me.Tag
    End If

End Sub

Private Sub acbOperaciones_ActionClick(ByVal Index As Integer, _
                                       ByVal ActionType As Integer, _
                                       ByVal ActionName As String)
    iMante = Index
End Sub

Private Sub cmdBuscar_Click()

    Dim bRet As Boolean

    'cCONNECT = "Provider=sqloledb;Server=SERVER_DEV;Database=HIALPESA;Integrated Security=SSPI;Use Procedure for Prepare=0;Auto Translate=FALSE;Packet Size=4096;Use Encryption for Data=FALSE;Tag with column collation when possible=FALSE"
    bRet = BUSCAR

    If Me.optCod_PurOrd.value = True And Not bRet And RTrim(Me.txtCod_PurOrd.Text) <> "" Then
        If Wizard(True) Then
            BUSCAR
        End If
    End If

End Sub

Public Function BUSCAR() As Boolean
        
    Dim obj   As clsTG_PurOrd

    Dim vbuff As Variant

    Dim irow  As Variant

    BUSCAR = False

    irow = Me.ssgrdDatos.Bookmark
    iRowsGrilla = ssgrdDatos.Rows
    Me.ssgrdDatos.Redraw = False

    SSDBGridSetGrid Me.ssgrdDatos

    Set obj = New clsTG_PurOrd
    obj.ConexionString = cCONNECT

    If Me.optCod_TemCli.value = True Then
        vbuff = obj.ViewDetalle(sCod_Cliente, "", Me.txtCod_TemCli.Text, "", "", vusu)
    ElseIf Me.optCod_PurOrd.value = True Then
        vbuff = obj.ViewDetalle(sCod_Cliente, Me.txtCod_PurOrd.Text, "", "", "", vusu)
    ElseIf Me.optCod_EstCli.value = True Then
        vbuff = obj.ViewDetalle(sCod_Cliente, "", "", Me.txtCod_EstCli.Text, "", vusu)
    Else
        vbuff = obj.ViewDetalle(sCod_Cliente, "", "", "", Me.txtCod_Ordpro.Text, vusu)
    End If

    If Not IsEmpty(vbuff) Then
        BUSCAR = True

        If RTrim(DevuelveCampo("SELECT ISNULL(FLG_CREAMV,'N') FROM TG_CONTROL ", cCONNECT)) = "S" Then
            If vbuff(42, 0) = "E" Then
                Aviso "CUIDADO!! PO. Solicitada NO EXISTE EN CLIENTE ACTUAL.", 1
                sCod_Cliente = vbuff(43, 0)
                txtAbr_Cliente.Text = vbuff(44, 0)
                txtNom_Cliente.Text = vbuff(45, 0)
            End If
        End If
    End If

    LibraryVBToSSDBGrid obj, vbuff, ssgrdDatos
    ssgrdDatos.ActiveRowStyleSet = "RowActive"
    ssgrdDatos.SelectTypeRow = ssSelectionTypeMultiSelectRange
    Set obj = Nothing

    RestoreRowSSDBGrid Me.ssgrdDatos, irow, iRowsGrilla
    Me.ssgrdDatos.Redraw = True

    If Me.Enabled Then
        Me.ssgrdDatos.SetFocus
    End If

    BuscarEStilos

    Exit Function

errores:
    Me.MousePointer = vbDefault

    If Not obj Is Nothing Then
        Set obj = Nothing
    End If

    ErrorHandler Err, Err.Description

End Function

Sub Plin(ByVal Text)

    If IsNull(Text) Then
        Text = ""
    End If

    Print #1, Text
    iLin = iLin + 1
End Sub

Private Sub txtCod_Cliente_Change()

End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, _
                                   ByVal ActionType As Integer, _
                                   ByVal ActionName As String)

    Dim oWizard As frmWizard

    Select Case ActionName

        Case "VERLOTE"
            
            If Me.ssgrdDatos.Rows > 0 Then
                BuscarEStilos
            End If

        Case "MODIFICAR"

            If sNivAccUsuario = "V" Or ssgrdDatos.Columns("NivAcc").Text = "V" Then
                Mensaje kMESSAGE_ERR_NOT_RIGHT_OPTION

                Exit Sub

            End If
            
            Set oWizard = New frmWizard
            Load oWizard
            oWizard.sAccionName = ActionName
            oWizard.dPor_ComisionCliente = dPor_ComisionCliente
            oWizard.sCod_Cliente = Me.sCod_Cliente
            oWizard.sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
            Set oWizard.oParent = Me
            oWizard.cmdNav(0).Visible = False
            oWizard.cmdNav(1).Visible = False
            oWizard.cmdNav(2).Visible = False
            oWizard.cmdNav(3).Visible = False
            oWizard.cmdNav(4).Visible = False
            oWizard.cmdAceptar.Visible = True
            oWizard.cmdAceptar.Top = oWizard.cmdNav(0).Top
            oWizard.cmdCancelar.Visible = True
            oWizard.cmdCancelar.Top = oWizard.cmdNav(0).Top
            oWizard.LoadPOC oWizard.sCod_Cliente, oWizard.sCod_PurOrd
            oWizard.dtpFec_DespachoAct.Visible = False
            oWizard.txtCod_Destino.Visible = False
            oWizard.txtDes_Destino.Visible = False
            oWizard.cmdCod_Destino.Visible = False
            oWizard.labels(3).Visible = False
            oWizard.labels(1).Visible = False
            oWizard.Show vbModal
            Set oWizard = Nothing

        Case "ELIMINAR"

            If sNivAccUsuario = "V" Or ssgrdDatos.Columns("NivAcc").Text = "V" Then
                Mensaje kMESSAGE_ERR_NOT_RIGHT_OPTION

                Exit Sub

            End If
        
            Set oWizard = New frmWizard
            Load oWizard
            oWizard.dPor_ComisionCliente = dPor_ComisionCliente
            oWizard.sAccionName = ActionName
            oWizard.sCod_Cliente = Me.sCod_Cliente
            oWizard.sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
            Set oWizard.oParent = Me
            oWizard.cmdNav(0).Visible = False
            oWizard.cmdNav(1).Visible = False
            oWizard.cmdNav(2).Visible = False
            oWizard.cmdNav(3).Visible = False
            oWizard.cmdNav(4).Visible = False
            oWizard.cmdAceptar.Visible = True
            oWizard.cmdAceptar.Top = oWizard.cmdNav(0).Top
            oWizard.cmdCancelar.Visible = True
            oWizard.cmdCancelar.Top = oWizard.cmdNav(0).Top
            oWizard.LoadPOC oWizard.sCod_Cliente, oWizard.sCod_PurOrd
            oWizard.Show vbModal
            Set oWizard = Nothing

        Case "IMPRIMIR"
            Load frmSelReporte
            Set frmSelReporte.oParent = Me
            frmSelReporte.Show vbModal
            Set frmSelReporte = Nothing

            If Tipo_Rep = "" Then

                Exit Sub

            End If

            Dim Ruta, Usu As String

            Dim oo As Object

            On Error GoTo ImprimirErr

            'Esto para usar el otro reporte
            If Tipo_RepAcum = "SI" Then
                Ruta = App.Path & "\PODetalleTotal.xlt"
            Else
                Ruta = App.Path & "\PODetalle.xlt"
            End If

            Usu = "Usuario : " & vusu
            Set oo = CreateObject("excel.application")
            oo.Workbooks.Open Ruta
            oo.Visible = True
            oo.DisplayAlerts = False

            'Esto para usar el otro reporte
            If Tipo_RepAcum = "SI" Then
                oo.Run "GeneraReporte", Me.sCod_Cliente, Me.ssgrdDatos.Columns("Cod_PurOrd").Text, Usu, Tipo_Rep, "S", cCONNECT
            Else
                oo.Run "GeneraReporte", Me.sCod_Cliente, Me.ssgrdDatos.Columns("Cod_PurOrd").Text, Usu, Tipo_Rep, Tipo_RepAcum, cCONNECT
            End If
            
            Tipo_Rep = ""
            Set oo = Nothing

        Case "COPIAR"

            Dim strVerif As String

            If Me.ssgrdDatos.Columns("Cod_PurOrd").Text = "" Or IsNull(Me.ssgrdDatos.Columns("Cod_PurOrd").Text) Then

                Exit Sub

            End If

            strVerif = "SELECT dbo.uf_verif_purord ('" & Me.sCod_Cliente & "','" & Me.ssgrdDatos.Columns("Cod_PurOrd").Text & "')"
            strVerif = DevuelveCampo(strVerif, cCONNECT)

            If strVerif <> "0" Then
                MsgBox "No se puede realizaar la operación. PO tiene estilos cerrados", vbInformation, "Copiar PO"

                Exit Sub

            End If

            'Ventana para Ingresar nuevo PO
            sPONew = ""

            Dim oCopia As New frmNewPO

            Load oCopia
            Set oCopia.oParent = Me
            oCopia.txtIdCliente.Text = Me.sCod_Cliente
            oCopia.txtNomCliente.Text = txtNom_Cliente.Text
            oCopia.Show 1
            Set oCopia = Nothing

            If sPONew = "" Then

                Exit Sub

            End If

            'Realiza Copia de PO
            If Not CopiaPO() Then
                MsgBox "Error. No se pudo copiar el PO", vbInformation, "Copiar PO"

                Exit Sub

            End If

            BUSCAR
            MsgBox "Copia PO creada satisfactoriamente" & vbCr & "Nuevo #PO [" & sPONew & "]", vbInformation, "Copiar PO"

        Case "GENERAROP"

            If ssgrdDatos.Rows > 0 Then
            
                If ssgrdDatos.Columns("Cod_PurOrd").Text = "VP" Then
                    MsgBox "Esta clase de PO del registro no permite acceder a esta acción. Sirvase verificar", vbInformation, "Genera O/P"

                    Exit Sub

                End If
            
                Load frmGenerarOP
                Set frmGenerarOP.oParent = Me
                frmGenerarOP.sCod_Cliente = sCod_Cliente
                frmGenerarOP.sCod_PurOrd = ssgrdDatos.Columns("Cod_PurOrd").Text
                frmGenerarOP.Show vbModal
                Set frmGenerarOP = Nothing
            End If
        
        Case "PROFORMA"
            Reporte

        Case "IMPCONTROL"
            Call Reporte_Control

        Case "CAMBIARPOGEN"

            If ssgrdDatos.Rows > 0 Then
                Load frmChangePO
                frmChangePO.varCod_Cliente = sCod_Cliente
                frmChangePO.bNivelPO = True
                frmChangePO.varCod_EstCli = ""
                
                frmChangePO.varCod_LotPurOrd = ""
                frmChangePO.varCod_TemCli = ssgrdDatos.Columns("Cod_TemCli").Text
                
                frmChangePO.txtPO = ssgrdDatos.Columns("Cod_PurOrd").Text
                
                frmChangePO.txtCliente = Trim(Me.txtAbr_Cliente.Text) & " - " & Trim(Me.txtNom_Cliente.Text)
                frmChangePO.txtEstilo.Visible = False
                frmChangePO.Show 1
                Set frmChangePO = Nothing
                Call BUSCAR
            End If

        Case "CAMBIARCLASEPO"

            If ssgrdDatos.Rows > 0 Then
                frmCambiaClasePO.varCod_Cliente = sCod_Cliente
                frmCambiaClasePO.varNom_Cliente = txtNom_Cliente.Text
                frmCambiaClasePO.varCod_Purord = ssgrdDatos.Columns("Cod_PurOrd").Text
                frmCambiaClasePO.varCod_Clapurord = ssgrdDatos.Columns("Cod_ClaPurOrd").Text
                frmCambiaClasePO.Show 1
                Call BUSCAR
            End If

        Case "CAPRECOL"
        
            Load frmCambiaPrecioColor
            Set frmCambiaPrecioColor.oParent = Me
            frmCambiaPrecioColor.sCod_Cliente = Me.sCod_Cliente
            frmCambiaPrecioColor.sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
            frmCambiaPrecioColor.sCod_LotPurOrd = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
            frmCambiaPrecioColor.sCod_EstCli = ssgrdDatos2.Columns("Cod_EstCli").Text
            frmCambiaPrecioColor.sCod_TemCli = ssgrdDatos.Columns("Cod_TemCli").Text
            frmCambiaPrecioColor.BUSCAR
            frmCambiaPrecioColor.Show vbModal
            Set frmCambiaPrecioColor = Nothing

        Case "ACTUAFECHLLEG"
            Load Frm_Actualiza_FechaPO
            Set Frm_Actualiza_FechaPO.oParent = Me

            With Frm_Actualiza_FechaPO
                .sCod_Cliente = Me.sCod_Cliente
                .sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text

                If Me.ssgrdDatos.Columns("Fec_Hora_Llegada_PO").Text = "" Then
                    .DTPicker1.value = Format(Now, "dd/MM/yyyy HH:MM")
                    .DTPicker1.value = Null
                Else
                    .DTPicker1.value = Format(Me.ssgrdDatos.Columns("Fec_Hora_Llegada_PO").Text, "dd/MM/yyyy HH:MM")
                End If

                .Show vbModal
            End With

            Set Frm_Actualiza_FechaPO = Nothing
    End Select

    Exit Sub

ImprimirErr:
    ErrorHandler Err, "Imprimir"
    Set oo = Nothing
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, _
                                   ByVal ActionType As Integer, _
                                   ByVal ActionName As String)

    On Error GoTo errores

    'cCONNECT = "Provider=sqloledb;Server=SERVER_DEV;Database=HIALPESA;Integrated Security=SSPI;Use Procedure for Prepare=0;Auto Translate=FALSE;Packet Size=4096;Use Encryption for Data=FALSE;Tag with column collation when possible=FALSE"

    Dim oWizard     As frmWizard

    Dim oColor      As frmMantColor

    Dim otalla      As frmMantPurOrdTal

    Dim sLote       As String

    Dim sCod_EstCli As String

    Dim NVEZ        As Integer

    Select Case ActionName

        Case "MODIFICARDET"
            
            If sNivAccUsuario = "V" Or ssgrdDatos.Columns("NivAcc").Text = "V" Then
                Mensaje kMESSAGE_ERR_NOT_RIGHT_OPTION

                Exit Sub

            End If
    
            If ssgrdDatos2.Columns("Flg_Status").Text = "C" Then
                Mensaje kMESSAGE_ERR_LOTEST_CLOSED

                Exit Sub

            End If
        
            Set oWizard = New frmWizard
            Load oWizard
            oWizard.sLote = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
            oWizard.sCod_EstCli = ssgrdDatos2.Columns("Cod_EstCli").Text
            oWizard.sAccionName = "MODIFICAR"
            oWizard.sModoWizard = "   ESTDAT"
            oWizard.dPor_ComisionCliente = dPor_ComisionCliente
            oWizard.sCod_Cliente = Me.sCod_Cliente
            oWizard.sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
            Set oWizard.oParent = Me
            oWizard.LoadPOC oWizard.sCod_Cliente, oWizard.sCod_PurOrd
        
            oWizard.fraStep(0).Enabled = False
            oWizard.txtCod_EstCliLOT.Enabled = False
            oWizard.cmdCod_EstCli.Enabled = False
            oWizard.txtAbr_FabricaLOT.Enabled = False
            oWizard.txtNom_FabricaLOT.Enabled = False
            oWizard.txtPrecioLOT.Enabled = False
            oWizard.txtCod_DestinoLOT.Enabled = False
            oWizard.txtDes_DestinoLOT.Enabled = False
            oWizard.txtPrecioLOT.Enabled = False
            oWizard.txtCod_DivPreLOT.Enabled = False
            oWizard.dtpFec_DespachoActLOT.Enabled = False
            'oWizard.txtPor_ComisionLOT.Enabled = False
            oWizard.chkPrecioIgual.value = "0"
            'oWizard.SetStep 0, 2
            oWizard.ValidStep 0, 2, True
            oWizard.SetStep 1, 2
            oWizard.ValidStep 1, 2, True
                    
            oWizard.PU_strCodClaPO = Me.ssgrdDatos.Columns("Cod_ClaPurOrd").Text
                    
            oWizard.cmdMatrizDestinoEmpaque.Visible = True
            oWizard.Show vbModal
            Set oWizard = Nothing
        
            'If bChangedPODetalleDestino Then
            
            '    FunctButt2_ActionClick 0, 0, "MODIFICARDET"
            '    bChangedPODetalleDestino = False
            'End If
        
        Case "INCORPORAR"

            If sNivAccUsuario = "V" Or ssgrdDatos.Columns("NivAcc").Text = "V" Then
                Mensaje kMESSAGE_ERR_NOT_RIGHT_OPTION

                Exit Sub

            End If
    
            Set oWizard = New frmWizard
            Load oWizard
            oWizard.sAccionName = ActionName
            oWizard.sModoWizard = "   ESTDAT"
            oWizard.dPor_ComisionCliente = dPor_ComisionCliente
            oWizard.sCod_Cliente = Me.sCod_Cliente
            oWizard.sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
            Set oWizard.oParent = Me
            oWizard.LoadPOC oWizard.sCod_Cliente, oWizard.sCod_PurOrd
            oWizard.txtCod_DestinoLOT.Text = oWizard.txtCod_Destino.Text
            oWizard.txtDes_DestinoLOT.Text = oWizard.txtDes_Destino.Text
            oWizard.sCod_DestinoLOT = oWizard.txtCod_Destino.Text
        
            oWizard.txtAbr_FabricaLOT.Text = oWizard.txtAbr_Fabrica.Text
            oWizard.txtNom_FabricaLOT.Text = oWizard.txtNom_Fabrica.Text
            oWizard.sCod_FabricaLot = oWizard.sCod_Fabrica
        
            oWizard.fraStep(0).Enabled = False
            'oWizard.SetStep 0, 2
            oWizard.ValidStep 0, 2, True
            oWizard.SetStep 1, 2
            'oWizard.ValidStep 1, 2 , true
        
            oWizard.PU_strCodClaPO = Me.ssgrdDatos.Columns("Cod_ClaPurOrd").Text
        
            oWizard.Show vbModal
            Set oWizard = Nothing
        
        Case "ELIMINARDET"

            If sNivAccUsuario = "V" Or ssgrdDatos.Columns("NivAcc").Text = "V" Then
                Mensaje kMESSAGE_ERR_NOT_RIGHT_OPTION

                Exit Sub

            End If
    
            DeleteLotPurOrd

        Case "COLORES"

            If sNivAccUsuario = "V" Or ssgrdDatos.Columns("NivAcc").Text = "V" Then
                Mensaje kMESSAGE_ERR_NOT_RIGHT_OPTION

                Exit Sub

            End If
    
            Set oColor = New frmMantColor
            Load oColor
            oColor.VALIDAR_ACCESO_MERMA
            Set oColor.oParent = Me
            oColor.sCod_Cliente = Me.sCod_Cliente
            oColor.sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
            oColor.sCod_EstCli = Me.ssgrdDatos2.Columns("Cod_EstCli").Text
            oColor.Inicializar
            oColor.Cargar_Datos
            oColor.txtabrecli.Text = txtAbr_Cliente.Text
            oColor.Show vbModal
            Set oColor = Nothing

        Case "TALLAS"

            If sNivAccUsuario = "V" Or ssgrdDatos.Columns("NivAcc").Text = "V" Then
                Mensaje kMESSAGE_ERR_NOT_RIGHT_OPTION

                Exit Sub

            End If

            frmMantPurOrdTal.Cliente = DevuelveCampo("select cod_cliente from tg_cliente where abr_cliente='" & txtAbr_Cliente & "'", cCONNECT)
            frmMantPurOrdTal.txtNom_Cliente.Text = txtNom_Cliente
            frmMantPurOrdTal.PO = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
            frmMantPurOrdTal.txtCod_PurOrd.Text = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
            frmMantPurOrdTal.Estilo = Me.ssgrdDatos2.Columns("Cod_EstCli").Text
            frmMantPurOrdTal.txtCod_EstCli.Text = Me.ssgrdDatos2.Columns("Cod_EstCli").Text
            frmMantPurOrdTal.Cargar_Data
            frmMantPurOrdTal.Show

        Case "DETALLEXTALLA"
                
            Set oWizard = New frmWizard
            Load oWizard
            oWizard.sLote = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
            oWizard.sCod_EstCli = ssgrdDatos2.Columns("Cod_EstCli").Text
            oWizard.sAccionName = ActionName
            oWizard.sModoWizard = "   ESTDAT"
            oWizard.dPor_ComisionCliente = dPor_ComisionCliente
            oWizard.sCod_Cliente = Me.sCod_Cliente
            oWizard.sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
            Set oWizard.oParent = Me
            oWizard.LoadPOC oWizard.sCod_Cliente, oWizard.sCod_PurOrd
                
            oWizard.fraStep(0).Enabled = False
            oWizard.txtCod_EstCliLOT.Enabled = False
            oWizard.txtAbr_FabricaLOT.Enabled = False
            oWizard.txtNom_FabricaLOT.Enabled = False
            oWizard.txtPrecioLOT.Enabled = False
            oWizard.chkPrecioIgual.value = "0"
            oWizard.SetStep 0, 2
            oWizard.ValidStep 0, 2, False
            oWizard.SetStep 1, 2
            oWizard.SetStep 2, 2
            oWizard.ValidStep 2, 2, False
            oWizard.SetStep 3, 2
            oWizard.ValidStep 3, 2, False
        
            oWizard.cmdNav(0).Visible = False
            oWizard.cmdNav(1).Visible = False
            oWizard.cmdNav(2).Visible = False
            oWizard.cmdNav(3).Visible = False
            oWizard.cmdNav(4).Visible = False
            oWizard.cmdAceptar.Visible = True
            oWizard.cmdAceptar.Top = oWizard.cmdNav(0).Top
            oWizard.cmdCancelar.Visible = True
            oWizard.cmdCancelar.Top = oWizard.cmdNav(0).Top
        
            oWizard.LibraryVBToMatrizResumen oWizard.SSgrdDatosPrec, True, False, True, False, True, False, True, False, True, False, True, True, False, False, True, True
            oWizard.Show vbModal
            Set oWizard = Nothing

        Case "DETALLEXCOLOR"
            Load frmColorDetail
            Set frmColorDetail.oParent = Me
            frmColorDetail.sCod_Cliente = sCod_Cliente
            frmColorDetail.sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
            frmColorDetail.sCod_LotPurOrd = Me.ssgrdDatos2.Columns("Cod_LotPurOrd").Text
            frmColorDetail.sCod_EstCli = Me.ssgrdDatos2.Columns("Cod_EstCli").Text
            frmColorDetail.BUSCAR
            frmColorDetail.Show vbModal
            Set frmColorDetail = Nothing

        Case "IMPRIMIRDET"
            Load frmSelReporte
            frmSelReporte.optTotal.Visible = False
            frmSelReporte.optSimple.Left = frmSelReporte.optSimple.Left + 900
            frmSelReporte.optAgrupado.Left = frmSelReporte.optAgrupado.Left + 900
                
            Set frmSelReporte.oParent = Me
            frmSelReporte.Show vbModal
            Set frmSelReporte = Nothing

            If Tipo_Rep = "" Then

                Exit Sub

            End If

            Dim Ruta, Usu As String

            Dim oo    As Object

            Dim iFila As Integer

            On Error GoTo errores

            Ruta = App.Path & "\PODetalle.xlt"
            Usu = "Usuario : " & vusu
            Set oo = CreateObject("excel.application")
            oo.Workbooks.Open Ruta
            oo.Visible = True
            oo.DisplayAlerts = False

            iFila = oo.Run("ReporteMatriz", Usu, Me.sCod_Cliente, Me.ssgrdDatos.Columns("Cod_PurOrd").Text, Me.ssgrdDatos2.Columns("Cod_LotPurOrd").Text, Me.ssgrdDatos2.Columns("Cod_EstCli").Text, Tipo_Rep, Tipo_RepAcum, cCONNECT, 7, True)
            Tipo_Rep = ""
            Set oo = Nothing

        Case "VIEWLOTE"
            Set oWizard = New frmWizard
            Load oWizard
            oWizard.sLote = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
            oWizard.sCod_EstCli = ssgrdDatos2.Columns("Cod_EstCli").Text
            oWizard.sAccionName = ActionName
            oWizard.sModoWizard = "   ESTDAT"
            oWizard.dPor_ComisionCliente = dPor_ComisionCliente
            oWizard.sCod_Cliente = Me.sCod_Cliente
            oWizard.sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
            Set oWizard.oParent = Me
            oWizard.LoadPOC oWizard.sCod_Cliente, oWizard.sCod_PurOrd
                
            oWizard.fraStep(0).Enabled = False
            oWizard.txtCod_EstCliLOT.Enabled = False
            oWizard.txtAbr_FabricaLOT.Enabled = False
            oWizard.txtNom_FabricaLOT.Enabled = False
            oWizard.txtPrecioLOT.Enabled = False
            oWizard.txtCod_DestinoLOT.Enabled = False
            oWizard.txtDes_DestinoLOT.Enabled = False
            oWizard.txtPrecioLOT.Enabled = False
            oWizard.dtpFec_DespachoActLOT.Enabled = False
            oWizard.txtPor_ComisionLOT.Enabled = False
            oWizard.chkPrecioIgual.value = "0"
            oWizard.SetStep 0, 2
            oWizard.ValidStep 0, 2, False
            oWizard.SetStep 1, 2
            oWizard.ValidStep 1, 2, False
            oWizard.fraStep(1).Enabled = False
            oWizard.cmdNav(0).Visible = False
            oWizard.cmdNav(1).Visible = False
            oWizard.cmdNav(2).Visible = False
            oWizard.cmdNav(3).Visible = False
            oWizard.cmdNav(4).Visible = False
            oWizard.cmdAceptar.Visible = True
            oWizard.cmdAceptar.Enabled = False
            oWizard.cmdAceptar.Top = oWizard.cmdNav(0).Top
            oWizard.cmdCancelar.Visible = True
            oWizard.cmdCancelar.Top = oWizard.cmdNav(0).Top
            oWizard.Show vbModal
            Set oWizard = Nothing

        Case "GENERALUPDATE"
    
            Dim COD As String

            COD = DevuelveCampo("SELECT DBO.TG_MUESTRA_SEGPRECIOS_FOB('" & vusu & "')", cCONNECT)
        
            If Trim(COD) = "" Then MsgBox "No Tiene Acceso para ver ni modificar Precio": Exit Sub
        
            If sNivAccUsuario = "V" Or ssgrdDatos.Columns("NivAcc").Text = "V" Then
                Mensaje kMESSAGE_ERR_NOT_RIGHT_OPTION

                Exit Sub

            End If
    
            If ssgrdDatos2.Columns("Flg_Status").Text = "C" Then
                Mensaje kMESSAGE_ERR_LOTEST_CLOSED

                Exit Sub

            End If

            Load frmUpdateDatGenLotEst
        
            Set frmUpdateDatGenLotEst.oParent = Me
            '        frmUpdateDatGenLotEst.sCod_Cliente = sCod_Cliente
            '        frmUpdateDatGenLotEst.sCod_PurOrd = ssgrdDatos.Columns("Cod_PurORd").Text
            '        frmUpdateDatGenLotEst.sCod_LotPurORd = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
            '        frmUpdateDatGenLotEst.sCod_EstCli = ssgrdDatos2.Columns("Cod_EStCli").Text
            '        frmUpdateDatGenLotEst.txtCod_DestinoLOT.Text = ssgrdDatos2.Columns("Cod_Destino").Text
            '        frmUpdateDatGenLotEst.txtDes_DestinoLOT.Text = ssgrdDatos2.Columns("DES_Destino").Text
            '        frmUpdateDatGenLotEst.dtpFec_DespachoActLOT.Value = ssgrdDatos2.Columns("Fec_DespachoAct").Text
            '        frmUpdateDatGenLotEst.txtPor_ComisionLOT.Text = ssgrdDatos2.Columns("Por_Comision").Text
            '        frmUpdateDatGenLotEst.txtPrecioLOT.Text = ssgrdDatos2.Columns("Precio").Text
            '        frmUpdateDatGenLotEst.txtCod_DivPreLOT.Text = ssgrdDatos2.Columns("Cod_DivPre").Text
        
            frmUpdateDatGenLotEst.Txtcod_Lead = ssgrdDatos2.Columns("TIP_LEAD_TIME").Text
            frmUpdateDatGenLotEst.TxtLead = ssgrdDatos2.Columns("DES_TIP_LEAD_TIME").Text
        
            frmUpdateDatGenLotEst.sCod_Cliente = sCod_Cliente
            frmUpdateDatGenLotEst.sCod_PurOrd = ssgrdDatos.Columns("Cod_PurORd").Text
            frmUpdateDatGenLotEst.sCod_LotPurOrd = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
            frmUpdateDatGenLotEst.sCod_EstCli = ssgrdDatos2.Columns("Cod_EStCli").Text
            frmUpdateDatGenLotEst.txtCod_DestinoLOT.Text = ssgrdDatos2.Columns("Cod_Destino").Text
            frmUpdateDatGenLotEst.txtDes_DestinoLOT.Text = ssgrdDatos2.Columns("DES_Destino").Text
        
            frmUpdateDatGenLotEst.dtpFec_DespachoOriLOT.value = ssgrdDatos2.Columns("Fec_DespachoOri").Text

            Dim sEditarExFactory As String

            sEditarExFactory = DevuelveCampo("SELECT Flg_Editar_ExFactory FROM TG_CONTROl", cCONNECT)
            frmUpdateDatGenLotEst.dtpFec_DespachoOriLOT.Enabled = IIf(sEditarExFactory = "S", True, False)
        
            frmUpdateDatGenLotEst.txtPor_ComisionLOT.Text = ssgrdDatos2.Columns("Por_Comision").Text
            'frmUpdateDatGenLotEst.dtpFec_DespachoOriLOT.value = ssgrdDatos2.Columns("Fec_DespachoOri").Text
            frmUpdateDatGenLotEst.txtDes_General.Text = Trim(ssgrdDatos.Columns("Des_General").Text)
            frmUpdateDatGenLotEst.txtPrecioLOT.Text = ssgrdDatos2.Columns("Precio").Text
            frmUpdateDatGenLotEst.txtCod_DivPreLOT.Text = ssgrdDatos2.Columns("Cod_DivPre").Text
            frmUpdateDatGenLotEst.sFlg_Regular = ssgrdDatos.Columns("Flg_Regular").Text
        
            If ssgrdDatos2.Columns("Flg_ComisionEnPorcentaje").Text = "S" Then
                frmUpdateDatGenLotEst.optComisionEnPorcentaje.value = True
                frmUpdateDatGenLotEst.optComisionEnImporte.value = False
            Else
                frmUpdateDatGenLotEst.optComisionEnImporte.value = True
                frmUpdateDatGenLotEst.optComisionEnPorcentaje.value = False
            End If
        
            frmUpdateDatGenLotEst.txtImp_Comision = ssgrdDatos2.Columns("IMP_Comision").Text
        
            If frmUpdateDatGenLotEst.sFlg_Regular = "S" Then
                frmUpdateDatGenLotEst.fraNORegular.Visible = False
            Else
                frmUpdateDatGenLotEst.dtpFec_RecCliLOT.value = ssgrdDatos2.Columns("Fec_RecCli").Text
                frmUpdateDatGenLotEst.txtPrecio_RecCliLOT.Text = ssgrdDatos2.Columns("Precio_RecCli").Text
                frmUpdateDatGenLotEst.fraNORegular.Visible = True
            End If
        
            frmUpdateDatGenLotEst.txtPrecio_Cotizado.Text = DevuelveCampo("select dbo.tg_valida_precio_cotizacion_estilo_cliente('" & sCod_Cliente & "','" & ssgrdDatos.Columns("Cod_TEMCLI").Text & "','" & ssgrdDatos2.Columns("Cod_EStCli").Text & "'," & frmUpdateDatGenLotEst.txtPrecioLOT.Text & ")", cCONNECT)
            frmUpdateDatGenLotEst.txtUtilidadCotizada.Text = DevuelveCampo("select dbo.tg_muestra_utilidad_cotizacion_estilo_cliente('" & sCod_Cliente & "','" & ssgrdDatos.Columns("Cod_TEMCLI").Text & "','" & ssgrdDatos2.Columns("Cod_EStCli").Text & "')", cCONNECT)

            'Pre Pack
            If ssgrdDatos2.Columns("Flg_PrePack").Text = "S" Then
                frmUpdateDatGenLotEst.optPrePackSi = True
            Else
                frmUpdateDatGenLotEst.optPrePackNo = True
            End If
        
            frmUpdateDatGenLotEst.Show vbModal
            Set frmUpdateDatGenLotEst = Nothing
        
            Call BuscarEStilos
        
        Case "ABRIR"

            If sNivAccUsuario = "V" Or ssgrdDatos.Columns("NivAcc").Text = "V" Then
                Mensaje kMESSAGE_ERR_NOT_RIGHT_OPTION

                Exit Sub

            End If
    
            Dim sVerif, sQuery As String

            sQuery = "SELECT dbo.uf_verif_lotest_abierto('" & Me.sCod_Cliente & "','" & ssgrdDatos.Columns("Cod_PurOrd").Text & "','" & ssgrdDatos2.Columns("Cod_LotPurOrd").Text & "','" & ssgrdDatos2.Columns("Cod_EstCli").Text & "')"
            sVerif = DevuelveCampo(sQuery, cCONNECT)

            Select Case sVerif

                Case 4
                    MsgBox "Estilo Abierto. No se puede realizar esta Operación", vbInformation, "Abrir Estilo"

                    Exit Sub

                Case 1
                    MsgBox "Estilo pertenece a mes cerrado. No se puede realizar esta Operación", vbInformation, "Abrir Estilo"

                    Exit Sub

                Case 3
                    MsgBox "Estilo tiene Lotes Generados. No se puede realizar esta Operación", vbInformation, "Abrir Estilo"

                    Exit Sub

                Case -1
                    MsgBox "Ocurrio un error inesperado", vbInformation, "Abrir Estilo"
            End Select

            'Procedimiento para Abrir Estilo
            If Not Abre_LotEst() Then
                MsgBox "Error. No se pudo Abrir Estilo", vbInformation, "Abrir Estilo"

                Exit Sub

            End If

            MsgBox "Estado del Estilo fue cambiado a :" & vbCr & "Abierto", vbInformation, "Abrir Estilo"

        Case "VEROPS"
            Load frmViewOPs
            frmViewOPs.sCod_Cliente = sCod_Cliente
            frmViewOPs.sCod_PurOrd = ssgrdDatos.Columns("Cod_PurOrd").Text
            frmViewOPs.sCod_LotPurOrd = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
            frmViewOPs.sCod_EstCli = ssgrdDatos2.Columns("Cod_EstCli").Text
            Set frmViewOPs.oParent = Me
            frmViewOPs.BuscarOps
            frmViewOPs.Show vbModal
            Set frmViewOPs = Nothing

        Case "CAMBIAPO"
    
            'varCliente = RstBusca("Cliente").value
            'varFabrica = RstBusca("Fabrica").value
            'varPO = DbLista.Columns("Cod_PurOrd").Text
            'varEstCli = DbLista.Columns("Cod_EstCli").Text
    
            Load frmChangePO
            frmChangePO.varCod_Cliente = sCod_Cliente
            frmChangePO.varCod_EstCli = ssgrdDatos2.Columns("Cod_EstCli").Text
            frmChangePO.varCod_LotPurOrd = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
            frmChangePO.varCod_TemCli = ssgrdDatos.Columns("Cod_TemCli").Text
        
            frmChangePO.txtPO = ssgrdDatos.Columns("Cod_PurOrd").Text
            'frmChangePO.txtCliente = frmChangePO.varCod_Cliente & " - " & RstBusca("Cliente").value
            frmChangePO.txtCliente = Trim(Me.txtAbr_Cliente.Text) & " - " & Trim(Me.txtNom_Cliente.Text)
            frmChangePO.txtEstilo = ssgrdDatos2.Columns("Cod_EstCli").Text
        
            frmChangePO.Show 1
            Set frmChangePO = Nothing
            Call BuscarEStilos
        
            'CARGAR_DATA DevuelveCampo("select cod_cliente from tg_cliente where nom_cliente='" & cmbCliente & "'", cCONNECT), DevuelveCampo("select cod_fabrica from tg_fabrica where Nom_Fabrica='" & CmbFabrica & "'", cCONNECT), Year(Me.dtMes.value), Month(dtMes.value), Year(Me.dtMesF.value), Month(dtMesF.value), cboCod_Estcli.Text
            'Call BuscaCampoGrilla(RstBusca, varCliente, varFabrica, varPO, varEstCli)

        Case "VERIFICADESTINOS"
            VerificaDestinos
        
        Case "UPDSINPACK"

            If sNivAccUsuario = "V" Or ssgrdDatos.Columns("NivAcc").Text = "V" Then
                Mensaje kMESSAGE_ERR_NOT_RIGHT_OPTION

                Exit Sub

            End If

            If ssgrdDatos2.Columns("Flg_Status").Text = "C" Then
                Mensaje kMESSAGE_ERR_LOTEST_CLOSED

                Exit Sub

            End If
        
            With frmCambioSinglePrePack

                If ssgrdDatos2.Columns("Flg_PrePack").Text = "S" Then .optPrePackSi = True Else .optPrePackNo = True
                .sCodigoCliente = Me.ssgrdDatos.Columns("COD_CLIENTE").Text
                .sCodigoPurOrd = Me.ssgrdDatos.Columns("COD_PURORD").Text
                .sCodigoLotePurOrd = Me.ssgrdDatos2.Columns("COD_LOTPURORD").Text
                .sCodigoEstadoCliente = Me.ssgrdDatos2.Columns("COD_ESTCLI").Text
                .Show vbModal
            End With

            Call BuscarEStilos
        
    End Select

    Exit Sub

errores:
    ErrorHandler Err, Err.Description
End Sub

Private Sub FunctButt3_ActionClick(ByVal Index As Integer, _
                                   ByVal ActionType As Integer, _
                                   ByVal ActionName As String)

    Select Case ActionName

        Case "AVANENCAJADO"
            Load frmAvanEncajado

            With frmAvanEncajado
                .vCod_Cliente = Me.sCod_Cliente
                .vCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
                .vcod_lotpurord = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                .vcod_estcli = ssgrdDatos2.Columns("Cod_EstCli").Text
                .Label1.Caption = "P.O.: " & Trim(Me.ssgrdDatos.Columns("Cod_PurOrd").Text)
                .Label2.Caption = "Estilo : " & Trim(ssgrdDatos2.Columns("Cod_EstCli").Text)
                .CARGA_GRID
                .Show 1
            End With

    End Select

End Sub

Private Sub FunctButt4_ActionClick(ByVal Index As Integer, _
                                   ByVal ActionType As Integer, _
                                   ByVal ActionName As String)

    If Me.ssgrdDatos2.Rows > 0 Then

        Select Case ActionName

            Case "ASIGNANRODESPACHO"
                Load frmAsignaNroDespacho
                frmAsignaNroDespacho.sCod_Cliente = Me.sCod_Cliente
                frmAsignaNroDespacho.sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
                frmAsignaNroDespacho.sCod_LotPurOrd = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                frmAsignaNroDespacho.sCod_EstCli = ssgrdDatos2.Columns("Cod_EstCli").Text
                frmAsignaNroDespacho.CargaNroDespachoActual
                frmAsignaNroDespacho.Show vbModal
                Set frmAsignaNroDespacho = Nothing

            Case "PO"
                Load frmTG_PurOrdDestinos
                Set frmTG_PurOrdDestinos.oParent = Me
                frmTG_PurOrdDestinos.sFlgOpcion_Nueva = "S"
                frmTG_PurOrdDestinos.sAccionName = "MODIFICAR"
                frmTG_PurOrdDestinos.sModoWizard = "   ESTDAT"
                frmTG_PurOrdDestinos.sCod_Cliente = Me.sCod_Cliente
                frmTG_PurOrdDestinos.sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
                frmTG_PurOrdDestinos.sCod_LotPurOrd = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                frmTG_PurOrdDestinos.sCod_EstCli = ssgrdDatos2.Columns("Cod_EstCli").Text
                frmTG_PurOrdDestinos.sCod_TemCli = ssgrdDatos.Columns("Cod_TemCli").Text
                frmTG_PurOrdDestinos.BUSCAR
                frmTG_PurOrdDestinos.Show vbModal
                Set frmTG_PurOrdDestinos = Nothing

            Case "DATOS ADICIONALES"
                Load frmDatosAdicionales
                Set frmDatosAdicionales.oParent = Me
                frmDatosAdicionales.sFlgOpcion_Nueva = "S"
                frmDatosAdicionales.sCod_Cliente = Me.sCod_Cliente
                frmDatosAdicionales.sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
                frmDatosAdicionales.sCod_LotPurOrd = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                frmDatosAdicionales.sCod_EstCli = ssgrdDatos2.Columns("Cod_EstCli").Text
                frmDatosAdicionales.sCod_TemCli = ssgrdDatos.Columns("Cod_TemCli").Text
                frmDatosAdicionales.BUSCAR
                frmDatosAdicionales.Show vbModal
                Set frmDatosAdicionales = Nothing

            Case "COMENTARIO"
                'Load FrmComentario
                Set FrmComentario.oParent = Me
                'FrmComentario.sFlgOpcion_Nueva = "S"
                FrmComentario.sCod_Cliente = Me.sCod_Cliente
                FrmComentario.sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
                FrmComentario.sCod_LotPurOrd = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                FrmComentario.sCod_EstCli = ssgrdDatos2.Columns("Cod_EstCli").Text
                'FrmComentario.Buscar
                FrmComentario.Cargar_Data
                FrmComentario.Show vbModal
                Set FrmComentario = Nothing

            Case "POASOCIADAS"
                Load frmPoAsociadas
                'frmPoAsociadas.Caption = "PO Asociadas " + "Cliente:" + Me.ssgrdDatos.Columns("Nom_cliente").Text + " PO:" + Me.ssgrdDatos.Columns("Cod_PurOrd").Text + " Lote:" + ssgrdDatos2.Columns("Cod_LotPurOrd").Text + " Numero:" + ssgrdDatos2.Columns("Cod_EstCli").Text
                frmPoAsociadas.Caption = "PO Asociadas " + "Cliente:" + Me.txtNom_Cliente.Text + "  PO:" + Me.ssgrdDatos.Columns("Cod_PurOrd").Text + " Lote:" + ssgrdDatos2.Columns("Cod_LotPurOrd").Text + " Numero:" + ssgrdDatos2.Columns("Cod_EstCli").Text
                frmPoAsociadas.COD_CLIENTE = Me.ssgrdDatos.Columns("Cod_cliente").Text
                frmPoAsociadas.cod_purord = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
                frmPoAsociadas.cod_lotpurord = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                frmPoAsociadas.cod_estcli = ssgrdDatos2.Columns("Cod_EstCli").Text
                frmPoAsociadas.BUSCAR
                frmPoAsociadas.Show vbModal
                Set frmPoAsociadas = Nothing

            Case "IMPRIMIRPOASO"
                mReporte

            Case "CAMBFECFINPROD"

                If ssgrdDatos.Rows > 0 Then
                    Load frmCambioFecFinProduccion
                    frmCambioFecFinProduccion.COD_CLIENTE = sCod_Cliente
                    frmCambioFecFinProduccion.cod_purord = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
                    frmCambioFecFinProduccion.cod_lotpurord = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                    frmCambioFecFinProduccion.cod_estcli = ssgrdDatos2.Columns("Cod_EstCli").Text
                    
                    frmCambioFecFinProduccion.txtCliente.Text = Trim(Me.txtAbr_Cliente.Text) & " - " & Trim(Me.txtNom_Cliente.Text)
                    frmCambioFecFinProduccion.txtPO.Text = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
                    frmCambioFecFinProduccion.txtEstilo.Text = Me.ssgrdDatos2.Columns("Cod_EstCli").Text
                    frmCambioFecFinProduccion.dtpFecFinProd.value = ssgrdDatos2.Columns("Fec_DespachoAct").Text
                    frmCambioFecFinProduccion.Show 1
                    Set frmCambioFecFinProduccion = Nothing
                    Call BUSCAR
                End If

            Case "DATOSFINAN"

                Dim svpo     As String

                Dim svcommit As String

                Dim strSql   As String

                Load frmDatosFinanzas
                frmDatosFinanzas.varCod_Cliente = sCod_Cliente
                frmDatosFinanzas.varCod_EstCli = ssgrdDatos2.Columns("Cod_EstCli").Text
                frmDatosFinanzas.varCod_LotPurOrd = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                frmDatosFinanzas.varCod_TemCli = ssgrdDatos.Columns("Cod_TemCli").Text
                Set frmDatosFinanzas.oParent = Me
                        
                '                Strsql = "SELECT FLG_IFINANCIERA_PO FROM TG_PURORD WHERE COD_CLIENTE = '" & sCod_Cliente & "' AND COD_PURORD = '" & Me.ssgrdDatos.Columns("Cod_PurOrd").Text & "'"
                '                svpo = DevuelveCampo(Strsql, cCONNECT)
                '
                '                If svpo = "N" Then
                '                    frmDatosFinanzas.chkop.value = 0
                '                Else
                '                    frmDatosFinanzas.chkop.value = 1
                '                End If
                '
                '                Strsql = "SELECT FLG_IFINANCIERA_COMMIT FROM TG_PURORD WHERE COD_CLIENTE = '" & sCod_Cliente & "' AND COD_PURORD = '" & Me.ssgrdDatos.Columns("Cod_PurOrd").Text & "'"
                '                svcommit = DevuelveCampo(Strsql, cCONNECT)
                '
                '                If svcommit = "N" Then
                '                    frmDatosFinanzas.chkcommit.value = 0
                '                Else
                '                    frmDatosFinanzas.chkcommit.value = 1
                '                End If

                If ssgrdDatos.Columns(47).Text = "N" Then
                    frmDatosFinanzas.chkcommit.value = 0
                Else
                    frmDatosFinanzas.chkcommit.value = 1
                End If
                
                If ssgrdDatos.Columns(46).Text = "N" Then
                    frmDatosFinanzas.chkop.value = 0
                Else
                    frmDatosFinanzas.chkop.value = 1
                End If

                frmDatosFinanzas.txtPO = ssgrdDatos.Columns("Cod_PurOrd").Text
                frmDatosFinanzas.txtCliente = Trim(Me.txtAbr_Cliente.Text) & " - " & Trim(Me.txtNom_Cliente.Text)
                frmDatosFinanzas.txtEstilo = ssgrdDatos2.Columns("Cod_EstCli").Text
                
                frmDatosFinanzas.Show 1
                Set frmDatosFinanzas = Nothing
                
            Case "EXFACTORY"

                If ssgrdDatos.Rows > 0 Then
                    Load FrmExFactory
                    FrmExFactory.COD_CLIENTE = sCod_Cliente
                    FrmExFactory.cod_purord = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
                    FrmExFactory.cod_lotpurord = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                    FrmExFactory.cod_estcli = ssgrdDatos2.Columns("Cod_EstCli").Text
                    FrmExFactory.dtpFec.value = ssgrdDatos2.Columns("Fec_DespachoOri_Reprogramada").Text
                    FrmExFactory.Show 1
                    Set FrmExFactory = Nothing
                    Call BUSCAR
                End If
                
            Case "EFAFAC"

                If ssgrdDatos.Rows = 0 Then Exit Sub
                Load frmExFactoryFacturacion

                With frmExFactoryFacturacion
                    .COD_CLIENTE = sCod_Cliente
                    .PO = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
                    .LOTE_PO = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                    .ESTILO_CLIENTE = ssgrdDatos2.Columns("Cod_EstCli").Text
                    .dtpFec.value = ssgrdDatos2.Columns("Fec_ExFactory_Ajustada_Facturacion").Text
                    .Show 1
                End With

                Set frmExFactoryFacturacion = Nothing
                Call BUSCAR
        End Select

    End If

End Sub

Private Sub FunctButt5_ActionClick(ByVal Index As Integer, _
                                   ByVal ActionType As Integer, _
                                   ByVal ActionName As String)

    If Me.ssgrdDatos2.Rows > 0 Then

        Select Case ActionName

            Case "FACTURAS"
                Load frmImprimirFacturasPacking
                frmImprimirFacturasPacking.Caption = "Facturas"
                frmImprimirFacturasPacking.sStore = "CN_VENTAS_MUESTRA_FACTURAS_DESDE_LOTEST"
                frmImprimirFacturasPacking.sCod_Cliente = Me.sCod_Cliente
                frmImprimirFacturasPacking.sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
                frmImprimirFacturasPacking.sCod_LotPurOrd = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                frmImprimirFacturasPacking.sCod_EstCli = ssgrdDatos2.Columns("Cod_EstCli").Text
                frmImprimirFacturasPacking.CARGAGRILLA
                frmImprimirFacturasPacking.Show vbModal
                Set frmImprimirFacturasPacking = Nothing

            Case "PACKINGLIST"
                Load frmImprimirFacturasPacking
                Set frmTG_PurOrdDestinos.oParent = Me
                frmImprimirFacturasPacking.Caption = "Packing List"
                frmImprimirFacturasPacking.sStore = "CN_VENTAS_MUESTRA_PACKING_DESDE_LOTEST"
                frmImprimirFacturasPacking.sCod_Cliente = Me.sCod_Cliente
                frmImprimirFacturasPacking.sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
                frmImprimirFacturasPacking.sCod_LotPurOrd = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                frmImprimirFacturasPacking.sCod_EstCli = ssgrdDatos2.Columns("Cod_EstCli").Text
                frmImprimirFacturasPacking.CARGAGRILLA
                frmImprimirFacturasPacking.Show vbModal
                Set frmImprimirFacturasPacking = Nothing

            Case "MODIFCANT"
                Load FrmModificar_Cantidades
                FrmModificar_Cantidades.sCod_Cliente = sCod_Cliente
                FrmModificar_Cantidades.sCod_PurOrd = ssgrdDatos.Columns("Cod_PurOrd").Text
                FrmModificar_Cantidades.sCod_LotPurOrd = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                FrmModificar_Cantidades.sCod_EstCli = ssgrdDatos2.Columns("Cod_EstCli").Text
                FrmModificar_Cantidades.Cargar
                FrmModificar_Cantidades.Show 1
                Set FrmModificar_Cantidades = Nothing

            Case "VERFINANCIAMIENTO"
                Call Reporte_VERFINANCIAMIENTO

            Case "REVPRECIO"

                If ssgrdDatos2.Rows = 0 Then Exit Sub
                Load Frm_LotestAdicional
                Frm_LotestAdicional.sCod_Cliente = sCod_Cliente
                Frm_LotestAdicional.sCod_PurOrd = ssgrdDatos.Columns("Cod_PurOrd").Text
                Frm_LotestAdicional.sCod_LotPurOrd = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                Frm_LotestAdicional.sCod_EstCli = ssgrdDatos2.Columns("Cod_EstCli").Text
                Frm_LotestAdicional.Cargar
                Frm_LotestAdicional.Show 1
                Set Frm_LotestAdicional = Nothing
                
            Case "GRAFICO"

                If ssgrdDatos2.Rows = 0 Then Exit Sub
                Load frmShowTG_PurOrd_EstCli
                'cCONNECT = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=HIALPESA;Data Source=HIALPESA4"
                frmShowTG_PurOrd_EstCli.strCodCliente = sCod_Cliente
                frmShowTG_PurOrd_EstCli.txtPO = ssgrdDatos.Columns("Cod_PurOrd").Text
                frmShowTG_PurOrd_EstCli.CARGA_GRID_ESTILO_CLIENTE
                frmShowTG_PurOrd_EstCli.Show 1
                Set frmShowTG_PurOrd_EstCli = Nothing
        End Select

    End If

End Sub

Sub Reporte_VERFINANCIAMIENTO()

    On Error GoTo Fin

    Dim oo     As Object, vRutaLogo As String

    Dim strSql As String

    Dim oRs    As ADODB.Recordset
      
    strSql = "EXEC TG_SM_PURORD_LOTEST_FINANCIAMIENTO '" & sCod_Cliente & "','" & txtCod_PurOrd.Text & "'"
    Set oRs = CargarRecordSetDesconectado(strSql, cCONNECT)
    
    If oRs.RecordCount = 0 Then
        MsgBox "No Existen Registros"

        Exit Sub

    End If
    
    strSql = "SELECT Ruta_Logo From SEGURIDAD..SEG_EMPRESAS " & "WHERE Cod_Empresa = '" & vemp & "'"
    vRutaLogo = DevuelveCampo(strSql, cCONNECT)
    
    Screen.MousePointer = 11
    
    Set oo = CreateObject("excel.application")
    
    'oo.Workbooks.Open vRuta & "\rpt_MuestraLotEstFinanciamiento.xlt"
    oo.Workbooks.Open "\\hlpserver01\Desarrollo\HIALPESA\Pedidos" & "\rpt_MuestraLotEstFinanciamiento.xlt"
    'oo.Workbooks.Open "C:\Program Files\Gestion de Pedidos" & "\rpt_MuestraLotEstFinanciamiento.xlt"
    oo.DisplayAlerts = False
    oo.Visible = True
    oo.Run "REPORTE", vRutaLogo, oRs
    'oo.Run "REPORTE", vRutaLogo, oRs
    Screen.MousePointer = vbNormal
    Set oo = Nothing

    Exit Sub

Fin:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
End Sub

Private Sub optCod_EstCli_Click()
    Me.txtCod_PurOrd.Text = ""
    Me.txtCod_TemCli.Text = ""
    Me.txtNom_TemCli.Text = ""
    Me.txtCod_Ordpro.Text = ""
    Me.txtDes_estpro.Text = ""
    
    Me.fraEstCli.Visible = True
    Me.fraOP.Visible = False
    Me.fraPurOrd.Visible = False
    Me.fraTemporada.Visible = False
    
End Sub

Private Sub optCod_OrdPro_Click()

    Me.txtCod_EstCli.Text = ""
    Me.txtCod_TemCli.Text = ""
    Me.txtNom_TemCli.Text = ""
    Me.txtCod_PurOrd.Text = ""

    Me.fraEstCli.Visible = False
    Me.fraOP.Visible = True
    Me.fraPurOrd.Visible = False
    Me.fraTemporada.Visible = False
End Sub

Private Sub optCod_PurOrd_Click()
    Me.txtCod_EstCli.Text = ""
    Me.txtCod_TemCli.Text = ""
    Me.txtNom_TemCli.Text = ""
    Me.txtCod_Ordpro.Text = ""
    Me.txtDes_estpro.Text = ""
    
    Me.fraEstCli.Visible = False
    Me.fraOP.Visible = False
    Me.fraPurOrd.Visible = True
    Me.fraTemporada.Visible = False
    
End Sub

Private Sub optCod_TemCli_Click()
    Me.txtCod_PurOrd.Text = ""
    Me.txtCod_EstCli.Text = ""
    Me.txtCod_Ordpro.Text = ""
    Me.txtDes_estpro.Text = ""
    
    Me.fraEstCli.Visible = False
    Me.fraOP.Visible = False
    Me.fraPurOrd.Visible = False
    Me.fraTemporada.Visible = True
    Me.txtCod_TemCli.SetFocus
End Sub

Private Sub ssgrdDatos_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)

    If Me.ssgrdDatos.Rows > 0 Then
        'If Val(LastRow) <> Me.ssgrdDatos.Row Then
        BuscarEStilos
    End If

End Sub

Private Sub txtAbr_Cliente_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        'FunctButt1.FunctionsUser =
        sFlag = "ABR_CLIENTE"

        If Filtrar(sFlag, Me, txtAbr_Cliente, txtNom_Cliente) Then
            optCod_PurOrd.value = True
            Me.txtCod_PurOrd.SetFocus
            'Me.txtCod_TemCli.SetFocus
        End If
    End If

End Sub

Private Sub txtCod_EstCli_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.optCod_EstCli.value = True
    
    If KeyCode = vbKeyReturn Then
        BUSCAR
    End If

End Sub

Private Sub txtCod_Ordpro_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.optCod_OrdPro.value = True
End Sub

Private Sub txtCod_Ordpro_KeyPress(KeyAscii As Integer)

    Dim varCod_EstPro As String

    Dim strSql        As String

    If KeyAscii = 13 Then
        txtCod_Ordpro.Text = Right("00000" & txtCod_Ordpro.Text, 5)
        strSql = "SELECT DISTINCT(Cod_Estpro) FROM tg_lotestpro where cod_Cliente = '" & Me.sCod_Cliente & "' and Cod_OrdPro = '" & Me.txtCod_Ordpro.Text & "'"
        varCod_EstPro = DevuelveCampo(strSql, cCONNECT)
        strSql = "SELECT Des_estpro FROM ES_ESTPRO where Cod_EstPro = '" & varCod_EstPro & "'"
        Me.txtDes_estpro.Text = DevuelveCampo(strSql, cCONNECT)
        BUSCAR
    End If

End Sub

Private Sub txtCod_Ordpro_LostFocus()

    Dim varCod_EstPro As String

    Dim strSql        As String

    txtCod_Ordpro.Text = Right("00000" & txtCod_Ordpro.Text, 5)
    strSql = "SELECT DISTINCT(Cod_Estpro) FROM tg_lotestpro where cod_Cliente = '" & Me.sCod_Cliente & "' and Cod_OrdPro = '" & Me.txtCod_Ordpro.Text & "'"
    varCod_EstPro = DevuelveCampo(strSql, cCONNECT)
    strSql = "SELECT Des_estpro FROM ES_ESTPRO where Cod_EstPro = '" & varCod_EstPro & "'"
    Me.txtDes_estpro.Text = DevuelveCampo(strSql, cCONNECT)
End Sub

Private Sub txtCod_PurOrd_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim oMsg As clsMessages

    Me.optCod_PurOrd.value = True
        
    If KeyCode = vbKeyReturn Then
        If RTrim(DevuelveCampo("SELECT ISNULL(FLG_CREAMV,'N') FROM TG_CONTROL ", cCONNECT)) = "S" Then
            If RTrim(txtCod_PurOrd) = "" Then
                txtCod_PurOrd = "MV_" & Mid(ComputerName, 1, 5) & "_" & Format(Now, "DDMMYY") & Format(Now, "HHMM")
                Set oMsg = New clsMessages
                oMsg.Codigo = MESSAGECODE.kMESSAGE_ASK_NUEVO_PURORD
                oMsg.OptionalText = "MUESTRA DE VENTA : " & txtCod_PurOrd
                
                If oMsg.ShowMesage(iLanguage) Then
                    
                    If Wizard(False) Then
                        BUSCAR
                    End If
                End If

            Else

                If Not BUSCAR Then
                    If Wizard(True) Then
                        BUSCAR
                    End If
                End If
            End If

        Else

            If Not BUSCAR Then
                If Wizard(True) Then
                    BUSCAR
                End If
            End If
        End If
    End If

End Sub

Private Sub txtCod_TemCli_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.optCod_TemCli.value = True

    If KeyCode = vbKeyReturn Then
        sFlag = "COD_TEMCLI"

        If Filtrar(sFlag, Me, txtCod_TemCli, txtNom_TemCli) Then
            BUSCAR
        End If
    End If

End Sub

Public Sub BuscarEStilos()

    Dim obj   As clsTG_PurOrd

    Dim vbuff As Variant

    Dim irow  As Variant

    irow = Me.ssgrdDatos2.Bookmark
    iRowsGrilla = ssgrdDatos2.Rows
    Me.ssgrdDatos2.Redraw = False
        
    SSDBGridSetGrid Me.ssgrdDatos2
        
    Set obj = New clsTG_PurOrd
    obj.ConexionString = cCONNECT
    vbuff = obj.ViewEstilos(sCod_Cliente, Me.ssgrdDatos.Columns("Cod_PurOrd").Text, vusu)
        
    LibraryVBToSSDBGrid obj, vbuff, ssgrdDatos2
    ssgrdDatos2.ActiveRowStyleSet = "RowActive"
    ssgrdDatos2.SelectTypeRow = ssSelectionTypeMultiSelectRange
    Set obj = Nothing
    Me.ssgrdDatos2.SplitterPos = 1
    Me.ssgrdDatos2.SplitterVisible = True
    RestoreRowSSDBGrid Me.ssgrdDatos2, irow, iRowsGrilla
    Me.ssgrdDatos2.Redraw = True
       
    Exit Sub

errores:
    Me.MousePointer = vbDefault

    If Not obj Is Nothing Then
        Set obj = Nothing
    End If

    ErrorHandler Err, Err.Description

End Sub

Private Function Wizard(ByVal bQuestion As Boolean) As Boolean

    On Error GoTo errores

    Dim oMensaje As clsMessages

    Dim oWizard  As frmWizard

    If optCod_PurOrd.value = True Then
        If RTrim(Me.txtCod_PurOrd.Text) = "" Then
            Mensaje kMESSAGE_ERR_NOTEMPTY

            If Me.txtCod_PurOrd.Enabled Then
                Me.txtCod_PurOrd.SetFocus
            End If

            Exit Function

        End If
    End If
    
    If sNivAccUsuario = "V" Then
        Mensaje kMESSAGE_ERR_NOT_RIGHT_OPTION

        Exit Function

    End If
    
    If bQuestion Then
        Set oMensaje = New clsMessages
        oMensaje.Codigo = MESSAGECODE.kMESSAGE_ASK_NUEVO_PURORD
           
        If Not oMensaje.ShowMesage(iLanguage) Then

            Exit Function

        End If
    End If
    
    sCod_PurOrd = Me.txtCod_PurOrd.Text
        
    Wizard = True
    
    Set oWizard = New frmWizard
    Load oWizard
    oWizard.sAccionName = "ADICIONAR"
    oWizard.dPor_ComisionCliente = dPor_ComisionCliente
    oWizard.sModoWizard = "POCESTDAT"
    oWizard.sCod_Cliente = Me.sCod_Cliente
    oWizard.sCod_PurOrd = Me.sCod_PurOrd
    Set oWizard.oParent = Me
    'oWizard
    oWizard.Show vbModal
    Set oWizard = Nothing
    Wizard = True
    
    Exit Function

errores:
    ErrorHandler Err, Err.Description
End Function

Private Function DeleteLotPurOrd() As Boolean

    On Error GoTo errores

    Dim vbuff

    Dim objPO          As clsTG_LotColTal

    Dim sCod_LotPurOrd As String

    Dim oMensaje       As clsMessages
    
    Dim oWizard        As frmWizard

    Set oMensaje = New clsMessages
    oMensaje.Codigo = MESSAGECODE.kMESSAGE_ASK_DELETE_LOTEST
    
    If Not oMensaje.ShowMesage(iLanguage) Then

        Exit Function

    End If
    
    sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
    sCod_LotPurOrd = Me.ssgrdDatos2.Columns("Cod_LotPurOrd").Text
                
    Set objPO = New clsTG_LotColTal
    objPO.ConexionString = cCONNECT
        
    objPO.DeleteLotPurOrd sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd
    Set objPO = Nothing
    
    BUSCAR

    Exit Function

errores:

    If Not objPO Is Nothing Then
        Set objPO = Nothing
    End If

    ErrorHandler Err, "DeleteLotPurOrd"
End Function

Function CopiaPO() As Boolean

    Dim cn As New ADODB.Connection

    CopiaPO = True

    On Error GoTo CopiaPOErr

    cn.ConnectionString = cCONNECT
    cn.Open
    cn.Execute "SG_CopiaPO '" & Me.sCod_Cliente & "','" & Me.ssgrdDatos.Columns("Cod_PurOrd").Text & "','" & sPONew & "'"

    Exit Function

CopiaPOErr:
    ErrorHandler Err, "CopiaPO"
    Set cn = Nothing
    CopiaPO = False
End Function

Function Abre_LotEst() As Boolean

    Dim cn As New ADODB.Connection

    Abre_LotEst = True

    On Error GoTo Abre_LotEstErr

    cn.ConnectionString = cCONNECT
    cn.Open
    cn.Execute "SG_Abrir_LotEst '" & Me.sCod_Cliente & "','" & Me.ssgrdDatos.Columns("Cod_PurOrd").Text & "','" & ssgrdDatos2.Columns("Cod_LotPurOrd").Text & "','" & ssgrdDatos2.Columns("Cod_EstCli").Text & "'"

    Exit Function

Abre_LotEstErr:
    ErrorHandler Err, "Abre_LotEst"
    Set cn = Nothing
    Abre_LotEst = False
End Function

Function VerificaDestinos() As Boolean

    Dim cn As New ADODB.Connection

    VerificaDestinos = True

    On Error GoTo VerificaDestinosErr

    Dim rsData As ADODB.Recordset

    Set rsData = GetDataSet(cCONNECT, "SG_VERIFICA_LotEst_DESTINOS '" & Me.sCod_Cliente & "','" & Me.ssgrdDatos.Columns("Cod_PurOrd").Text & "','" & ssgrdDatos2.Columns("Cod_LotPurOrd").Text & "','" & ssgrdDatos2.Columns("Cod_EstCli").Text & "' , '" & ssgrdDatos.Columns("Cod_TEMCLI").Text & "'")
            
    If Not rsData.EOF Then
        If rsData(0).value = "0" Then
            Load frmVerificaMatrizDetalle
            Set frmVerificaMatrizDetalle.oParent = Me
            frmVerificaMatrizDetalle.sCod_Cliente = sCod_Cliente
            frmVerificaMatrizDetalle.sCod_PurOrd = Me.ssgrdDatos.Columns("Cod_PurOrd").Text
            frmVerificaMatrizDetalle.sCod_LotPurOrd = ssgrdDatos2.Columns("Cod_LotPurOrd").Text
            frmVerificaMatrizDetalle.sCod_TemCli = ssgrdDatos2.Columns("Cod_EstCli").Text
            Set frmVerificaMatrizDetalle.rsData = rsData
        
            frmVerificaMatrizDetalle.BUSCAR
            frmVerificaMatrizDetalle.Show vbModal
            Set frmVerificaMatrizDetalle = Nothing
        Else
            BuscarEStilos
            Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
        End If
    End If

    Exit Function

VerificaDestinosErr:
    ErrorHandler Err, "Abre_LotEst"
    Set cn = Nothing
    VerificaDestinos = False
End Function

Public Sub SetFormCliente(ByVal sNivAccUsuario As String)
    'If sNivAccUsuario = "V" Then
    '    FunctButt1.ChangeProperty "ENABLED", 1, False
    '    FunctButt1.ChangeProperty "ENABLED", 2, False
    '
    '    FunctButt2.ChangeProperty "ENABLED", 0, False
    '    FunctButt2.ChangeProperty "ENABLED", 1, False
    '    FunctButt2.ChangeProperty "ENABLED", 2, False
    '    FunctButt2.ChangeProperty "ENABLED", 3, False
    '    FunctButt2.ChangeProperty "ENABLED", 4, False
    '    FunctButt2.ChangeProperty "ENABLED", 9, False
    'End If

End Sub

Sub mReporte()

    On Error GoTo ErrorImpresion

    Dim oo       As Object

    Dim sEmpresa As String

    sEmpresa = DevuelveCampo("SELECT DES_EMPRESA FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA='" & vemp & "'", cCONNECT)
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open App.Path & "\RptPosAsociadas.XLT"
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "REPORTE", sCod_Cliente, txtNom_Cliente, txtCod_TemCli, txtNom_TemCli, cCONNECT, sEmpresa
    Set oo = Nothing

    Exit Sub

ErrorImpresion:
    Set oo = Nothing
    ErrorHandler Err, "Reporte"
End Sub
