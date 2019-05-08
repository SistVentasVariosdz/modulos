VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmAdicionaDocumVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adiciona Documento Venta"
   ClientHeight    =   8655
   ClientLeft      =   405
   ClientTop       =   690
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   8385
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   5640
      TabIndex        =   41
      Top             =   8040
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmAdicionaDocumVentas.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame frExportacion 
      BorderStyle     =   0  'None
      Height          =   2100
      Left            =   120
      TabIndex        =   67
      Top             =   5400
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox txtcajas 
         Height          =   285
         Left            =   4560
         MaxLength       =   25
         TabIndex        =   80
         Top             =   1800
         Width           =   1080
      End
      Begin VB.CheckBox chkFlete 
         Alignment       =   1  'Right Justify
         Caption         =   " Incluyen &FLETE"
         Height          =   255
         Left            =   5940
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtDes_Destino 
         Height          =   285
         Left            =   2880
         TabIndex        =   23
         Top             =   0
         Width           =   2385
      End
      Begin VB.TextBox txtCod_Destino 
         Height          =   285
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   22
         Top             =   0
         Width           =   600
      End
      Begin VB.TextBox txtDua 
         Height          =   285
         Left            =   2160
         MaxLength       =   25
         TabIndex        =   31
         Top             =   1080
         Width           =   3120
      End
      Begin VB.TextBox txtEmbarque_Cod 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4635
         MaxLength       =   4
         TabIndex        =   29
         Top             =   720
         Width           =   600
      End
      Begin VB.TextBox txtEmbarque_Des 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5400
         TabIndex        =   30
         Top             =   735
         Width           =   2265
      End
      Begin VB.CheckBox chkSeguro 
         Alignment       =   1  'Right Justify
         Caption         =   " Incluyen &Seguro"
         Height          =   255
         Left            =   5940
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1695
      End
      Begin NumBoxProject.NumBox Imp_Flete 
         Height          =   285
         Left            =   2160
         TabIndex        =   28
         Tag             =   "SET/VALID"
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         TypeVal         =   2
         Mask            =   "9,999,999,999.99"
         Formato         =   "#,###,###,###.##"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   3
         Text            =   "0.00"
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   2
      End
      Begin NumBoxProject.NumBox txtPeso_Bruto 
         Height          =   285
         Left            =   2160
         TabIndex        =   25
         Tag             =   "SET/VALID"
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         TypeVal         =   2
         Mask            =   "9,999,999,999.99"
         Formato         =   "#,###,###,###.##"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   3
         Text            =   "0.00"
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   2
      End
      Begin NumBoxProject.NumBox txtPeso_Neto 
         Height          =   285
         Left            =   4260
         TabIndex        =   26
         Tag             =   "SET/VALID"
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         TypeVal         =   2
         Mask            =   "9,999,999,999.99"
         Formato         =   "#,###,###,###.##"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   3
         Text            =   "0.00"
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   2
      End
      Begin NumBoxProject.NumBox txtShip_Date 
         Height          =   285
         Left            =   6240
         TabIndex        =   24
         Top             =   0
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         TypeVal         =   3
         Mask            =   "99/99/9999"
         Formato         =   "dd/MM/yyyy"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   2
         Text            =   ""
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   0
      End
      Begin NumBoxProject.NumBox txtImp_Seguro 
         Height          =   285
         Left            =   6240
         TabIndex        =   27
         Tag             =   "SET/VALID"
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         TypeVal         =   2
         Mask            =   "9,999,999,999.99"
         Formato         =   "#,###,###,###.##"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   3
         Text            =   "0.00"
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   2
      End
      Begin NumBoxProject.NumBox txtFec_Numeracion 
         Height          =   285
         Left            =   2160
         TabIndex        =   32
         Top             =   1440
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         TypeVal         =   3
         Mask            =   "99/99/9999"
         Formato         =   "dd/MM/yyyy"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   2
         Text            =   ""
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   0
      End
      Begin NumBoxProject.NumBox txtFec_Embarque 
         Height          =   285
         Left            =   4560
         TabIndex        =   33
         Top             =   1440
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         TypeVal         =   3
         Mask            =   "99/99/9999"
         Formato         =   "dd/MM/yyyy"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   2
         Text            =   ""
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   0
      End
      Begin NumBoxProject.NumBox txtImp_FOB_Dol_Dua 
         Height          =   285
         Left            =   2160
         TabIndex        =   78
         Tag             =   "SET/VALID"
         Top             =   1785
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         TypeVal         =   2
         Mask            =   "9,999,999,999.99"
         Formato         =   "#,###,###,###.##"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   3
         Text            =   "0.00"
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   2
      End
      Begin VB.Label Label31 
         Caption         =   "Nº de Cajas :"
         Height          =   255
         Left            =   3480
         TabIndex        =   81
         Top             =   1815
         Width           =   975
      End
      Begin VB.Label Label30 
         Caption         =   "Importe FOB DUA $:"
         Height          =   255
         Left            =   15
         TabIndex        =   79
         Top             =   1785
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Importe FLETE :"
         Height          =   255
         Left            =   0
         TabIndex        =   77
         Top             =   735
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Destino :"
         Height          =   255
         Left            =   0
         TabIndex        =   76
         Top             =   15
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Peso Bruto :"
         Height          =   255
         Left            =   0
         TabIndex        =   75
         Top             =   375
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Peso Neto :"
         Height          =   255
         Left            =   3360
         TabIndex        =   74
         Top             =   375
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Ship Date :"
         Height          =   255
         Left            =   5400
         TabIndex        =   73
         Top             =   15
         Width           =   855
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Seguro :"
         Height          =   195
         Left            =   5520
         TabIndex        =   72
         Top             =   375
         Width           =   600
      End
      Begin VB.Label Label23 
         Caption         =   "Dua :"
         Height          =   255
         Left            =   0
         TabIndex        =   71
         Top             =   1095
         Width           =   495
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Fec. Numeracion :"
         Height          =   195
         Left            =   0
         TabIndex        =   70
         Top             =   1455
         Width           =   1305
      End
      Begin VB.Label Label25 
         Caption         =   "Fec  Embarque:"
         Height          =   255
         Left            =   3360
         TabIndex        =   69
         Top             =   1455
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "Tipo Embarque :"
         Height          =   255
         Left            =   3360
         TabIndex        =   68
         Top             =   735
         Width           =   1215
      End
   End
   Begin VB.Frame frMain 
      Height          =   7800
      Left            =   120
      TabIndex        =   42
      Top             =   0
      Width           =   8175
      Begin VB.TextBox txtTipoFac 
         Height          =   285
         Left            =   2280
         TabIndex        =   83
         Top             =   7200
         Width           =   1095
      End
      Begin VB.CheckBox chkExonerado 
         Alignment       =   1  'Right Justify
         Caption         =   "Exonerado de IGV"
         Height          =   255
         Left            =   5970
         TabIndex        =   82
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtDes_TipVenta 
         Height          =   285
         Left            =   3000
         TabIndex        =   3
         Top             =   720
         Width           =   2505
      End
      Begin VB.TextBox txtCod_TipVenta 
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   2
         Top             =   720
         Width           =   600
      End
      Begin VB.CheckBox chkDetraccion 
         Alignment       =   1  'Right Justify
         Caption         =   "Detraccion"
         Height          =   255
         Left            =   5970
         TabIndex        =   66
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtObservacion 
         Height          =   285
         Left            =   2280
         TabIndex        =   21
         Top             =   4680
         Width           =   5505
      End
      Begin VB.CheckBox chkExportacion 
         Alignment       =   1  'Right Justify
         Caption         =   "Exportacion"
         Height          =   255
         Left            =   5970
         TabIndex        =   6
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtGlosa 
         Height          =   645
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   4005
         Width           =   5505
      End
      Begin VB.TextBox txtSer_Docum 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtCod_TipDoc 
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   0
         Top             =   360
         Width           =   600
      End
      Begin VB.TextBox txtDes_TipDoc 
         Height          =   285
         Left            =   3000
         TabIndex        =   1
         Top             =   360
         Width           =   2505
      End
      Begin VB.TextBox txtNro_DocInter 
         Height          =   285
         Left            =   2280
         TabIndex        =   19
         Top             =   3690
         Width           =   5505
      End
      Begin VB.TextBox txtNum_Docum 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4320
         MaxLength       =   8
         TabIndex        =   5
         Top             =   1080
         Width           =   1200
      End
      Begin VB.TextBox txtNum_Ruc 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6240
         MaxLength       =   11
         TabIndex        =   9
         Top             =   1470
         Width           =   1545
      End
      Begin VB.TextBox txtNro_Ordener 
         Height          =   285
         Left            =   2280
         TabIndex        =   18
         Top             =   3330
         Width           =   5505
      End
      Begin VB.TextBox txtNro_Guias 
         Height          =   285
         Left            =   2280
         TabIndex        =   17
         Top             =   2970
         Width           =   5505
      End
      Begin VB.TextBox txtDes_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2685
         TabIndex        =   8
         Top             =   1470
         Width           =   2865
      End
      Begin VB.TextBox txtCod_Moneda 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   13
         Top             =   2190
         Width           =   600
      End
      Begin VB.TextBox txtDes_Moneda 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3000
         TabIndex        =   14
         Top             =   2205
         Width           =   3345
      End
      Begin VB.TextBox txtCod_ConPag 
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   15
         Top             =   2610
         Width           =   600
      End
      Begin VB.TextBox txtDes_ConPag 
         Height          =   285
         Left            =   3000
         TabIndex        =   16
         Top             =   2610
         Width           =   3345
      End
      Begin NumBoxProject.NumBox inpFec_EmiDoc 
         Height          =   285
         Left            =   2280
         TabIndex        =   10
         Top             =   1830
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         TypeVal         =   3
         Mask            =   "99/99/9999"
         Formato         =   "dd/MM/yyyy"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   2
         Text            =   ""
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   0
      End
      Begin NumBoxProject.NumBox InpFec_RegDoc 
         Height          =   285
         Left            =   4920
         TabIndex        =   11
         Top             =   1830
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         TypeVal         =   3
         Mask            =   "99/99/9999"
         Formato         =   "dd/MM/yyyy"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   2
         Text            =   ""
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   0
      End
      Begin NumBoxProject.NumBox TxtTipo_Cambio 
         Height          =   285
         Left            =   6810
         TabIndex        =   12
         Tag             =   "SET/VALID"
         Top             =   1830
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         TypeVal         =   2
         Mask            =   "9,999,999,999.999"
         Formato         =   "#,###,###,###.###"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   3
         Text            =   "0.000"
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   3
      End
      Begin VB.TextBox txtCod_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   7
         Text            =   "C"
         Top             =   1470
         Width           =   360
      End
      Begin VB.TextBox txtDes_TipAnex 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3000
         MaxLength       =   11
         TabIndex        =   64
         Top             =   1470
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Frame frReferencia 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         TabIndex        =   61
         Top             =   5040
         Visible         =   0   'False
         Width           =   7815
         Begin VB.TextBox txtDes_NotaAbono 
            Height          =   285
            Left            =   2880
            TabIndex        =   39
            Top             =   0
            Width           =   2505
         End
         Begin VB.TextBox txtCod_NotaAbono 
            Height          =   285
            Left            =   2160
            MaxLength       =   4
            TabIndex        =   38
            Top             =   0
            Width           =   600
         End
         Begin VB.TextBox txtReferencia 
            Height          =   285
            Left            =   2160
            TabIndex        =   40
            Top             =   360
            Width           =   5505
         End
         Begin VB.Label Label20 
            Caption         =   "Motivo de Nota de Abono :"
            Height          =   255
            Left            =   0
            TabIndex        =   65
            Top             =   0
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Doc Referencia :"
            Height          =   255
            Left            =   0
            TabIndex        =   62
            Top             =   375
            Width           =   1695
         End
      End
      Begin VB.Frame frOtros 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         TabIndex        =   58
         Top             =   5040
         Visible         =   0   'False
         Width           =   7695
         Begin NumBoxProject.NumBox Imp_Gastos_Finacieros 
            Height          =   285
            Left            =   2160
            TabIndex        =   36
            Tag             =   "SET/VALID"
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            TypeVal         =   2
            Mask            =   "9,999,999,999.99"
            Formato         =   "#,###,###,###.##"
            AllowedMask     =   -1
            MaskLen         =   10
            Aling           =   3
            Text            =   "0.00"
            CanEmpty        =   -1
            ShowError       =   0
            Locked          =   0   'False
            Enabled         =   -1  'True
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DecimalNumber   =   2
         End
         Begin NumBoxProject.NumBox Imp_Otros 
            Height          =   285
            Left            =   2160
            TabIndex        =   37
            Tag             =   "SET/VALID"
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            TypeVal         =   2
            Mask            =   "9,999,999,999.99"
            Formato         =   "#,###,###,###.##"
            AllowedMask     =   -1
            MaskLen         =   10
            Aling           =   3
            Text            =   "0.00"
            CanEmpty        =   -1
            ShowError       =   0
            Locked          =   0   'False
            Enabled         =   -1  'True
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DecimalNumber   =   2
         End
         Begin VB.Label LblOtros 
            Caption         =   "Gastos Financieros :"
            Height          =   255
            Left            =   0
            TabIndex        =   60
            Top             =   15
            Width           =   1695
         End
         Begin VB.Label Label22 
            Caption         =   "Otros :"
            Height          =   255
            Left            =   0
            TabIndex        =   59
            Top             =   375
            Width           =   1215
         End
      End
      Begin VB.Label Label32 
         Caption         =   "Tipo Fac"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   7200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Observacion :"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   4695
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "Glosa :"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   4020
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Tipo Venta :"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   735
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Tipo Documento :"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   375
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Doc Interno :"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   3705
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Ordenes / Pedidos :"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   3345
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Guias :"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   2985
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Serie :"
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   1125
         Width           =   450
      End
      Begin VB.Label Label5 
         Caption         =   "Numero :"
         Height          =   255
         Left            =   3390
         TabIndex        =   50
         Top             =   1095
         Width           =   855
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   1515
         Width           =   570
      End
      Begin VB.Label Label28 
         Caption         =   "R.U.C."
         Height          =   255
         Left            =   5640
         TabIndex        =   48
         Top             =   1485
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Registro :"
         Height          =   255
         Left            =   4080
         TabIndex        =   47
         Top             =   1845
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha Emisión :"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1845
         Width           =   1335
      End
      Begin VB.Label Label27 
         Caption         =   "T./C.:"
         Height          =   255
         Left            =   6360
         TabIndex        =   45
         Top             =   1845
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Moneda :"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   2205
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Forma Pago :"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   2625
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmAdicionaDocumVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public codigo As String, Descripcion As String, strOption As String, strNum_Corre As String, strCod_Anxo As String
Dim strSQL As String

Sub Busca_Opcion_Anexo(strCampo1 As String, strCampo2 As String, StrTabla As String, txtCod As TextBox, txtDes As TextBox, Opcion As Integer)

On Error GoTo Fin

Dim rstAux As ADODB.Recordset
    strSQL = "select Cod_Anxo as Cod,Des_Anexo as Nombre,Num_Ruc as Ruc,Origen from cn_anexoscontables where cod_tipanex = 'C' and "

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    Select Case Opcion
    Case 1: strSQL = strSQL & strCampo1 & " like '%" & txtCod & "%'"
    Case 2: strSQL = strSQL & strCampo2 & " like '%" & txtDes & "%'"
    End Select
    txtCod = ""
    txtDes = ""
    With frmBusqGeneral
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        
        codigo = ".."
        .DGridLista.Columns("Cod").Visible = False
        .DGridLista.Columns("Origen").Visible = False
        .DGridLista.Columns("Nombre").Width = 4575
        .DGridLista.Columns("RUC").Width = 1695
        Set rstAux = .DGridLista.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If codigo <> "" And rstAux.RecordCount > 0 Then
            strCod_Anxo = Trim(rstAux!Cod)
            txtDes = Trim(rstAux!Nombre)
            txtCod = Trim(rstAux!Ruc)
            If rstAux!origen = "E" Then chkExportacion.Value = 1 Else chkExportacion.Value = 0
            Select Case Opcion
            Case 1: SendKeys "{TAB}"
            Case 2: SendKeys "{TAB}": SendKeys "{TAB}"
            End Select
        Else
            SendKeys "{TAB}"
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Resume
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & Opcion & ")"
End Sub

Sub Busca_Opcion(strCampo1 As String, strCampo2 As String, StrTabla As String, txtCod As TextBox, txtDes As TextBox, Opcion As Integer)

On Error GoTo Fin

Dim rstAux As ADODB.Recordset
    strSQL = "Select " & strCampo1 & " AS Cod," & strCampo2 & " as Descripcion from " & StrTabla

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    Select Case Opcion
    Case 1: strSQL = strSQL & strCampo1 & " like '%" & txtCod & "%'"
    Case 2: strSQL = strSQL & strCampo2 & " like '%" & txtDes & "%'"
    End Select
    txtCod = ""
    txtDes = ""
    With frmBusqGeneral
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        
        codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If codigo <> "" And rstAux.RecordCount > 0 Then
            txtCod = Trim(rstAux!Cod)
            txtDes = Trim(rstAux!Descripcion)
            Select Case Opcion
            Case 1: SendKeys "{TAB}": SendKeys "{TAB}"
            Case 2: SendKeys "{TAB}"
            End Select
        Else
            SendKeys "{TAB}"
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & Opcion & ")"
End Sub

Private Sub chkExportacion_Click()
  Cambio_FR
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

On Error GoTo dprDepurar

Select Case ActionName

Case Is = "GRABAR"
  If MsgBox("Desea Grabar esta Grabar la Factura " & txtSer_Docum & "-" & txtNum_Docum, vbYesNo, "AVISO") = vbYes Then
    Grabar
    Unload Me
  End If
Case Is = "CANCELAR"
  Unload Me
End Select

Exit Sub

Resume

dprDepurar:

errores err.Number

End Sub

Sub Grabar()
 
Dim rs As Object
Set rs = CreateObject("ADODB.Recordset")

strSQL = "Ventas_Up_Man '" & strOption & "','" & strNum_Corre & "','" & txtCod_TipDoc & "','" & txtSer_Docum & "','" _
        & txtNum_Docum & "','" & txtCod_TipAne & "','" & strCod_Anxo & "','" & txtCod_ConPag & "','" & inpFec_EmiDoc.Text & "','" _
        & InpFec_RegDoc.Text & "','" & txtCod_Moneda & "'," & Imp_Gastos_Finacieros.Text & "," & Imp_Otros.Text & ",'" _
        & Des_Apos(txtGlosa) & "','" & vusu & "','" & txtNro_Guias & "','" & txtNro_Ordener & "','" & txtNro_DocInter & "','" _
        & txtCod_TipVenta.Text & "','" & IIf(chkExportacion, "S", "N") & "'," & Imp_Flete.Text & "," & txtTipo_Cambio.Text & ",'" _
        & IIf(chkFlete, "S", "N") & "','" & Des_Apos(txtReferencia) & "','" & Des_Apos(TxtObservacion.Text) & "','" _
        & txtCod_Destino & "','" & txtShip_Date.Text & "'," & txtPeso_Bruto.Text & "," & txtPeso_Neto.Text & ",'" _
        & txtCod_NotaAbono & "'," & txtImp_Seguro.Text & ",'" & txtEmbarque_Cod.Text & "','" & txtDua.Text & "','" _
        & txtFec_Numeracion.Text & "','" & txtFec_Embarque.Text & "','" & IIf(chkSeguro, "S", "N") & "','" & IIf(chkDetraccion, "S", "N") & "','" & txtImp_FOB_Dol_Dua.Text & "','" & txtcajas.Text & "','" & IIf(chkExonerado, "S", "N") & "','" & Trim(txtTipoFac.Text) & "'"
        
Set rs = CargarRecordSetDesconectado(strSQL, cCONNECT)

If Not rs.EOF And Not rs.BOF Then
strNum_Corre = rs!Num_Corre
frmAdicionaDetalleDocum.strNum_Corre_Detalle = strNum_Corre

End If
       
End Sub

Sub Cambio_FR()
  Imp_Gastos_Finacieros.Text = 0
  Imp_Otros.Text = 0
  Imp_Flete.Text = 0
  txtPeso_Bruto.Text = 0
  txtShip_Date.Text = ""
  txtPeso_Neto.Text = 0
  chkFlete.Value = 0
  chkSeguro.Value = 0
  frOtros.Visible = False
  frExportacion.Visible = False
  frReferencia.Visible = False
    
 
  If txtCod_TipDoc = "NC" Or txtCod_TipDoc = "ND" Then
    frReferencia.Visible = True
  End If
  
  If chkExportacion Then
    frExportacion.Visible = True
  Else
    frOtros.Visible = True
  End If
  
End Sub

Private Sub Imp_Flete_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Imp_Gastos_Finacieros_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub Imp_Otros_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub inpFec_EmiDoc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub inpFec_EmiDoc_LostFocus()
  txtTipo_Cambio.Text = DevuelveCampo("select Tipo_Venta from cn_tipocambio where fecha = '" & inpFec_EmiDoc.Text & "'", cCONNECT)
End Sub

Private Sub InpFec_RegDoc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCod_ConPag_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_CondVent", "Des_CondVent", "Lg_CondVent where ", txtCod_ConPag, txtDes_ConPag, 1)
End Sub

Private Sub txtCod_Destino_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Destino", "Des_Destino", "tg_destino where ", txtCod_Destino, txtDes_Destino, 1)
End Sub

Private Sub txtCod_Moneda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Moneda", "Nom_Moneda", "TG_Moneda where ", txtCod_Moneda, txtDes_Moneda, 1)
End Sub

Private Sub txtCod_NotaAbono_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Mot_Nota", "Descripcion", "Cn_Ventas_Motivos_Notas_Abonos where Cod_TipDoc ='" & txtCod_TipDoc & "' and ", txtCod_NotaAbono, txtDes_NotaAbono, 1)
End Sub

Private Sub txtCod_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAne, txtDes_TipAnex, 1)
End Sub

Private Sub txtCod_TipDoc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_TipDoc", "Des_TipDoc", "CN_TiposDocum where Flg_Doc_Ventas = '*' and ", txtCod_TipDoc, txtDes_TipDoc, 1)
    Cambio_FR
    If txtCod_TipDoc = "BV" Then txtCod_TipAne = ""
  End If
  
End Sub

Private Sub txtCod_TipDoc_LostFocus()
  Cambio_FR
End Sub

Private Sub txtCod_TipVenta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Tipo_Venta", "Descripcion", "Cn_Tipos_Venta where ", txtCod_TipVenta, txtDes_TipVenta, 1)
'    If gfVerificar_ExisteRegistroTabla("Cn_Ventas_Motivos_Notas_Abonos", "Cod_TipDoc ='" & txtCod_TipDoc & "'", cCONNECT) = eNoExiste Then
End Sub

Private Sub txtDes_ConPag_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_CondVent", "Des_CondVent", "Lg_CondVent where ", txtCod_ConPag, txtDes_ConPag, 2)
End Sub

Private Sub txtDes_Destino_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Destino", "Des_Destino", "tg_destino where ", txtCod_Destino, txtDes_Destino, 2)
End Sub

Private Sub txtDes_Moneda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Moneda", "Nom_Moneda", "TG_Moneda where ", txtCod_Moneda, txtDes_Moneda, 2)
End Sub

Private Sub txtDes_NotaAbono_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Mot_Nota", "Descripcion", "Cn_Ventas_Motivos_Notas_Abonos where Cod_TipDoc ='" & txtCod_TipDoc & "' and ", txtCod_NotaAbono, txtDes_NotaAbono, 2)
End Sub

Private Sub txtDes_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 2, Me)
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtDes_TipDoc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_TipDoc", "Des_TipDoc", "CN_TiposDocum where Flg_Doc_Ventas = '*' and ", txtCod_TipDoc, txtDes_TipDoc, 2)
    Cambio_FR
  End If
End Sub

Private Sub txtDes_TipVenta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Tipo_Venta", "Descripcion", "Cn_Tipos_Venta where ", txtCod_TipVenta, txtDes_TipVenta, 2)
End Sub

Private Sub txtEmbarque_Cod_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Tip_Embarque", "Des_TipEmbarque", "TG_TipoEmbarque where ", txtEmbarque_Cod, txtEmbarque_Des, 1)
End Sub

Private Sub txtEmbarque_Des_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Tip_Embarque", "Des_TipEmbarque", "TG_TipoEmbarque where ", txtEmbarque_Cod, txtEmbarque_Des, 1)
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtImp_Seguro_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNro_DocInter_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNro_Guias_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNro_Ordener_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNum_Docum_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
  If KeyAscii >= 48 And KeyAscii <= 57 _
    Or KeyAscii = 8 Or KeyAscii = 13 Then Else KeyAscii = 0
End Sub

Private Sub txtNum_Docum_LostFocus()
  txtNum_Docum = Format(txtNum_Docum, "00000000")
End Sub

Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 1, Me)
End Sub

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtPeso_Bruto_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtPeso_Neto_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtSer_Docum_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
  If KeyAscii >= 48 And KeyAscii <= 57 _
    Or KeyAscii = 8 Or KeyAscii = 13 Then Else KeyAscii = 0
End Sub

Private Sub txtSer_Docum_LostFocus()
  txtSer_Docum = Format(txtSer_Docum, "000")
End Sub

Private Sub txtShip_Date_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtTipo_Cambio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Public Sub Busca()
   Dim oTipo As New frmBusqGeneral
   Dim rs As Object
   Set rs = CreateObject("ADODB.Recordset")
   Set oTipo.oParent = Me
      oTipo.sQuery = "EXEC Ventas_Muestra_Motivo_Notas_Abono_Credito  '" & txtCod_TipDoc & "'"
      oTipo.Cargar_Datos
      oTipo.DGridLista.Columns(2).Width = 3500
      oTipo.Show 1
      If codigo <> "" Then
         Me.txtCod_TipVenta = Trim(codigo)
         Me.txtDes_TipVenta = Trim(Descripcion)
           codigo = "": Descripcion = ""
      End If
        Set oTipo = Nothing
        Set rs = Nothing
End Sub

