VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmAdicionaDocumVentasExport 
   Caption         =   "Adiciona Documento Venta Prendas"
   ClientHeight    =   10620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   10620
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frMain 
      Height          =   10005
      Left            =   60
      TabIndex        =   39
      Top             =   0
      Width           =   8475
      Begin VB.TextBox txtCod_Vendor 
         Height          =   315
         Left            =   2265
         MaxLength       =   20
         TabIndex        =   36
         Top             =   9555
         Width           =   2130
      End
      Begin VB.TextBox txtCod_Class 
         Height          =   315
         Left            =   5640
         MaxLength       =   10
         TabIndex        =   37
         Top             =   9570
         Width           =   1125
      End
      Begin VB.TextBox txtNom_Embarque 
         Height          =   315
         Left            =   2280
         TabIndex        =   33
         Top             =   7245
         Width           =   2340
      End
      Begin VB.TextBox txtDes_Embarque 
         Height          =   345
         Left            =   2910
         TabIndex        =   32
         Top             =   6840
         Width           =   4860
      End
      Begin VB.TextBox txtCod_Embarque 
         Height          =   345
         Left            =   2280
         TabIndex        =   44
         Top             =   6840
         Width           =   585
      End
      Begin VB.TextBox txtPie_Pagina1 
         Height          =   885
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   34
         Top             =   7620
         Width           =   5940
      End
      Begin VB.TextBox txtPie_Pagina2 
         Height          =   885
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   35
         Top             =   8595
         Width           =   5940
      End
      Begin VB.TextBox txtDes_TipVenta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2970
         TabIndex        =   69
         Top             =   555
         Width           =   2505
      End
      Begin VB.TextBox txtCod_TipVenta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   68
         Top             =   555
         Width           =   600
      End
      Begin VB.TextBox txtCod_LugEnt 
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   23
         Top             =   4965
         Width           =   600
      End
      Begin VB.TextBox txtDes_LugEnt 
         Height          =   285
         Left            =   2970
         TabIndex        =   24
         Top             =   4965
         Width           =   3345
      End
      Begin VB.TextBox txtCod_TipoFact 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   18
         Top             =   3240
         Width           =   600
      End
      Begin VB.TextBox txtDes_TipoFact 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2970
         TabIndex        =   19
         Top             =   3240
         Width           =   3345
      End
      Begin VB.TextBox txtNum_CartaCredito 
         Height          =   285
         Left            =   2280
         TabIndex        =   17
         Top             =   2895
         Width           =   2115
      End
      Begin VB.TextBox txtAbr_Cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   8
         Top             =   1545
         Width           =   570
      End
      Begin VB.TextBox txtNom_Cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   9
         Top             =   1530
         Width           =   3075
      End
      Begin VB.TextBox txtDes_TipAnex 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         MaxLength       =   11
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtCod_TipAne 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "C"
         Top             =   1200
         Width           =   360
      End
      Begin VB.Frame frOtros 
         BorderStyle     =   0  'None
         Height          =   1425
         Left            =   120
         TabIndex        =   45
         Top             =   5325
         Width           =   8175
         Begin VB.TextBox txtImp_Desaduanaje 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6540
            TabIndex        =   42
            Text            =   "0"
            Top             =   345
            Width           =   1125
         End
         Begin VB.TextBox txtImp_Transporte_Pais_Destino 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6540
            TabIndex        =   43
            Text            =   "0"
            Top             =   705
            Width           =   1125
         End
         Begin VB.TextBox txtDes_Termino_Venta 
            Height          =   285
            Left            =   2850
            TabIndex        =   26
            Top             =   0
            Width           =   4815
         End
         Begin VB.TextBox txtCod_Termino_Venta 
            Height          =   285
            Left            =   2160
            TabIndex        =   25
            Top             =   0
            Width           =   585
         End
         Begin NumBoxProject.NumBox Imp_Gastos_Finacieros 
            Height          =   285
            Left            =   2160
            TabIndex        =   27
            Tag             =   "SET/VALID"
            Top             =   330
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
            TabIndex        =   28
            Tag             =   "SET/VALID"
            Top             =   675
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
         Begin NumBoxProject.NumBox imp_Flete 
            Height          =   285
            Left            =   4080
            TabIndex        =   40
            Tag             =   "SET/VALID"
            Top             =   345
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
         Begin NumBoxProject.NumBox imp_Seguro 
            Height          =   285
            Left            =   4080
            TabIndex        =   41
            Tag             =   "SET/VALID"
            Top             =   690
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
         Begin NumBoxProject.NumBox Imp_Descuento 
            Height          =   285
            Left            =   2160
            TabIndex        =   29
            Tag             =   "SET/VALID"
            Top             =   1100
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
         Begin NumBoxProject.NumBox porc_comision 
            Height          =   285
            Left            =   4440
            TabIndex        =   30
            Tag             =   "SET/VALID"
            Top             =   1100
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
         Begin NumBoxProject.NumBox imp_comision 
            Height          =   285
            Left            =   6600
            TabIndex        =   31
            Tag             =   "SET/VALID"
            Top             =   1100
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
         Begin VB.Label Label32 
            Caption         =   "Importe Comision :"
            Height          =   375
            Left            =   5760
            TabIndex        =   82
            Top             =   1035
            Width           =   1095
         End
         Begin VB.Label Label31 
            Caption         =   "% Comision :"
            Height          =   252
            Left            =   3480
            TabIndex        =   81
            Top             =   1100
            Width           =   1212
         End
         Begin VB.Label Label30 
            Caption         =   "Desaduanaje"
            Height          =   255
            Left            =   5400
            TabIndex        =   80
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Label29 
            Caption         =   "Transporte en Pais Destino"
            Height          =   435
            Left            =   5400
            TabIndex        =   79
            Top             =   600
            Width           =   1110
         End
         Begin VB.Label Label18 
            Caption         =   "Terminos de Ventas"
            Height          =   285
            Left            =   -30
            TabIndex        =   78
            Top             =   30
            Width           =   1920
         End
         Begin VB.Label Dscto 
            Caption         =   "Descuento :"
            Height          =   255
            Left            =   0
            TabIndex        =   71
            Top             =   1100
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Flete:"
            Height          =   255
            Left            =   3450
            TabIndex        =   63
            Top             =   375
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "Seguro :"
            Height          =   255
            Left            =   3450
            TabIndex        =   62
            Top             =   705
            Width           =   615
         End
         Begin VB.Label Label22 
            Caption         =   "Otros :"
            Height          =   255
            Left            =   0
            TabIndex        =   47
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label LblOtros 
            Caption         =   "Gastos Financieros :"
            Height          =   255
            Left            =   0
            TabIndex        =   46
            Top             =   390
            Width           =   1695
         End
      End
      Begin VB.TextBox txtDes_ConPag 
         Height          =   285
         Left            =   3000
         TabIndex        =   16
         Top             =   2550
         Width           =   3345
      End
      Begin VB.TextBox txtCod_ConPag 
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   15
         Top             =   2550
         Width           =   600
      End
      Begin VB.TextBox txtDes_Moneda 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   14
         Top             =   2220
         Width           =   3345
      End
      Begin VB.TextBox txtCod_Moneda 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   13
         Top             =   2205
         Width           =   600
      End
      Begin VB.TextBox txtDes_TipAne 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2685
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1200
         Width           =   2865
      End
      Begin VB.TextBox txtNro_Guias 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   20
         Top             =   3585
         Width           =   5505
      End
      Begin VB.TextBox txtNro_Ordener 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   21
         Top             =   3930
         Width           =   5505
      End
      Begin VB.TextBox txtNum_Ruc 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6240
         MaxLength       =   11
         TabIndex        =   7
         Top             =   1200
         Width           =   1545
      End
      Begin VB.TextBox txtNum_Docum 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         MaxLength       =   8
         TabIndex        =   3
         Top             =   885
         Width           =   1200
      End
      Begin VB.TextBox txtDes_TipDoc 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   1
         Top             =   240
         Width           =   2505
      End
      Begin VB.TextBox txtCod_TipDoc 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   0
         Top             =   240
         Width           =   600
      End
      Begin VB.TextBox txtSer_Docum 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   2
         Top             =   885
         Width           =   600
      End
      Begin VB.TextBox txtGlosa 
         Height          =   630
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   4275
         Width           =   5505
      End
      Begin NumBoxProject.NumBox inpFec_EmiDoc 
         Height          =   285
         Left            =   2280
         TabIndex        =   10
         Top             =   1890
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
         Top             =   1890
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
         Top             =   1890
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
         Enabled         =   0   'False
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   3
      End
      Begin VB.Label Label25 
         Caption         =   "Cod.Vendor"
         Height          =   255
         Left            =   150
         TabIndex        =   77
         Top             =   9630
         Width           =   1485
      End
      Begin VB.Label Label24 
         Caption         =   "Class"
         Height          =   315
         Left            =   4560
         TabIndex        =   76
         Top             =   9600
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "Nombre Transporte"
         Height          =   270
         Left            =   135
         TabIndex        =   75
         Top             =   7305
         Width           =   1485
      End
      Begin VB.Label Label21 
         Caption         =   "Modo de Transporte"
         Height          =   315
         Left            =   105
         TabIndex        =   74
         Top             =   6900
         Width           =   1590
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Pie Factura 1:"
         Height          =   195
         Left            =   150
         TabIndex        =   73
         Top             =   7680
         Width           =   990
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Pie Factura 2:"
         Height          =   195
         Left            =   165
         TabIndex        =   72
         Top             =   8655
         Width           =   990
      End
      Begin VB.Label lblNum_Corre 
         Height          =   300
         Left            =   5625
         TabIndex        =   70
         Top             =   255
         Width           =   2310
      End
      Begin VB.Label Label17 
         Caption         =   "Lugar de Entrega :"
         Height          =   255
         Left            =   105
         TabIndex        =   67
         Top             =   4980
         Width           =   1740
      End
      Begin VB.Label Label16 
         Caption         =   "Tipo de Facturación :"
         Height          =   255
         Left            =   105
         TabIndex        =   66
         Top             =   3255
         Width           =   1740
      End
      Begin VB.Label Label15 
         Caption         =   "Carta de Crédito:"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   2910
         Width           =   1365
      End
      Begin VB.Label Label10 
         Caption         =   "Cliente Comercial"
         Height          =   255
         Left            =   150
         TabIndex        =   64
         Top             =   1575
         Width           =   1590
      End
      Begin VB.Label Label3 
         Caption         =   "Forma Pago :"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   2565
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Moneda :"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   2220
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "T./C.:"
         Height          =   255
         Left            =   6360
         TabIndex        =   59
         Top             =   1905
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha Emisión :"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1905
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Registro :"
         Height          =   255
         Left            =   4080
         TabIndex        =   57
         Top             =   1905
         Width           =   735
      End
      Begin VB.Label Label28 
         Caption         =   "R.U.C."
         Height          =   255
         Left            =   5640
         TabIndex        =   56
         Top             =   1215
         Width           =   615
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Consignatario:"
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   1245
         Width           =   1530
      End
      Begin VB.Label Label5 
         Caption         =   "Numero :"
         Height          =   255
         Left            =   3390
         TabIndex        =   54
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Serie :"
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   930
         Width           =   450
      End
      Begin VB.Label Label6 
         Caption         =   "Guias :"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Ordenes / Pedidos :"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   3945
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Tipo Documento :"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   255
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Tipo Venta :"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   585
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Glosa :"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   4290
         Width           =   1455
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3067
      TabIndex        =   38
      Top             =   10065
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmAdicionaDocumVentasExport.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmAdicionaDocumVentasExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public codigo As String, Descripcion As String, strOption As String, strNum_Corre As String, strCod_Anxo As String
Dim strSQL As String

Sub Busca_Opcion_Anexo(strCampo1 As String, strCampo2 As String, StrTabla As String, txtCod As TextBox, txtDes As TextBox, opcion As Integer)

On Error GoTo Fin

Dim rstAux As ADODB.Recordset
    strSQL = "select Cod_Anxo as Cod,Des_Anexo as Nombre,Num_Ruc as Ruc,Origen from cn_anexoscontables where cod_tipanex = 'C' and "

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    Select Case opcion
    Case 1: strSQL = strSQL & strCampo1 & " like '%" & txtCod & "%'"
    Case 2: strSQL = strSQL & strCampo2 & " like '%" & txtDes & "%'"
    End Select
    txtCod = ""
    txtDes = ""
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = strSQL
        .CARGAR_DATOS
        
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
            
            Select Case opcion
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
    "Búsqueda de Descuento (" & opcion & ")"
End Sub

Sub Busca_Opcion(strCampo1 As String, strCampo2 As String, StrTabla As String, txtCod As TextBox, txtDes As TextBox, opcion As Integer)

On Error GoTo Fin

Dim rstAux As ADODB.Recordset
    strSQL = "Select " & strCampo1 & " AS Cod," & strCampo2 & " as Descripcion from " & StrTabla

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    Select Case opcion
    Case 1: strSQL = strSQL & strCampo1 & " like '%" & txtCod & "%'"
    Case 2: strSQL = strSQL & strCampo2 & " like '%" & txtDes & "%'"
    End Select
    txtCod = ""
    txtDes = ""
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = strSQL
        .CARGAR_DATOS
        
        codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If codigo <> "" And rstAux.RecordCount > 0 Then
            txtCod = Trim(rstAux!Cod)
            txtDes = Trim(rstAux!Descripcion)
            Select Case opcion
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
    "Búsqueda de Descuento (" & opcion & ")"
End Sub


Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

On Error GoTo dprDepurar

Select Case ActionName

Case Is = "GRABAR"
  If MsgBox("Desea Grabar Factura " & txtSer_Docum & "-" & txtNum_Docum, vbYesNo, "AVISO") = vbYes Then
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
 
Dim RS As Object
Set RS = CreateObject("ADODB.Recordset")

On Error GoTo errx
       
strSQL = "Ventas_Up_Man_Exportacion  '$', '$','$','$', '$','$','$','$',$,$,'$','$','$','$','$',$,$,$,'$','$','$','$','$','$','$',$,$,$,$"
strSQL = VBsprintf(strSQL, strOption, lblNum_Corre, txtCod_TipDoc, txtSer_Docum, txtNum_Docum, _
txtCod_ConPag, inpFec_EmiDoc.Text, InpFec_RegDoc.Text, Imp_Gastos_Finacieros.Text, _
Imp_Otros.Text, txtGlosa, vusu, txtCod_TipoFact, txtCod_LugEnt, txtNum_CartaCredito, _
Imp_Flete.Text, imp_Seguro.Text, Imp_Descuento.Text, txtCod_Termino_Venta.Text, txtCod_Embarque.Text, _
txtNom_Embarque.Text, txtPie_Pagina1.Text, txtPie_Pagina2.Text, txtCod_Vendor.Text, txtCod_Class.Text, _
txtImp_Desaduanaje.Text, txtImp_Transporte_Pais_Destino.Text, Me.porc_comision.Text, Me.imp_comision.Text)
              
Set RS = CargarRecordSetDesconectado(strSQL, cCONNECT)

If Not RS.EOF And Not RS.BOF Then strNum_Corre = RS!Num_Corre

Exit Sub
errx:
    errores err.Number
      
End Sub


Private Sub Imp_Descuento_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
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

Private Sub imp_Seguro_KeyPress(KeyAscii As Integer)
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

Private Sub txt_PesoNeto_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub porc_comision_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub imp_comision_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCod_Class_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCod_ConPag_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_CondVent", "Des_CondVent", "Lg_CondVent where ", txtCod_ConPag, txtDes_ConPag, 1)
End Sub


Private Sub txtCod_LugEnt_KeyPress(KeyAscii As Integer)
      If KeyAscii = vbKeyReturn Then
        BuscaLugEnt 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCod_Moneda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Moneda", "Nom_Moneda", "TG_Moneda where ", txtCod_Moneda, txtDes_Moneda, 1)
End Sub

Private Sub txtCod_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAne, txtDes_TipAnex, 1)
End Sub

Private Sub txtCod_TipDoc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_TipDoc", "Des_TipDoc", "CN_TiposDocum where Flg_Doc_Ventas = '*' and ", txtCod_TipDoc, txtDes_TipDoc, 1)
    
  End If
  
End Sub

Private Sub txtCod_TipoFact_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCod_TipVenta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If gfVerificar_ExisteRegistroTabla("Cn_Ventas_Motivos_Notas_Abonos", "Cod_TipDoc ='" & txtCod_TipDoc & "'", cCONNECT) = eNoExiste Then
      Call Busca_Opcion("Cod_Tipo_Venta", "Descripcion", "Cn_Tipos_Venta where ", txtCod_TipVenta, txtDes_TipVenta, 1)
    Else
      Call Busca_Opcion("Cod_Mot_Nota", "Descripcion", "Cn_Ventas_Motivos_Notas_Abonos where Cod_TipDoc ='" & txtCod_TipDoc & "' and ", txtCod_TipVenta, txtDes_TipVenta, 1)
    End If
  End If
End Sub

Private Sub txtCod_Vendor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDes_ConPag_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_CondVent", "Des_CondVent", "Lg_CondVent where ", txtCod_ConPag, txtDes_ConPag, 2)
End Sub


Private Sub txtDes_Embarque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDes_LugEnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDes_Moneda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_Moneda", "Nom_Moneda", "TG_Moneda where ", txtCod_Moneda, txtDes_Moneda, 2)
End Sub

Private Sub txtDes_Termino_Venta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDes_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion_Anexo("Num_Ruc", "Des_Anexo", " Cn_AnexosContables where cod_tipanex = 'C' and ", txtNum_Ruc, txtDes_TipAne, 2)
End Sub


Private Sub txtDes_TipoFact_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDes_TipVenta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If gfVerificar_ExisteRegistroTabla("Cn_Ventas_Motivos_Notas_Abonos", "Cod_TipDoc ='" & txtCod_TipDoc & "'", cCONNECT) = eNoExiste Then
      Call Busca_Opcion("Cod_Tipo_Venta", "Descripcion", "Cn_Tipos_Venta where ", txtCod_TipVenta, txtDes_TipVenta, 2)
    Else
      Call Busca_Opcion("Cod_Mot_Nota", "Descripcion", "Cn_Ventas_Motivos_Notas_Abonos where Cod_TipDoc ='" & txtCod_TipDoc & "' and ", txtCod_TipVenta, txtDes_TipVenta, 2)
    End If
  End If
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNro_DocInter_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtImp_Desaduanaje_GotFocus()
    SelectionText txtImp_Desaduanaje
End Sub

Private Sub txtImp_Desaduanaje_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtImp_Transporte_Pais_Destino_GotFocus()
    SelectionText txtImp_Transporte_Pais_Destino
End Sub

Private Sub txtImp_Transporte_Pais_Destino_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtNro_Guias_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNro_Ordener_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNum_CartaCredito_KeyPress(KeyAscii As Integer)
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
  If KeyAscii = 13 Then Call Busca_Opcion_Anexo("Num_Ruc", "Des_Anexo", " Cn_AnexosContables where cod_tipanex = 'C' and ", txtNum_Ruc, txtDes_TipAne, 1)
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



Private Sub txtSecuencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        
        BuscaLugEnt 1
        SendKeys "{TAB}"
    End If
End Sub

Public Sub BuscaLugEnt(opcion As String)
Dim rstAux As ADODB.Recordset
    strSQL = "SELECT Secuencia, RTRIM(Linea1) + ' ' + RTRIM(Linea2) + " & _
             "RTRIM(Linea3) AS Linea1 FROM TG_CLIENTE_LUGENT " & _
             "WHERE Cod_Cliente = '" & txtAbr_Cliente.Tag & "' AND "
    
    txtCod_LugEnt = Trim(txtCod_LugEnt)
    txtDes_LugEnt = Trim(txtDes_LugEnt)
    
    Select Case opcion
    Case 1: strSQL = strSQL & "CONVERT(varchar(8), Secuencia ) like '%" & txtCod_LugEnt & "%'"
    Case 2: strSQL = strSQL & "RTRIM(txtDes_LugEnt ) + ' ' + RTRIM(Linea2) + " & _
             "RTRIM(Linea3) LIKE '%" & txtDes_LugEnt & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    'frmBusqGeneral3.Show vbModal
    
    frmBusqGeneral3.gexLista.Columns("Secuencia").Visible = False
    frmBusqGeneral3.gexLista.Columns("Secuencia").Width = 570
    frmBusqGeneral3.gexLista.Columns("Linea1").Width = 2370
    
    frmBusqGeneral3.gexLista.Columns("Secuencia").Caption = "Secuencia"
    frmBusqGeneral3.gexLista.Columns("Linea1").Caption = "Lug.Entr."
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtCod_LugEnt = ""
    txtDes_LugEnt = ""
    
    If codigo <> "" Then
        txtCod_LugEnt = codigo
        txtDes_LugEnt = Descripcion
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    codigo = ""
    Descripcion = ""
End Sub




Private Sub txtCod_CondVent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaCondVent 1
        SendKeys "{TAB}"
    End If
End Sub

Public Sub BuscaCondVent(opcion As String)
Dim rstAux As ADODB.Recordset

    strSQL = "SELECT Cod_CondVent, Des_CondVent FROM lg_condvent WHERE "
    
    txtCod_ConPag = Trim(txtCod_ConPag)
    txtDes_ConPag = Trim(txtDes_ConPag)
    
    Select Case opcion
    Case 1: strSQL = strSQL & "Cod_condVent like '%" & txtCod_ConPag & "%'"
    Case 2: strSQL = strSQL & "Des_condVent LIKE '%" & txtDes_ConPag & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    'frmBusqGeneralJanus.Show vbModal
    
    frmBusqGeneral3.gexLista.Columns("Cod_CondVent").Width = 700
    frmBusqGeneral3.gexLista.Columns("Des_CondVent").Width = 2000
    
    frmBusqGeneral3.gexLista.Columns("Cod_CondVent").Caption = "Cond.Vta"
    frmBusqGeneral3.gexLista.Columns("Des_condVent").Caption = "Descrip."
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtCod_ConPag = ""
    txtDes_ConPag = ""
    
    If codigo <> "" Then
        txtCod_ConPag = codigo
        txtDes_ConPag = Descripcion
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
        
    codigo = ""
    Descripcion = ""
End Sub


Private Sub txtCod_Termino_Venta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaTerminoVent 1
        SendKeys "{TAB}"
    End If
End Sub

Public Sub BuscaTerminoVent(opcion As String)
Dim rstAux As ADODB.Recordset

    strSQL = "SELECT Cod_Termino_Venta, Des_Termino_Venta FROM CN_Termino_Venta WHERE "
    
    txtCod_Termino_Venta = Trim(txtCod_Termino_Venta)
    txtDes_Termino_Venta = Trim(txtDes_Termino_Venta)
    
    Select Case opcion
    Case 1: strSQL = strSQL & "Cod_Termino_Venta like '%" & txtCod_Termino_Venta & "%'"
    Case 2: strSQL = strSQL & "Des_Termino_Venta LIKE '%" & txtDes_Termino_Venta & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    
    frmBusqGeneral3.gexLista.Columns("Cod_Termino_Venta").Width = 700
    frmBusqGeneral3.gexLista.Columns("Des_Termino_Venta").Width = 2000
    
    frmBusqGeneral3.gexLista.Columns("Cod_Termino_Venta").Caption = "Termino.Venta"
    frmBusqGeneral3.gexLista.Columns("Des_Termino_Venta").Caption = "Descrip."
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtCod_Termino_Venta = ""
    txtDes_Termino_Venta = ""
    
    If codigo <> "" Then
        txtCod_Termino_Venta = codigo
        txtDes_Termino_Venta = Descripcion
    End If
    
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
        
    codigo = ""
    Descripcion = ""
End Sub


Private Sub txtCod_Embarque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaModoTransporte 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtNom_embarque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub


Public Sub BuscaModoTransporte(opcion As String)
Dim rstAux As ADODB.Recordset

    strSQL = "SELECT Cod_Embarque, Des_Embarque FROM TG_TIPEMB WHERE "
    
    txtCod_Embarque = Trim(txtCod_Embarque)
    txtDes_Embarque = Trim(txtDes_Embarque)
    
    Select Case opcion
    Case 1: strSQL = strSQL & "Cod_Embarque like '%" & txtCod_Embarque & "%'"
    Case 2: strSQL = strSQL & "Des_Embarque LIKE '%" & txtDes_Embarque & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    
    frmBusqGeneral3.gexLista.Columns("Cod_Embarque").Width = 700
    frmBusqGeneral3.gexLista.Columns("Des_Embarque").Width = 2000
    
    frmBusqGeneral3.gexLista.Columns("Cod_Embarque").Caption = "Embarque"
    frmBusqGeneral3.gexLista.Columns("Des_Embarque").Caption = "Descrip."
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtCod_Embarque = ""
    txtDes_Embarque = ""
    
    If codigo <> "" Then
        txtCod_Embarque = codigo
        txtDes_Embarque = Descripcion
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    codigo = ""
    Descripcion = ""
End Sub


