VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmLetra 
   Caption         =   "LETRAS CLIENTES"
   ClientHeight    =   7815
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15465
   Icon            =   "frmLetra.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   15465
   Begin VB.Frame Fra_Fecha 
      Caption         =   "Fecha Recepcion"
      Height          =   1455
      Left            =   5280
      TabIndex        =   28
      Top             =   4680
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   2160
         TabIndex        =   32
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Cmd_Aceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   1080
         TabIndex        =   31
         Top             =   960
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   1920
         TabIndex        =   30
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   75169793
         CurrentDate     =   41130
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Recepcion"
         Height          =   195
         Left            =   480
         TabIndex        =   29
         Top             =   360
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Busqueda Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1995
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   15315
      Begin VB.TextBox Txt_Cod_Usuario 
         Height          =   285
         Left            =   1920
         TabIndex        =   34
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Txt_DesUsuario 
         Height          =   285
         Left            =   2880
         TabIndex        =   33
         Top             =   1560
         Width           =   5415
      End
      Begin VB.OptionButton OptAMN 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Asiento"
         Height          =   390
         Left            =   7575
         TabIndex        =   27
         Top             =   1155
         Width           =   900
      End
      Begin VB.TextBox TxtAnioReg 
         BackColor       =   &H80000014&
         Height          =   300
         Left            =   10365
         MaxLength       =   4
         TabIndex        =   26
         Top             =   1185
         Width           =   540
      End
      Begin VB.TextBox TxtMesReg 
         BackColor       =   &H80000014&
         Height          =   300
         Left            =   10950
         MaxLength       =   2
         TabIndex        =   25
         Top             =   1185
         Width           =   360
      End
      Begin VB.TextBox TxtNumReg 
         Height          =   300
         Left            =   11370
         TabIndex        =   24
         Top             =   1185
         Width           =   525
      End
      Begin VB.TextBox TxtDes_TipoDiario 
         Height          =   300
         Left            =   9045
         TabIndex        =   23
         Top             =   1185
         Width           =   1290
      End
      Begin VB.TextBox TxtCod_TipoDiario 
         Height          =   285
         Left            =   8475
         TabIndex        =   22
         Top             =   1185
         Width           =   405
      End
      Begin VB.CheckBox chkPendientes 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Todas"
         Height          =   255
         Left            =   8760
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin NumBoxProject.NumBox txtCodGrupo 
         Height          =   285
         Left            =   6645
         TabIndex        =   6
         Top             =   1185
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         TypeVal         =   1
         Mask            =   "9999999999"
         Formato         =   "#,###,###,###"
         AllowedMask     =   0
         MaskLen         =   10
         Aling           =   3
         Text            =   "0"
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   0
      End
      Begin VB.OptionButton optCod_Grupo 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Grupo"
         Height          =   270
         Left            =   5910
         TabIndex        =   18
         Top             =   1200
         Width           =   765
      End
      Begin VB.TextBox txtNum_Ruc 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1920
         MaxLength       =   11
         TabIndex        =   0
         Top             =   780
         Width           =   1185
      End
      Begin VB.TextBox txtCod_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3480
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "C"
         Top             =   780
         Width           =   360
      End
      Begin VB.OptionButton optNum_Corre 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Correlativo"
         Height          =   270
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox txtCorrelativo 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   1920
         MaxLength       =   12
         TabIndex        =   4
         Top             =   1200
         Width           =   1560
      End
      Begin VB.OptionButton optNroPropio 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Nro Letra"
         Height          =   270
         Left            =   3600
         TabIndex        =   13
         Top             =   1200
         Width           =   1005
      End
      Begin VB.TextBox txtNumPropio 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   4680
         MaxLength       =   8
         TabIndex        =   5
         Top             =   1200
         Width           =   1200
      End
      Begin VB.OptionButton optProv 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cliente"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   810
         Value           =   -1  'True
         Width           =   1065
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   495
         Left            =   13920
         TabIndex        =   3
         Top             =   720
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.TextBox txtDes_TipAne 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3885
         TabIndex        =   2
         Top             =   780
         Width           =   4425
      End
      Begin VB.TextBox txtDes_TipAnex 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4200
         MaxLength       =   11
         TabIndex        =   17
         Top             =   780
         Visible         =   0   'False
         Width           =   1545
      End
      Begin MSComCtl2.DTPicker txtFec_Ini 
         Height          =   315
         Left            =   2190
         TabIndex        =   8
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   75169793
         CurrentDate     =   37543
      End
      Begin MSComCtl2.DTPicker txtFec_Fin 
         Height          =   315
         Left            =   4350
         TabIndex        =   9
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   75169793
         CurrentDate     =   37543
      End
      Begin VB.OptionButton Opt_Vendedor 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Vendedor"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1560
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Rango de Fechas Desde :"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   270
         Width           =   1875
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Hasta :"
         Height          =   255
         Left            =   3720
         TabIndex        =   19
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ruc :"
         Height          =   180
         Left            =   1470
         TabIndex        =   16
         Tag             =   "Anexo Type"
         Top             =   825
         Width           =   435
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3210
         TabIndex        =   15
         Tag             =   "Anexo Type"
         Top             =   795
         Width           =   165
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   720
      Left            =   0
      TabIndex        =   7
      Top             =   7080
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   1270
      Custom          =   $"frmLetra.frx":030A
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1150
      ControlHeigth   =   700
      ControlSeparator=   10
   End
   Begin GridEX20.GridEX gexLetra 
      Height          =   5040
      Left            =   0
      TabIndex        =   21
      Top             =   2040
      Width           =   15330
      _ExtentX        =   27040
      _ExtentY        =   8890
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GridLineStyle   =   2
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      BorderStyle     =   2
      GroupByBoxVisible=   0   'False
      ImageWidth      =   0
      ImageHeight     =   0
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmLetra.frx":0785
      Column(2)       =   "frmLetra.frx":084D
      FormatStylesCount=   9
      FormatStyle(1)  =   "frmLetra.frx":08F1
      FormatStyle(2)  =   "frmLetra.frx":0A29
      FormatStyle(3)  =   "frmLetra.frx":0AD9
      FormatStyle(4)  =   "frmLetra.frx":0B8D
      FormatStyle(5)  =   "frmLetra.frx":0C65
      FormatStyle(6)  =   "frmLetra.frx":0D1D
      FormatStyle(7)  =   "frmLetra.frx":0DFD
      FormatStyle(8)  =   "frmLetra.frx":1209
      FormatStyle(9)  =   "frmLetra.frx":1619
      ImageCount      =   0
      PrinterProperties=   "frmLetra.frx":17A1
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   15
      Top             =   8490
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmLetra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public codigo As String, strCod_Anxo As String
Public Descripcion As String, strNum_Corre_Let_Renov As String, TipoAdd As String
Public Tipo As String
Public xCod_Grupo As Integer
Public strNum_Corre_Let As String


Private Sub chkPendientes_Click()
If chkPendientes Then
  txtFec_Fin.Enabled = True
  txtFec_Ini.Enabled = True
Else
  txtFec_Fin.Enabled = False
  txtFec_Ini.Enabled = False
End If

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Cmd_Aceptar_Click()
FechaRecepcion
End Sub

Private Sub Cmd_Cancelar_Click()
Fra_Fecha.Visible = False
End Sub

Private Sub Form_Load()
  FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
  txtFec_Ini = Date - 30
  txtFec_Fin = Date
  DTPicker1 = Now
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim vCorrelativo As String
Dim strSQL As String
    Select Case ActionName
        Case "NUEVO"
          Call Agrega_Letra(False)
        Case "MODIFICAR"
        
            If gexLetra.RowCount = 0 Then Exit Sub
            Load frmModLetraDatGen
            
            vCorrelativo = gexLetra.Value(gexLetra.Columns("Num_Corre").Index)
            
            With frmModLetraDatGen
              .LblSimbolo.Caption = DevuelveCampo("select isnull(simbolo,'') from tg_moneda where cod_moneda='" & gexLetra.Value(gexLetra.Columns("Moneda").Index) & "'", cCONNECT)
              .mvarNum_Corr = gexLetra.Value(gexLetra.Columns("Num_Corre").Index)
              .TxtNumero = gexLetra.Value(gexLetra.Columns("Nro_Letra").Index)
              .TxtMonto = gexLetra.Value(gexLetra.Columns("Imp_Total").Index)
              .inpFec_Emi.Text = gexLetra.Value(gexLetra.Columns("Fec_EmiDoc").Index)
              .inpFec_Venc.Text = IIf(IsNull(gexLetra.Value(gexLetra.Columns("Fec_VenDoc").Index)), "", gexLetra.Value(gexLetra.Columns("Fec_VenDoc").Index))
              .txtGlosa.Text = Trim(gexLetra.Value(gexLetra.Columns("Glosa").Index))
              .txtNumLetraBanco = RTrim(FixNulos(gexLetra.Value(gexLetra.Columns("Num_Letra_Banco").Index), vbString))
              .TxtCod_Banco.Text = Trim(gexLetra.Value(gexLetra.Columns("Cod_Banco").Index))
              .TxtNom_Banco.Text = Trim(gexLetra.Value(gexLetra.Columns("nom_banco").Index))
              .txtDes_TipAne.Text = Trim(gexLetra.Value(gexLetra.Columns("Des_Aval").Index))
              .txtNum_Ruc.Text = Trim(gexLetra.Value(gexLetra.Columns("Ruc_Aval").Index))
              .txtCod_TipAne.Text = Trim(gexLetra.Value(gexLetra.Columns("Cod_Tipanex_Aval").Index))
              .strCod_Anxo = Trim(gexLetra.Value(gexLetra.Columns("Cod_Anxo_Aval").Index))
              .txtFecha_Banco_Desc.Text = IIf(IsNull(gexLetra.Value(gexLetra.Columns("Fec_Aceptacion_Letra_Banco").Index)), "", gexLetra.Value(gexLetra.Columns("Fec_Aceptacion_Letra_Banco").Index))
              .txtTercero_CodTipAnexo.Text = Trim(gexLetra.Value(gexLetra.Columns("Cod_Tipanex_Tercero").Index))
              .txtTercero_NomAnexo.Text = Trim(gexLetra.Value(gexLetra.Columns("Des_Ter").Index))
              .txtTercero_NumRuc.Text = Trim(gexLetra.Value(gexLetra.Columns("Ruc_Ter").Index))
              .strTercero_Cod_Anxo = Trim(gexLetra.Value(gexLetra.Columns("Cod_Anxo_Tercero").Index))
              
            If Trim(gexLetra.Value(gexLetra.Columns("Status").Index)) <> "P" Then
              .txtGlosa.Text = Trim(gexLetra.Value(gexLetra.Columns("Glosa").Index))
            End If

              .Show 1
            
            End With
            
            Buscar
            Call gexLetra.Find(gexLetra.Columns("Num_Corre").Index, jgexEqual, vCorrelativo)
            
        Case "ELIMINARLETRA"
          If gexLetra.RowCount = 0 Then Exit Sub
          If MsgBox("Esta seguro de Eliminar este Grupo de Letras ", vbYesNo, "IMPORTANTE") = vbYes Then
            ELIMINAR_LETRA
            Buscar
          End If
        Case "CAMBIODEESTADO"
        
            If gexLetra.RowCount = 0 Then Exit Sub
              
            If gexLetra.Value(gexLetra.Columns("Status").Index) = "P" Then
                With frmLetraControlStatus
                  .Caption = .Caption & " Nro " & gexLetra.Value(gexLetra.Columns("Nro_Letra").Index)
                  .strNum_Corre = gexLetra.Value(gexLetra.Columns("Num_Corre").Index)
                  .Show 1
                  Buscar
                  gexLetra.SetFocus
                End With
              ElseIf gexLetra.Value(gexLetra.Columns("Status").Index) = "A" Or gexLetra.Value(gexLetra.Columns("Status").Index) = "C" Then
                With frmLetraControlStatusAbono
                  .Caption = .Caption & " Nro " & gexLetra.Value(gexLetra.Columns("Nro_Letra").Index)
                  .strNum_Corre = gexLetra.Value(gexLetra.Columns("Num_Corre").Index)
                  .TxtCod_Banco.Text = Trim(gexLetra.Value(gexLetra.Columns("Cod_Banco").Index))
                  .TxtNom_Banco.Text = Trim(gexLetra.Value(gexLetra.Columns("nom_banco").Index))
                  .txtNumLetraBanco = RTrim(FixNulos(gexLetra.Value(gexLetra.Columns("Num_Letra_Banco").Index), vbString))
                  .Show 1
                  Buscar
                  gexLetra.SetFocus
                End With
             ElseIf gexLetra.Value(gexLetra.Columns("Status").Index) = "G" Then
                With frmLetraControlStatusCartera
                  .Caption = .Caption & " Nro " & gexLetra.Value(gexLetra.Columns("Nro_Letra").Index)
                  .strNum_Corre = gexLetra.Value(gexLetra.Columns("Num_Corre").Index)
                  .TxtCod_Banco.Text = Trim(gexLetra.Value(gexLetra.Columns("Cod_Banco").Index))
                  .TxtNom_Banco.Text = Trim(gexLetra.Value(gexLetra.Columns("nom_banco").Index))
                  .txtNumLetraBanco = RTrim(FixNulos(gexLetra.Value(gexLetra.Columns("Num_Letra_Banco").Index), vbString))
                  .Show 1
                  Buscar
                  gexLetra.SetFocus
                End With
              Else
                MsgBox "Estado Actual de la Letra no se puede Cambiar por este opcion", vbInformation, "IMPORTANTEN"
              End If
            
        Case "PROTESTOLETRAS"
          If gexLetra.Value(gexLetra.Columns("Status").Index) = "B" Or gexLetra.Value(gexLetra.Columns("Status").Index) = "D" Or gexLetra.Value(gexLetra.Columns("Status").Index) = "G" Then
            With frmLetraControlStatusProtesto
              .strNum_Corre = gexLetra.Value(gexLetra.Columns("Num_Corre").Index)
              .Show 1
              If .lfSalvar Then Buscar
            End With
          Else
            MsgBox "Estado Actual de la Letra no se puede Protestar", vbInformation, "IMPORTANTEN"
          End If
        Case "VERDETALLE"
          If gexLetra.RowCount = 0 Then Exit Sub
          Load frmDetLetra
          frmDetLetra.vNum_Correlativo = gexLetra.Value(gexLetra.Columns("Num_Corre").Index)
          frmDetLetra.Caption = "Detalle " & " LETRA Nº " & gexLetra.Value(gexLetra.Columns("Nro_Letra").Index) & " " & Trim(gexLetra.Value(gexLetra.Columns("Cliente").Index))
          frmDetLetra.CARGA_GRID
          frmDetLetra.Show 1
          Buscar
        Case "IMPRIMIR"
          If gexLetra.RowCount = 0 Then Exit Sub
          Imprimir gexLetra.Value(gexLetra.Columns("Num_Corre").Index), gexLetra.Value(gexLetra.Columns("Imp_Total").Index), False, Left(gexLetra.Value(gexLetra.Columns("Nro_Letra").Index), 2)
        Case "RENOVACION"
          If Agrega_Letra(True) Then
            If strNum_Corre_Let_Renov <> "" Then
              With frmLetraControlStatusAbono
                .TxtCod_Banco.Text = DevuelveCampo("select cod_banco from cn_ventas where num_corre = '" & strNum_Corre_Let_Renov & "'", cCONNECT)
                .txtNumLetraBanco.Text = DevuelveCampo("select Num_Letra_Banco from cn_ventas where num_corre = '" & strNum_Corre_Let_Renov & "'", cCONNECT)
                .txtFecha.Text = gexLetra.Value(gexLetra.Columns("Fec_EmiDoc").Index)
                .txtNumLetraBanco = RTrim(FixNulos(gexLetra.Value(gexLetra.Columns("Num_Letra_Banco").Index), vbString))
                .TxtNom_Banco.Text = Trim(gexLetra.Value(gexLetra.Columns("nom_banco").Index))
                .strNum_Corre = gexLetra.Value(gexLetra.Columns("Num_Corre").Index)
                .Show 1
                Buscar
              End With
            End If
          End If
        Case "SALIR"
          Unload Me
        Case "REPCANJE"
            Call ReporteCanje
        Case "REVERTIRESTADO"
          Reversion
        Case "RECEPCION"
            If gexLetra.RowCount = 0 Then Exit Sub
            If gexLetra.Value(gexLetra.Columns("Fecha_Recepcion").Index) Then
                 DTPicker1.CheckBox = False
                DTPicker1 = gexLetra.Value(gexLetra.Columns("Fecha_Recepcion").Index)
            Else
                DTPicker1.CheckBox = True
                DTPicker1 = gexLetra.Value(gexLetra.Columns("Fecha_Recepcion").Index)
            End If
            Fra_Fecha.Visible = True
        Case "VOUCHER"
          MuestraVoucher2
    End Select

End Sub

Private Sub MuestraVoucher2()

On Error GoTo errx
Dim sSql As String
Dim rsAsientos As Object
Set rsAsientos = CreateObject("ADODB.Recordset")


If gexLetra.RowCount = 0 Then Exit Sub

sSql = "FI_Muestra_Data_Asientos_Letra_x_Cobrar '$'"
sSql = VBsprintf(sSql, gexLetra.Value(gexLetra.Columns("Num_Corre").Index))

'SELECT MIN(NUM_CORRE) FROM CN_VENTAS WHERE Cod_Grupo_Letra

Set rsAsientos = GetDataSet(cCONNECT, sSql)

With rsAsientos
  
  If .BOF Or .EOF Then
    MsgBox "No se le ha Generado Voucher", vbInformation, "AVISO"
    Exit Sub
  End If

  Load frmShowVoucher
  frmShowVoucher.sCod_TipoDiario = !Cod_TipoDiario
  frmShowVoucher.sano = !Ano_Contable
  frmShowVoucher.smes = !Mes_Contable
  frmShowVoucher.lNum_Registro = !Num_Registro
  frmShowVoucher.Num_Corre = !num_corre_para_letras
  frmShowVoucher.dImporte = gexLetra.Value(gexLetra.Columns("Imp_Total").Index)
  frmShowVoucher.sFlg_Status = gexLetra.Value(gexLetra.Columns("Estatus_Letra").Index)
  frmShowVoucher.Buscar
  frmShowVoucher.Show vbModal
  Set frmShowVoucher = Nothing
  
End With

Set rsAsientos = Nothing

Exit Sub

Resume
errx:
    errores err.Number

End Sub


Sub Reversion()

On Error GoTo hand
Dim sSql As String

sSql = "Cn_Ventas_Revierte_Pendiente_Letra '" & gexLetra.Value(gexLetra.Columns("Num_Corre").Index) & "'"
ExecuteCommandSQL cCONNECT, sSql
Buscar

Exit Sub

hand:
errores err.Number

End Sub
Private Function Agrega_Letra(dRenovacion As Boolean) As Boolean

Dim rs As Object
Set rs = CreateObject("ADODB.Recordset")

If strCod_Anxo = "" Then
    MsgBox "Seleccione un Cliente", vbInformation, Me.Caption
    Agrega_Letra = False
    Exit Function
End If
Load frmLetraAdd
Set frmLetraAdd.oParent = Me
With frmLetraAdd
  .strOption = "I"
  .varCod_anxo = strCod_Anxo
  .varCod_TipAnex = Trim(txtCod_TipAne.Text)
  ''''If Not dRenovacion Then .txtNumero_Propio = DevuelveCampo("select max(convert(numeric(13,0),num_docum_Ventas)) from cn_ventas where cod_tipdoc = '81' and isnumeric(num_docum_Ventas) = 1 and convert(numeric(13,0),num_docum_Ventas) between 10000 and 99999", cCONNECT)
  
  If Not dRenovacion Then .txtNumero_Propio = DevuelveCampo("select isnull(max(convert(numeric(13,0),num_docum_Ventas)),0) from cn_ventas where cod_tipdoc = '81' and isnumeric(num_docum_Ventas) = 1 and convert(numeric(13,0),num_docum_Ventas) between 10000 and 99999", cCONNECT)
  
  Set rs = CargarRecordSetDesconectado("Ventas_Obtiene_Anexo_Aval '" & Trim(txtCod_TipAne.Text) & "','" & strCod_Anxo & "'", cCONNECT)
  With rs
    If Not (.BOF And .EOF) Then
      frmLetraAdd.strCod_Anxo = !Cod_Anxo
      frmLetraAdd.txtCod_TipAne = !Cod_TipAnex
      frmLetraAdd.txtNum_Ruc.Text = !Num_Ruc
      frmLetraAdd.txtDes_TipAne.Text = !DES_ANEXO
    End If
  End With
  
  .Caption = IIf(dRenovacion, "Renovacion ", "Adicion ") & " de Letras al Cliente " & txtDes_TipAne
  .strRenovacion = IIf(dRenovacion, "X", "")
  .frOrigen.Visible = dRenovacion
  .Show 1
  'xCod_Grupo = frmLetraAdd.intCod_Grupo
  'strNum_Corre_Let_Renov = frmLetraAdd.strNum_Corre_Let_Renov
End With

If xCod_Grupo <> 0 Then
  optCod_Grupo.Value = True
  txtCodGrupo.Text = xCod_Grupo
  Buscar
  Call gexLetra.Find(gexLetra.Columns("Num_Corre").Index, jgexEqual, strNum_Corre_Let)
  Agrega_Letra = True
Else
  Agrega_Letra = False
End If

End Function

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
  Buscar
End Sub

Private Sub gexLetra_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
Cancel = True
End Sub

Private Sub gexLetra_RowFormat(RowBuffer As GridEX20.JSRowData)
'    If RowBuffer.Value(gexLetra.Columns("PROVISIONADO").Index) = "P" Then
'        RowBuffer.CellStyle(gexLetra.Columns("PROVISIONADO").Index) = "PROVISIONADO"
'    End If
'
'    If RowBuffer.Value(gexLetra.Columns("PROVISIONADO").Index) = " " Then
'        RowBuffer.CellStyle(gexLetra.Columns("CORRELATIVO").Index) = "NOPROVSIONADO"
'    End If
End Sub

Private Sub Opt_Vendedor_Click()
    Limpia_Busqueda
    Txt_Cod_Usuario.SetFocus

End Sub

Private Sub OptAMN_Click()
    Limpia_Busqueda
    TxtCod_TipoDiario.SetFocus
End Sub

Private Sub optCod_Grupo_Click()
  Limpia_Busqueda
  txtCodGrupo.Enabled = True
  txtCodGrupo.Text = 0
  txtCodGrupo.SetFocus
End Sub

Private Sub optNroPropio_Click()
  Limpia_Busqueda
  txtNumPropio.Enabled = True
  txtNumPropio.SetFocus
End Sub

Private Sub optNum_Corre_Click()
  Limpia_Busqueda
  txtCorrelativo.Enabled = True
  txtCorrelativo.SetFocus
End Sub

Private Sub optProv_Click()
Limpia_Busqueda

txtNum_Ruc.Enabled = True
txtCod_TipAne.Enabled = True
txtCod_TipAne = "C"
txtDes_TipAne.Enabled = True
txtNum_Ruc.SetFocus

End Sub
Sub Limpia_Busqueda()
  txtNum_Ruc = ""
  txtCod_TipAne = ""
  txtDes_TipAne = ""
  txtCodGrupo.Text = 0
  txtCorrelativo = ""
  txtNumPropio = ""
  Txt_Cod_Usuario = ""
  Txt_DesUsuario.Text = ""
  txtNum_Ruc.Enabled = False
  txtCod_TipAne.Enabled = False
  txtDes_TipAne.Enabled = False
  txtCodGrupo.Enabled = False
  txtCorrelativo.Enabled = False
  txtNumPropio.Enabled = False
  strCod_Anxo = ""

  TxtCod_TipoDiario.Text = ""
  TxtAnioReg.Text = ""
  TxtMesReg.Text = ""
  TxtNumReg.Text = 0
End Sub

Private Sub Txt_Cod_Usuario_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Busca_Trabajador
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Txt_DesUsuario_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Busca_Trabajador
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCod_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAne, txtDes_TipAnex, 1, Me)
End Sub

Private Sub txtCodGrupo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then FunctButt2.SetFocus
End Sub

Private Sub txtCorrelativo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  txtCorrelativo = Format(txtCorrelativo, "000000000000")
  FunctButt2.SetFocus
End If
End Sub

Sub Buscar()

On Error GoTo hand
'Dim Rs As ADODB.Recordset
Dim rs As Object
Dim vBookmark  As Variant
Dim lRows As Long

lRows = gexLetra.RowCount
vBookmark = gexLetra.Row
Set rs = CreateObject("ADODB.Recordset")
rs.CursorLocation = adUseClient
rs.Open "Ventas_Muestra_Letras '" & txtFec_Ini & "','" & txtFec_Fin & "','" & IIf(chkPendientes, "", "P") & "','" & Trim(txtCod_TipAne.Text) & "', '" & strCod_Anxo & "','" & txtNumPropio & "','" & txtCorrelativo & "'," & txtCodGrupo.Text & ",'" & TxtCod_TipoDiario.Text & "','" & TxtAnioReg.Text & "','" & TxtMesReg.Text & "','" & TxtNumReg & "','" & Left(Txt_Cod_Usuario, 1) & "','" & Right(Txt_Cod_Usuario, 4) & "'", cCONNECT
Set gexLetra.ADORecordset = rs
ConfiguraGrid
If gexLetra.RowCount > lRows Then
    gexLetra.Row = gexLetra.RowCount
Else
    gexLetra.Row = vBookmark
End If

Set rs = Nothing

Exit Sub
Resume
hand:
ErrorHandler err, "Buscar"
Set rs = Nothing
End Sub

Sub ConfiguraGrid()
Dim fmtCon As JSFmtCondition

    gexLetra.Columns("status").Visible = False
    gexLetra.Columns("Cod_Tipanex_Aval").Visible = False
    gexLetra.Columns("Ruc_Aval").Visible = False
    gexLetra.Columns("Des_Aval").Visible = False
    gexLetra.Columns("Cod_Banco").Visible = False
    
    gexLetra.Columns("Num_Corre").ColumnType = jgexIconAndText
    gexLetra.Columns("Num_Corre").Caption = "Correlativo"
    gexLetra.Columns("Num_Corre").Width = 1215
    gexLetra.Columns("Nro_Letra").Width = 1320
    gexLetra.Columns("Estatus_Letra").Width = 1215
    gexLetra.Columns("Fec_EmiDoc").Width = 1170
    gexLetra.Columns("Fec_VenDoc").Width = 1170
    gexLetra.Columns("Moneda").Width = 720
    gexLetra.Columns("Imp_Cancelado").Width = 1245
    gexLetra.Columns("Imp_Total").Width = 1260
    gexLetra.Columns("Fec_Ult_Pago").Width = 1155
    gexLetra.Columns("Grupo").Width = 570
    
    gexLetra.Columns("Ruc").Width = 1260
    gexLetra.Columns("Cliente").Width = 2805
  
    gexLetra.Columns("Imp_Total").Format = "###,###.00"
    gexLetra.Columns("Imp_Cancelado").Format = "###,###.00"
    
    If optProv Then
      gexLetra.Columns("Ruc").Visible = False
      gexLetra.Columns("Cliente").Visible = False
    End If
    

    'gexLetra.Columns("Correlativo").HeaderIcon = 1

    Set fmtCon = gexLetra.FmtConditions.Add(gexLetra.Columns("Status").Index, jgexNotEqual, "P")
    fmtCon.FormatStyle.FontBold = True
    'gexLetra.ColumnHeaderHeight = 500

    'gexLetra.FrozenColumns = 2
End Sub

Private Sub txtDes_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 2, Me)
  
End Sub

Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Call Busca_Opcion_Anexo1("Num_Ruc", "Des_Anexo", txtCod_TipAne, txtNum_Ruc, txtDes_TipAne, txtCod_TipAne, 1, Me)
     SendKeys "{TAB}"
     SendKeys "{TAB}"
  End If
End Sub

Private Sub txtNumPropio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  FunctButt2.SetFocus
End If
End Sub

Private Sub ELIMINAR_LETRA()
On Error GoTo hand
Dim sSql As String

sSql = "Ventas_Revierte_Letras '" & gexLetra.Value(gexLetra.Columns("Num_Corre").Index) & "'"
ExecuteCommandSQL cCONNECT, sSql

Exit Sub

hand:
errores err.Number

End Sub

Private Sub gexLetra_DblClick()
    Dim i As Integer
    For i = 1 To gexLetra.Columns.Count
        Debug.Print gexLetra.Name & ".Columns(" & Chr(34) & gexLetra.Columns(i).Caption & Chr(34) & ").width = " & CStr(gexLetra.Columns(i).Width)
    Next
End Sub


Private Sub TxtNumReg_LostFocus()
    If TxtNumReg.Text <> "" Then
        TxtNumReg.Text = StrZero(TxtNumReg.Text, 4)
    End If
End Sub
Private Sub TxtCod_TipoDiario_KeyPress(KeyAscii As Integer)
    OptAMN.Value = True
    If KeyAscii = 13 Then
        Call Busca_SubDiario("1")
    End If
    
End Sub



Private Sub TxtDes_TipoDiario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_SubDiario("2")
End If
End Sub

Sub Busca_SubDiario(Tipo As String)
Dim oTipo As New frmBusqGeneral3
Dim iCol As Long
Dim rstAux As Object
Set rstAux = CreateObject("ADODB.Recordset")
Dim strSQL As String

Set oTipo.oParent = Me

If Tipo = "1" Then
    strSQL = "SELECT cod_tipodiario as Codigo, Des_TipoDiario as Descripcion, flg_canjefacturasporpagarconletras as Flg from cn_tipodiario where cod_tipodiario like '" & Trim(TxtCod_TipoDiario.Text) & "%'"
Else
    strSQL = "SELECT cod_tipodiario as Codigo, Des_TipoDiario as Descripcion, flg_canjefacturasporpagarconletras as Flg from cn_tipodiario where des_tipodiario like '%" & Trim(TxtDes_TipoDiario.Text) & "%'"
End If
With oTipo
    Set .oParent = Me
    .sQuery = strSQL
    .Cargar_Datos
    .Caption = "Selccionar SubDiario"
    codigo = ".."
    Set rstAux = .gexLista.ADORecordset
    
    .gexLista.Columns("Codigo").Width = 700
    .gexLista.Columns("Descripcion").Width = 5000

'    For iCol = 3 To .gexLista.Columns.count
'        .gexLista.Columns(iCol).Visible = False
'    Next iCol
    
    If rstAux.RecordCount = 1 Then
        codigo = Trim(rstAux!codigo)
        Descripcion = Trim(rstAux!Descripcion)
        TipoAdd = Trim(rstAux!flg)
    End If
    
    If rstAux.RecordCount > 1 Then .Show vbModal
    
    If codigo <> "" And rstAux.RecordCount > 0 Then
        TxtCod_TipoDiario = codigo
        TxtDes_TipoDiario = Descripcion
        TxtAnioReg.SetFocus
    End If
End With

codigo = "": Descripcion = ""
Unload oTipo
Set oTipo = Nothing
rstAux.Close
Set rstAux = Nothing
End Sub



Private Sub TxtAnioReg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtMesReg.SetFocus
End If
End Sub

Private Sub TxtMesReg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtNumReg.SetFocus
End If
End Sub

Private Sub TxtMesReg_LostFocus()
    TxtMesReg.Text = Format(Trim(TxtMesReg.Text), "00")
End Sub



Sub FechaRecepcion()

On Error GoTo hand
Dim sSql As String

sSql = "Vt_Actualiza_Fecha_Recepcion '" & gexLetra.Value(gexLetra.Columns("Num_Corre").Index) & "','" & Format(DTPicker1, "dd/mm/yyyy") & "'"
ExecuteCommandSQL cCONNECT, sSql
Buscar
Fra_Fecha.Visible = False
Exit Sub

hand:
errores err.Number

End Sub

Public Sub ReporteCanje()
On Error GoTo ErrorImpresion
Dim RS1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset

If gexLetra.RowCount = 0 Then Exit Sub

VB.Screen.MousePointer = vbHourglass

Set RS1 = GetRecordset(cCONNECT, "exec cn_reporte_canje_documentos '" & gexLetra.Value(gexLetra.Columns("GRUPO").Index) & "','1'")
Set Rs2 = GetRecordset(cCONNECT, "exec cn_reporte_canje_documentos '" & gexLetra.Value(gexLetra.Columns("GRUPO").Index) & "','2'")
Dim oo As Object
Set oo = CreateObject("excel.application")

oo.Workbooks.Open vRuta & "\ReporteDetalleLetra.xlt"
oo.Visible = True
oo.Run "REPORTE", RS1, UCase(Me.Caption), Rs2


Screen.MousePointer = vbNormal
oo.Visible = True
Set oo = Nothing

Exit Sub
Resume
ErrorImpresion:
    Screen.MousePointer = vbNormal
    Set oo = Nothing
    Error err.Number
End Sub



Public Sub Busca_Trabajador()
On Error GoTo Fin
Dim iCol As Long
Dim rstAux As New ADODB.Recordset
Dim Opcion As String
Dim strSQL As String
strSQL = "Tg_Sm_Muestra_Operario_Caracteristica '001'"
    With frmBusqGeneralOperario
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("Codigo").Caption = "Codigo"
        .DGridLista.Columns("Codigo").Width = 900
        .DGridLista.Columns("Apellido_Paterno").Caption = "Ape Paterno"
        .DGridLista.Columns("Apellido_Paterno").Width = 1500
        .DGridLista.Columns("Apellido_Materno").Caption = "Ape Materno"
        .DGridLista.Columns("Apellido_Materno").Width = 1500
        .DGridLista.Columns("Nombre_Trabajador").Caption = "Nombres"
        .DGridLista.Columns("Nombre_Trabajador").Width = 1500
        
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If codigo <> "" And rstAux.RecordCount > 0 Then
            Txt_Cod_Usuario = Trim(rstAux!codigo)
            Txt_Cod_Usuario.Tag = Left(Trim(rstAux!codigo), 1)
            Txt_DesUsuario = Trim(rstAux!Apellido_Paterno) + " " + Trim(rstAux!Apellido_Materno) + " " + Trim(rstAux!Nombre_Trabajador)
            Txt_DesUsuario.Tag = Right(Trim(rstAux!codigo), 4)
            'stip_Trabajador = Left(rstAux!codigo, 1)
            'scod_trabajador = Right(rstAux!codigo, 4)
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
    "Búsqueda de Color (" & Opcion & ")"
End Sub


