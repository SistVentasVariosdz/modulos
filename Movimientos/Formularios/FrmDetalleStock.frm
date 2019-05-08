VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmDetalleStock 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de movimientos"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   12525
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   4860
      Left            =   11160
      TabIndex        =   2
      Top             =   120
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   8573
      Custom          =   $"FrmDetalleStock.frx":0000
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1250
      ControlHeigth   =   630
      ControlSeparator=   70
   End
   Begin VB.Frame Fralista 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   -30
      TabIndex        =   0
      Tag             =   "List"
      Top             =   0
      Width           =   11175
      Begin VB.Frame FraPO 
         BackColor       =   &H00FFC0C0&
         Height          =   2175
         Left            =   6000
         TabIndex        =   9
         Top             =   1680
         Visible         =   0   'False
         Width           =   3700
         Begin VB.TextBox txtPO 
            Height          =   285
            Left            =   1150
            TabIndex        =   12
            Top             =   700
            Width           =   2175
         End
         Begin VB.CommandButton Aceptar 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   400
            TabIndex        =   11
            Top             =   1500
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2000
            TabIndex        =   10
            Top             =   1500
            Width           =   1335
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "NP"
            Height          =   255
            Left            =   600
            TabIndex        =   13
            Top             =   700
            Width           =   400
         End
      End
      Begin VB.Frame FraValorizar 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Valorizar Transferencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   2000
         TabIndex        =   3
         Top             =   1800
         Visible         =   0   'False
         Width           =   3735
         Begin VB.TextBox TxtDolares 
            Height          =   285
            Left            =   2160
            TabIndex        =   5
            Top             =   720
            Width           =   975
         End
         Begin FunctionsButtons.FunctButt FunctButt2 
            Height          =   510
            Left            =   600
            TabIndex        =   6
            Top             =   1200
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   900
            Custom          =   $"FrmDetalleStock.frx":0277
            Orientacion     =   0
            Style           =   0
            Language        =   0
            TypeImageList   =   0
            ControlWidth    =   1155
            ControlHeigth   =   490
            ControlSeparator=   110
         End
         Begin VB.TextBox TxtSoles 
            Height          =   285
            Left            =   2160
            TabIndex        =   4
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Importe Dolares"
            Height          =   255
            Left            =   600
            TabIndex        =   8
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Importe Soles"
            Height          =   255
            Left            =   600
            TabIndex        =   7
            Top             =   480
            Width           =   1335
         End
      End
      Begin GridEX20.GridEX DGridLista 
         Height          =   5100
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10905
         _ExtentX        =   19235
         _ExtentY        =   8996
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         BackColorBkg    =   12648384
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "FrmDetalleStock.frx":030D
         Column(2)       =   "FrmDetalleStock.frx":03D5
         FormatStylesCount=   6
         FormatStyle(1)  =   "FrmDetalleStock.frx":0479
         FormatStyle(2)  =   "FrmDetalleStock.frx":05B1
         FormatStyle(3)  =   "FrmDetalleStock.frx":0661
         FormatStyle(4)  =   "FrmDetalleStock.frx":0715
         FormatStyle(5)  =   "FrmDetalleStock.frx":07ED
         FormatStyle(6)  =   "FrmDetalleStock.frx":08A5
         ImageCount      =   0
         PrinterProperties=   "FrmDetalleStock.frx":0985
      End
   End
End
Attribute VB_Name = "FrmDetalleStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cod_tipmov As String, Cod_ClaOrdComp As String, Cod_Almacen As String, Num_MovStk As String
Public Ser_OrdComp As String, Cod_OrdComp As String
Public Paso As Boolean
Public sflg_adicionales As String
Public varValida_Factura As Boolean
Public varCod_Fabrica As String
Public varCod_OrdPro As String
Public varNum_SecOrd As String
Public varflg_secord As String
Public var_tipo As String, varTallaProv As String

Dim strSQL As String
Dim vCod_OrdPro As String
Public CODIGO As String, DESCRIPCION As String
Public vFlg_Almacen_Tejeduria As String, vcod_cencost As String, vFlg_Despacho_Acabado As String
Public vFLG_CREA_COMBINACION_ITEMS_TEJEDURIA As String
Public FLG_TRANSFERENCIA_EXTERNA As String
Public sFecmovstk  As Date
Dim indicegrilla  As Long
Public num_guia As String
Function ValidaFlag() As Boolean
ValidaFlag = True
If DevuelveCampo("select Flg_StatusVAL from Lg_MoviStk where Cod_Almacen='" & Trim(Right(FrmMovAlmacen.Almacen, 3)) & "' and Num_MovStk='" & Num_MovStk & "'", cConnect) = "S" Then
    MsgBox "Este registro no puede ser eliminado", vbInformation
   ValidaFlag = False
End If
End Function

Public Sub Datos(Accion As String, EsAccion As Boolean)
On Error GoTo hand

strSQL = "UP_Lg_MovstkItem '" & UCase(Accion) & "','" & Cod_Almacen & "','" & Num_MovStk & "','','','','','','','',0"

If EsAccion = False Then
    Set DGridLista.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    
    Me.DGridLista.Columns("Des_Item").Visible = False
    Me.DGridLista.Columns("Des_Comb").Visible = False
    Me.DGridLista.Columns("Des_Color").Visible = False
    Me.DGridLista.Columns("Des_Destino").Visible = False
    Me.DGridLista.Columns("Num_Secuencia").Visible = False
    Me.DGridLista.Columns("Cod_Item").Visible = False
    Me.DGridLista.Columns("Cod_Comb").Visible = False
    Me.DGridLista.Columns("Cod_OrdComp").Visible = False
    Me.DGridLista.Columns("Ser_OrdComp").Visible = False
    Me.DGridLista.Columns("Sec_OrdComp").Visible = False
    Me.DGridLista.Columns("cod_Color").Visible = False
    Me.DGridLista.Columns("Cod_Talla").Visible = False
    'Me.DGridLista.Columns("flg_adicionales").Visible = False
    Me.DGridLista.Columns("Descripcion").Visible = False
    Me.DGridLista.Columns("Cod_maquina_tejeduria").Visible = False
    Me.DGridLista.Columns("des_maquina_tejeduria").Width = 1400
    Me.DGridLista.Columns("peso_kgs").Width = 1000
    Me.DGridLista.RowSelected(indicegrilla) = True
  
End If


Exit Sub
hand:
ErrorHandler err, "Datos"
End Sub

Sub CARGA_DATOS()
If DGridLista.RowCount = 0 Then Exit Sub
FrmAddMovimAlmDet.TxtItem = DGridLista.Value(DGridLista.Columns("cod_item").Index)
FrmAddMovimAlmDet.TxtDesitem = DGridLista.Value(DGridLista.Columns("Des_Item").Index)
FrmAddMovimAlmDet.TxtCod_Comb.Text = Trim(DGridLista.Value(DGridLista.Columns("cod_comb").Index))
FrmAddMovimAlmDet.TxtDes_comb.Text = Trim(DGridLista.Value(DGridLista.Columns("des_comb").Index))
FrmAddMovimAlmDet.CmbColor = DGridLista.Value(DGridLista.Columns("cod_Color").Index)
FrmAddMovimAlmDet.TxtDetalle = DGridLista.Value(DGridLista.Columns("Des_Color").Index)
FrmAddMovimAlmDet.Txtcod_Destino = DGridLista.Value(DGridLista.Columns("cod_Destino").Index)
FrmAddMovimAlmDet.TxtDes_Destino = DGridLista.Value(DGridLista.Columns("Des_Destino").Index)
FrmAddMovimAlmDet.TxtCod_Medida = DGridLista.Value(DGridLista.Columns("cod_talla").Index)
FrmAddMovimAlmDet.TxtDes_Medida = DGridLista.Value(DGridLista.Columns("descripcion").Index)
FrmAddMovimAlmDet.TxtCantidad = CDbl(DGridLista.Value(DGridLista.Columns("Cantidad").Index))
FrmAddMovimAlmDet.TxtCodProv = DGridLista.Value(DGridLista.Columns("Cod_Prov").Index)
FrmAddMovimAlmDet.Num_Secuencia = DGridLista.Value(DGridLista.Columns("Num_Secuencia").Index)
FrmAddMovimAlmDet.Sec_OrdComp = Trim(DGridLista.Value(DGridLista.Columns("Sec_OrdComp").Index))
FrmAddMovimAlmDet.Cant = DGridLista.Value(DGridLista.Columns("Cantidad").Index)
FrmAddMovimAlmDet.item = DGridLista.Value(DGridLista.Columns("cod_item").Index)
FrmAddMovimAlmDet.CombinacionX = DGridLista.Value(DGridLista.Columns("cod_comb").Index)
FrmAddMovimAlmDet.TallaX = DGridLista.Value(DGridLista.Columns("Cod_Talla").Index)
FrmAddMovimAlmDet.Color = DGridLista.Value(DGridLista.Columns("cod_color").Index)
FrmAddMovimAlmDet.TxtCod_EstCli = DGridLista.Value(DGridLista.Columns("Cod_EstCli").Index)
FrmAddMovimAlmDet.txtPeso = CDbl(DGridLista.Value(DGridLista.Columns("peso_kgs").Index))

If FrmAddMovimAlmDet.sflg_adicionales = "*" Then

FrmAddMovimAlmDet.TxtOP = DGridLista.Value(DGridLista.Columns("Cod_ORDPRO").Index)
FrmAddMovimAlmDet.TxtEstilo = DevuelveCampo("SELECT b.Des_EstPro FROM   ES_OrdPro  a , ES_EstPRo b WHERE  a.Cod_EstPro = b.Cod_EstPRo AND a.Cod_Fabrica= '001' AND a.Cod_OrdPro = '" & DGridLista.Value(DGridLista.Columns("Cod_ORDPRO").Index) & "'", cConnect)

End If

End Sub



Private Sub Aceptar_Click()

If txtPO.Text <> "" Then
 vCod_OrdPro = Trim(txtPO.Text)
 Load FrmAviosPendientesxDespachar
 FrmAviosPendientesxDespachar.varCod_Fabrica = Me.varCod_Fabrica
 FrmAviosPendientesxDespachar.varCod_OrdPro = vCod_OrdPro
 FrmAviosPendientesxDespachar.varCod_TipMov = Me.cod_tipmov
 FrmAviosPendientesxDespachar.varCOD_ALMACEN = Me.Cod_Almacen
 FrmAviosPendientesxDespachar.varNUM_MOVSTK = Me.Num_MovStk
 FrmAviosPendientesxDespachar.CARGA_GRID
 FrmAviosPendientesxDespachar.Show 1
 Set FrmAviosPendientesxDespachar = Nothing
Else
MsgBox ("Ingrese una orden NP")
End If

End Sub

Private Sub Command2_Click()
txtPO.Text = ""
FraPO.Visible = False
End Sub



 
Private Sub DGridLista_Click()
indicegrilla = DGridLista.Row
End Sub

Private Sub Form_Load()
On Error GoTo hand

'Datos "V", False
Me.varTallaProv = ""

'If vemp1 = "09" Then
  'FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp1, Me.Name) '
'End If

 indicegrilla = 1
Exit Sub
hand:
ErrorHandler err, "Form_Load()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim fnuevo As Boolean
Dim Secuencia As String

Select Case ActionName
Case "ADICIONAR"
    If Me.varflg_secord = "*" Then
        Load frmReqCfOrdPro
        frmReqCfOrdPro.varCod_Fabrica = Me.varCod_Fabrica
        frmReqCfOrdPro.varCod_OrdPro = Me.varCod_OrdPro
        frmReqCfOrdPro.varNum_SecOrd = Me.varNum_SecOrd
        frmReqCfOrdPro.varCod_TipMov = Me.cod_tipmov
        frmReqCfOrdPro.varCOD_ALMACEN = Me.Cod_Almacen
        frmReqCfOrdPro.varNUM_MOVSTK = Me.Num_MovStk
        'frmReqCfOrdPro.varSec_OrdComp = Me.Ser_OrdComp
        frmReqCfOrdPro.CARGA_GRID
        frmReqCfOrdPro.Show 1
        Set frmReqCfOrdPro = Nothing
    ElseIf FLG_TRANSFERENCIA_EXTERNA = "S" Then
     
     Call ADICIONAR_REGISTRO_ESTOP
    
    ElseIf Me.vFlg_Despacho_Acabado = "S" Then
    
         If Me.cod_tipmov = "S08" Then
        
          FraPO.Visible = True
          txtPO.SetFocus
         
         Else
        Load FrmAviosPendientesxDespachar
        FrmAviosPendientesxDespachar.varCod_Fabrica = Me.varCod_Fabrica
        FrmAviosPendientesxDespachar.varCod_OrdPro = Me.varCod_OrdPro
        FrmAviosPendientesxDespachar.varCod_TipMov = Me.cod_tipmov
        FrmAviosPendientesxDespachar.varCOD_ALMACEN = Me.Cod_Almacen
        FrmAviosPendientesxDespachar.varNUM_MOVSTK = Me.Num_MovStk
        FrmAviosPendientesxDespachar.CARGA_GRID
        FrmAviosPendientesxDespachar.Show 1
        Set FrmAviosPendientesxDespachar = Nothing
         End If
         
       
    Else
        If Me.Cod_ClaOrdComp <> "" Then
            Load FrmReqCompra
            FrmReqCompra.cod_tipmov = Me.cod_tipmov
            FrmReqCompra.Ser_OrdComp = Ser_OrdComp
            FrmReqCompra.Cod_OrdComp = Cod_OrdComp
            FrmReqCompra.BUSCAR
            FrmReqCompra.Show 1
            Set FrmReqCompra = Nothing
        Else
           
            
            FrmAddMovimAlmDet.sflg_adicionales = sflg_adicionales
            Load FrmAddMovimAlmDet
            Set FrmAddMovimAlmDet.oParent = Me
            FrmAddMovimAlmDet.Cod_Almacen = Me.Cod_Almacen
            FrmAddMovimAlmDet.Num_MovStk = Me.Num_MovStk
            FrmAddMovimAlmDet.cod_tipmov = Me.cod_tipmov
            FrmAddMovimAlmDet.Ser_OrdComp = Me.Ser_OrdComp
            FrmAddMovimAlmDet.Cod_OrdComp = Me.Cod_OrdComp
            FrmAddMovimAlmDet.var_tipo = Me.var_tipo
            FrmAddMovimAlmDet.varNum_SecOrd = Me.varNum_SecOrd
            FrmAddMovimAlmDet.Limpia
            FrmAddMovimAlmDet.Habilita
            FrmAddMovimAlmDet.Estado = "I"
            FrmAddMovimAlmDet.vFlg_Almacen_Tejeduria = Me.vFlg_Almacen_Tejeduria
            If UCase(Trim(Me.var_tipo)) = "S" And Trim(Me.vcod_cencost) <> "" And UCase(vFlg_Almacen_Tejeduria) = "S" Then
                FrmAddMovimAlmDet.TxtCod_Maquina.Visible = True
                FrmAddMovimAlmDet.TxtDes_Maquina.Visible = True
                FrmAddMovimAlmDet.Etiqueta(0).Visible = True
            Else
                FrmAddMovimAlmDet.TxtCod_Maquina.Visible = False
                FrmAddMovimAlmDet.TxtDes_Maquina.Visible = False
                FrmAddMovimAlmDet.Etiqueta(0).Visible = False
            End If
            If UCase(Trim(Me.vFLG_CREA_COMBINACION_ITEMS_TEJEDURIA)) = "S" Then
                FrmAddMovimAlmDet.Etiqueta(1).Visible = True
                FrmAddMovimAlmDet.TxtGlosa.Visible = True
            Else
                FrmAddMovimAlmDet.Etiqueta(1).Visible = False
                FrmAddMovimAlmDet.TxtGlosa.Visible = False
            End If
            
            If Trim(DevuelveCampo("SELECT COD_TIPMOVREL FROM LG_TIPOSMOV WHERE Cod_TipMov='" & cod_tipmov & "'", cConnect)) <> "" Then
                FrmAddMovimAlmDet.CmdTransferir.Enabled = True
                FrmAddMovimAlmDet.cmdTransfMismoItem.Enabled = True
            Else
                FrmAddMovimAlmDet.CmdTransferir.Enabled = False
                FrmAddMovimAlmDet.cmdTransfMismoItem.Enabled = False
            End If
            FrmAddMovimAlmDet.Caption = FrmAddMovimAlmDet.Caption & "  " & Me.Num_MovStk
            If Me.vFlg_Almacen_Tejeduria = "S" Then
                FrmAddMovimAlmDet.Deshabilita_Tej
            End If
            FrmAddMovimAlmDet.Show vbModal
            Set FrmAddMovimAlmDet = Nothing
        End If
    End If
    Datos "v", False
Case "MODIFICAR"
    If DGridLista.RowCount = 0 Then Exit Sub
    If Me.varValida_Factura = False Then
        MsgBox "No se puede modificar el registro por que el Mov. Almacen posee una factura relacionada.", vbInformation, "Mensaje"
        Exit Sub
    End If
    FrmAddMovimAlmDet.sflg_adicionales = sflg_adicionales
    Load FrmAddMovimAlmDet
    FrmAddMovimAlmDet.Estado = "U"
    FrmAddMovimAlmDet.Cod_Almacen = Me.Cod_Almacen
    FrmAddMovimAlmDet.Num_MovStk = Me.Num_MovStk
    FrmAddMovimAlmDet.cod_tipmov = Me.cod_tipmov
    FrmAddMovimAlmDet.Ser_OrdComp = Me.Ser_OrdComp
    FrmAddMovimAlmDet.Cod_OrdComp = Me.Cod_OrdComp
    FrmAddMovimAlmDet.var_tipo = Me.var_tipo
    FrmAddMovimAlmDet.varNum_SecOrd = Me.varNum_SecOrd
    CARGA_DATOS
    FrmAddMovimAlmDet.Deshabilita
    FrmAddMovimAlmDet.vFlg_Almacen_Tejeduria = Me.vFlg_Almacen_Tejeduria
    If UCase(Trim(Me.var_tipo)) = "S" And Trim(Me.vcod_cencost) <> "" And vFlg_Almacen_Tejeduria = "S" Then
        FrmAddMovimAlmDet.TxtCod_Maquina.Visible = True
        FrmAddMovimAlmDet.TxtDes_Maquina.Visible = True
        FrmAddMovimAlmDet.Etiqueta(0).Visible = True
        FrmAddMovimAlmDet.TxtCod_Maquina.Text = DGridLista.Value(DGridLista.Columns("cod_maquina_tejeduria").Index)
        FrmAddMovimAlmDet.TxtDes_Maquina.Text = DGridLista.Value(DGridLista.Columns("des_maquina_tejeduria").Index)
        FrmAddMovimAlmDet.TxtCod_Maquina.Enabled = True
        FrmAddMovimAlmDet.TxtDes_Maquina.Enabled = True
    Else
        FrmAddMovimAlmDet.TxtCod_Maquina.Visible = False
        FrmAddMovimAlmDet.TxtDes_Maquina.Visible = False
        FrmAddMovimAlmDet.Etiqueta(0).Visible = False
    End If
    If Me.vFlg_Almacen_Tejeduria = "S" Then
        FrmAddMovimAlmDet.Deshabilita_Tej
    End If
    FrmAddMovimAlmDet.TxtCantidad.Enabled = True
    FrmAddMovimAlmDet.txtPeso.Enabled = True
    FrmAddMovimAlmDet.Show vbModal
    Set FrmAddMovimAlmDet = Nothing
    Datos "v", False
Case "ELIMINAR"

    If DGridLista.RowCount = 0 Then Exit Sub
    If Me.varValida_Factura = False Then
        MsgBox "No se puede eliminar el registro por que el Mov. Almacen posee una factura relacionada.", vbInformation, "Mensaje"
        Exit Sub
    End If
    
If Trim(Cod_Almacen) = "69" Then
    
 If MsgBox("Esta seguro de eliminar las prendas seleccionadas", vbInformation + vbYesNo, "IMPORTANTE") = vbYes Then
    strSQL = "LG_UP_ELIMINA_LG_MOVISTKITEM '" & Me.Cod_Almacen & "','" & Me.Num_MovStk & "','" & DGridLista.Value(DGridLista.Columns("NUM_SECUENCIA").Index) & "','" & vusu & "'"
    Call ExecuteSQL(cConnect, strSQL)
 End If
Else
    FrmAddMovimAlmDet.sflg_adicionales = sflg_adicionales
    Load FrmAddMovimAlmDet
    FrmAddMovimAlmDet.Cod_Almacen = Me.Cod_Almacen
    FrmAddMovimAlmDet.Num_MovStk = Me.Num_MovStk
    FrmAddMovimAlmDet.cod_tipmov = Me.cod_tipmov
    FrmAddMovimAlmDet.Ser_OrdComp = Me.Ser_OrdComp
    FrmAddMovimAlmDet.Cod_OrdComp = Me.Cod_OrdComp
    FrmAddMovimAlmDet.var_tipo = Me.var_tipo
    FrmAddMovimAlmDet.varNum_SecOrd = Me.varNum_SecOrd
    'FrmAddMovimAlmDet.OP = DGridLista.Columns("cod_maquina_tejeduria").Index
    FrmAddMovimAlmDet.Estado = "D"
    If UCase(Trim(Me.var_tipo)) = "S" And Trim(Me.vcod_cencost) <> "" Then
        FrmAddMovimAlmDet.TxtCod_Maquina.Visible = True
        FrmAddMovimAlmDet.TxtDes_Maquina.Visible = True
        FrmAddMovimAlmDet.Etiqueta(0).Visible = True
        FrmAddMovimAlmDet.TxtCod_Maquina.Text = DGridLista.Value(DGridLista.Columns("cod_maquina_tejeduria").Index)
        FrmAddMovimAlmDet.TxtDes_Maquina.Text = DGridLista.Value(DGridLista.Columns("des_maquina_tejeduria").Index)
    Else
        FrmAddMovimAlmDet.TxtCod_Maquina.Visible = False
        FrmAddMovimAlmDet.TxtDes_Maquina.Visible = False
        FrmAddMovimAlmDet.Etiqueta(0).Visible = False
    End If
    CARGA_DATOS
    FrmAddMovimAlmDet.Deshabilita
    If Me.vFlg_Almacen_Tejeduria = "S" Then
        FrmAddMovimAlmDet.Deshabilita_Tej
    End If
    FrmAddMovimAlmDet.Show vbModal
    Set FrmAddMovimAlmDet = Nothing
    Datos "v", False

End If
    
Case "VALORIZAR"
    FraValorizar.Visible = True
    TxtSoles.SetFocus
    TxtSoles.Text = DevuelveCampo("select imp_factura from lg_movistkitem where cod_almacen ='" & Me.Cod_Almacen & "' and num_movstk='" & Me.Num_MovStk & "' and num_secuencia ='" & DGridLista.Value(DGridLista.Columns("Num_Secuencia").Index) & "'", cConnect)
    TxtDolares.Text = DevuelveCampo("select imp_factura_dolares from lg_movistkitem where cod_almacen ='" & Me.Cod_Almacen & "' and num_movstk='" & Me.Num_MovStk & "' and num_secuencia ='" & DGridLista.Value(DGridLista.Columns("Num_Secuencia").Index) & "'", cConnect)
Case "SALIDAS"
    If DGridLista.RowCount = 0 Then Exit Sub
    Load FrmVerSalidasxDua
    FrmVerSalidasxDua.Ser_OrdComp = Me.Ser_OrdComp
    FrmVerSalidasxDua.Cod_OrdComp = Me.Cod_OrdComp
    FrmVerSalidasxDua.Sec_OrdComp = DGridLista.Value(DGridLista.Columns("Sec_OrdComp").Index)
    FrmVerSalidasxDua.CARGA_GRID
    FrmVerSalidasxDua.Show vbModal
    Set FrmVerSalidasxDua = Nothing
Case "AUDCALRECEPAVIOS"
    'If DGridLista.RowCount = 0 Then Exit Sub
    'Load frmAuditoriaCalidadRecepcionAvios
    'frmAuditoriaCalidadRecepcionAvios.Cod_Almacen = Cod_Almacen
    'frmAuditoriaCalidadRecepcionAvios.Num_MovStk = Num_MovStk
    'frmAuditoriaCalidadRecepcionAvios.Num_Secuencia = DGridLista.Value(DGridLista.Columns("Num_Secuencia").Index)
    'frmAuditoriaCalidadRecepcionAvios.Cantidad = DGridLista.Value(DGridLista.Columns("Cantidad").Index)
    'frmAuditoriaCalidadRecepcionAvios.Show 1
    'Set frmAuditoriaCalidadRecepcionAvios = Nothing
    'Secuencia = DGridLista.Value(DGridLista.Columns("Item").Index)
    'Datos "V", False
    'fnuevo = DGridLista.Find(DGridLista.Value(DGridLista.Columns("Secuencia").Index), jgexGreaterThanOrEqualTo, Secuencia)
Case "SALIR"
    Unload Me
End Select
End Sub

Private Sub ADICIONAR_REGISTRO_ESTOP()
    On Error GoTo SALTO_ERROR
    Dim sCadSQL As String
    
    
    sCadSQL = "SELECT FEC_EMISION_VOUCHER FROM LG_MOVISTK  WHERE COD_ALMACEN ='" & Cod_Almacen & "' AND NUM_MOVSTK = '" & Num_MovStk & "'"
    If Not IsNull(DevuelveCampo(sCadSQL, cConnect)) Then
        MsgBox "Voucher ya fue impreso, revertir Voucher para modificar", vbCritical + vbInformation, "Aviso"
        Exit Sub
    End If

    Dim oRs As New Recordset

    sCadSQL = "EXEC CF_MUESTRA_DATOS_TIPMOV '" & cod_tipmov & "'"
    Set oRs = CargarRecordSetDesconectado(sCadSQL, cConnect)

    If oRs.RecordCount = 0 Then Exit Sub

        Dim sTip_PtMP As String
        Dim sTip_Accion As String
        Dim sTipo_MovConfec  As String
        Dim sCod_TipOrdPro As String
        Dim sCod_TipAnx  As String
        Dim sFlg_norealizado  As String
        Dim flg_transferencia  As String
        Dim FLG_TRANSFERENCIA_EXTERNA  As String
        Dim sCod_ClaMov As String


    
    With oRs
        sTip_PtMP = .Fields("Tip_PTMP").Value
        sTip_Accion = .Fields("Tip_Accion").Value
        sTipo_MovConfec = .Fields("Cod_MovCost").Value
        sCod_TipOrdPro = .Fields("Cod_TipOrdPro").Value
        sCod_TipAnx = .Fields("Cod_TipAnx").Value
        sFlg_norealizado = .Fields("Flg_NoRealizado").Value
        flg_transferencia = .Fields("flg_transferencia").Value
        FLG_TRANSFERENCIA_EXTERNA = .Fields("FLG_TRANSFERENCIA_EXTERNA").Value
        sCod_ClaMov = .Fields("Cod_Clamov").Value
    End With
    
   ' If sguia <> "" And sCod_ClaMov = "S" And sTip_Accion = "E" And Flg_Protos <> "*" Then
   '     Aviso "Movimiento no se puede Modificar, Guia ya fue Impresa", 2
   '     Exit Sub
   ' End If
       
    With FrmAaMuestraPrendas
        Set .oParent = Me
        .sCod_Almacen = Cod_Almacen
        .snum_movistkActual = Num_MovStk
        .num_guia = num_guia
        .sTipo_MovConfec = sTipo_MovConfec
        .sTip_Accion = sTip_Accion
        .sFec_movActual = sFecmovstk
        .sCod_TipMov = cod_tipmov
        .Cod_ClaMov = sCod_ClaMov
        .Show vbModal
    End With
    Load FrmAaMuestraPrendas
    Set FrmAaMuestraPrendas = Nothing

    Exit Sub
SALTO_ERROR:
    MsgBox err.Description, vbCritical, Me.Caption
'
End Sub


Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    Call Grabar
    FraValorizar.Visible = False
Case "CANCELAR"
    FraValorizar.Visible = False
End Select
End Sub







Private Sub TxtDolares_GotFocus()
SelectionText TxtDolares
End Sub

Private Sub TxtDolares_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtDolares, KeyAscii, True, 2)
End If
End Sub

Private Sub txtPO_Keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPO = Format(txtPO.Text, "00000")
    txtPO = Trim(DevuelveCampo("Select dbo.uf_devuelvecodigo(5," & IIf(Trim(txtPO) = "", 0, txtPO) & ")", cConnect))
    Aceptar.SetFocus
End If
End Sub
 



Private Sub TxtSoles_GotFocus()
SelectionText TxtSoles
End Sub

Private Sub TxtSoles_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtSoles, KeyAscii, True, 2)
End If
End Sub

Sub Grabar()
On Error GoTo errGrabar

If Trim(TxtSoles.Text) = "" Then
    TxtSoles.Text = "0"
End If

If Trim(TxtDolares.Text) = "" Then
    TxtDolares.Text = "0"
End If

strSQL = "LG_MovistkItem_Valoriza_Transferencia '" & Me.Cod_Almacen & "','" & Me.Num_MovStk & "','" & DGridLista.Value(DGridLista.Columns("Num_Secuencia").Index) & "'," & _
        CDbl(TxtSoles.Text) & "," & CDbl(TxtDolares.Text)
ExecuteSQL cConnect, strSQL

TxtSoles.Text = ""
TxtDolares.Text = ""
Exit Sub
errGrabar:
    ErrorHandler err, "Grabar"
End Sub
