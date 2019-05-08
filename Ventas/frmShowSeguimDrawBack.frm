VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmShowSeguimDrawBack 
   Caption         =   "Seguimiento Draw Back"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraGeneral 
      Caption         =   "Argumentos de Búsqueda General"
      Height          =   2085
      Left            =   45
      TabIndex        =   13
      Top             =   30
      Width           =   11745
      Begin VB.OptionButton optAll 
         Caption         =   "Pendientes + En Trámite"
         Height          =   255
         Left            =   180
         TabIndex        =   35
         Top             =   1710
         Width           =   2310
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   555
         Left            =   9600
         TabIndex        =   26
         Top             =   420
         Width           =   1185
      End
      Begin VB.TextBox txtNum_Dias 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   2730
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "60"
         Top             =   270
         Width           =   540
      End
      Begin VB.Frame Frame1 
         Caption         =   "Fecha"
         Height          =   690
         Left            =   4995
         TabIndex        =   21
         Top             =   405
         Width           =   3915
         Begin MSComCtl2.DTPicker dtpFecIni 
            Height          =   315
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Format          =   23986177
            CurrentDate     =   37543
         End
         Begin MSComCtl2.DTPicker dtpFecFin 
            Height          =   315
            Left            =   2310
            TabIndex        =   23
            Top             =   240
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Format          =   23986177
            CurrentDate     =   37543
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta :"
            Height          =   255
            Left            =   1680
            TabIndex        =   24
            Top             =   270
            Width           =   615
         End
      End
      Begin VB.TextBox txtFlg_Status 
         BackColor       =   &H80000014&
         Height          =   315
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   1
         Top             =   630
         Width           =   360
      End
      Begin VB.TextBox txtDes_Status 
         BackColor       =   &H80000014&
         Height          =   315
         Left            =   2385
         MaxLength       =   30
         TabIndex        =   2
         Top             =   630
         Width           =   2445
      End
      Begin VB.OptionButton optEstado 
         Caption         =   "Por Estado"
         Height          =   255
         Left            =   180
         TabIndex        =   20
         Top             =   720
         Width           =   1605
      End
      Begin VB.OptionButton optDocum 
         Caption         =   "Documento Específico:"
         Height          =   360
         Left            =   180
         TabIndex        =   19
         Top             =   1185
         Width           =   1560
      End
      Begin VB.TextBox txtNum_Desde 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   6945
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1155
         Width           =   1440
      End
      Begin VB.TextBox txtSer_Desde 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   5565
         MaxLength       =   3
         TabIndex        =   5
         Top             =   1155
         Width           =   540
      End
      Begin VB.TextBox txtCod_TipDoc2 
         BackColor       =   &H80000014&
         Height          =   315
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "FA"
         Top             =   1155
         Width           =   360
      End
      Begin VB.TextBox txtDes_TipDoc2 
         BackColor       =   &H80000014&
         Height          =   315
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1155
         Width           =   2445
      End
      Begin VB.OptionButton optCriticos 
         Caption         =   "Críticos"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   285
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.Label Label1 
         Caption         =   "Nro Dias"
         Height          =   210
         Left            =   1995
         TabIndex        =   25
         Top             =   330
         Width           =   810
      End
      Begin VB.Label Label7 
         Caption         =   "Serie: "
         Height          =   210
         Left            =   5025
         TabIndex        =   18
         Top             =   1230
         Width           =   510
      End
      Begin VB.Label Label9 
         Caption         =   "Número :"
         Height          =   225
         Left            =   6195
         TabIndex        =   17
         Tag             =   "Number"
         Top             =   1230
         Width           =   645
      End
   End
   Begin VB.Frame fraDrawBack 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cambio de Estado"
      Height          =   3000
      Left            =   1170
      TabIndex        =   7
      Top             =   3435
      Visible         =   0   'False
      Width           =   9660
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Seleccione Nuevo Estado"
         Height          =   1605
         Left            =   630
         TabIndex        =   28
         Top             =   840
         Width           =   6225
         Begin VB.OptionButton optRetornaraPendiente 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Retornar a Pendiente de Envío"
            Height          =   240
            Left            =   135
            TabIndex        =   32
            Top             =   1230
            Width           =   2625
         End
         Begin VB.OptionButton optRetornaraEnTramite 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Retornar a En Trámite"
            Height          =   240
            Left            =   120
            TabIndex        =   31
            Top             =   900
            Width           =   2625
         End
         Begin VB.OptionButton optCobrado 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Cobrado"
            Height          =   240
            Left            =   135
            TabIndex        =   30
            Top             =   585
            Width           =   2625
         End
         Begin VB.OptionButton optEnviaraAduana 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Enviar a Aduana (En Trámite)"
            Height          =   240
            Left            =   135
            TabIndex        =   29
            Top             =   285
            Width           =   2625
         End
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   315
            Left            =   4410
            TabIndex        =   33
            Top             =   630
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Format          =   23986177
            CurrentDate     =   37543
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Fecha"
            Height          =   270
            Left            =   4410
            TabIndex        =   34
            Top             =   390
            Width           =   1500
         End
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   480
         Left            =   7095
         TabIndex        =   27
         Top             =   1170
         Width           =   1740
      End
      Begin VB.TextBox txtDescrip_Tipdoc_DB 
         BackColor       =   &H80000014&
         Height          =   330
         Left            =   1650
         MaxLength       =   30
         TabIndex        =   10
         Top             =   375
         Width           =   1980
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   480
         Left            =   7080
         TabIndex        =   9
         Top             =   1695
         Width           =   1740
      End
      Begin VB.TextBox txtEstado 
         BackColor       =   &H80000014&
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   4470
         MaxLength       =   30
         TabIndex        =   8
         Top             =   360
         Width           =   4380
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Documento:"
         Height          =   315
         Left            =   675
         TabIndex        =   12
         Tag             =   "Document Type"
         Top             =   450
         Width           =   1020
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Estado: "
         Height          =   315
         Left            =   3810
         TabIndex        =   11
         Tag             =   "Document Type"
         Top             =   420
         Width           =   1020
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4110
      Left            =   45
      TabIndex        =   15
      Top             =   2175
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   7250
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmShowSeguimDrawBack.frx":0000
      Column(2)       =   "frmShowSeguimDrawBack.frx":00C8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmShowSeguimDrawBack.frx":016C
      FormatStyle(2)  =   "frmShowSeguimDrawBack.frx":02A4
      FormatStyle(3)  =   "frmShowSeguimDrawBack.frx":0354
      FormatStyle(4)  =   "frmShowSeguimDrawBack.frx":0408
      FormatStyle(5)  =   "frmShowSeguimDrawBack.frx":04E0
      FormatStyle(6)  =   "frmShowSeguimDrawBack.frx":0598
      FormatStyle(7)  =   "frmShowSeguimDrawBack.frx":0678
      FormatStyle(8)  =   "frmShowSeguimDrawBack.frx":0724
      ImageCount      =   0
      PrinterProperties=   "frmShowSeguimDrawBack.frx":07D4
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   8055
      TabIndex        =   16
      Top             =   6390
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   900
      Custom          =   $"frmShowSeguimDrawBack.frx":09AC
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   270
      Top             =   6525
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmShowSeguimDrawBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNum_Corre As String
Dim OP_Opcion  As String
Dim sSql As String
Public Codigo As String
Public Descripcion As String


Private Sub cmdEnviaraEnTramite_Click()

End Sub

Private Sub cmdAceptar_Click()
    If Not optCobrado And Not optEnviaraAduana And Not optRetornaraEnTramite And Not optRetornaraPendiente Then
        Aviso "Seleccione Estado Destino", 1
        Exit Sub
    End If
    
    CambiarStatus sSql, DTPFecha
End Sub

Private Sub cmdSalir_Click()
    Me.fraDrawBack.Visible = False
End Sub

Private Sub Form_Load()
    OP_Opcion = "1"
    dtpFecIni = Date - 30
    dtpFecFin = Date
    DTPFecha = Date
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "DRAWBACK"
            sSql = ""
            optEnviaraAduana.Value = False
            optCobrado.Value = False
            optRetornaraEnTramite.Value = False
            optRetornaraPendiente.Value = False
            
            sNum_Corre = GridEX1.Value(GridEX1.Columns("NUM_CORRE").Index)
            txtDescrip_Tipdoc_DB = GridEX1.Value(GridEX1.Columns("COD_TIPDOC").Index) & " " & GridEX1.Value(GridEX1.Columns("DOCUMENTO").Index)
            txtEstado = GridEX1.Value(GridEX1.Columns("DES_STATUS").Index)
            Me.fraDrawBack.Visible = True
        Case "IMPRIMIR"
            Reporte
        Case "SALIR"
            Unload Me
    End Select
End Sub

Private Sub cmdBuscar_Click()
  Buscar
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub Buscar()

Dim sSql As String

sSql = "CN_Ventas_Seguimiento_DrawBack '$',$,'$','$','$','$','$','$'"
sSql = VBsprintf(sSql, OP_Opcion, txtNum_Dias.Text, txtFlg_Status.Text, txtCod_TipDoc2, txtSer_Desde, txtNum_Desde, dtpFecIni, dtpFecFin)

GridEX1.ClearFields

GridEX1.DefaultGroupMode = jgexDGMExpanded
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSql, cCONNECT)


Configurar

End Sub
Sub Configurar()

GridEX1.ContinuousScroll = True

End Sub


Private Sub optAll_Click()
    OP_Opcion = "4"
    cmdBuscar.SetFocus
End Sub

Private Sub optCobrado_Click()
    sSql = "CN_VENTAS_CAMBIAR_STATUS_REGISTRO_COBRO_DRAWBACK"
End Sub

Private Sub optCriticos_Click()
        OP_Opcion = "1"
        txtNum_Dias.SetFocus
End Sub

Private Sub optDocum_Click()
    OP_Opcion = "3"
    txtCod_TipDoc2.SetFocus
End Sub

Private Sub optEnviaraAduana_Click()
    sSql = "CN_VENTAS_CAMBIAR_STATUS_REGISTRO_ENVIO_DRAWBACK"

End Sub

Private Sub optEstado_Click()
    OP_Opcion = "2"
    txtFlg_Status.SetFocus
End Sub


Private Sub CambiarStatus(sSql As String, sFecha As String)
On Error GoTo errx
Dim rs As ADODB.Recordset
Dim vResp As Variant

sSql = sSql & "'$','$','$','$'"

vResp = MsgBox("Desea Cambiar de Estado al Documento Indicado : " & txtDescrip_Tipdoc_DB & " ? ", vbOKCancel + vbQuestion, "Confirmación")

If vResp <> vbOK Then
    Exit Sub
End If


sSql = VBsprintf(sSql, sNum_Corre, vusu, ComputerName(), sFecha)

ExecuteCommandSQL cCONNECT, sSql
Mensaje kMESSAGE_INF_PROCESS_SATISFACTO

fraDrawBack.Visible = False
Buscar

Exit Sub
errx:
    errores Err.Number
End Sub

Private Sub optRetornaraEnTramite_Click()
    sSql = "CN_VENTAS_CAMBIAR_STATUS_REGISTRO_COBRO_A_ENVIO_DRAWBACK"
End Sub

Private Sub optRetornaraPendiente_Click()
    sSql = "CN_VENTAS_CAMBIAR_STATUS_REGISTRO_PENDIENTE_DRAWBACK"
End Sub

Sub Reporte()
On Error GoTo error
Dim sSql As String
Dim oo As Object
Dim Ruta As String

If GridEX1.RowCount = 0 Then Exit Sub
Ruta = vRuta & "\RptSeguimDrawBack.xlt"

Set oo = CreateObject("excel.application")
oo.Workbooks.Open Ruta
oo.Visible = True
oo.DisplayAlerts = False
        
oo.Run "Reporte", GridEX1.ADORecordset
Set oo = Nothing

Exit Sub
error:
    errores Err.Number
End Sub


Private Sub txtCod_TipDoc2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        
        Set frmBusqGeneral.oParent = Me
            frmBusqGeneral.SQuery = "SELECT COD_TIPDOC AS CODIGO, DES_TIPDOC AS DESCRIPCION , DOC_SUNAT AS TIPO FROM CN_TIPOSDOCUM WHERE COD_TIPDOC LIKE '%" & Trim(txtCod_TipDoc2.Text) & "%'"
            frmBusqGeneral.Cargar_Datos
            If frmBusqGeneral.DGridLista.RowCount > 1 Then
                frmBusqGeneral.Show 1
            Else
                frmBusqGeneral.cmdAceptar_Click
            End If
        If Codigo <> "" Then
            txtCod_TipDoc2.Text = Codigo
            txtDes_TipDoc2.Text = Descripcion
            If Me.Visible Then
                txtSer_Desde.SetFocus
            End If
            
        Else
            txtCod_TipDoc2.Text = ""
            txtDes_TipDoc2.Text = ""
        End If
        Codigo = ""
        Descripcion = ""
        
    End If
End Sub


Private Sub txtFlg_Status_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
                
        Set frmBusqGeneral.oParent = Me
            frmBusqGeneral.SQuery = "SELECT FLG_STATUS AS CODIGO, DES_STATUS AS DESCRIPCION FROM CN_VENTAS_Status_DrawBack  WHERE FLG_STATUS like '%" & Trim(txtFlg_Status.Text) & "%'"
            frmBusqGeneral.Cargar_Datos
            If frmBusqGeneral.DGridLista.RowCount > 1 Then
                frmBusqGeneral.Show 1
            Else
                frmBusqGeneral.cmdAceptar_Click
            End If
        If Codigo <> "" Then
            txtFlg_Status.Text = Codigo
            txtDes_Status.Text = Descripcion
            If Me.Visible Then
                dtpFecIni.SetFocus
            End If
            
        Else
            txtFlg_Status.Text = ""
            txtDes_Status.Text = ""
        End If
        Codigo = ""
        Descripcion = ""
        
    End If
End Sub


Private Sub txtNum_Desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdBuscar.SetFocus
    End If
End Sub

Private Sub txtNum_Dias_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdBuscar.SetFocus
    End If
End Sub

Private Sub txtSer_Desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtNum_Desde.SetFocus
    End If
End Sub
