VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmShowTG_Embarque_Detalle 
   Caption         =   "Detalle Embarque Prendas"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   3000
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txt_npac 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "N° Packing"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   3855
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   6800
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmShowTG_Embarque_Detalle.frx":0000
      Column(2)       =   "frmShowTG_Embarque_Detalle.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmShowTG_Embarque_Detalle.frx":016C
      FormatStyle(2)  =   "frmShowTG_Embarque_Detalle.frx":02A4
      FormatStyle(3)  =   "frmShowTG_Embarque_Detalle.frx":0354
      FormatStyle(4)  =   "frmShowTG_Embarque_Detalle.frx":0408
      FormatStyle(5)  =   "frmShowTG_Embarque_Detalle.frx":04E0
      FormatStyle(6)  =   "frmShowTG_Embarque_Detalle.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmShowTG_Embarque_Detalle.frx":0678
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   2910
      Left            =   9180
      TabIndex        =   1
      Top             =   30
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   5133
      Custom          =   $"frmShowTG_Embarque_Detalle.frx":0850
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmShowTG_Embarque_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lNum_Embarque As Long

Private Sub cmdBuscar_Click()
    Buscar
End Sub


Public Function Buscar() As Boolean
On Error GoTo errores
Dim ssql As String
Dim vBookmark As Variant

ssql = "TG_Embarques_Prendas_Muestra '$'"
ssql = VBsprintf(ssql, lNum_Embarque)
  
vBookmark = GridEX1.Row
GridEX1.ClearFields

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(ssql, cCONNECT)

GridEX1.Row = vBookmark


GridEX1.ContinuousScroll = True

GridEX1.FrozenColumns = 3
Exit Function

errores:
    errores err.Number
End Function


Private Sub cmdAceptar_Click()
On Error GoTo Fall
Dim stSQL As String
Dim packing As String
Dim Adors As ADODB.Recordset
    packing = (Me.txt_npac.Text)
    If packing = "" Then
        packing = 0
    End If
    stSQL = "TG_Genera_Embarques_Prendas_por_Packing " & lNum_Embarque & "," & packing & ""
    Set Adors = CargarRecordSetDesconectado(stSQL, cCONNECT)
            
    Mensaje kMESSAGE_INF_PROCESS_SATISFACTO
    Buscar
    Me.txt_npac.Text = ""
    Me.Frame1.Visible = False
    Set Adors = Nothing
Fall:
Set Adors = Nothing
MsgBox err.Number
End Sub

Private Sub cmdCancelar_Click()
Me.Frame1.Visible = False
Me.txt_npac.Text = ""
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ADICIONAR"
            Load frmTG_Embarque_Prendas
            frmTG_Embarque_Prendas.lNum_Embarque = lNum_Embarque
            Set frmTG_Embarque_Prendas.oParent = Me
            frmTG_Embarque_Prendas.sAccion = "I"
            frmTG_Embarque_Prendas.Show vbModal
            Set frmTG_Embarque_Prendas = Nothing
        Case "MODIFICAR"
            Load frmTG_Embarque_Prendas
            frmTG_Embarque_Prendas.lNum_Embarque = lNum_Embarque
            frmTG_Embarque_Prendas.lSec_Embarque = GridEX1.Value(GridEX1.Columns("SEC_EMBARQUE").Index)
            Set frmTG_Embarque_Prendas.oParent = Me
            CargarData
            frmTG_Embarque_Prendas.sAccion = "U"
            frmTG_Embarque_Prendas.Show vbModal
            Set frmTG_Embarque_Prendas = Nothing
        Case "ELIMINAR"
            Load frmTG_Embarque_Prendas
            frmTG_Embarque_Prendas.lNum_Embarque = lNum_Embarque
            frmTG_Embarque_Prendas.lSec_Embarque = GridEX1.Value(GridEX1.Columns("SEC_EMBARQUE").Index)
            Set frmTG_Embarque_Prendas.oParent = Me
            CargarData
            frmTG_Embarque_Prendas.sAccion = "D"
            frmTG_Embarque_Prendas.Show vbModal
            Set frmTG_Embarque_Prendas = Nothing
        Case "PACKING"
            Me.Frame1.Visible = True
            Me.txt_npac.SetFocus
        Case "SALIR"
            Unload Me
    End Select
End Sub

Private Sub CargarData()
frmTG_Embarque_Prendas.lSec_Embarque = GridEX1.Value(GridEX1.Columns("SEC_EMBARQUE").Index)
frmTG_Embarque_Prendas.txtCod_Fabrica.Text = GridEX1.Value(GridEX1.Columns("COD_FABRICA").Index)
frmTG_Embarque_Prendas.txtCod_OrdPro.Text = GridEX1.Value(GridEX1.Columns("ORDENES").Index)
frmTG_Embarque_Prendas.txtAbr_Cliente.Text = GridEX1.Value(GridEX1.Columns("ABR_CLIENTE").Index)
frmTG_Embarque_Prendas.txtAbr_Cliente.Tag = GridEX1.Value(GridEX1.Columns("COD_CLIENTE").Index)
frmTG_Embarque_Prendas.txtNom_Cliente.Text = GridEX1.Value(GridEX1.Columns("NOM_CLIENTE").Index)
frmTG_Embarque_Prendas.txtCod_PurOrd.Text = GridEX1.Value(GridEX1.Columns("COD_PURORD").Index)
frmTG_Embarque_Prendas.txtCod_LotPurOrd.Text = GridEX1.Value(GridEX1.Columns("COD_LOTPURORD").Index)
frmTG_Embarque_Prendas.txtCod_EstCli.Text = GridEX1.Value(GridEX1.Columns("COD_ESTCLI").Index)
frmTG_Embarque_Prendas.txtNum_Prendas_Prog.Text = GridEX1.Value(GridEX1.Columns("NUM_PRENDAS_PROG").Index)
frmTG_Embarque_Prendas.txtPre_Unitario.Text = GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index)
frmTG_Embarque_Prendas.txtNum_Cajas_Prog.Text = GridEX1.Value(GridEX1.Columns("Num_Cajas_Prog").Index)
frmTG_Embarque_Prendas.txtPeso_Bruto_Prog.Text = GridEX1.Value(GridEX1.Columns("Peso_Bruto_Prog").Index)
frmTG_Embarque_Prendas.txtPeso_Neto_Prog.Text = GridEX1.Value(GridEX1.Columns("Peso_Neto_Prog").Index)
frmTG_Embarque_Prendas.txtCubicaje_Prog.Text = GridEX1.Value(GridEX1.Columns("Cubicaje_Prog").Index)

frmTG_Embarque_Prendas.txtNum_Prendas.Text = GridEX1.Value(GridEX1.Columns("NUM_PRENDAS").Index)
frmTG_Embarque_Prendas.txtNum_Cajas.Text = GridEX1.Value(GridEX1.Columns("Num_Cajas").Index)
frmTG_Embarque_Prendas.txtPeso_Bruto.Text = GridEX1.Value(GridEX1.Columns("Peso_Bruto").Index)
frmTG_Embarque_Prendas.txtPeso_Neto.Text = GridEX1.Value(GridEX1.Columns("Peso_Neto").Index)
frmTG_Embarque_Prendas.txtCubicaje.Text = GridEX1.Value(GridEX1.Columns("Cubicaje").Index)

frmTG_Embarque_Prendas.txtarancelaria1.Text = GridEX1.Value(GridEX1.Columns("Num_Partida_Arancelaria").Index)
frmTG_Embarque_Prendas.txtarancelaria2.Text = GridEX1.Value(GridEX1.Columns("Sec_Partida_Arancelaria").Index)
frmTG_Embarque_Prendas.txtarancelaria3.Text = GridEX1.Value(GridEX1.Columns("Num_Categoria_Internacional").Index)

frmTG_Embarque_Prendas.txtCod_Fabrica.Enabled = False
frmTG_Embarque_Prendas.txtCod_OrdPro.Enabled = False
frmTG_Embarque_Prendas.txtAbr_Cliente.Enabled = False
frmTG_Embarque_Prendas.txtNom_Cliente.Enabled = False
frmTG_Embarque_Prendas.txtCod_PurOrd.Enabled = False
frmTG_Embarque_Prendas.txtCod_LotPurOrd.Enabled = False
frmTG_Embarque_Prendas.txtCod_EstCli.Enabled = False
frmTG_Embarque_Prendas.txtarancelaria1.Enabled = False
frmTG_Embarque_Prendas.txtarancelaria2.Enabled = False
frmTG_Embarque_Prendas.txtarancelaria3.Enabled = False

End Sub

Private Sub txt_npac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
