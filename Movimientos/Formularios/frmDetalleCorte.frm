VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmDetalleCorte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Corte"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   9690
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraValorizar 
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
      Left            =   3105
      TabIndex        =   26
      Top             =   1395
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox TxtDolares 
         Height          =   285
         Left            =   2160
         TabIndex        =   28
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TxtSoles 
         Height          =   285
         Left            =   2160
         TabIndex        =   27
         Top             =   360
         Width           =   975
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   510
         Left            =   600
         TabIndex        =   29
         Top             =   1200
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmDetalleCorte.frx":0000
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label7 
         Caption         =   "Importe Dolares"
         Height          =   255
         Left            =   600
         TabIndex        =   31
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Soles"
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   30
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2025
      Left            =   75
      TabIndex        =   10
      Top             =   3330
      Width           =   9540
      Begin VB.CommandButton cmdGetInfo 
         Height          =   285
         Left            =   2700
         Picture         =   "frmDetalleCorte.frx":0096
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Seleccionar Ordenes por Partida"
         Top             =   225
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdPesosBal 
         Caption         =   "..."
         Height          =   285
         Left            =   8595
         TabIndex        =   24
         Top             =   255
         Width           =   345
      End
      Begin VB.TextBox TxtCantidad 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7080
         TabIndex        =   4
         Text            =   "0"
         Top             =   225
         Width           =   945
      End
      Begin VB.TextBox TxtBultos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7080
         TabIndex        =   6
         Text            =   "0"
         Top             =   915
         Width           =   945
      End
      Begin VB.TextBox TxtObs 
         Enabled         =   0   'False
         Height          =   645
         Left            =   7095
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Text            =   "frmDetalleCorte.frx":03A0
         Top             =   1245
         Width           =   2385
      End
      Begin VB.TextBox txtCan_Movimiento_2daunimed 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   5
         Text            =   "0"
         Top             =   585
         Width           =   945
      End
      Begin VB.TextBox TxtDesitem 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2055
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   540
         Width           =   2490
      End
      Begin VB.TextBox TxtItem 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1110
         MaxLength       =   8
         TabIndex        =   2
         Top             =   540
         Width           =   945
      End
      Begin VB.TextBox txtCO_CodOrdPro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1575
         TabIndex        =   1
         Top             =   225
         Width           =   945
      End
      Begin VB.Label Label6 
         Caption         =   "KG."
         Height          =   225
         Left            =   8070
         TabIndex        =   23
         Top             =   315
         Width           =   270
      End
      Begin VB.Label lblPartida 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1125
         TabIndex        =   22
         Top             =   975
         Width           =   2565
      End
      Begin VB.Label Label5 
         Caption         =   "Lote:"
         Height          =   180
         Left            =   195
         TabIndex        =   21
         Top             =   990
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad Paños:"
         Height          =   240
         Left            =   5820
         TabIndex        =   20
         Top             =   645
         Width           =   1185
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1110
         TabIndex        =   19
         Top             =   1650
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Calidad:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   18
         Top             =   1650
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comb:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   17
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1125
         TabIndex        =   16
         Top             =   1320
         Width           =   4095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad Mov:"
         Height          =   195
         Index           =   7
         Left            =   5820
         TabIndex        =   15
         Top             =   315
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad Bultos:"
         Height          =   195
         Index           =   8
         Left            =   5820
         TabIndex        =   14
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   9
         Left            =   5820
         TabIndex        =   13
         Top             =   1290
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tela:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   645
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "OCorte-Req.Tela:"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   11
         Top             =   270
         Width           =   1665
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3225
      Left            =   75
      TabIndex        =   9
      Top             =   75
      Width           =   9540
      Begin GridEX20.GridEX gexDetCorte 
         Height          =   2940
         Left            =   75
         TabIndex        =   0
         Top             =   180
         Width           =   9390
         _ExtentX        =   16563
         _ExtentY        =   5186
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmDetalleCorte.frx":03A6
         Column(2)       =   "frmDetalleCorte.frx":046E
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmDetalleCorte.frx":0512
         FormatStyle(2)  =   "frmDetalleCorte.frx":064A
         FormatStyle(3)  =   "frmDetalleCorte.frx":06FA
         FormatStyle(4)  =   "frmDetalleCorte.frx":07AE
         FormatStyle(5)  =   "frmDetalleCorte.frx":0886
         FormatStyle(6)  =   "frmDetalleCorte.frx":093E
         ImageCount      =   0
         PrinterProperties=   "frmDetalleCorte.frx":0A1E
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   3030
      TabIndex        =   8
      Top             =   5430
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmDetalleCorte.frx":0BF6
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   495
      Left            =   135
      TabIndex        =   32
      Top             =   5445
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      Custom          =   "0~0~VALORIZAR~Verdadero~Verdadero~&Valorizar Transferencias~0~0~1~~0~Falso~Falso~&Valorizar Transferencias~"
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmDetalleCorte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_Almacen As String
Public sCod_TipMov As String
Public sNum_MovStk As String
Public sCod_ClaOrdComp As String

Public Paso As Boolean
Public Codigo As String
Public Descripcion As String
Public sCOD_COMB As String
Public scod_color As String
Public sCOD_TALLA As String
Public sCod_TipOrdTra As String
Public Scod_ordtra As String

Dim StrSql As String
Dim Estado As String

Private Sub cmdGetInfo_Click()
    frmPartidaCortes.sCod_TipMov = sCod_TipMov
    frmPartidaCortes.sCod_Almacen = sCod_Almacen
    frmPartidaCortes.sNum_MovStk = sNum_MovStk
    frmPartidaCortes.Show vbModal
    CARGA_GRID
End Sub

Private Sub cmdPesosBal_Click()
    frmPesosBal.Show vbModal
    TxtCantidad = frmPesosBal.lblTotal
    Unload frmPesosBal
End Sub



Private Sub gexDetCorte_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    txtCO_CodOrdPro.Text = gexDetCorte.Value(gexDetCorte.Columns("corte").Index)
    TxtItem.Text = gexDetCorte.Value(gexDetCorte.Columns("cod_tela").Index)
    TxtDesitem.Text = gexDetCorte.Value(gexDetCorte.Columns("tela").Index)
    lblPartida.Caption = gexDetCorte.Value(gexDetCorte.Columns("lote").Index)
    sCOD_COMB = gexDetCorte.Value(gexDetCorte.Columns("cod_comb").Index)
    Label3.Caption = gexDetCorte.Value(gexDetCorte.Columns("combinacion").Index)
    scod_color = gexDetCorte.Value(gexDetCorte.Columns("cod_color").Index)
    Label2.Caption = gexDetCorte.Value(gexDetCorte.Columns("calidad").Index)
    sCOD_TALLA = gexDetCorte.Value(gexDetCorte.Columns("talla").Index)
    sCod_TipOrdTra = gexDetCorte.Value(gexDetCorte.Columns("cod_tipordtra").Index)
    Scod_ordtra = gexDetCorte.Value(gexDetCorte.Columns("cod_ordtra").Index)
    TxtCantidad.Text = gexDetCorte.Value(gexDetCorte.Columns("cant mov.").Index)
    txtCan_Movimiento_2daunimed.Text = gexDetCorte.Value(gexDetCorte.Columns("paños").Index)
    TxtBultos.Text = gexDetCorte.Value(gexDetCorte.Columns("bultos").Index)
    TxtObs.Text = gexDetCorte.Value(gexDetCorte.Columns("observaciones").Index)
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo hand
Select Case ActionName
    Case "ADICIONAR"
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            LimpiaCampos
            HabilitaCampos True
            Estado = "NUEVO"
            txtCO_CodOrdPro.SetFocus
    Case "MODIFICAR"
'        If Me.varValida_Factura = False Then
'            MsgBox "No se puede modificar el registro por que el Mov. Almacen posee una factura relacionada.", vbInformation, "Mensaje"
'            Exit Sub
'        End If
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Estado = "MODIFICAR"
        Call HabilitaCampos(True)
        txtCO_CodOrdPro.Enabled = False
        TxtItem.Enabled = False
        TxtDesitem.Enabled = False
        TxtCantidad.SetFocus
    Case "ELIMINAR"
'        If Me.varValida_Factura = False Then
'            MsgBox "No se puede eliminar el registro por que el Mov. Almacen posee una factura relacionada.", vbInformation, "Mensaje"
'            Exit Sub
'        End If
        If (MsgBox("Esta seguro que desea eliminar el registro", vbYesNo, "Consulta")) = vbYes Then
            SALVAR_DATOS "e"
        End If
        LimpiaCampos
        CARGA_GRID
        Call HabilitaCampos(False)
    Case "GRABAR"
        If Trim(TxtCantidad) = "" Or TxtCantidad = "0" Then MsgBox "Llene la cantidad", vbInformation: Exit Sub
        'If Trim(TxtBultos) = "" Or TxtBultos  "0" Then MsgBox "Llene la cantidad de bultos", vbInformation: Exit Sub
        If Trim(TxtBultos) = "" Then TxtBultos = "0"
        If Trim(TxtItem) = "" Then MsgBox "Debe seleccionar la Tela", vbInformation: Exit Sub
        
        'Aqui haremos una validacion sobre cantidades
'        StrSql = "EXEC SM_ENCUENTRA_TIPOFAMTELA '" & Me.TxtItem.Text & "',''"
'        varCod_TipFamTela = DevuelveCampo(StrSql, cConnect)
'        If varCod_TipFamTela = "N" Then
'            If Val(Me.txtCan_Movimiento_2daunimed.Text) < 0 Then
'                MsgBox "La 2da cantidad no puede ser menor que 0", vbInformation, "Mensaje"
'                Me.txtCan_Movimiento_2daunimed.SetFocus
'                Exit Sub
'            End If
'        Else
            If Val(Me.txtCan_Movimiento_2daunimed.Text) < 0 Then
                MsgBox "La cantidad de Paños no puede ser menor que 0", vbInformation, "Mensaje"
                Me.txtCan_Movimiento_2daunimed.SetFocus
                Exit Sub
            End If
'        End If

        If Estado = "NUEVO" Then
            SALVAR_DATOS "i"
        Else
            SALVAR_DATOS "m"
        End If
        LimpiaCampos
        Call HabilitaCampos(False)
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        CARGA_GRID
    Case "DESHACER"
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        LimpiaCampos
        CARGA_GRID
        Call HabilitaCampos(False)
    Case "SALIR"
        Unload Me
End Select
Exit Sub
hand:
ErrorHandler err, "MantFunc1_ActionClick"
End Sub

Private Sub TxtBultos_GotFocus()
    Call SelectionText(TxtBultos)
End Sub

Private Sub txtBultos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then AVANZA (13)
End Sub

Private Sub txtCan_Movimiento_2daunimed_GotFocus()
    Call SelectionText(txtCan_Movimiento_2daunimed)
End Sub

Private Sub txtCan_Movimiento_2daunimed_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then AVANZA (13)
End Sub

Private Sub TxtCantidad_GotFocus()
    Call SelectionText(TxtCantidad)
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then AVANZA (13)
End Sub

Private Sub txtCo_CodOrdPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AVANZA 13
    Else
        Call SoloNumeros(txtCO_CodOrdPro, KeyAscii, False, 0, 5)
    End If
End Sub

Private Sub txtCO_CodOrdPro_LostFocus()
    txtCO_CodOrdPro.Text = Format(Trim(txtCO_CodOrdPro.Text), "00000")
End Sub

Sub BuscaTelas()
    
    StrSql = "EXEC CO_SM_MUESTRA_TELAS_DEL_CORTE '" & txtCO_CodOrdPro.Text & "','" & _
             sCod_TipMov & "', '" & sCod_Almacen & "', '" & sNum_MovStk & "'"
    Set frmBusqGeneral.oParent = Me
    frmBusqGeneral.sQuery = StrSql
    frmBusqGeneral.Cargar_Datos
    frmBusqGeneral.gexList.Columns("cod_tela").Visible = False
    frmBusqGeneral.gexList.Columns("cod_tipordtra").Visible = False
    frmBusqGeneral.gexList.Columns("cod_ordtra").Visible = False
    frmBusqGeneral.gexList.Columns("cod_comb").Visible = False
    frmBusqGeneral.gexList.Columns("cod_color").Visible = False
    frmBusqGeneral.gexList.Columns("cod_medida").Visible = False
    frmBusqGeneral.gexList.Columns("cod_proveedor").Visible = False
    frmBusqGeneral.gexList.Columns("cod_calidad").Visible = False
    
    frmBusqGeneral.gexList.Columns("partida").Width = 1000
    frmBusqGeneral.gexList.Columns("tela").Width = 2500
    frmBusqGeneral.gexList.Columns("color").Width = 2000
    frmBusqGeneral.gexList.Columns("saldo").Width = 700
    frmBusqGeneral.Show 1
    If Paso = True Then
        TxtItem = Codigo
        TxtDesitem = Descripcion
        StrSql = "SELECT cod_clamov FROM lg_tiposmov WHERE Cod_TipMov = '" & sCod_TipMov & "'"
        If Trim(DevuelveCampo(StrSql, cConnect)) = "E" Then
            TxtCantidad.Text = 0
        End If
        TxtCantidad.SetFocus
    End If
End Sub

Sub CARGA_GRID()
On Error GoTo ErrCargaGrid
' cambio vl panos orden directa
    StrSql = "SELECT Flg_Panos_Orden_Directa FROM lg_ALMACEN WHERE Cod_Almacen = '" & sCod_Almacen & "'"

    If Trim(DevuelveCampo(StrSql, cConnect)) = "S" Then
        StrSql = "EXEC UP_Lg_MoviStkCorte_Directo '" & sCod_Almacen & "','" & sNum_MovStk & "'"
    Else
        StrSql = "EXEC UP_Lg_MoviStkCorte '" & sCod_Almacen & "','" & sNum_MovStk & "'"
    End If
    
' cambio vl panos orden direct
    Set gexDetCorte.ADORecordset = CargarRecordSetDesconectado(StrSql, cConnect)
    ConfigurarGrid
    
Exit Sub
ErrCargaGrid:
    ErrorHandler err, "CARGA_GRID"
End Sub

Sub ConfigurarGrid()
    gexDetCorte.Columns("cod_proveedor").Visible = False
    gexDetCorte.Columns("cod_tela").Visible = False
    gexDetCorte.Columns("cod_comb").Visible = False
    gexDetCorte.Columns("cod_color").Visible = False
    gexDetCorte.Columns("cod_tipordtra").Visible = False
    gexDetCorte.Columns("cod_ordtra").Visible = False
    
    gexDetCorte.Columns("secuencia").Width = 900
    gexDetCorte.Columns("corte").Width = 700
    gexDetCorte.Columns("talla").Width = 600
    gexDetCorte.Columns("lote").Width = 900
    gexDetCorte.Columns("tela").Width = 2000
    gexDetCorte.Columns("paños").Width = 700
    gexDetCorte.Columns("bultos").Width = 700
    gexDetCorte.Columns("Cant Mov.").Width = 800
    gexDetCorte.Columns("calidad").Width = 700
End Sub

Sub HabilitaCampos(sEstado As Boolean)
    txtCO_CodOrdPro.Enabled = sEstado
    TxtItem.Enabled = sEstado
    TxtDesitem.Enabled = sEstado
    TxtCantidad.Enabled = sEstado
    txtCan_Movimiento_2daunimed.Enabled = sEstado
    TxtBultos.Enabled = sEstado
    TxtObs.Enabled = sEstado
    cmdPesosBal.Enabled = sEstado
    cmdGetInfo.Enabled = (Not sEstado)
End Sub

Sub LimpiaCampos()
    txtCO_CodOrdPro.Text = ""
    TxtItem.Text = ""
    TxtDesitem.Text = ""
    TxtCantidad.Text = 0
    txtCan_Movimiento_2daunimed.Text = 0
    TxtBultos.Text = 0
    TxtObs.Text = ""
    lblPartida.Caption = ""
    Label3.Caption = ""
    Label2.Caption = ""
    
    sCOD_COMB = ""
    scod_color = ""
    sCOD_TALLA = ""
    sCod_TipOrdTra = ""
    Scod_ordtra = ""
End Sub

Private Sub TxtItem_KeyPress(KeyAscii As Integer)
    BuscaTelas
End Sub

Private Sub TxtObs_GotFocus()
    Call SelectionText(TxtBultos)
End Sub

Private Sub TxtObs_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Sub SALVAR_DATOS(sTipo As String)
On Error GoTo ErrSalvar
' cambio vl panos orden directa
    StrSql = "SELECT Flg_Panos_Orden_Directa FROM lg_ALMACEN WHERE Cod_Almacen = '" & sCod_Almacen & "'"

    If Trim(DevuelveCampo(StrSql, cConnect)) = "S" Then
            StrSql = "EXEC UP_ACT_STOCKSTELCOR_DIRECTO '" & sTipo & "','" & sCod_Almacen & "','" & _
            sNum_MovStk & "','" & gexDetCorte.Value(gexDetCorte.Columns("secuencia").Index) & "','" & _
            Trim(txtCO_CodOrdPro.Text) & "','" & sCod_TipOrdTra & "','" & Scod_ordtra & "','" & _
            Trim(TxtItem.Text) & "','" & sCOD_COMB & "','" & _
            scod_color & "','" & sCOD_TALLA & "','" & _
            Trim(Label2.Caption) & "'," & TxtCantidad.Text & "," & _
            TxtBultos.Text & "," & txtCan_Movimiento_2daunimed.Text & ",'" & _
            Trim(TxtObs.Text) & "','" & vusu & "'"
    
    Else
            StrSql = "EXEC UP_ACT_STOCKSTELCOR '" & sTipo & "','" & sCod_Almacen & "','" & _
            sNum_MovStk & "','" & gexDetCorte.Value(gexDetCorte.Columns("secuencia").Index) & "','" & _
            Trim(txtCO_CodOrdPro.Text) & "','" & sCod_TipOrdTra & "','" & Scod_ordtra & "','" & _
            Trim(TxtItem.Text) & "','" & sCOD_COMB & "','" & _
            scod_color & "','" & sCOD_TALLA & "','" & _
            Trim(Label2.Caption) & "'," & TxtCantidad.Text & "," & _
            TxtBultos.Text & "," & txtCan_Movimiento_2daunimed.Text & ",'" & _
            Trim(TxtObs.Text) & "','" & vusu & "'"
    End If
            
    Call ExecuteSQL(cConnect, StrSql)
    
' cambio vl panos orden directa
Exit Sub
ErrSalvar:
    ErrorHandler err, "Salvar_Datos"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "VALORIZAR"
        FraValorizar.Visible = True
        TxtSoles.SetFocus
        
        TxtSoles.Text = DevuelveCampo("select imp_factura from lg_movistktelcor where cod_almacen ='" & Me.sCod_Almacen & "' and num_movstk='" & Me.sNum_MovStk & "' and num_secuencia ='" & gexDetCorte.Value(gexDetCorte.Columns("secuencia").Index) & "'", cConnect)
        TxtDolares.Text = DevuelveCampo("select imp_factura_dolares from lg_movistktelcor where cod_almacen ='" & Me.sCod_Almacen & "' and num_movstk='" & Me.sNum_MovStk & "' and num_secuencia ='" & gexDetCorte.Value(gexDetCorte.Columns("secuencia").Index) & "'", cConnect)
    End Select
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
Dim StrSql As String

If Trim(TxtSoles.Text) = "" Then
    TxtSoles.Text = "0"
End If

If Trim(TxtDolares.Text) = "" Then
    TxtDolares.Text = "0"
End If

StrSql = "EXEC LG_MovistkItem_Valoriza_Transferencia '" & Me.sCod_Almacen & "','" & Me.sNum_MovStk & "','" & gexDetCorte.Value(gexDetCorte.Columns("secuencia").Index) & "'," & _
        CDbl(TxtSoles.Text) & "," & CDbl(TxtDolares.Text)
ExecuteSQL cConnect, StrSql

TxtSoles.Text = ""
TxtDolares.Text = ""
Exit Sub
errGrabar:
    ErrorHandler err, "Grabar"
End Sub


