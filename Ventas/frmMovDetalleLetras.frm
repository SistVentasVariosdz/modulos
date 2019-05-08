VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMovDetalleLetras 
   Caption         =   "Adionar Letras a Planilla"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   10650
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frCliente 
      Height          =   810
      Left            =   75
      TabIndex        =   13
      Top             =   0
      Width           =   10485
      Begin VB.TextBox txtCod_Anexo 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   2970
         MaxLength       =   4
         TabIndex        =   2
         Top             =   255
         Width           =   600
      End
      Begin VB.TextBox txtDes_Anexo 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   3570
         MaxLength       =   30
         TabIndex        =   3
         Top             =   255
         Width           =   4155
      End
      Begin VB.TextBox txtCod_TipAne 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   2595
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "C"
         Top             =   255
         Width           =   360
      End
      Begin VB.TextBox txtRuc 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   1200
         MaxLength       =   11
         TabIndex        =   0
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   405
         Left            =   8310
         TabIndex        =   4
         Top             =   210
         Width           =   1335
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   390
         TabIndex        =   15
         Top             =   285
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "/"
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
         Left            =   2460
         TabIndex        =   14
         Tag             =   "Anexo Type"
         Top             =   270
         Width           =   90
      End
   End
   Begin VB.Frame fraOperaciones 
      Caption         =   "Operaciones                                                                            Operaciones a Generar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4485
      Left            =   90
      TabIndex        =   5
      Top             =   825
      Width           =   10485
      Begin VB.CommandButton cmdDelAllOpe 
         Caption         =   "<<"
         Height          =   480
         Left            =   5040
         TabIndex        =   9
         Top             =   3120
         Width           =   435
      End
      Begin VB.CommandButton cmdDelOpe 
         Caption         =   "<"
         Height          =   480
         Left            =   5040
         TabIndex        =   8
         Top             =   2655
         Width           =   435
      End
      Begin VB.CommandButton cmdAddOpe 
         Caption         =   ">"
         Height          =   480
         Left            =   5040
         TabIndex        =   7
         Top             =   2190
         Width           =   435
      End
      Begin VB.CommandButton cmdAddAllOpe 
         Caption         =   ">>"
         Height          =   480
         Left            =   5040
         TabIndex        =   6
         Top             =   1725
         Width           =   435
      End
      Begin GridEX20.GridEX GridEXDispon 
         Height          =   4200
         Left            =   120
         TabIndex        =   10
         Top             =   195
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   7408
         Version         =   "2.0"
         HoldSortSettings=   -1  'True
         DefaultGroupMode=   1
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ImageCount      =   2
         ImagePicture1   =   "frmMovDetalleLetras.frx":0000
         ImagePicture2   =   "frmMovDetalleLetras.frx":031A
         DataMode        =   1
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "frmMovDetalleLetras.frx":0634
         Column(2)       =   "frmMovDetalleLetras.frx":06FC
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmMovDetalleLetras.frx":07A0
         FormatStyle(2)  =   "frmMovDetalleLetras.frx":08D8
         FormatStyle(3)  =   "frmMovDetalleLetras.frx":0988
         FormatStyle(4)  =   "frmMovDetalleLetras.frx":0A3C
         FormatStyle(5)  =   "frmMovDetalleLetras.frx":0B14
         FormatStyle(6)  =   "frmMovDetalleLetras.frx":0BCC
         FormatStyle(7)  =   "frmMovDetalleLetras.frx":0CAC
         FormatStyle(8)  =   "frmMovDetalleLetras.frx":0D58
         ImageCount      =   2
         ImagePicture(1) =   "frmMovDetalleLetras.frx":0E08
         ImagePicture(2) =   "frmMovDetalleLetras.frx":1122
         PrinterProperties=   "frmMovDetalleLetras.frx":143C
      End
      Begin GridEX20.GridEX GridEXSelec 
         Height          =   4200
         Left            =   5535
         TabIndex        =   11
         Top             =   210
         Width           =   4890
         _ExtentX        =   8625
         _ExtentY        =   7408
         Version         =   "2.0"
         HoldSortSettings=   -1  'True
         DefaultGroupMode=   1
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ImageCount      =   2
         ImagePicture1   =   "frmMovDetalleLetras.frx":1614
         ImagePicture2   =   "frmMovDetalleLetras.frx":192E
         DataMode        =   1
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "frmMovDetalleLetras.frx":1C48
         Column(2)       =   "frmMovDetalleLetras.frx":1D10
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmMovDetalleLetras.frx":1DB4
         FormatStyle(2)  =   "frmMovDetalleLetras.frx":1EEC
         FormatStyle(3)  =   "frmMovDetalleLetras.frx":1F9C
         FormatStyle(4)  =   "frmMovDetalleLetras.frx":2050
         FormatStyle(5)  =   "frmMovDetalleLetras.frx":2128
         FormatStyle(6)  =   "frmMovDetalleLetras.frx":21E0
         FormatStyle(7)  =   "frmMovDetalleLetras.frx":22C0
         FormatStyle(8)  =   "frmMovDetalleLetras.frx":236C
         ImageCount      =   2
         ImagePicture(1) =   "frmMovDetalleLetras.frx":241C
         ImagePicture(2) =   "frmMovDetalleLetras.frx":2736
         PrinterProperties=   "frmMovDetalleLetras.frx":2A50
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   8145
      TabIndex        =   12
      Top             =   5385
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmMovDetalleLetras.frx":2C28
      Orientacion     =   0
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmMovDetalleLetras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public codigo As String, Descripcion As String
Public sCOD_BANCO As String
Public sFlg_Status_Letras As String
Public sCod_Tipanex As String
Public sCod_Anxo As String
Public rsGridOperaciones As ADODB.Recordset
Public scadena As String
Public num As String
Public sCuenta As String
Dim rsGridDisponibles As ADODB.Recordset
Dim rsGridSeleccionados As ADODB.Recordset
Dim strNum_Corre_Let_Renov As String, TipoAdd As String
Public Tipo As String, xCod_Grupo As Integer, strNum_Corre_Let As String
Dim strSQL  As String
Public sNum_Planilla_Letra As Integer
Dim iRowAnterior As Long
Dim iColAnterior As Long
Dim bClickColSelec As Boolean
Dim bCargaGRid As Boolean
Dim bPuedeAutorizar  As Boolean
Dim sTipoDocAutorizar As String
Dim strOpcion As String
Public strCod_Anxo As String
Dim lvSW As Boolean


Private Sub cmdAddAllOpe_Click()
Dim iRowSave As Long
Dim iRowactual As Long
Dim iRowSave2 As Long
Dim sTendido As String

    iRowSave = GridEXDispon.Row

    If GridEXDispon.RowCount > 0 Then
        GridEXDispon.Row = 1
    End If

    iRowactual = GridEXDispon.Row

    Do While True
        GridEXDispon.Row = 1
        cmdAddOpe_Click
        If GridEXDispon.Row = 0 Then
            Exit Sub
        End If
    Loop
    GridEXDispon.Row = iRowSave
End Sub

Private Sub cmdAddOpe_Click()
    GridEX_To_GridEX GridEXDispon, GridEXSelec, rsGridDisponibles, rsGridSeleccionados
End Sub


Private Sub GridEX_To_GridEX(ByRef GridexOrg As GridEx, ByRef GridexDst As GridEx, ByRef rsGridOrg As ADODB.Recordset, ByRef rsGridDst As ADODB.Recordset)
    Dim iRow As Long
    iRow = GridexOrg.Row
    If GridexOrg.RowCount > 0 Then
        If Not GridexDst.Find(GridexDst.Columns("num_corre").Index, jgexEqual, GridexOrg.Value(GridexOrg.Columns("num_corre").Index)) Then

            If RTrim(GridexOrg.Value(GridexOrg.Columns("num_corre").Index)) <> "" Then
                rsGridOrg.Find "num_corre = " & GridexOrg.Value(GridexOrg.Columns("num_corre").Index), , , 1
                'rsGridOrg.Find "Key = " & "0001", , , 1
            End If

            Set GridexDst.ADORecordset = Nothing
            rsGridDst.AddNew

            rsGridDst("cod_tipdoc").Value = GridexOrg.Value(GridexOrg.Columns("cod_tipdoc").Index)
            rsGridDst("ser_docum").Value = GridexOrg.Value(GridexOrg.Columns("ser_docum").Index)

            rsGridDst("num_docum_ventas").Value = GridexOrg.Value(GridexOrg.Columns("num_docum_ventas").Index)
            rsGridDst("cod_moneda").Value = GridexOrg.Value(GridexOrg.Columns("cod_moneda").Index)
            rsGridDst("imp_total").Value = GridexOrg.Value(GridexOrg.Columns("imp_total").Index)
'
            rsGridDst("num_corre").Value = GridexOrg.Value(GridexOrg.Columns("num_corre").Index)


            Set GridexDst.ADORecordset = rsGridDst

            RefrescaGrid GridexDst, rsGridDst

            If GridexOrg.RowCount > 0 Then

                rsGridOrg.Delete
                Set GridexOrg.ADORecordset = Nothing
                RefrescaGrid GridexOrg, rsGridOrg
            End If


        Else
            If GridexOrg.RowCount > 0 Then
                If RTrim(GridexOrg.Value(GridexOrg.Columns("num_corre").Index)) <> "" Then
                    rsGridOrg.Find "num_corre = " & GridexOrg.Value(GridexOrg.Columns("num_corre").Index), , , 1
                End If

                rsGridOrg.Delete
                Set GridexOrg.ADORecordset = Nothing
                RefrescaGrid GridexOrg, rsGridOrg
            End If

        End If
    End If
End Sub


Private Sub cmdBuscar_Click()
'Call CargarOperaciones

Dim sSQL As String
Dim ssql1 As String
Dim vBookmark As Variant

If sCOD_BANCO = 4 Then

sSQL = "Ventas_Generar_Detalle_Letras  '$','$','$','$','$','$','$','$'"
sSQL = VBsprintf(sSQL, 1, sNum_Planilla_Letra, sCOD_BANCO, sFlg_Status_Letras, txtCod_TipAne.Text, txtCod_Anexo.Text, txtRuc.Text, sCuenta)

Else

sSQL = "Ventas_Generar_Detalle_Letras  '$','$','$','$','$','$','$','$'"
sSQL = VBsprintf(sSQL, 1, sNum_Planilla_Letra, sCOD_BANCO, sFlg_Status_Letras, txtCod_TipAne.Text, txtCod_Anexo.Text, txtRuc.Text, sCuenta)

End If
Set GridEXDispon.ADORecordset = CargarRecordSetDesconectado(sSQL, cCONNECT)
Set rsGridDisponibles = GridEXDispon.ADORecordset
    GridEXDispon.Refresh

    Dim colTemp As JSColumn

    GridEXDispon.ColumnHeaderHeight = 500
    GridEXDispon.Columns("num_corre").Visible = False
    GridEXDispon.Columns("cod_tipdoc").Caption = "Tipo Doc."
    GridEXDispon.Columns("ser_docum").Caption = "Serie"
    GridEXDispon.Columns("num_docum_ventas").Caption = "Num. Doc."
    GridEXDispon.Columns("cod_moneda").Caption = "Moneda"
    GridEXDispon.Columns("imp_total").Caption = "Imp. Total"
    GridEXDispon.Columns("fec_emidoc").Caption = "Fec. Emi. Doc."
    GridEXDispon.Columns("fec_vendoc").Caption = "Fec. Ven. Doc."

    GridEXDispon.Columns("cod_tipdoc").Width = 500
    GridEXDispon.Columns("ser_docum").Width = 500
    GridEXDispon.Columns("num_docum_ventas").Width = 900
    GridEXDispon.Columns("cod_moneda").Width = 1000
    GridEXDispon.Columns("imp_total").Width = 1000
    GridEXDispon.Columns("fec_emidoc").Width = 1200
    GridEXDispon.Columns("fec_vendoc").Width = 1200

    GridEXDispon.ContinuousScroll = True

    GridEXDispon.Refresh

ssql1 = "Ventas_Generar_Detalle_Letras  '$','$','$','$','$','$','$','$'"
ssql1 = VBsprintf(ssql1, 2, sNum_Planilla_Letra, sCOD_BANCO, sFlg_Status_Letras, txtCod_TipAne.Text, txtCod_Anexo.Text, "", sCuenta)

Set GridEXSelec.ADORecordset = CargarRecordSetDesconectado(ssql1, cCONNECT)
Set rsGridSeleccionados = GridEXSelec.ADORecordset

    GridEXSelec.Refresh


    GridEXSelec.ColumnHeaderHeight = 500
    GridEXSelec.Columns("num_corre").Visible = False
    GridEXSelec.Columns("cod_tipdoc").Caption = "Tipo Doc."
    GridEXSelec.Columns("ser_docum").Caption = "Serie"
    GridEXSelec.Columns("num_docum_ventas").Caption = "Num. Doc."
    GridEXSelec.Columns("cod_moneda").Caption = "Moneda"
    GridEXSelec.Columns("imp_total").Caption = "Imp. Total"
    GridEXSelec.Columns("fec_emidoc").Caption = "Fec. Emi. Doc."
    GridEXSelec.Columns("fec_vendoc").Caption = "Fec. Ven. Doc."

    GridEXSelec.Columns("cod_tipdoc").Width = 500
    GridEXSelec.Columns("ser_docum").Width = 500
    GridEXSelec.Columns("num_docum_ventas").Width = 900
    GridEXSelec.Columns("cod_moneda").Width = 1000
    GridEXSelec.Columns("imp_total").Width = 1000
    GridEXSelec.Columns("fec_emidoc").Width = 1200
    GridEXSelec.Columns("fec_vendoc").Width = 1200

    GridEXSelec.ContinuousScroll = True

    GridEXSelec.Refresh

End Sub

Private Sub cmdDelAllOpe_Click()
Dim iRowSave As Long
Dim iRowactual As Long
Dim iRowSave2 As Long
Dim sTend As String

iRowSave = GridEXSelec.Row
sTend = GridEXSelec.Value(GridEXSelec.Columns("NUM_CORRE").Index)

If GridEXSelec.RowCount > 0 Then
    GridEXSelec.Row = 1
End If

iRowactual = GridEXSelec.Row

Do While True
    GridEXSelec.Row = 1
    cmdDelOpe_Click
    If GridEXSelec.RowCount > 0 Then
        GridEXSelec.Row = 1
        iRowactual = GridEXSelec.Row
    End If

    If GridEXSelec.RowCount = 0 Then
        GridEXSelec.Row = iRowSave
        Exit Sub
    End If
Loop
GridEXSelec.Row = iRowSave
End Sub

Private Sub cmdDelOpe_Click()
 GridEX_To_GridEX GridEXSelec, GridEXDispon, GridEXSelec.ADORecordset, GridEXDispon.ADORecordset
End Sub

Private Sub Form_Load()

    strSQL = "Ventas_Generar_Detalle_Letras  '$','$','$','$','$','$','$','$'"
    strSQL = VBsprintf(strSQL, 1, sNum_Planilla_Letra, sCOD_BANCO, sFlg_Status_Letras, txtCod_TipAne.Text, txtCod_Anexo.Text, "", sCuenta)

    Set GridEXDispon.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)
    Set GridEXSelec.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

    Set rsGridDisponibles = GridEXDispon.ADORecordset

    GridEXDispon.Refresh

    GridEXDispon.ColumnHeaderHeight = 500
    GridEXDispon.Columns("num_corre").Visible = False
    GridEXDispon.Columns("cod_tipdoc").Caption = "Tipo Doc."
    GridEXDispon.Columns("ser_docum").Caption = "Serie"
    GridEXDispon.Columns("num_docum_ventas").Caption = "Num. Doc."
    GridEXDispon.Columns("cod_moneda").Caption = "Moneda"
    GridEXDispon.Columns("imp_total").Caption = "Imp. Total"
    GridEXDispon.Columns("fec_emidoc").Caption = "Fec. Emi. Doc."
    GridEXDispon.Columns("fec_vendoc").Caption = "Fec. Ven. Doc."

    GridEXDispon.Columns("cod_tipdoc").Width = 500
    GridEXDispon.Columns("ser_docum").Width = 500
    GridEXDispon.Columns("num_docum_ventas").Width = 900
    GridEXDispon.Columns("cod_moneda").Width = 1000
    GridEXDispon.Columns("imp_total").Width = 1000
    GridEXDispon.Columns("fec_emidoc").Width = 1200
    GridEXDispon.Columns("fec_vendoc").Width = 1200

    GridEXDispon.ContinuousScroll = True

    GridEXDispon.Refresh

    Set rsGridSeleccionados = GridEXSelec.ADORecordset

    GridEXSelec.Refresh

    GridEXSelec.ColumnHeaderHeight = 500
    GridEXSelec.Columns("num_corre").Visible = False
    GridEXSelec.Columns("cod_tipdoc").Caption = "Tipo Doc."
    GridEXSelec.Columns("ser_docum").Caption = "Serie"
    GridEXSelec.Columns("num_docum_ventas").Caption = "Num. Doc."
    GridEXSelec.Columns("cod_moneda").Caption = "Moneda"
    GridEXSelec.Columns("imp_total").Caption = "Imp. Total"
    GridEXSelec.Columns("fec_emidoc").Caption = "Fec. Emi. Doc."
    GridEXSelec.Columns("fec_vendoc").Caption = "Fec. Ven. Doc."

    GridEXSelec.Columns("cod_tipdoc").Width = 500
    GridEXSelec.Columns("ser_docum").Width = 500
    GridEXSelec.Columns("num_docum_ventas").Width = 900
    GridEXSelec.Columns("cod_moneda").Width = 1000
    GridEXSelec.Columns("imp_total").Width = 1000
    GridEXSelec.Columns("fec_emidoc").Width = 1200
    GridEXSelec.Columns("fec_vendoc").Width = 1200

    GridEXSelec.ContinuousScroll = True

    GridEXSelec.Refresh

End Sub


Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

On Error GoTo Fin
Dim ssql1 As String
Dim sSQL As String
Dim iReg As Integer

    Select Case ActionName
    Case "ACEPTAR"
        For iReg = 1 To GridEXSelec.RowCount
            GridEXSelec.Row = iReg

            sSQL = "VN_DETALLE_LETRAS '$','$','$','$','$'"
            sSQL = VBsprintf(sSQL, 1, GridEXSelec.Value(GridEXSelec.Columns("NUM_CORRE").Index), sNum_Planilla_Letra, sCuenta, sCOD_BANCO)
            ExecuteCommandSQL cCONNECT, sSQL

        Next
        cmdBuscar_Click
        Unload Me
        
        

    Case "CANCELAR"
        Unload Me
    End Select
    
    
    Exit Sub
    Resume
   
Fin:
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Letras"

End Sub


Private Sub RefrescaGrid(ByRef GridEx As GridEx, ByRef rsGrid As ADODB.Recordset)
    Dim oGroup As JSGroup
    
    Set GridEx.ADORecordset = rsGrid
    GridEx.Refresh
        
    GridEx.ColumnHeaderHeight = 500
    GridEx.Columns("num_corre").Visible = False
    GridEx.Columns("cod_tipdoc").Caption = "Tipo Doc."
    GridEx.Columns("ser_docum").Caption = "Serie"
    GridEx.Columns("num_docum_ventas").Caption = "Num. Doc."
    GridEx.Columns("cod_moneda").Caption = "Moneda"
    GridEx.Columns("imp_total").Caption = "Imp. Total"
    GridEx.Columns("fec_emidoc").Caption = "Fec. Emi. Doc."
    GridEx.Columns("fec_vendoc").Caption = "Fec. Ven. Doc."

    GridEx.Columns("cod_tipdoc").Width = 500
    GridEx.Columns("ser_docum").Width = 500
    GridEx.Columns("num_docum_ventas").Width = 900
    GridEx.Columns("cod_moneda").Width = 1000
    GridEx.Columns("imp_total").Width = 1000
    GridEx.Columns("fec_emidoc").Width = 1200
    GridEx.Columns("fec_vendoc").Width = 1200
        
    GridEx.Refresh
End Sub
Private Sub RefrescaGridOperaciones()
    Dim oGroup As JSGroup

    If GridEXSelec.Groups.Count > 0 Then
        GridEXSelec.Groups.Remove GridEXSelec.Groups.Count
        GridEXSelec.Refresh
    End If
    Set GridEXSelec.ADORecordset = rsGridOperaciones

    GridEXSelec.DefaultGroupMode = jgexDGMExpanded
    GridEXSelec.ExpandAll

    GridEXSelec.Columns("num_corre").Visible = False
    GridEXSelec.Columns("cod_tipdoc").Caption = "Tipo Doc."
    GridEXSelec.Columns("ser_docum").Caption = "Serie"
    GridEXSelec.Columns("num_docum_ventas").Caption = "Num. Doc."
    GridEXSelec.Columns("cod_moneda").Caption = "Moneda"
    GridEXSelec.Columns("imp_total").Caption = "Imp. Total"

    GridEXSelec.ContinuousScroll = True
    GridEXSelec.RefreshGroups
    GridEXSelec.Refresh
End Sub


Private Sub txtRuc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then

    BUSCARUC 1
  End If
End Sub




Private Sub BUSCARUC(opcion As Integer)

On Error GoTo Fin
Dim strSQL As String
Dim oTipo As New frmBusqGeneral

    strSQL = "SELECT num_ruc as Ruc,Des_Anexo Descripcion FROM CN_AnexosContables "
    txtRuc = Trim(txtRuc)

    strSQL = strSQL & " where num_ruc like '%" & txtRuc & "%' and Cod_TipAnex ='C'"

    txtRuc = ""

    Set oTipo.oParent = Me

    oTipo.SQuery = strSQL
    oTipo.CARGAR_DATOS
    oTipo.DGridLista.Columns(1).Width = 4350.047
    oTipo.Show 1
    If codigo <> "" Then
      txtRuc = Trim(codigo)
      txtDes_Anexo = Trim(Descripcion)

      strSQL = "SELECT Cod_TipAnEx FROM CN_AnexosContables WHERE num_ruc = '" & txtRuc.Text & "' and Cod_TipAnex ='C'"
      txtCod_TipAne.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
      strSQL = "SELECT Cod_Anxo FROM CN_AnexosContables WHERE num_ruc = '" & txtRuc.Text & "' and Cod_TipAnex ='C'"
      txtCod_Anexo.Text = Trim(DevuelveCampo(strSQL, cCONNECT))

      cmdBuscar.SetFocus
    End If
    Set oTipo = Nothing

Exit Sub
Resume
Fin:
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda (" & opcion & ")"
End Sub

Private Sub txtCod_TipAne_KeyPress(KeyAscii As Integer)
  GridEXDispon.ClearFields
  If KeyAscii = vbKeyReturn Then
      If Trim(txtCod_TipAne.Text) <> "" Then
          Call BUSCA_TIPO_ANEXO(1, 1)
           SendKeys "{TAB}"
      Else
          Call BUSCA_TIPO_ANEXO(2, 1)
           SendKeys "{TAB}"
      End If
  End If
End Sub



Sub BUSCA_TIPO_ANEXO(Tipo As Integer, Ubic As Integer)
    Select Case Tipo
        Case 1:
                If Ubic = 1 Then
                    strSQL = "SELECT DES_TIPANEX FROM CN_TipoAnexoContable WHERE COD_TIPANEX = '" & txtCod_TipAne.Text & "'"
                    txtCod_Anexo.SetFocus
                Else
                End If
        Case 2:
                Dim oTipo As New frmBusqGeneral
                Dim RS As Object
                Set RS = CreateObject("ADODB.Recordset")
                Set oTipo.oParent = Me
                If Ubic = 1 Then
                    oTipo.SQuery = "SELECT COD_TIPANEX as Código, DES_TIPANEX as Descripción FROM CN_TipoAnexoContable "
                Else
                End If
                oTipo.CARGAR_DATOS
                oTipo.Show 1
                If codigo <> "" Then
                    If Ubic = 1 Then
                        txtCod_TipAne.Text = Trim(codigo)
                        txtCod_Anexo.SetFocus
                    Else
                    End If
                End If
                Set oTipo = Nothing
                Set RS = Nothing

    End Select
End Sub

Private Sub txtDes_Anexo_KeyPress(KeyAscii As Integer)
    GridEXDispon.ClearFields
    If KeyAscii = vbKeyReturn Then
        If Trim(txtDes_Anexo.Text) <> "" Then
            If Len(Trim(txtDes_Anexo)) > 2 Then
                Call BUSCA_ANEXO(2, 1)
            Else
                Aviso "Debe ingresar al menos 3 caracteres del Nombre requerido", 1
                Exit Sub
            End If
        Else
            Aviso "Debe ingresar al menos 3 caracteres del Nombre requerido", 1
            Exit Sub
        End If
    End If
End Sub


Sub BUSCA_ANEXO(Tipo As Integer, Ubic As Integer)

Dim iLen As Integer
    Select Case Tipo
        Case 1:
                If Ubic = 1 Then
                    strSQL = "SELECT MIN(DATALENGTH(COD_ANXO)) FROM CN_AnexosContables"
                    iLen = Trim(DevuelveCampo(strSQL, cCONNECT))

                    txtCod_Anexo.Text = Right(Repl("0", iLen) & txtCod_Anexo, iLen)


                     strSQL = "SELECT Des_Anexo FROM CN_AnexosContables WHERE Cod_TipAnEX = '" & txtCod_TipAne.Text & "' AND Cod_Anxo = '" & txtCod_Anexo.Text & "'"
                     txtDes_Anexo.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
                     SendKeys "{TAB}"

                     Exit Sub

                Else
                End If
        Case 2:

                Dim oTipo As New frmBusqGeneral
                Dim RS As Object
                Set RS = CreateObject("ADODB.Recordset")
                Set oTipo.oParent = Me
                If Ubic = 1 Then
                    oTipo.SQuery = "SELECT Cod_Anxo as Código, Des_Anexo as Descripción FROM CN_AnexosContables WHERE Cod_TipAnEX = '" & txtCod_TipAne.Text & "' AND Des_Anexo like '%" & Trim(txtDes_Anexo.Text) & "%'"
                Else
                End If
                oTipo.CARGAR_DATOS
                oTipo.Top = txtDes_Anexo.Top + txtDes_Anexo.Height
                oTipo.Left = txtDes_Anexo.Left
                oTipo.DGridLista.Columns(1).Width = 1000
                oTipo.Show 1
                If codigo <> "" Then
                    If Ubic = 1 Then
                        txtCod_Anexo.Text = Trim(codigo)
                        txtDes_Anexo.Text = Trim(Descripcion)
                        strSQL = "SELECT num_ruc FROM CN_AnexosContables WHERE Cod_TipAnEX = '" & txtCod_TipAne.Text & "' AND Cod_Anxo = '" & txtCod_Anexo.Text & "'"
                        txtRuc = Trim(DevuelveCampo(strSQL, cCONNECT))

                        SendKeys "{TAB}"
                    Else
                    End If
                End If
                Set oTipo = Nothing
                Set RS = Nothing

    End Select

End Sub


