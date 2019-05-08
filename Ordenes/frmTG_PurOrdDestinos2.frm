VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmTG_PurOrdDestinos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccione P.O. Destino"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraGenerarCajas 
      Caption         =   "Completar Información:"
      Height          =   1515
      Left            =   2550
      TabIndex        =   12
      Top             =   570
      Visible         =   0   'False
      Width           =   2820
      Begin VB.TextBox txtNum_Packing 
         Height          =   285
         Left            =   1515
         TabIndex        =   13
         Top             =   345
         Width           =   960
      End
      Begin FunctionsButtons.FunctButt funcGenerarCajas 
         Height          =   510
         Left            =   165
         TabIndex        =   15
         Top             =   810
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmTG_PurOrdDestinos2.frx":0000
         Orientacion     =   0
         Style           =   1
         Language        =   1
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Packing"
         Height          =   195
         Left            =   465
         TabIndex        =   14
         Top             =   375
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdGeneraCajasMultipleDest 
      Caption         =   "Generar Cajas"
      Height          =   450
      Left            =   3945
      TabIndex        =   11
      Top             =   3675
      Width           =   1530
   End
   Begin VB.CommandButton cmdRegistrarCantidades 
      Caption         =   "Registrar Cantidades"
      Height          =   450
      Left            =   2220
      TabIndex        =   10
      Top             =   3675
      Width           =   1530
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   450
      Left            =   5700
      TabIndex        =   9
      Top             =   3675
      Width           =   1530
   End
   Begin VB.Frame fraPoDestino 
      Caption         =   "Nuevo P.O./Destino"
      Height          =   1905
      Left            =   1980
      TabIndex        =   2
      Top             =   1275
      Visible         =   0   'False
      Width           =   4365
      Begin VB.TextBox txtCod_PurOrd_Destino 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1095
         MaxLength       =   20
         TabIndex        =   5
         Top             =   270
         Width           =   3120
      End
      Begin VB.TextBox txtCod_AlmacenCliente 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1095
         MaxLength       =   3
         TabIndex        =   4
         Top             =   735
         Width           =   630
      End
      Begin VB.TextBox txtDes_AlmacenCliente 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1830
         MaxLength       =   20
         TabIndex        =   3
         Top             =   735
         Width           =   2340
      End
      Begin FunctionsButtons.FunctButt funcPoDestino 
         Height          =   510
         Left            =   930
         TabIndex        =   6
         Top             =   1215
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmTG_PurOrdDestinos2.frx":0097
         Orientacion     =   0
         Style           =   1
         Language        =   1
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "P.O."
         Height          =   195
         Left            =   255
         TabIndex        =   8
         Top             =   345
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
         Height          =   195
         Left            =   255
         TabIndex        =   7
         Top             =   795
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdDestinos 
      Caption         =   "Agregar P.O. / Destino"
      Height          =   450
      Left            =   495
      TabIndex        =   1
      Top             =   3675
      Width           =   1530
   End
   Begin GridEX20.GridEX Gridex1 
      Height          =   3480
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   6138
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmTG_PurOrdDestinos2.frx":012E
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmTG_PurOrdDestinos2.frx":0480
      Column(2)       =   "frmTG_PurOrdDestinos2.frx":0548
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmTG_PurOrdDestinos2.frx":05EC
      FormatStyle(2)  =   "frmTG_PurOrdDestinos2.frx":0724
      FormatStyle(3)  =   "frmTG_PurOrdDestinos2.frx":07D4
      FormatStyle(4)  =   "frmTG_PurOrdDestinos2.frx":0888
      FormatStyle(5)  =   "frmTG_PurOrdDestinos2.frx":0960
      FormatStyle(6)  =   "frmTG_PurOrdDestinos2.frx":0A18
      FormatStyle(7)  =   "frmTG_PurOrdDestinos2.frx":0AF8
      FormatStyle(8)  =   "frmTG_PurOrdDestinos2.frx":0FB0
      ImageCount      =   1
      ImagePicture(1) =   "frmTG_PurOrdDestinos2.frx":13FC
      PrinterProperties=   "frmTG_PurOrdDestinos2.frx":174E
   End
End
Attribute VB_Name = "frmTG_PurOrdDestinos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Codigo As String

Public Descripcion As String

Public sAccionName As String

Public sModoWizard As String

Public sCod_Cliente As String

Public sCod_PurOrd As String

Public scod_LotPurOrd As String

Public sCod_TemCli As String

Public oParent As Object
Public sCod_EStCli As String

Dim Strsql  As String

Public sFlgOpcion_Nueva As String

Private Sub cmdDestinos_Click()
    Me.txtCod_PurOrd_Destino.Text = ""
    Me.txtCod_AlmacenCliente.Text = ""
    Me.txtCod_PurOrd_Destino.Text = sCod_PurOrd
    Me.fraPoDestino.Visible = True
    Me.cmdSalir.Visible = False
    Me.cmdDestinos.value = False
End Sub

Private Sub cmdGeneraCajasMultipleDest_Click()
    Me.fraGenerarCajas.Visible = True
End Sub

Private Sub cmdRegistrarCantidades_Click()
    Load frmMatrizDestinoEmpaque
    Set frmMatrizDestinoEmpaque.oParent = Me
    frmMatrizDestinoEmpaque.sFlgOpcion_Nueva = Me.sFlgOpcion_Nueva
    frmMatrizDestinoEmpaque.sAccionName = sAccionName
    frmMatrizDestinoEmpaque.sModoWizard = sModoWizard
    frmMatrizDestinoEmpaque.sCod_Cliente = sCod_Cliente
    frmMatrizDestinoEmpaque.sCod_PurOrd = sCod_PurOrd
    frmMatrizDestinoEmpaque.scod_LotPurOrd = scod_LotPurOrd
    frmMatrizDestinoEmpaque.sCod_EStCli = sCod_EStCli
    frmMatrizDestinoEmpaque.sSec_PurOrd = Gridex1.value(Gridex1.Columns("Sec_PurOrd").Index)
    frmMatrizDestinoEmpaque.sCod_TemCli = sCod_TemCli
    frmMatrizDestinoEmpaque.sCod_AlmacenCliente = Gridex1.value(Gridex1.Columns("Cod_AlmacenCliente").Index)
    frmMatrizDestinoEmpaque.sDes_AlmacenCliente = Gridex1.value(Gridex1.Columns("NOM_ALMACENCLIENTE").Index)
    
    frmMatrizDestinoEmpaque.BUSCAR
    frmMatrizDestinoEmpaque.Show vbModal
    Set frmMatrizDestinoEmpaque = Nothing
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub funcGenerarCajas_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            If Gridex1.RowCount = 0 Then Exit Sub
            GenerarCajasMultipleDestinos
        Case "CANCELAR"
            Me.fraGenerarCajas.Visible = False
    End Select
End Sub

Private Sub funcPoDestino_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            GrabarNuevoPODestino
        Case "CANCELAR"
            Me.fraPoDestino.Visible = False
            Me.cmdSalir.Visible = True
            Me.cmdDestinos.Visible = True
    End Select
End Sub

Private Sub Gridex1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = True
End Sub

Private Sub txtCod_AlmacenCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(txtCod_AlmacenCliente.Text) = "" Then
            BUSCA_ALMACEN (3)
        Else
            BUSCA_ALMACEN (1)
        End If
    End If

End Sub

Public Sub BUSCA_ALMACEN(Tipo As Integer)
    Select Case Tipo
        Case 1:
                    Strsql = "SELECT Descripcion FROM TG_CLIENTE_ALMACEN WHERE Cod_Cliente='" & sCod_Cliente & "' AND Cod_AlmacenCliente = '" & Trim(Me.txtCod_AlmacenCliente.Text) & "'"
                    Me.txtDes_AlmacenCliente.Text = Trim(DevuelveCampo(Strsql, cCONNECT))
                    funcPoDestino.SetFocus
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.sQuery = "SELECT Cod_AlmacenCliente as Codigo, Descripcion FROM TG_CLIENTE_ALMACEN WHERE Cod_Cliente='" & sCod_Cliente & "' AND Descripcion like '%" & Trim(txtDes_AlmacenCliente.Text) & "%' order by 2"
                    Else
                        oTipo.sQuery = "SELECT Cod_AlmacenCliente as Codigo, Descripcion FROM TG_CLIENTE_ALMACEN WHERE Cod_Cliente='" & sCod_Cliente & "' order by 2"
                    End If
                    
                    oTipo.Cargar_Datos
                    oTipo.Show 1
                    If Codigo <> "" Then
                         Me.txtCod_AlmacenCliente.Text = Trim(Codigo)
                         Me.txtDes_AlmacenCliente.Text = Trim(Descripcion)
                         funcPoDestino.SetFocus
                         Codigo = "": Descripcion = ""
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
    End Select
End Sub



Private Function GrabarNuevoPODestino() As Boolean
On Error GoTo errx
Dim ssql As String

If RTrim(txtCod_PurOrd_Destino.Text) = "" Then
    txtCod_PurOrd_Destino.SetFocus
    Exit Function
End If

If RTrim(txtCod_AlmacenCliente.Text) = "" Then
    txtCod_AlmacenCliente.SetFocus
    Exit Function
End If

ssql = "EXEC UP_TG_PURORD_DESTINOS '$' , '$' , '$' , '$', '$' ,'I'"
ssql = VBsprintf(ssql, sCod_Cliente, sCod_PurOrd, "", txtCod_PurOrd_Destino.Text, txtCod_AlmacenCliente.Text)

ExecuteCommandSQL cCONNECT, ssql
GrabarNuevoPODestino = True
BUSCAR

Me.fraPoDestino.Visible = False
Me.cmdSalir.Visible = True
Me.cmdDestinos.Visible = True

Exit Function
errx:
    Err.Raise Err.Number, Err.source, Err.Description
End Function

Public Function BUSCAR() As Boolean
On Error GoTo errx
Dim ssql As String
Dim vBookmark As Variant
Dim colTemp  As JSColumn
Dim irow As Long
Dim iCol As Long

ssql = "SM_MUESTRA_TG_PURORD_DESTINOS1 '$','$' "
ssql = VBsprintf(ssql, sCod_Cliente, sCod_PurOrd)

vBookmark = Gridex1.Row
Gridex1.ClearFields

Set Gridex1.ADORecordset = CargarRecordSetDesconectado(ssql, cCONNECT)

Gridex1.Columns("SEC_PURORD").Caption = "Sec.P.O."
Gridex1.Columns("COD_PURORD_DESTINO").Caption = "P.O.Destino"
Gridex1.Columns("COD_ALMACENCLIENTE").Caption = "Destino"
Gridex1.Columns("NOM_ALMACENCLIENTE").Caption = "Nombre Destino"

Gridex1.Row = vBookmark

Gridex1.ContinuousScroll = True
Exit Function

errx:
    Err.Raise Err.Number, Err.source, Err.Description
End Function



Public Function CargarRecordSetDesconectado(ByVal ssql As String, ByVal cCONNECT As String) As ADODB.Recordset
Dim rsBD As ADODB.Recordset
Dim rsGridEx As ADODB.Recordset
Dim ofield As Object
Dim oCon As ADODB.Connection

    Set oCon = New ADODB.Connection
    
    oCon.CursorLocation = adUseClient
    oCon.Open cCONNECT
    oCon.CommandTimeout = 900
    
    Set rsBD = New ADODB.Recordset
    Set rsBD.ActiveConnection = oCon
     
    rsBD.CursorLocation = adUseClient
    rsBD.CursorType = adOpenStatic
    
    rsBD.Open ssql

    Set rsGridEx = New ADODB.Recordset
    rsGridEx.CursorLocation = adUseClient
    Set rsGridEx.ActiveConnection = Nothing

    For Each ofield In rsBD.Fields
        If RTrim(ofield.Name) <> "" Then
            rsGridEx.Fields.Append ofield.Name, ofield.Type, ofield.DefinedSize, adFldIsNullable
            rsGridEx.Fields(ofield.Name).NumericScale = rsBD.Fields(ofield.Name).NumericScale
            rsGridEx.Fields(ofield.Name).DefinedSize = rsBD.Fields(ofield.Name).DefinedSize
            rsGridEx.Fields(ofield.Name).Precision = rsBD.Fields(ofield.Name).Precision
        End If
    Next
    rsGridEx.Open
           
    If rsBD.RecordCount Then
        rsBD.MoveFirst
        Do While Not rsBD.EOF
            rsGridEx.AddNew
            For Each ofield In rsBD.Fields
                If RTrim(ofield.Name) <> "" Then
                    rsGridEx.Fields(ofield.Name).value = FixData(rsBD.Fields(ofield.Name).value, rsBD.Fields(ofield.Name))
                End If
            Next
            rsGridEx.Update
            rsBD.MoveNext
        Loop
    End If

    Set CargarRecordSetDesconectado = rsGridEx
    
End Function

Public Function SetGeneralGridEX(ByRef GridEx As GridEX20.GridEx, ByVal iFixsCols As Integer, ByVal iTipoColorBack As Integer)

    If iFixsCols > 0 Then
        GridEx.FrozenColumns = iFixsCols
    End If
    
    If iTipoColorBack = 1 Then
        GridEx.BackColor = &H80000018
        GridEx.BackColorBkg = &H80000018
        GridEx.GridLines = jgexGLVertical
        GridEx.GridLineStyle = jgexGLSSmallDots
    Else
        GridEx.BackColor = &H80000005
        GridEx.BackColorBkg = &H80000005
        GridEx.GridLines = jgexGLBoth
        GridEx.GridLineStyle = jgexGLSSmallDots
    End If
    
End Function



Function FixData(wtexto As Variant, ofield As ADODB.FIELD)
   If IsNull(wtexto) Or Len(Trim(wtexto)) = 0 Then
   
      Select Case ofield.Type
        Case adBigInt, adCurrency, adDecimal, adDouble, adInteger, adNumeric, adSingle
            wtexto = 0
        Case adBoolean
            wtexto = False
        Case adDate
            wtexto = Empty
        Case adChar, adVarChar
            wtexto = ""
      End Select
   End If
   FixData = wtexto
End Function

Private Sub txtNum_Packing_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        
    End If
End Sub

Private Sub GenerarCajasMultipleDestinos()
On Error GoTo errx
Dim ssql As String

ssql = "UP_GENERA_PACKING_MULTIPLE_DESTINOS '$', '$','$','$' ,$ "
ssql = VBsprintf(ssql, sCod_Cliente, sCod_PurOrd, scod_LotPurOrd, Gridex1.value(Gridex1.Columns("Sec_PurOrd").Index), Val(txtNum_Packing.Text))

ExecuteCommandSQL cCONNECT, ssql
Mensaje kMESSAGE_INF_PROCESS_SATISFACTO

Me.fraGenerarCajas.Visible = False

Exit Sub
errx:
    errores Err.Number
End Sub
