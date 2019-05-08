VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "gridex20.ocx"
Begin VB.Form FrmCambioPrecioPorTalla 
   Caption         =   "Cambio de Precio Por Talla"
   ClientHeight    =   4155
   ClientLeft      =   75
   ClientTop       =   375
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FRAMODIFICAR 
      Caption         =   "Modificar Precio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2715
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtprecio 
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Width           =   930
      End
      Begin FunctionsButtons.FunctButt FunctButt3 
         Height          =   510
         Left            =   720
         TabIndex        =   2
         Top             =   840
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"FrmCambioPrecioPorTalla.frx":0000
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Precio :"
         Height          =   195
         Left            =   840
         TabIndex        =   3
         Top             =   420
         Width           =   540
      End
   End
   Begin GridEX20.GridEX Gridex1 
      Height          =   3480
      Left            =   30
      TabIndex        =   4
      Top             =   30
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   6138
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "FrmCambioPrecioPorTalla.frx":008A
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "FrmCambioPrecioPorTalla.frx":03DC
      Column(2)       =   "FrmCambioPrecioPorTalla.frx":04A4
      FormatStylesCount=   8
      FormatStyle(1)  =   "FrmCambioPrecioPorTalla.frx":0548
      FormatStyle(2)  =   "FrmCambioPrecioPorTalla.frx":0680
      FormatStyle(3)  =   "FrmCambioPrecioPorTalla.frx":0730
      FormatStyle(4)  =   "FrmCambioPrecioPorTalla.frx":07E4
      FormatStyle(5)  =   "FrmCambioPrecioPorTalla.frx":08BC
      FormatStyle(6)  =   "FrmCambioPrecioPorTalla.frx":0974
      FormatStyle(7)  =   "FrmCambioPrecioPorTalla.frx":0A54
      FormatStyle(8)  =   "FrmCambioPrecioPorTalla.frx":0F0C
      ImageCount      =   1
      ImagePicture(1) =   "FrmCambioPrecioPorTalla.frx":1358
      PrinterProperties=   "FrmCambioPrecioPorTalla.frx":16AA
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   7110
      TabIndex        =   5
      Top             =   3570
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmCambioPrecioPorTalla.frx":1882
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmCambioPrecioPorTalla"
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


Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "MODIFICAR"
            txtprecio.Text = Gridex1.value(Gridex1.Columns("precio").Index)
            FRAMODIFICAR.Visible = True
            txtprecio.SetFocus
        Case "SALIR"
            Unload Me
    End Select
End Sub

Private Sub FunctButt3_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "GRABAR"
            Grabar
            FRAMODIFICAR.Visible = False
            BUSCAR
        Case "SALIR"
            FRAMODIFICAR.Visible = False
    End Select
End Sub

Private Sub Grabar()
On Error GoTo errx
Dim ssql As String, sprecio As Double
Dim pc As String
Dim fecha As Date
pc = ComputerName()
fecha = Now()
vusu = vusu

If txtprecio.Text = "" Then
    sprecio = 0
Else
    sprecio = txtprecio.Text
End If


ssql = "SM_TG_LOTESTUPDATEDATGEN_NEW_POR_TALLA '$', '$','$','$','$','$','$','$','$','$',$ "
ssql = VBsprintf(ssql, Trim(sCod_Cliente), Trim(sCod_PurOrd), Trim(scod_LotPurOrd), Trim(sCod_EStCli), _
Trim(Gridex1.value(Gridex1.Columns("COD_COLCLI").Index)), _
Trim(Gridex1.value(Gridex1.Columns("COD_TALLA").Index)), _
Trim(sCod_TemCli), _
Trim(pc), Trim(fecha), vusu, sprecio)

ExecuteCommandSQL cCONNECT, ssql
Mensaje kMESSAGE_INF_PROCESS_SATISFACTO


Exit Sub
errx:
    errores Err.Number
End Sub


Private Sub Gridex1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = True
End Sub


Public Function BUSCAR() As Boolean
On Error GoTo errx
Dim ssql As String
Dim vBookmark As Variant
Dim colTemp  As JSColumn
Dim irow As Long
Dim iCol As Long

ssql = "SM_MUESTRA_PRECIOS_TG_LOTCOLTAL '$','$' ,'$','$'"
ssql = VBsprintf(ssql, Trim(sCod_Cliente), Trim(sCod_PurOrd), Trim(scod_LotPurOrd), Trim(sCod_EStCli))

vBookmark = Gridex1.Row
Gridex1.ClearFields

Set Gridex1.ADORecordset = CargarRecordSetDesconectado(ssql, cCONNECT)

Gridex1.Columns("COD_CLIENTE").Caption = "Cod"
Gridex1.Columns("ABR_CLIENTE").Caption = "Cliente"
Gridex1.Columns("COD_PURORD").Caption = "Po"
Gridex1.Columns("COD_LOTPURORD").Caption = "Lote"
Gridex1.Columns("COD_ESTCLI").Caption = "Estilo Cliente"
Gridex1.Columns("COD_COLCLI").Caption = "Color Cliente"
Gridex1.Columns("COD_TALLA").Caption = "Talla"
Gridex1.Columns("PRECIO").Caption = "Precio"


Gridex1.Columns("COD_CLIENTE").Width = 800
Gridex1.Columns("ABR_CLIENTE").Width = 800
Gridex1.Columns("COD_PURORD").Width = 2000
Gridex1.Columns("COD_LOTPURORD").Width = 800
Gridex1.Columns("COD_ESTCLI").Width = 1000
Gridex1.Columns("COD_COLCLI").Width = 2000
Gridex1.Columns("COD_TALLA").Width = 800
Gridex1.Columns("PRECIO").Width = 800

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
        GridEx.Gridlines = jgexGLVertical
        GridEx.GridLineStyle = jgexGLSSmallDots
    Else
        GridEx.BackColor = &H80000005
        GridEx.BackColorBkg = &H80000005
        GridEx.Gridlines = jgexGLBoth
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

Private Sub txtprecio_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
         SoloNumeros txtprecio, KeyAscii, True, 3, 4
'        FunctButt3.SetFocus
'    End If

End Sub


