VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMatrizDestinoEmpaque 
   Caption         =   "Ingresar Color/Talla a Nivel P.O. Destino:"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   9795
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTipEmpaque 
      Caption         =   "Nuevo Tipo de Empaque"
      Height          =   1860
      Left            =   2580
      TabIndex        =   5
      Top             =   3705
      Visible         =   0   'False
      Width           =   4575
      Begin VB.TextBox txtDes_TipEmpaque 
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
         Left            =   1755
         MaxLength       =   50
         TabIndex        =   8
         Top             =   720
         Width           =   2685
      End
      Begin VB.TextBox txtCod_TipEmpaque 
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
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   7
         Top             =   330
         Width           =   570
      End
      Begin FunctionsButtons.FunctButt funcTip_Empaque 
         Height          =   510
         Left            =   1035
         TabIndex        =   9
         Top             =   1200
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmMatrizDestinoEmpaque.frx":0000
         Orientacion     =   0
         Style           =   1
         Language        =   1
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nombre de Empaque: "
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   750
         Width           =   1590
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Empaque: "
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   360
         Width           =   1350
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir Sin Grabar"
      Height          =   450
      Left            =   5820
      TabIndex        =   4
      Top             =   6840
      Width           =   1530
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar Cantidades"
      Height          =   450
      Left            =   4110
      TabIndex        =   3
      Top             =   6840
      Width           =   1530
   End
   Begin VB.CommandButton cmdTip_Empaque 
      Caption         =   "Agregar Tipo de             Empaque"
      Height          =   450
      Left            =   2430
      TabIndex        =   2
      Top             =   6855
      Width           =   1530
   End
   Begin VB.Frame Frame1 
      Height          =   6600
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   9705
      Begin GridEX20.GridEX GridEX1 
         Height          =   6330
         Left            =   60
         TabIndex        =   1
         Top             =   165
         Width           =   9540
         _ExtentX        =   16828
         _ExtentY        =   11165
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         Options         =   8
         RecordsetType   =   1
         GroupByBoxVisible=   0   'False
         ImageCount      =   1
         ImagePicture1   =   "frmMatrizDestinoEmpaque.frx":0097
         DataMode        =   1
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "frmMatrizDestinoEmpaque.frx":03E9
         Column(2)       =   "frmMatrizDestinoEmpaque.frx":04B1
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmMatrizDestinoEmpaque.frx":0555
         FormatStyle(2)  =   "frmMatrizDestinoEmpaque.frx":068D
         FormatStyle(3)  =   "frmMatrizDestinoEmpaque.frx":073D
         FormatStyle(4)  =   "frmMatrizDestinoEmpaque.frx":07F1
         FormatStyle(5)  =   "frmMatrizDestinoEmpaque.frx":08C9
         FormatStyle(6)  =   "frmMatrizDestinoEmpaque.frx":0981
         FormatStyle(7)  =   "frmMatrizDestinoEmpaque.frx":0A61
         FormatStyle(8)  =   "frmMatrizDestinoEmpaque.frx":0F19
         ImageCount      =   1
         ImagePicture(1) =   "frmMatrizDestinoEmpaque.frx":1365
         PrinterProperties=   "frmMatrizDestinoEmpaque.frx":16B7
      End
   End
End
Attribute VB_Name = "frmMatrizDestinoEmpaque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sAccionName         As String

Public sModoWizard         As String

Public sCod_Cliente        As String

Public sCod_PurOrd         As String

Public sCod_LotPurOrd      As String

Public sSec_PurOrd         As String

Public sCod_TemCli         As String

Public sCod_AlmacenCliente As String

Public sDes_AlmacenCliente As String

Public oParent             As Object

''Public sCod_ColCli As String
Public sCod_EstCli         As String
''Public sCod_Talla As String

Public dNum_PreReq         As Long

Public Codigo              As String

Public Descripcion         As String

Dim strSql                 As String

Public sFlgOpcion_Nueva    As String

Private Sub cmdGrabar_Click()
    GrabarMatrizDestino False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdTip_Empaque_Click()
    Me.fraTipEmpaque.Visible = True
    Me.cmdGrabar.Visible = False
    Me.cmdSalir.Visible = False
    Me.cmdTip_Empaque.Visible = False
    Me.txtCod_TipEmpaque.Text = ""
    Me.txtDes_TipEmpaque.Text = ""
    
End Sub

Private Sub cmdGrabarySalir_Click()
    GrabarMatrizDestino True
End Sub

Private Sub funcTip_Empaque_ActionClick(ByVal Index As Integer, _
                                        ByVal ActionType As Integer, _
                                        ByVal ActionName As String)

    Select Case ActionName

        Case "ACEPTAR"
            GrabarNuevoTipoEmpaque

        Case "CANCELAR"
            Me.fraTipEmpaque.Visible = False
            Me.cmdGrabar.Visible = True
            Me.cmdSalir.Visible = True
            Me.cmdTip_Empaque.Visible = True
    End Select

End Sub

Private Function GrabarNuevoTipoEmpaque() As Boolean

    On Error GoTo errx

    Dim sSQl As String

    If RTrim(txtCod_TipEmpaque.Text) = "" Then
        txtCod_TipEmpaque.SetFocus

        Exit Function

    End If

    sSQl = "EXEC UP_TG_Cliente_TipEmpaque '$' , '$' , '$' ,'I'"
    sSQl = VBsprintf(sSQl, sCod_Cliente, txtCod_TipEmpaque.Text, txtDes_TipEmpaque.Text)

    ExecuteCommandSQL cCONNECT, sSQl
    GrabarNuevoTipoEmpaque = True
    BUSCAR

    Me.fraTipEmpaque.Visible = False
    Me.cmdGrabar.Visible = True
    Me.cmdSalir.Visible = True
    Me.cmdTip_Empaque.Visible = True

    Exit Function

errx:
    Err.Raise Err.Number, Err.source, Err.Description
End Function

Private Function GrabarMatrizDestino(ByVal bGrabarySailr As Boolean) As Boolean

    On Error GoTo errx

    Dim sSQl             As String

    Dim vRowBookmark     As Variant

    Dim vColBookmark     As Variant

    Dim irow             As Integer

    Dim iCol             As Integer

    Dim vNewPrendas      As Long

    Dim vOldPrendas      As Long

    Dim vDiferencia      As Long

    Dim sValorInicializa As String

    Dim oMessage         As clsMessages

    Dim vProporcion      As Long

    vRowBookmark = Me.GridEX1.Row
    vColBookmark = Me.GridEX1.col

    sValorInicializa = "0"

    For irow = 1 To GridEX1.RowCount
        GridEX1.Row = irow

        For iCol = 1 To GridEX1.Columns.count

            If Mid(GridEX1.Columns(iCol).Key, 1, 6) = "PRENDA" Then
                vOldPrendas = GridEX1.value(iCol - 1)
                vNewPrendas = GridEX1.value(iCol)
                vDiferencia = vNewPrendas  '(SIEMPRE PASO LA CANTIDAD NUEVA)
            End If
        
            If Mid(GridEX1.Columns(iCol).Key, 1, 6) = "PROPOR" Then
                vProporcion = GridEX1.value(iCol)
            
                sSQl = "EXEC UP_TG_LotColTal_Destinos_Empaque '$' , '$' , '$' ,'$','$','$','$','$', $ , $ , 'PRENDA' , '$' ,'$','$','$'"
                sSQl = VBsprintf(sSQl, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, GridEX1.value(GridEX1.Columns("Cod_ColCli").Index), GridEX1.value(GridEX1.Columns("Cod_Talla").Index), sSec_PurOrd, Mid(GridEX1.Columns(iCol).Key, 8, 3), vDiferencia, 0, sValorInicializa, vusu, ComputerName, Me.sFlgOpcion_Nueva, vProporcion)
        
                ExecuteCommandSQL cCONNECT, sSQl
            
                If sValorInicializa = "0" Then
                    sValorInicializa = "1"
                End If
            End If

        Next
    Next

    GrabarMatrizDestino = True
    Mensaje kMESSAGE_INF_PROCESS_SATISFACTO

    oParent.oParent.oParent.BUSCAR
    Unload Me

    Exit Function

errx:
    sValorInicializa = "0"
    'Err.Raise Err.Number, Err.source, Err.Description
    errores Err.Number
    GridEX1.Refresh
    BUSCAR
End Function

Public Function BUSCAR() As Boolean

    On Error GoTo errx

    Dim sSQl      As String

    Dim vBookmark As Variant

    Dim colTemp   As JSColumn

    Dim irow      As Long

    Dim iCol      As Long

    Me.Caption = Me.Caption & sCod_AlmacenCliente & " " & sDes_AlmacenCliente

    sSQl = "SM_MUESTRA_TG_LotColTal_Destinos_Empaque '$','$','$','$','$', '$'"
    sSQl = VBsprintf(sSQl, sCod_Cliente, sCod_PurOrd, sCod_LotPurOrd, sCod_EstCli, sSec_PurOrd, sCod_TemCli)

    vBookmark = GridEX1.Row
    GridEX1.ClearFields

    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQl, cCONNECT)

    If GridEX1.RowCount > 0 Then
        GridEX1.Columns("COD_COLCLI").Caption = "Color Cliente"
        GridEX1.Columns("NOM_COLOR").Caption = "Nombre de Color "
        GridEX1.Columns("COD_TALLA").Caption = "Talla"
    
        For irow = 1 To GridEX1.RowCount
            GridEX1.Row = irow

            For iCol = 1 To GridEX1.Columns.count

                Select Case Mid(GridEX1.Columns(iCol).Key, 1, 6)

                    Case "PRENDA"
                        GridEX1.Columns(iCol).Caption = Mid(GridEX1.Columns(iCol).Key, 12)
                        GridEX1.Columns(iCol).Width = 1200

                    Case "PROPOR"
                        GridEX1.Columns(iCol).Caption = "Propor " & Mid(GridEX1.Columns(iCol).Key, 12)
                        GridEX1.Columns(iCol).Width = 1600

                    Case "PRENDO", "PRECIO", "IMPORT"
                        GridEX1.Columns(iCol).Visible = False

                    Case Else
                        GridEX1.Columns(iCol).Visible = True
                        GridEX1.Columns(iCol).Width = 1000
                End Select

            Next
        Next

    End If

    GridEX1.Row = vBookmark
    GridEX1.FrozenColumns = 3
    GridEX1.ContinuousScroll = True

    Exit Function

errx:
    Err.Raise Err.Number, Err.source, Err.Description
End Function

Public Function CargarRecordSetDesconectado(ByVal sSQl As String, _
                                            ByVal cCONNECT As String) As ADODB.Recordset

    Dim rsBD     As ADODB.Recordset

    Dim rsGridEx As ADODB.Recordset

    Dim ofield   As Object

    Dim oCon     As ADODB.Connection

    Set oCon = New ADODB.Connection
    
    oCon.CursorLocation = adUseClient
    oCon.Open cCONNECT
    oCon.CommandTimeout = 900
    
    Set rsBD = New ADODB.Recordset
    Set rsBD.ActiveConnection = oCon
     
    rsBD.CursorLocation = adUseClient
    rsBD.CursorType = adOpenStatic
    
    rsBD.Open sSQl

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

Public Function SetGeneralGridEX(ByRef GridEx As GridEX20.GridEx, _
                                 ByVal iFixsCols As Integer, _
                                 ByVal iTipoColorBack As Integer)

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

Private Sub Gridex1_BeforeColEdit(ByVal ColIndex As Integer, _
                                  ByVal Cancel As GridEX20.JSRetBoolean)

    If Mid(GridEX1.Columns(ColIndex).Key, 1, 6) <> "PRENDA" And Mid(GridEX1.Columns(ColIndex).Key, 1, 6) <> "PROPOR" Then
        Cancel = True
    End If

End Sub

