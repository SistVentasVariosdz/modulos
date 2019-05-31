VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmVerificaMatrizDetalle 
   Caption         =   "Cambio de Estado Matriz Destinos/Empaques"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   450
      Left            =   7350
      TabIndex        =   1
      Top             =   4350
      Width           =   1530
   End
   Begin GridEX20.GridEX Gridex1 
      Height          =   4200
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   7408
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmVerificaMatrizDetalle.frx":0000
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmVerificaMatrizDetalle.frx":0352
      Column(2)       =   "frmVerificaMatrizDetalle.frx":041A
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmVerificaMatrizDetalle.frx":04BE
      FormatStyle(2)  =   "frmVerificaMatrizDetalle.frx":05F6
      FormatStyle(3)  =   "frmVerificaMatrizDetalle.frx":06A6
      FormatStyle(4)  =   "frmVerificaMatrizDetalle.frx":075A
      FormatStyle(5)  =   "frmVerificaMatrizDetalle.frx":0832
      FormatStyle(6)  =   "frmVerificaMatrizDetalle.frx":08EA
      FormatStyle(7)  =   "frmVerificaMatrizDetalle.frx":09CA
      FormatStyle(8)  =   "frmVerificaMatrizDetalle.frx":0E82
      ImageCount      =   1
      ImagePicture(1) =   "frmVerificaMatrizDetalle.frx":12CE
      PrinterProperties=   "frmVerificaMatrizDetalle.frx":1620
   End
End
Attribute VB_Name = "frmVerificaMatrizDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Codigo         As String

Public Descripcion    As String

Public sAccionName    As String

Public sModoWizard    As String

Public sCod_Cliente   As String

Public sCod_PurOrd    As String

Public sCod_LotPurOrd As String

Public sCod_TemCli    As String

Public oParent        As Object

Public sCod_EstCli    As String

Public rsData         As ADODB.Recordset

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Function BUSCAR() As Boolean

    On Error GoTo errx

    Dim sSQl      As String

    Dim vBookmark As Variant

    Dim colTemp   As JSColumn

    Dim irow      As Long

    Dim iCol      As Long

    vBookmark = GridEX1.Row
    GridEX1.ClearFields

    Set GridEX1.ADORecordset = rsData 'CargarRecordSetDesconectado(ssql, cCONNECT)

    If GridEX1.RowCount > 0 Then
        GridEX1.Columns("Status").Visible = False
        GridEX1.Columns("Cod_ColCli").Width = 1000
        GridEX1.Columns("Cod_Talla").Width = 1000
        GridEX1.Columns("NOM_COLOR").Width = 2000
        GridEX1.Columns("NUM_PREREQ").Width = 1500
        GridEX1.Columns("NUM_PRENDASDETALLE").Width = 1500
    
        GridEX1.Columns("Cod_ColCli").Caption = "Color Cliente"
        GridEX1.Columns("Cod_Talla").Caption = "Talla"
        GridEX1.Columns("NOM_COLOR").Caption = "Nombre Color"
        GridEX1.Columns("NUM_PREREQ").Caption = "Prenda Col/Tall"
        GridEX1.Columns("NUM_PRENDASDETALLE").Caption = "Prenda Dest/Empaq"
    End If

    GridEX1.Row = vBookmark

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
    Cancel = True
End Sub

