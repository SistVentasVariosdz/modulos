VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBusqGeneral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Select"
   Begin GridEX20.GridEX gexList 
      Height          =   3840
      Left            =   150
      TabIndex        =   0
      Top             =   75
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   6773
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      TabKeyBehavior  =   1
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      RowHeaders      =   -1  'True
      DataMode        =   1
      BackColorBkg    =   -2147483624
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmBusqGeneral.frx":0000
      Column(2)       =   "frmBusqGeneral.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmBusqGeneral.frx":016C
      FormatStyle(2)  =   "frmBusqGeneral.frx":02A4
      FormatStyle(3)  =   "frmBusqGeneral.frx":0354
      FormatStyle(4)  =   "frmBusqGeneral.frx":0408
      FormatStyle(5)  =   "frmBusqGeneral.frx":04E0
      FormatStyle(6)  =   "frmBusqGeneral.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmBusqGeneral.frx":0678
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3045
      TabIndex        =   1
      Tag             =   "&OK"
      Top             =   4215
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4290
      TabIndex        =   2
      Tag             =   "&Cancel"
      Top             =   4215
      Width           =   1185
   End
End
Attribute VB_Name = "frmBusqGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public sQuery As String
'Dim Rs_Carga As New ADODB.Recordset
Public CODIGO As String
Public DESCRIPCION As String
Public Paso As Boolean

Sub Cargar_Datos()
On Error GoTo Cargar_DatosErr

'Rs_Carga.ActiveConnection = cConnect
'Rs_Carga.CursorType = adOpenStatic
'Rs_Carga.CursorLocation = adUseClient
'Rs_Carga.LockType = adLockReadOnly
'Rs_Carga.Open sQuery
Set Me.gexList.ADORecordset = CargarRecordSetDesconectado(sQuery, cConnect)
If gexList.Columns.Count = 2 Then
    gexList.Columns(1).Width = 1200
    gexList.Columns(2).Width = 5000
End If
Exit Sub
Cargar_DatosErr:
    ErrorHandler err, "Cargar_Datos"
End Sub

'Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then DGridlista_DblClick
'End Sub


'Private Sub DGridlista_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'If DGridLista.RowContaining(y) >= 1 And DGridLista.RowContaining(y) <= Rs_Carga.RecordCount Then
'    DGridLista.Bookmark = DGridLista.RowBookmark(DGridLista.RowContaining(y))
'End If
'End Sub
'Private Sub Form_Load()
'Call FormSet(Me)
'FormateaGrid DGridLista
'DGridLista.Columns(1).Width = 4000
'End Sub
'Private Sub Form_Unload(Cancel As Integer)
'
'    Set Rs_Carga = Nothing
'End Sub
Public Sub cmdAceptar_Click()
    gexList_DblClick
End Sub
Public Sub cmdCancelar_Click()
    With oParent
        .CODIGO = ""
        .DESCRIPCION = ""
        '.Paso = False
    End With

Unload Me
End Sub

Private Sub gexList_DblClick()
On Error Resume Next
If gexList.RowCount > 0 Then
    With oParent
        .CODIGO = gexList.Value(gexList.Columns(1).Index)
        .DESCRIPCION = gexList.Value(gexList.Columns(2).Index)
        If oParent.Name = "FrmDetalleHilTel" Then
            .Cod_color = gexList.Value(gexList.Columns("Cod Color").Index)
        End If
        .Paso = True
    End With
    
    If gexList.Columns.Count > 3 Then
        If oParent.Name = "FrmDetalleTelaCa" Then
            With oParent
                .Cod_Comb = IIf(IsNull(gexList.Value(gexList.Columns("combinacion").Index)), "", Left(gexList.Value(gexList.Columns("combinacion").Index), 3))
                .Cod_color = Left(gexList.Value(gexList.Columns("color").Index), 6)
                .Cod_Talla = gexList.Value(gexList.Columns("talla").Index)
            End With
        ElseIf oParent.Name = "FrmDetalleTelCru" Then
            With oParent
                .Cod_Comb = IIf(IsNull(gexList.Value(gexList.Columns("combinacion").Index)), "", Left(gexList.Value(gexList.Columns("combinacion").Index), 3))
                '.Cod_color = Left(Rs_Carga("color"), 6)
                .Cod_Talla = gexList.Value(gexList.Columns("talla").Index)
                '.Cod_Calidad = Rs_Carga("calidad")
            End With
        ElseIf oParent.Name = "FrmMovAlmacen" Then
            With oParent
                .TxtGuia = gexList.Value(gexList.Columns("SER_GUIA").Index) & "-" & gexList.Value(gexList.Columns("NUMERO_GUIA").Index)
                '.sCod_AlmacenOrigen = gexList.Value(gexList.Columns("COD_ALMACEN").Index)
                '.sNum_MovStkOrigen = gexList.Value(gexList.Columns("NUM_MOVSTK").Index)
            End With
        ElseIf oParent.Name = "FrmAddMovimAlm" Then
            With oParent
                .TxtGuia = gexList.Value(gexList.Columns("SER_GUIA").Index) & "-" & gexList.Value(gexList.Columns("NUMERO_GUIA").Index)
                '.sCod_AlmacenOrigen = gexList.Value(gexList.Columns("COD_ALMACEN").Index)
                '.sNum_MovStkOrigen = gexList.Value(gexList.Columns("NUM_MOVSTK").Index)
            End With
        ElseIf oParent.Name = "frmDetalleCorte" Then
            With oParent
                .sCOD_COMB = gexList.Value(gexList.Columns("cod_comb").Index)
                .scod_color = gexList.Value(gexList.Columns("cod_color").Index)
                .Label2.Caption = gexList.Value(gexList.Columns("cod_calidad").Index)
                .sCOD_TALLA = gexList.Value(gexList.Columns("COD_MEDIDA").Index)
                .sCod_TipOrdTra = gexList.Value(gexList.Columns("cod_tipordtra").Index)
                .Scod_ordtra = gexList.Value(gexList.Columns("cod_ordtra").Index)
            End With
            
    End If
    End If
End If
Unload Me
End Sub

Private Sub gexList_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn
        gexList_DblClick
    Case vbKeyEscape
        CmdCancelar.SetFocus
    Case Else
        gexList.Find 1, jgexContains, Chr(KeyCode)
End Select
End Sub
