VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FUNCBUTT.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form frmShowTX_Ordtra_ItemsRollos 
   Caption         =   "Rollos Ingresados"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin GridEX20.GridEX GridEX1 
      Height          =   6210
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   10954
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmShowTX_Ordtra_ItemsRollos.frx":0000
      RowHeaders      =   -1  'True
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmShowTX_Ordtra_ItemsRollos.frx":0352
      Column(2)       =   "frmShowTX_Ordtra_ItemsRollos.frx":041A
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmShowTX_Ordtra_ItemsRollos.frx":04BE
      FormatStyle(2)  =   "frmShowTX_Ordtra_ItemsRollos.frx":05F6
      FormatStyle(3)  =   "frmShowTX_Ordtra_ItemsRollos.frx":06A6
      FormatStyle(4)  =   "frmShowTX_Ordtra_ItemsRollos.frx":075A
      FormatStyle(5)  =   "frmShowTX_Ordtra_ItemsRollos.frx":0832
      FormatStyle(6)  =   "frmShowTX_Ordtra_ItemsRollos.frx":08EA
      FormatStyle(7)  =   "frmShowTX_Ordtra_ItemsRollos.frx":09CA
      FormatStyle(8)  =   "frmShowTX_Ordtra_ItemsRollos.frx":0E82
      ImageCount      =   1
      ImagePicture(1) =   "frmShowTX_Ordtra_ItemsRollos.frx":12CE
      PrinterProperties=   "frmShowTX_Ordtra_ItemsRollos.frx":1620
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   1110
      Left            =   6255
      TabIndex        =   1
      Top             =   15
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1958
      Custom          =   $"frmShowTX_Ordtra_ItemsRollos.frx":17F8
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmShowTX_Ordtra_ItemsRollos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sTituliAbrOP As String
Public Codigo As String
Public Descripcion As String
Public sCod_TipOrdTra As String
Public sCod_OrdTra As String
Public sNum_Secuencia As String

Public oParent As Object


Private Sub Form_Load()
    Dim sSeguridad  As String
    sSeguridad = get_botones1(Me, vper, vemp, Me.Name)
    
    'Me.FunctButt1.FunctionsUser = sSeguridad
        
    iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))
End Sub

Public Function Buscar() As Boolean
On Error GoTo errores
Dim sSQl As String
Dim vBookmark As Variant

sSQl = "SM_MUESTRA_Tx_OrdTra_Items_Rollos '" & sCod_TipOrdTra & "','" & sCod_OrdTra & "','" & sNum_Secuencia & "'"

vBookmark = GridEX1.Row
GridEX1.ClearFields

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQl, cConnect)
GridEX1.Row = vBookmark

GridEX1.ContinuousScroll = True
GridEX1.FrozenColumns = 1
GridEX1.AllowEdit = True

Exit Function

errores:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ELIMINAR"
            If GridEX1.RowCount > 0 Then
                EliminarRollo
            End If
        Case "SALIR"
            Unload Me
    End Select
End Sub

Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = True
End Sub


Private Sub GridEX1_DblClick()
    Dim i As Integer
    For i = 1 To GridEX1.Columns.Count
        Debug.Print GridEX1.Name & ".Columns(" & Chr(34) & GridEX1.Columns(i).Caption & Chr(34) & ").width = " & CStr(GridEX1.Columns(i).Width)
    Next
End Sub

Public Function EliminarRollo() As Boolean
On Error GoTo errores
Dim sSQl As String
Dim oMensaje As clsMessages

Set oMensaje = New clsMessages

oMensaje.Codigo = CodeMsg.kMESSAGE_ASK_PROCESS
oMensaje.OptionalText = "Desea Eliminar Rollo Nro" & GridEX1.Value(GridEX1.Columns("NUM_ROLLO").Index)

If Not oMensaje.ShowMesage(iLanguage) Then
    Exit Function
End If

sSQl = "UP_ELIMINA_ROLLO '" & _
    sCod_TipOrdTra & "'," & _
    sCod_OrdTra & "," & _
    sNum_Secuencia & "," & _
    GridEX1.Value(GridEX1.Columns("NUM_ROLLO").Index) & ",'" & vusu & "'"

ExecuteSQL cConnect, sSQl
Buscar
EliminarRollo = True

Exit Function
erorres:
    Err.Raise Err.Number, Err.Source, Err.Description
Exit Function

errores:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


