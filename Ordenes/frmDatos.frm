VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDatos 
   Caption         =   "Datos"
   ClientHeight    =   6000
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin FunctionsButtons.FunctButt acbForm 
      Height          =   510
      Left            =   3000
      TabIndex        =   5
      Top             =   5385
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "7~0~ACEPTAR~True~True~&Aceptar~0~0~4~~0~True~False~&Ok~~8~0~CANCELAR~True~True~&Cancelar~0~0~3~~0~False~True~&Cancel~"
      Orientacion     =   0
      Style           =   1
      Language        =   1
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.CheckBox chkOk 
      Height          =   360
      Left            =   2190
      TabIndex        =   4
      Top             =   5490
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   6480
      Visible         =   0   'False
      Width           =   1230
   End
   Begin MSFlexGridLib.MSFlexGrid ssgrdDatos 
      Height          =   5250
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   9260
      _Version        =   393216
      Rows            =   1
      Cols            =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin VB.Label lblValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1215
      TabIndex        =   3
      Top             =   5460
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label lblCaption 
      Caption         =   "Label1"
      Height          =   420
      Left            =   90
      TabIndex        =   2
      Top             =   5475
      Visible         =   0   'False
      Width           =   1080
   End
End
Attribute VB_Name = "frmDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const Desing = True
'
Option Explicit
Public oParent As Object
Private vBuffer As Variant
Private obj As Object
Private bEdit As Boolean
Public OK  As Boolean
Private bSel As Boolean
Public DataFound As Boolean
Public Event BeforeChangeText(ByVal irow As Long, ByVal iCol As Long, ByRef Value As Variant)
Public Event Click()
Public Event DblClick()
Private Sub BuffToGrid(grdControl As MSFlexGridLib.MSFlexGrid)
Dim iNro As Integer
Dim iFils, iCols As Integer
Dim i, j As Integer
Dim tRow, tCol As Integer
Dim Value As Variant
If IsEmpty(vBuffer) = False Then
 DataFound = True
 grdControl.Redraw = False
 tRow = grdControl.Row
 tCol = grdControl.Col
 iFils = UBound(vBuffer, 2)
 iCols = UBound(vBuffer, 1)
 grdControl.Cols = iCols + 1
 For i = 0 To iFils
  grdControl.AddItem ""
  grdControl.Row = grdControl.Rows - 1
  For j = 0 To iCols
   Value = vBuffer(j, i)
   RaiseEvent BeforeChangeText(i, j, Value)
   grdControl.Col = j: grdControl.Text = "" & Value 'vBuffer(j, i)
  Next j
 Next i
 grdControl.Row = tRow
 If bEdit = False Then
  grdControl.Col = 0
  grdControl.ColSel = grdControl.Cols - 1
 Else
  grdControl.Col = tCol
 End If
 grdControl.Redraw = True
Else
 Set obj = Nothing
End If
End Sub

Public Property Set RefObject(ByRef mobj As Object)
 Set obj = mobj
End Property
Public Property Get RefObject() As Object
 Set RefObject = obj
End Property

Public Property Get Buffer() As Variant
 Buffer = vBuffer
End Property

Public Property Let Buffer(vValue As Variant)
 vBuffer = vValue
 BuffToGrid ssgrdDatos
 'If Not (RefObject Is Nothing) Then
 '   Set RefObject = Nothing
 'End If
End Property
Public Property Let Edit(ByVal mEdit As Boolean)
bEdit = mEdit
If bEdit = True Then
 ssgrdDatos.SelectionMode = flexSelectionFree
Else
 ssgrdDatos.SelectionMode = flexSelectionByRow
End If
End Property
Public Property Get Edit() As Boolean
Edit = bEdit
End Property

Public Property Let FormatString(ByVal sFormatString As String)
ssgrdDatos.FormatString = sFormatString
End Property
Public Property Get FormatString() As String
 FormatString = ssgrdDatos.FormatString
End Property
Public Property Let ColumnWidths(vColumnWidths As Variant)
Dim iCols As Integer
Dim i As Integer
iCols = UBound(vColumnWidths) + 1
If iCols > ssgrdDatos.Cols Then
 ssgrdDatos.Cols = iCols
End If
For i = 0 To iCols - 1
 ssgrdDatos.ColWidth(i) = vColumnWidths(i)
Next i
End Property


Private Sub acbForm_ActionClick(ByVal index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
If index = 0 Then
 Me.OK = True
Else
 Me.OK = False
End If
Me.Hide
End Sub


Private Sub chkOk_Click()
    oParent.OK
    acbForm_ActionClick 1, 0, ""
End Sub

Private Sub Form_DblClick()

RaiseEvent DblClick
#If Desing Then
Dim i As Integer
Dim sprint As String
Debug.Print Me.Name & ".Left = " & Me.Left
Debug.Print Me.Name & ".Top = " & Me.Top
Debug.Print Me.Name & ".Width = " & Me.Width
Debug.Print Me.Name & ".Height = " & Me.Height
sprint = Me.Name & ".ColumnWidths = Array("
For i = 0 To Me.ssgrdDatos.Cols - 1
 If i > 0 Then
  sprint = sprint & ","
 End If
 sprint = sprint & Me.ssgrdDatos.ColWidth(i)
Next i
sprint = sprint & ")"
Debug.Print sprint
#End If

End Sub

Private Sub Form_Load()
 Edit = False
End Sub
Public Function Data(ByVal Row As Long, ByVal Col As Long) As Variant
Dim tRow, tCol As Integer
tRow = ssgrdDatos.Row
tCol = ssgrdDatos.Col
If Row > ssgrdDatos.Rows Or Col > ssgrdDatos.Cols Then
 Data = Empty
 Exit Function
Else
 ssgrdDatos.Row = Row
 ssgrdDatos.Col = Col
 Data = ssgrdDatos.Text
 ssgrdDatos.Row = tRow
 ssgrdDatos.Col = tCol
End If
End Function
Public Function TextArray(ByVal index As Long) As Variant
If index < ssgrdDatos.Cols Then
 ssgrdDatos.Col = index
 TextArray = ssgrdDatos.Text
Else
 TextArray = Empty
End If
End Function
Public Function TextMatrix(ByVal Row As Long, ByVal Col As Long)
 TextMatrix = ssgrdDatos.TextMatrix(Row, Col)
End Function

Private Sub Form_Resize()
Dim tHeight As Integer
ssgrdDatos.Left = 0
ssgrdDatos.Top = 0
tHeight = Me.ScaleHeight - Me.acbForm.Height * 2
If tHeight < acbForm.Height * 2 Then
 tHeight = acbForm.Height * 2
 Me.Height = tHeight * 3
Else
 ssgrdDatos.Height = tHeight
 ssgrdDatos.Width = Me.ScaleWidth
 acbForm.Top = ssgrdDatos.Height + ((Me.ScaleHeight - ssgrdDatos.Height) - acbForm.Height) / 2
 acbForm.Left = ((Me.ScaleWidth) - acbForm.Width) / 2
End If
End Sub
Public Sub HeaderClick(ByVal mCol As Long)
Dim tCol As Long
Dim tColSel As Long

If ssgrdDatos.ColSel <> -1 Then
    tCol = ssgrdDatos.Col
    tColSel = ssgrdDatos.ColSel

    ssgrdDatos.Col = mCol
    ssgrdDatos.Sort = flexSortGenericAscending

    ssgrdDatos.Col = tCol
    ssgrdDatos.ColSel = tColSel
End If
End Sub
Private Sub ssgrdDatos_Click()
If ssgrdDatos.MouseRow < ssgrdDatos.FixedRows Then
 HeaderClick ssgrdDatos.MouseCol
End If
End Sub

Private Sub ssgrdDatos_DblClick()
RaiseEvent DblClick
If bEdit = True Then
 GridEdit Asc(" ")
Else
 If ssgrdDatos.Row > 0 Then
  acbForm_ActionClick 0, 0, "none"
 End If
End If
End Sub

Private Sub ssgrdDatos_KeyPress(KeyAscii As Integer)
If bEdit = False Then
 Select Case KeyAscii
  Case vbKeyReturn
     acbForm_ActionClick 0, 0, "KeyPress"
  Case vbKeyEscape
     acbForm_ActionClick 1, 0, "KeyPress"
  Case Else
     Call BUSCACAMPO_FLEX(0, Chr(KeyAscii), ssgrdDatos)
 End Select
Else
 GridEdit KeyAscii
End If
End Sub

Sub GridEdit(KeyAscii As Integer)
If bEdit = True Then
 txtEdit.FontName = ssgrdDatos.FontName
 txtEdit.FontSize = ssgrdDatos.FontSize
 Select Case KeyAscii
 Case 0 To Asc(" ")
  txtEdit = ssgrdDatos.Text
  bSel = True
 Case Else
  txtEdit = Chr(KeyAscii)
  txtEdit.SelStart = 1
  bSel = False
 End Select
 'position the edit box
 txtEdit.Left = ssgrdDatos.CellLeft + ssgrdDatos.Left
 txtEdit.Top = ssgrdDatos.CellTop + ssgrdDatos.Top
 txtEdit.Width = ssgrdDatos.CellWidth
 txtEdit.Height = ssgrdDatos.CellHeight
 txtEdit.Visible = True
 txtEdit.SetFocus
End If
End Sub

Private Sub ssgrdDatos_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 34 Then
  If Not (obj Is Nothing) Then
   vBuffer = obj.RowsDataSet()
   Call BuffToGrid(ssgrdDatos)
  End If
 End If
End Sub

Private Sub txtEdit_GotFocus()
If bSel = True Then
 txtEdit.SelStart = 0
 txtEdit.SelLength = Len(txtEdit.Text)
End If
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
 txtEdit.Visible = False
 ssgrdDatos.SetFocus
Case vbKeyReturn
 ssgrdDatos.Text = txtEdit.Text
 txtEdit.Visible = False
 ssgrdDatos.SetFocus
Case vbKeyDown
 txtEdit.Visible = False
 ssgrdDatos.Text = txtEdit.Text
 ssgrdDatos.SetFocus
 DoEvents
 If ssgrdDatos.Row < ssgrdDatos.Rows - 1 Then
  ssgrdDatos.Row = ssgrdDatos.Row + 1
 End If
Case vbKeyUp
 txtEdit.Visible = False
 ssgrdDatos.Text = txtEdit.Text
 ssgrdDatos.SetFocus
 DoEvents
 If ssgrdDatos.Row > ssgrdDatos.FixedRows Then
  ssgrdDatos.Row = ssgrdDatos.Row - 1
 End If
End Select
End Sub
Private Sub txtEdit_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub txtEdit_LostFocus()
 txtEdit.Visible = False
End Sub

'AHSP
Public Sub BUSCACAMPO_FLEX(varNumColumna As Integer, varValor As String, varFlexGrid As MSFlexGrid)
    Dim ProvvarFlexGrid As MSFlexGrid
    Dim varIndice As Integer
    Dim varNumFilas As Integer
    Dim varEncontro As Boolean
    varEncontro = False
    Set ProvvarFlexGrid = varFlexGrid
    ProvvarFlexGrid.Col = varNumColumna
    For varIndice = 1 To ProvvarFlexGrid.Rows - 1
        ProvvarFlexGrid.Row = varIndice
        If Mid(ProvvarFlexGrid.Text, 1, Len(varValor)) = UCase(varValor) Then
            varEncontro = True
            Exit For
        End If
    Next
    If varEncontro = False Then
        varFlexGrid.TopRow = 1
        varFlexGrid.Row = 1
    Else
        varFlexGrid.TopRow = varIndice
        varFlexGrid.Row = varIndice
    End If
    varFlexGrid.ZOrder (1)
    'varFlexGrid
    Set ProvvarFlexGrid = Nothing
End Sub


