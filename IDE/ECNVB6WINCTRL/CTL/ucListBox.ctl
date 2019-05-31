VERSION 5.00
Begin VB.UserControl ucListBox 
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4380
   ScaleHeight     =   2670
   ScaleWidth      =   4380
   ToolboxBitmap   =   "ucListBox.ctx":0000
   Begin VB.VScrollBar Bar 
      Height          =   2415
      Left            =   4080
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox PicBox 
      BackColor       =   &H00004000&
      Height          =   2415
      Left            =   0
      MousePointer    =   1  'Arrow
      ScaleHeight     =   2355
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.TextBox TxtList 
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "List"
         Top             =   270
         Width           =   975
      End
      Begin VB.CommandButton BtnTitle 
         Caption         =   "ucListBox"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   975
      End
   End
End
Attribute VB_Name = "ucListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim BoxTit(30) As String, BoxLenCol(30) As Integer
Dim BoxAlign(30) As Integer
Dim BoxQCol As Integer, WidthBox As Integer
Dim BoxQLin As Integer, HeightBox As Integer
Dim BoxBack As Long, BoxCaixa As Long, BoxFore As Long
Dim BoxNameFont As String, BoxSizeFont As Integer
Dim BoxStr() As String, BoxStrBack() As Long, BoxStrFore() As Long
Dim BoxStrNameFont() As String, BoxStrSizeFont() As Integer
Dim BoxStrBold() As Boolean, BoxLin As Integer, BoxCol As Integer
Dim BoxObj As Object, BoxBackSel As Long, BoxForeSel As Long
Dim BoxSel As Integer, BoxSelAnt As Integer
Dim BoxArg(30) As String, BoxIndex As Integer

Public Function Config(Obj As Object, Optional FontName As String = "Arial", _
Optional FontSize As Integer = 10, Optional QLin As Integer = 10, _
Optional BackColor As Long = &H4000&, Optional BoxColor As Long = &H400000, _
Optional TextColor As Long = &HFFFFFF, Optional BackSel As Long = &H800000, _
Optional ForeSel As Long = &HFFFFFF)
Dim Cont As Integer

BoxQCol = 0: BoxQLin = QLin
BoxLin = 0: BoxCol = 0
WidthBox = 0: HeightBox = 0
Cont = 0
Do While Cont < 20
  BoxTit(Cont) = ""
  BoxLenCol(Cont) = 0
  Cont = Cont + 1
Loop

BoxBack = BackColor
BoxCaixa = BoxColor
BoxFore = TextColor
BoxNameFont = FontName
BoxSizeFont = FontSize
Set BoxObj = Obj
BoxBackSel = BackSel
BoxForeSel = ForeSel
End Function

'S: Title of column
'T: Size of column
'Align: 0-Left 1-Right 2-Center
Public Function Title(S As String, T As Integer, Optional Align As Integer = 0)
BoxTit(BoxQCol) = S
BoxLenCol(BoxQCol) = T * 120
BoxAlign(BoxQCol) = Align
BoxQCol = BoxQCol + 1
End Function

'OpcTitle: True-with title False-without title
Public Function Activate(OpcTitle As Boolean)
Dim Cont As Integer, XLin As Integer, Clin As Integer

If OpcTitle = False Then
  TxtList(0).left = 0
  TxtList(0).top = 0
  BtnTitle(0).Visible = False
End If

PicBox.BackColor = BoxBack
TxtList(0).BackColor = BoxCaixa
TxtList(0).ForeColor = BoxFore
TxtList(0).left = BtnTitle(Cont).left
TxtList(0).text = ""
TxtList(0).Width = BoxLenCol(Cont) - 30
TxtList(0).Alignment = BoxAlign(Cont)
TxtList(0).FontName = BoxNameFont
TxtList(0).FontSize = BoxSizeFont
PicBox.FontName = BoxNameFont
PicBox.FontSize = BoxSizeFont
TxtList(0).Height = PicBox.TextHeight("X")
Cont = 0: Clin = 1
Do While Cont < BoxQCol
  BtnTitle(Cont).Caption = BoxTit(Cont)
  BtnTitle(Cont).Width = BoxLenCol(Cont)
  WidthBox = WidthBox + BtnTitle(Cont).Width
    
  'loads lines referring to the column
  XLin = 1: HeightBox = TxtList(0).Height + 30
  Do While XLin < BoxQLin
    Load TxtList(Clin)
    TxtList(Clin).Visible = True
    TxtList(Clin).top = TxtList(Clin - 1).top + TxtList(Clin - 1).Height + 30
    TxtList(Clin).left = BtnTitle(Cont).left
    TxtList(Clin).text = ""
    TxtList(Clin).Width = BoxLenCol(Cont) - 30
    TxtList(Clin).Alignment = BoxAlign(Cont)
    HeightBox = HeightBox + TxtList(Clin).Height + 30
    XLin = XLin + 1: Clin = Clin + 1
  Loop
  
  'load column
  Cont = Cont + 1
  If Cont < BoxQCol Then
    Load BtnTitle(Cont)
    If OpcTitle = False Then
      BtnTitle(Cont).Visible = False
    Else
      BtnTitle(Cont).Visible = True
    End If
    BtnTitle(Cont).left = BtnTitle(Cont - 1).left + BtnTitle(Cont - 1).Width
    
    'load first box (line) of column
    Load TxtList(Clin)
    TxtList(Clin).Visible = True
    TxtList(Clin).top = TxtList(0).top
    TxtList(Clin).left = BtnTitle(Cont).left
    TxtList(Clin).text = ""
    TxtList(Clin).Width = BoxLenCol(Cont) - 30
    TxtList(Clin).Alignment = BoxAlign(Cont)
    Clin = Clin + 1
  End If
Loop
Bar.Min = 0
Bar.Max = 0
Bar.LargeChange = BoxQLin
UserControl.Width = WidthBox + 300
PicBox.Width = WidthBox + 300
Bar.left = WidthBox + 30
If OpcTitle = True Then
  UserControl.Height = HeightBox + 300
  PicBox.Height = HeightBox + 300
  Bar.Height = HeightBox + 300
Else
  UserControl.Height = HeightBox + 30
  PicBox.Height = HeightBox + 30
  Bar.Height = HeightBox + 30
End If
ReDim BoxStr(BoxQCol, 1)
ReDim BoxStrBack(BoxQCol, 1)
ReDim BoxStrFore(BoxQCol, 1)
ReDim BoxStrNameFont(BoxQCol, 1)
ReDim BoxStrSizeFont(BoxQCol, 1)
ReDim BoxStrBold(BoxQCol, 1)
End Function

Private Sub Bar_Change()
ListBoxStr
End Sub

Private Sub TxtList_Click(Index As Integer)
ClickEvent
End Sub

Private Sub ClickEvent()
On Error Resume Next
BoxObj.Box_Click
End Sub

Private Sub TxtList_DblClick(Index As Integer)
DblClickEvent
End Sub

Private Sub DblClickEvent()
On Error Resume Next
BoxObj.Box_DblClick
End Sub

Private Sub TxtList_GotFocus(Index As Integer)
BoxSel = Index Mod BoxQLin 'selected line
SelBox BoxSel
BoxSelAnt = BoxSel         'previous selected line
End Sub

Private Sub TxtList_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim Cont As Integer, XCol As Integer

EventoKeyDown KeyCode, Shift

If KeyCode = vbKeyDelete Then
  Cont = BoxIndex
  Do While Cont < (BoxLin - 1)
    XCol = 0
    Do While XCol < BoxQCol
      BoxStr(XCol, Cont) = BoxStr(XCol, Cont + 1)
      BoxStrBack(XCol, Cont) = BoxStrBack(XCol, Cont + 1)
      BoxStrFore(XCol, Cont) = BoxStrFore(XCol, Cont + 1)
      BoxStrNameFont(XCol, Cont) = BoxStrNameFont(XCol, Cont + 1)
      BoxStrSizeFont(XCol, Cont) = BoxStrSizeFont(XCol, Cont + 1)
      BoxStrBold(XCol, Cont) = BoxStrBold(XCol, Cont + 1)
      XCol = XCol + 1
    Loop
    Cont = Cont + 1
  Loop
  XCol = 0
  Do While XCol < BoxQCol
    BoxStr(XCol, Cont) = ""
    BoxStrBack(XCol, Cont) = BoxCaixa
    BoxStrFore(XCol, Cont) = BoxFore
    XCol = XCol + 1
  Loop
  BoxLin = BoxLin - 1
  Bar.Max = Bar.Max - 1
  ListBoxStr
  SelBox BoxSel
  KeyCode = 0
End If
If KeyCode = vbKeyDown Then
  If BoxSel < BoxQLin - 1 Then
    TxtList(Index + 1).SetFocus
  Else
    If Bar.Value < BoxLin Then Bar.Value = Bar.Value + 1
  End If
  KeyCode = 0
End If
If KeyCode = vbKeyUp Then
  If BoxSel > 0 Then
    TxtList(Index - 1).SetFocus
  Else
    If Bar.Value > 0 Then Bar.Value = Bar.Value - 1
  End If
  KeyCode = 0
End If
If KeyCode = vbKeyPageDown Then
  If Bar.Value + BoxQLin < BoxLin Then Bar.Value = Bar.Value + BoxQLin
  KeyCode = 0
End If
If KeyCode = vbKeyPageUp Then
  If Bar.Value - BoxQLin >= 0 Then Bar.Value = Bar.Value - BoxQLin
  KeyCode = 0
End If
End Sub

Private Sub EventoKeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo FIM
BoxObj.Box_KeyDown KeyCode, Shift
FIM:
End Sub

Private Sub TxtList_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
  If BoxSel = BoxQLin - 1 Then SelBox BoxQLin - 1
End If
If KeyCode = vbKeyUp Then
  If BoxSel = 0 Then SelBox 0
End If
If KeyCode = vbKeyPageDown Or KeyCode = vbKeyPageUp Then
  SelBox BoxSel
End If
End Sub

Private Sub UserControl_Terminate()
Dim Cont As Integer, Clin As Integer

Cont = 1
Do While Cont < BoxQCol
  Unload BtnTitle(Cont)
  Cont = Cont + 1
Loop
Clin = 1
Do While Clin < BoxQLin * BoxQCol
  Unload TxtList(Clin)
  Clin = Clin + 1
Loop
ReDim BoxStr(BoxQCol, 1)
ReDim BoxStrBack(BoxQCol, 1)
ReDim BoxStrFore(BoxQCol, 1)
ReDim BoxStrNameFont(BoxQCol, 1)
ReDim BoxStrSizeFont(BoxQCol, 1)
ReDim BoxStrBold(BoxQCol, 1)
End Sub

Function Celula(Col As Integer, Lin As Integer, S As String, Back As Long, _
Fore As Long, FontName As String, FontSize As Integer, Negrito As Boolean)
Dim X As Integer

On Error GoTo FIM
X = (Col * BoxQLin) + Lin
TxtList(X).text = S
TxtList(X).BackColor = Back
TxtList(X).ForeColor = Fore
TxtList(X).FontName = FontName
TxtList(X).FontSize = FontSize
TxtList(X).FontBold = Negrito
Exit Function
FIM:
TxtList(X).BackColor = BoxCaixa
End Function

Function SelBox(Lin As Integer)
Dim X As Integer, Col As Integer, XLin As Long

On Error Resume Next

XLin = Bar.Value + BoxSelAnt
Col = 0
Do While Col < BoxQCol
  'Deselect previous line
  Celula Col, BoxSelAnt, BoxStr(Col, XLin), BoxStrBack(Col, XLin), BoxStrFore(Col, XLin), _
  BoxStrNameFont(Col, XLin), BoxStrSizeFont(Col, XLin), BoxStrBold(Col, XLin)
  
  X = (Col * BoxQLin) + Lin
  TxtList(X).BackColor = BoxBackSel
  TxtList(X).ForeColor = BoxForeSel
  'load BoxArg with the selected cells
  BoxArg(Col) = TxtList(X).text
  Col = Col + 1
Loop
BoxArg(Col) = ""
BoxIndex = Bar.Value + Lin
End Function

Function Add(S As String, Optional Back As Long = -1, _
Optional Fore As Long = -1, Optional FontName As String = "", _
Optional FontSize As Integer = 0, Optional Negrito As Boolean = False)
BoxStr(BoxCol, BoxLin) = S
If Back = -1 Then
  BoxStrBack(BoxCol, BoxLin) = BoxCaixa
Else
  BoxStrBack(BoxCol, BoxLin) = Back
End If
If Fore = -1 Then
  BoxStrFore(BoxCol, BoxLin) = BoxFore
Else
  BoxStrFore(BoxCol, BoxLin) = Fore
End If
If FontName = "" Then
  BoxStrNameFont(BoxCol, BoxLin) = BoxNameFont
Else
  BoxStrNameFont(BoxCol, BoxLin) = FontName
End If
If FontSize = 0 Then
  BoxStrSizeFont(BoxCol, BoxLin) = BoxSizeFont
Else
  BoxStrSizeFont(BoxCol, BoxLin) = FontSize
End If
BoxStrBold(BoxCol, BoxLin) = Negrito
BoxCol = BoxCol + 1
End Function

Function BoxNew()
BoxLin = BoxLin + 1
BoxCol = 0
ReDim Preserve BoxStr(BoxQCol, BoxLin + BoxQLin)
ReDim Preserve BoxStrBack(BoxQCol, BoxLin + BoxQLin)
ReDim Preserve BoxStrFore(BoxQCol, BoxLin + BoxQLin)
ReDim Preserve BoxStrNameFont(BoxQCol, BoxLin + BoxQLin)
ReDim Preserve BoxStrSizeFont(BoxQCol, BoxLin + BoxQLin)
ReDim Preserve BoxStrBold(BoxQCol, BoxLin + BoxQLin)
Bar.Max = Bar.Max + 1
ListBoxStr
End Function

Function ListBoxStr()
Dim XLin As Long, XCol As Integer, T As Long, X As Integer

If BoxLin = 0 Then Exit Function
ClearBox
XLin = Bar.Value: X = 0
T = Bar.Value + BoxQLin
Do While XLin < T
  XCol = 0
  Do While XCol < BoxQCol
    Celula XCol, X, BoxStr(XCol, XLin), BoxStrBack(XCol, XLin), BoxStrFore(XCol, XLin), _
    BoxStrNameFont(XCol, XLin), BoxStrSizeFont(XCol, XLin), BoxStrBold(XCol, XLin)
    XCol = XCol + 1
  Loop
  XLin = XLin + 1
  X = X + 1
Loop
End Function

Public Function Clear()
Dim Col As Integer, XLin As Long

'clear select bar if exist
On Error Resume Next
XLin = Bar.Value + BoxSelAnt
Col = 0
Do While Col < BoxQCol
  'deselect previous line
  Celula Col, BoxSelAnt, BoxStr(Col, XLin), BoxStrBack(Col, XLin), BoxStrFore(Col, XLin), _
  BoxStrNameFont(Col, XLin), BoxStrSizeFont(Col, XLin), BoxStrBold(Col, XLin)
  Col = Col + 1
Loop

ClearBox
ReDim BoxStr(BoxQCol, 1)
ReDim BoxStrBack(BoxQCol, 1)
ReDim BoxStrFore(BoxQCol, 1)
ReDim BoxStrNameFont(BoxQCol, 1)
ReDim BoxStrSizeFont(BoxQCol, 1)
ReDim BoxStrBold(BoxQCol, 1)
BoxLin = 0: BoxCol = 0
Bar.Min = 0
Bar.Max = 0
Bar.Value = 0
End Function

Private Function ClearBox()
Dim Cont As Integer, T As Integer

Cont = 0: T = BoxQLin * BoxQCol
Do While Cont < T
  TxtList(Cont).text = ""
  Cont = Cont + 1
Loop
End Function

Function Selected(X As Integer)
On Error GoTo FIM:
If X >= BoxQLin Then
  Bar.Value = (X - BoxQLin) + 1
  SelBox BoxQLin - 1
  BoxSelAnt = BoxQLin - 1
Else
  SelBox X
  BoxSelAnt = X
End If
Exit Function
FIM:
MsgBox "Invalid index!", vbCritical
End Function

Public Function Arg(Indice As Integer) As String
Arg = BoxArg(Indice)
End Function

Public Function ListCount() As Integer
ListCount = BoxLin
End Function

Public Function ListIndex() As Integer
ListIndex = BoxIndex
End Function
