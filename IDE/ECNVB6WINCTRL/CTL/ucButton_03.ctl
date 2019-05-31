VERSION 5.00
Begin VB.UserControl ucButton_03 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2370
   ScaleHeight     =   1320
   ScaleWidth      =   2370
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   1095
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      MousePointer    =   99  'Custom
      ScaleHeight     =   495
      ScaleWidth      =   1095
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "ucButton_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'File Name: AxAOLCmd.ctl
'Description:  This is a very customizable AOL button replication.

'API Calls
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'Constant for drawing the image
Private Const SRCCOPY = &HCC0020

'The events users of this control will have.
'MouseEnter and MouseLeave new to Version2.
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseEnter(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseLeave()
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'These allow us ot have ComboBoxes in the properties
Enum TheStyle
    StyleOne
    StyleTwo
End Enum

Enum TheAutoSize
    NoAutoSize
    ButtonToPic
    PicToButton
End Enum

'This is for storing the font
Private WithEvents TheFont As StdFont
Attribute TheFont.VB_VarHelpID = -1

'These are our storing variables, most store property values
Dim TheCaption As String, TheStyleX As TheStyle, HasFocus As Boolean
Dim MouseDowned As Boolean, TheEnabled As Boolean, TheGraphical As Boolean
Dim TheAutoSizeX As TheAutoSize, TheForeColor As OLE_COLOR, TheBackColor As OLE_COLOR
Dim TheStandardColors As Boolean, TheBorderLight As OLE_COLOR, TheBorderDark As OLE_COLOR
Dim TheBackColorClick As OLE_COLOR, TheX As Single, TheY As Single
Dim TheButton As Integer, TheShift As Integer, OldWndProc As Long

Private Sub DrawButton(Offset As Long, Clicked As Boolean)
    'This is where the drawing of the button occurs.
    'Everything is in here.  offset is so we can draw
    'the button down, and Clicked is so we know when
    'we are drawing the clicked state.
    Dim ClickColor As Long, SetWidth As Long, SetHeight As Long
    Dim SetLeft As Long, SetTop As Long, i As Long
    Dim x As Long, y As Long, Checker As Boolean
    'Clear our button so there are no remnants when we redraw it
    Picture1.Cls
    'Fill in the small space on the bottom left and bottom right
    'where we don't draw with the parent's color
    Picture1.BackColor = Ambient.BackColor
    'Draw the outline of the main part of the button
    Picture1.Line (Offset + 0, Offset + 0)-(Offset + 0, Offset + Picture1.Height - 60), BorderLight
    Picture1.Line (Offset + 0, Offset + 0)-(Offset + Picture1.Width - 60, Offset + 0), BorderLight
    Picture1.Line (Offset + 0, Offset + Picture1.Height - 60)-(Offset + Picture1.Width - 60, Offset + Picture1.Height - 60), BorderDark
    Picture1.Line (Offset + Picture1.Width - 60, Offset + 0)-(Offset + Picture1.Width - 60, Offset + Picture1.Height - 45), BorderDark
    'Draw the main area of the button
    If Clicked = True Then
        Picture1.Line (Offset + 15, Offset + 15)-(Offset + Picture1.Width - 75, Offset + Picture1.Height - 75), TheBackColorClick, BF
    Else
        Picture1.Line (Offset + 15, Offset + 15)-(Offset + Picture1.Width - 75, Offset + Picture1.Height - 75), TheBackColor, BF
    End If
    'Draw the 3 tiered shadow
    Picture1.Line (Offset + 75, Offset + Picture1.Height - 45)-(Offset + Picture1.Width - 30, Offset + Picture1.Height - 45), 8421504
    Picture1.Line (Offset + Picture1.Width - 45, Offset + 75)-(Offset + Picture1.Width - 45, Offset + Picture1.Height - 45), 8421504
    Picture1.Line (Offset + 75, Offset + Picture1.Height - 30)-(Offset + Picture1.Width - 30, Offset + Picture1.Height - 30), 10526880
    Picture1.Line (Offset + Picture1.Width - 30, Offset + 75)-(Offset + Picture1.Width - 30, Offset + Picture1.Height - 15), 10526880
    Picture1.Line (Offset + 75, Offset + Picture1.Height - 15)-(Offset + Picture1.Width - 15, Offset + Picture1.Height - 15), 12632256
    Picture1.Line (Offset + Picture1.Width - 15, Offset + 75)-(Offset + Picture1.Width - 15, Offset + Picture1.Height), 12632256
    'We need to draw the picture onto the button
    If TheGraphical = True Then
        'Don't draw unless there is a picture
        If Not Picture2.Picture = 0 Then
            'If the user wants the image stretched, we need to account for that
            If AutoSize = PicToButton Then
                'We get the width and the height of where it needs to be drawn
                SetWidth = (UserControl.Width - 75) / 15
                SetHeight = (UserControl.Height - 75) / 15
                'Draw it using the SreetchBlt call so we can resize it
                StretchBlt Picture1.hdc, 1 + Offset / 15, 1 + Offset / 15, SetWidth, SetHeight, Picture2.hdc, 0, 0, Picture2.Width / 15, Picture2.Height / 15, SRCCOPY
            Else
                'Get all the coordinates now so drawing is easier
                SetWidth = Picture2.Width / 15
                SetHeight = Picture2.Height / 15
                SetLeft = (UserControl.Width / 2 - Picture2.Width / 2 - 22) / 15 + Offset / 15
                SetTop = (UserControl.Height / 2 - Picture2.Height / 2 - 22) / 15 + Offset / 15
                'Check to see if the picture is bigger than our drawing surface.
                'If it is, we will need to clip parts off it
                If Picture2.Width > UserControl.Width - 75 Then
                    SetWidth = (UserControl.Width - 75) / 15
                    SetLeft = 1
                End If
                If Picture2.Height > UserControl.Height - 75 Then
                    SetHeight = (UserControl.Height - 75) / 15
                    SetTop = 1
                End If
                'Draw the picture
                BitBlt Picture1.hdc, SetLeft, SetTop, SetWidth, SetHeight, Picture2.hdc, 0, 0, SRCCOPY
            End If
        End If
    End If
    'This is where we draw the text.  First we find the exact place to draw it from
    Picture1.CurrentX = (Picture1.Width - Picture1.TextWidth(TheCaption)) / 2 - 22 + Offset
    Picture1.CurrentY = (Picture1.Height - Picture1.TextHeight(TheCaption)) / 2 - 22 + Offset
    'Set the color right
    Picture1.ForeColor = TheForeColor
    'And then wed draw it
    Picture1.Print TheCaption
    'This is where we draw the focus rectnagle
    If HasFocus = True Then
        'Check to see if we draw it clicked
        If Clicked = True Then
            'If the backcolor changes on click...
            If Not BackColor = BackColorClick Then
                'We change the DrawMode to 'Inverse'.  On closer
                'Inspection of AOL's button, the Focus Rectangle
                'changes color when clicked on IMs.  I found out
                'through trial and error that DrawMode 6 matches
                'exactly the color.
                Picture1.DrawMode = 6
            End If
        End If
        'Here is where we draw it.  Use step 30 to skip every
        'other pixel and then use PSet to draw the pixel.
        For i = 60 To Picture1.Width - 105 Step 30
            Picture1.PSet (i + Offset, 45 + Offset), 5608190
        Next
        For i = 60 To Picture1.Width - 120 Step 30
            Picture1.PSet (i + Offset, Picture1.Height - 105 + Offset), 5608190
        Next
        For i = 60 To Picture1.Height - 105 Step 30
            Picture1.PSet (45 + Offset, i + Offset), 5608190
        Next
        For i = 60 To Picture1.Height - 105 Step 30
            Picture1.PSet (Picture1.Width - 105 + Offset, i + Offset), 5608190
        Next
        'Set the DrawMode back to what it normally is
        Picture1.DrawMode = 13
    End If
    'If it is Disabled we need to draw a mask over it.
    'This mas is just a checkerboard of white pixels.
    'We use nested For...Next loops to accomplish this.
    If TheEnabled = False Then
        For x = 0 To Picture1.Width - 75 Step 15
            'Every other time we have it start one
            'pixel lower so achieve the checkboard effect.
            'We use 'Checker' to hold if it's time
            'to do this.
            If Checker = False Then
                For y = 0 To Picture1.Height - 75 Step 30
                    Picture1.PSet (x, y), &HFFFFFF
                Next
                Checker = True
            Else
                For y = 15 To Picture1.Height - 75 Step 30
                    Picture1.PSet (x, y), &HFFFFFF
                Next
                Checker = False
            End If
        Next
    End If
End Sub

Private Sub Picture1_Click()
    'Pass the click event only if it has the focus
    If HasFocus = True Then
        RaiseEvent Click
    End If
End Sub

Private Sub Picture1_DblClick()
    'Pass double click event
    RaiseEvent DblClick
End Sub

Private Sub Picture1_GotFocus()
    'Set our focus holding variable to True
    HasFocus = True
    'Redraw the button accordingly
    If MouseDowned = True Then
        If Style = StyleOne Then
            DrawButton 45, True
        Else
            DrawButton 0, True
        End If
    Else
        DrawButton 0, False
    End If
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    'Pass KewDown event
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
    'Pass KeyPress event
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
    'Pass KeyUp event
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Picture1_LostFocus()
    'Reset everything to non-focused state
    HasFocus = False
    MouseDowned = False
    DrawButton 0, False
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'If the mouse is clicked
    If Button = 1 Then
        'Captue the picture so we can
        'monitor when the mosue leaves
        SetCapture Picture1.Hwnd
        MouseDowned = True
        'If it has the focus (which is should)
        'draw the button in it's new state
        If HasFocus = True Then
            If Style = StyleOne Then
                DrawButton 45, True
            Else
                DrawButton 0, True
            End If
        End If
    End If
    'Pass MouseDown event
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'If our button is not already captured
    If GetCapture <> Picture1.Hwnd Then
'        'Raise MouseEnter because the mouse just entered
'        RaiseEvent MouseEnter(Button, Shift, x, y)
        'If the mouse is down
        If Button = 1 Then
            'If it has the focus
            If HasFocus = True Then
                'Capture the button so we know when it leaves
                SetCapture Picture1.Hwnd
                'Draw button correct way
                If Style = StyleOne Then
                    DrawButton 45, True
                Else
                    DrawButton 0, True
                End If
            End If
        End If
    End If
    'If the mouse is outside the button
    If x < 0 Or x > Picture1.Width Or y < 0 Or y > Picture1.Height Then
        'If our button is captured
        If GetCapture = Picture1.Hwnd Then
            'Release the capture
            ReleaseCapture
            'THe Mouse Left, raise the event
'            RaiseEvent MouseLeave(Button, Shift, x, y)
            'Redraw the button
            DrawButton 0, False
        End If
    End If
    'Pass MouseMove event
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'If if it was the left mouse button
    If Button = 1 Then
        'We Recapture the button for safety
        SetCapture Picture1.Hwnd
        'Reset our stuff
        MouseDowned = False
        DrawButton 0, False
    End If
    'Pass MouseUp event
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    'If the BackColor changes of the form
    If PropertyName = "BackColor" Then
        'Redraw the button to fill in new color
        DrawButton 0, False
    End If
End Sub

Private Sub UserControl_InitProperties()
    'Set some of the properties to their initial values
    Set TheFont = New StdFont
    TheFont.Name = "Tahoma"
    Set Font = TheFont
    Caption = Extender.Name
    Enabled = True
    ForeColor = &HFFFFFF
    BackColor = &HAA6D00
    StandardColors = True
    BorderLight = &HFFB691
    BorderDark = &HFFB691
    BackColorClick = &HAA6D00
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'Load our properties when we need to
    Caption = PropBag.ReadProperty("Caption", Extender.Name)
    Style = PropBag.ReadProperty("Style", StyleOne)
    Set Font = PropBag.ReadProperty("Font")
    Enabled = PropBag.ReadProperty("Enabled", True)
    Set Picture = PropBag.ReadProperty("Picture")
    Graphical = PropBag.ReadProperty("Graphical", False)
    AutoSize = PropBag.ReadProperty("AutoSize", NoAutoSize)
    ForeColor = PropBag.ReadProperty("ForeColor", &HFFFFFF)
    BackColor = PropBag.ReadProperty("BackColor", 11169024)
    StandardColors = PropBag.ReadProperty("StandardColors", True)
    BorderLight = PropBag.ReadProperty("BorderLight", &HFFB691)
    BorderDark = PropBag.ReadProperty("BorderDark", &HFFB691)
    BackColorClick = PropBag.ReadProperty("BackColorClick", &HAA6D00)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'Save our properties when we need to
    PropBag.WriteProperty "Caption", Caption, Extender.Name
    PropBag.WriteProperty "Style", Style, StyleOne
    PropBag.WriteProperty "Font", Font
    PropBag.WriteProperty "Enabled", Enabled, True
    PropBag.WriteProperty "Picture", Picture
    PropBag.WriteProperty "Graphical", Graphical, False
    PropBag.WriteProperty "AutoSize", AutoSize, NoAutoSize
    PropBag.WriteProperty "ForeColor", ForeColor, &HFFFFFF
    PropBag.WriteProperty "BackColor", BackColor, 11169024
    PropBag.WriteProperty "StandardColors", StandardColors, True
    PropBag.WriteProperty "BorderLight", BorderLight, &HFFB691
    PropBag.WriteProperty "BorderDark", BorderDark, &HFFB691
    PropBag.WriteProperty "BackColorClick", BackColorClick, &HAA6D00
End Sub

Private Sub UserControl_Resize()
    'Make sure it is at least a certain size
    If UserControl.Width < 300 Or UserControl.Height < 300 Then
        SIZE 300, 300
        Exit Sub
    End If
    'Make sure the picture is always as big as the button
    Picture1.Width = UserControl.Width
    Picture1.Height = UserControl.Height
    'If it is graphical
    If Graphical = True Then
        'If the button should be as big as the pic
        If TheAutoSizeX = ButtonToPic Then
            'Resize the button to right size
            If Picture2.Width > 120 And Picture2.Height > 120 Then
                UserControl.Height = Picture2.Height + 75
                UserControl.Width = Picture2.Width + 75
            End If
        End If
    End If
    'Redraw button
    DrawButton 0, False
End Sub

Public Property Get Caption() As String
    'Get the capture
    Caption = TheCaption
End Property

Public Property Let Caption(ByVal NewCaption As String)
    'Set caption and redraw
    TheCaption = NewCaption
    UserControl_Resize
    PropertyChanged "Caption"
End Property

Public Property Get Style() As TheStyle
    'Get style
    Style = TheStyleX
End Property

Public Property Let Style(ByVal NewStyle As TheStyle)
    'We might need to change some colors here
    'if StandardColors is true.  Redraw afterward
    TheStyleX = NewStyle
    If TheStandardColors = True Then
        If TheStyleX = StyleTwo Then
            TheBorderDark = &H0
            TheBackColorClick = &H8000000D
        Else
            TheBorderDark = &HFFB691
            TheBackColorClick = &HAA6D00
        End If
    End If
    UserControl_Resize
    PropertyChanged "Style"
End Property

Public Property Get Font() As StdFont
    'Get font
    Set Font = TheFont
End Property

Public Property Set Font(NewFont As StdFont)
    'Set font and redraw button
    Set TheFont = NewFont
    Set Picture1.Font = NewFont
    DrawButton 0, False
    PropertyChanged "Font"
End Property

Public Property Get Hwnd() As Long
    'Get Enabled
    Hwnd = Picture1.Hwnd
End Property

Public Property Get Enabled() As Boolean
    'Get Enabled
    Enabled = TheEnabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    'Set enabled and redraw
    TheEnabled = NewValue
    If Enabled = True Then
        UserControl.Enabled = True
    Else
        UserControl.Enabled = False
        ReleaseCapture
    End If
    DrawButton 0, False
    PropertyChanged "Enabled"
End Property

Public Property Get Picture() As Picture
    'Get picture
    Set Picture = Picture2.Picture
End Property

Public Property Set Picture(ByVal NewPicture As Picture)
    'Set picture and redraw
    Set Picture2.Picture = NewPicture
    UserControl_Resize
    PropertyChanged "Picture"
End Property

Public Property Get Graphical() As Boolean
    'Get graphical
    Graphical = TheGraphical
End Property

Public Property Let Graphical(ByVal NewValue As Boolean)
    'Change graphical and redraw
    TheGraphical = NewValue
    UserControl_Resize
    PropertyChanged "Graphical"
End Property

Public Property Get AutoSize() As TheAutoSize
    'Get AutoSize
    AutoSize = TheAutoSizeX
End Property

Public Property Let AutoSize(ByVal NewValue As TheAutoSize)
    'Change AutoSize and redraw
    TheAutoSizeX = NewValue
    UserControl_Resize
    PropertyChanged "AutoSize"
End Property

Public Property Get ForeColor() As OLE_COLOR
    'Get ForeColor
    ForeColor = TheForeColor
End Property

Public Property Let ForeColor(ByVal NewColor As OLE_COLOR)
    'Change ForeColor and StandardColors is false
    If TheStandardColors = True Then
        TheForeColor = &HFFFFFF
        Exit Property
    End If
    TheForeColor = NewColor
    DrawButton 0, False
    PropertyChanged "ForeColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    'Get BackColor
    BackColor = TheBackColor
End Property

Public Property Let BackColor(ByVal NewColor As OLE_COLOR)
    'Change BackColor is we can
    If TheStandardColors = True Then
        TheBackColor = &HAA6D00
        Exit Property
    End If
    TheBackColor = NewColor
    DrawButton 0, False
    PropertyChanged "BackColor"
End Property

Public Property Get StandardColors() As Boolean
    'Get StandardColors
    StandardColors = TheStandardColors
End Property

Public Property Let StandardColors(ByVal NewValue As Boolean)
    'Change other properties if needed and redraw
    TheStandardColors = NewValue
    If NewValue = True Then
        BackColor = &HAA6D00
        ForeColor = &HFFFFFF
        BorderLight = &HFFB691
        If TheStyleX = StyleTwo Then
            BorderDark = &H0
            BackColorClick = &H8000000D
        Else
            BorderDark = &HFFB691
            BackColorClick = &HAA6D00
        End If
    End If
    DrawButton 0, False
    PropertyChanged "StandardColors"
End Property

Public Property Get BorderLight() As OLE_COLOR
    'Get BorderLight
    BorderLight = TheBorderLight
End Property

Public Property Let BorderLight(ByVal NewColor As OLE_COLOR)
    'Change BorderLight if we can
    If TheStandardColors = True Then
        TheBorderLight = &HFFB691
        Exit Property
    End If
    TheBorderLight = NewColor
    DrawButton 0, False
    PropertyChanged "BorderLight"
End Property


Public Property Get BorderDark() As OLE_COLOR
    'Get BorderDark
    BorderDark = TheBorderDark
End Property

Public Property Let BorderDark(ByVal NewColor As OLE_COLOR)
    'Change BorderDark if we can
    If TheStandardColors = True Then
        If TheStyleX = StyleTwo Then
            TheBorderDark = &H0
        Else
            TheBorderDark = &HFFB691
        End If
        Exit Property
    End If
    TheBorderDark = NewColor
    DrawButton 0, False
    PropertyChanged "BorderDark"
End Property

Public Property Get BackColorClick() As OLE_COLOR
    'Get BackColorClick
    BackColorClick = TheBackColorClick
End Property

Public Property Let BackColorClick(ByVal NewColor As OLE_COLOR)
    'Change BackColorClick if we can
    If TheStandardColors = True Then
        If TheStyleX = StyleTwo Then
            TheBackColorClick = &H8000000D
        Else
            TheBackColorClick = &HAA6D00
        End If
        Exit Property
    End If
    TheBackColorClick = NewColor
    PropertyChanged "BackColorClick"
End Property
