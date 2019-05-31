VERSION 5.00
Begin VB.UserControl ucButton_01 
   ClientHeight    =   585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2505
   DefaultCancel   =   -1  'True
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   39
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   167
   Tag             =   "121001"
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   960
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Timer OverTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   360
      Top             =   0
   End
End
Attribute VB_Name = "ucButton_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'INFO:
'This Button Control is totally FREE to use, modify, or sell...., i.e. you can do
'whatever you want with it. You do not need to mention my name or anything.
'On the other hand, I would appreciate if you send me a FEEDBACK.
'This is an enhancement of another project of mine.
'I am sorry about the comments that do not exist.
'
'
'
'WHY THIS BUTTON?
'Well, this is just another button to which you can add pictures/icons. However,
'what makes this button different from the others is as follows:
'-You can customize the size of the picture/icon
'-You do not need a mask property to make the button picture transparent since
' it uses a picture control to achieve this.
'-It includes a URL Navigation function so that you can reach a web address
' or send an e-mail message just by clicking the button.
'-I added a 'sound' property so you'll hear a sound for hover & click events
'-You may say I am very picky but the most important feature of this button
' is that when you rapidly click the button, it responds to the clicks as a
' regular Microsoft VB CommandButton does. So far, I have not seen any
' VB-based user control that could respond to these clicks real-time.
' The ones I saw usually return 1 click graphically when you click
' 2 times rapidly.
'
'
'
'CREDITS:
'I benefited from a lot of people but I can not even remember their names.
'However, I remember Mr.Klaus H. Probst regarding the DrawEdge API, and
'Carles P.V. regarding ShowBorderOnFocus.
'
'
'12 October 2001
'Gurhan KARTAL
'http://gurhan.kartal.org (nothing much there :)
'gurhan@kartal.org
'
'
'
'
'Hope You like it!
'
'
'

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As textparametreleri) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function ShellExecute _
   Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

'-*-*-*-*-* SOUND  -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
Private Declare Function PlaySound Lib "winmm.dll" _
Alias "PlaySoundA" (ByVal lpszName As String, _
ByVal hModule As Long, ByVal dwFlags As Long) As Long

Const SND_ASYNC = &H1 'continue executing code even
'if sound isn't finished
Const SND_FILENAME = &H20000 '  name is a file name
Const SND_SYNC = &H0 'suspend execution until sound is finished
Const SND_NODEFAULT = &H2 'if file name is not found, don't play
'default sound
Const SND_LOOP = &H8 'loop the sound until next call to the
'function
Const SND_NOSTOP = &H10   'don't stop any currently playing sound
Const SND_NOWAIT = &H2000  'return immediately if driver is busy
'-*-*-*-*-* SOUND  BÝTER -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*


Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type textparametreleri
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Public Enum XBPicturePosition
    gbTOP = 0
    gbLEFT = 1
    gbRIGHT = 2
    gbBOTTOM = 3
End Enum
Public Enum XBPictureSize
    size16x16 = 0
    size32x32 = 1
    sizeDefault = 2
    sizeCustom = 3
End Enum

'XPDefault Color Stuff
Private Blue As Double
Private Green As Double
Private Red As Double
Private BlueS As Double
Private GreenS As Double
Private RGBs As String
Private l1 As Double
Private l2 As Double


Private mvarClientRect As RECT
Private mvarPictureRect As RECT
Private mvarCaptionRect As RECT
Dim mvarTempRect As RECT
Dim g_FocusRect As RECT
Dim alan As RECT
Dim g_TextRectUp As RECT, g_TextRectDown As RECT

Dim m_PictureOriginal As Picture
Dim m_PictureHover As Picture
Dim m_Caption As String
Dim m_PicturePosition As XBPicturePosition
Dim m_Picture As Picture
Dim m_PictureWidth As Long
Dim m_PictureHeight As Long
Dim m_PictureSize As XBPictureSize
Dim mvarDrawTextParams As textparametreleri
Dim g_HasFocus As Boolean
Dim g_MouseDown As Boolean, g_MouseIn As Boolean
Dim g_Button As Integer, g_Shift As Integer, g_X As Single, g_Y As Single
Dim g_KeyPressed As Boolean
Dim m_URL As String
Dim m_BorderEdged As Boolean
Dim m_Raised As Boolean
Dim m_ShowBorderOnFocus As Boolean
Dim m_ShowFocusRect As Boolean

Dim WithEvents g_Font As StdFont    'Use this to get rid of font problems
Attribute g_Font.VB_VarHelpID = -1

Const m_def_URL = ""
Const m_def_BorderEdged = 0
Const m_def_Raised = 0
Const m_def_ShowBorderOnFocus = False
Const m_def_ShowFocusRect = False
Const SW_SHOW = 1
Const mvarPadding As Long = 4
Const g_Light = &H80000016
Const g_Shadow = &H80000010
Const g_HighLight = &H80000014
Const g_DarkShadow = &H80000015

Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseIn(Shift As Integer)
Event MouseOut(Shift As Integer)
'**********************************************************************************
'Default Property Values:
Const m_def_SoundOver = "over.wav"
Const m_def_SoundClick = "click.wav"
Const m_def_ForeColor = &H80000012
Const m_def_BackColor = &H8000000F
Const m_def_XPDefaultColors = 0
Const m_def_XPColor_Pressed = &H80000014
Const m_def_XPColor_Hover = &H80000016
Const m_def_XPStyle = 1
'Property Variables:
Dim m_SoundOver As String
Dim m_SoundClick As String
Dim m_ForeColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_XPDefaultColors As Boolean
Dim m_XPColor_Pressed As OLE_COLOR
Dim m_XPColor_Hover As OLE_COLOR
Dim m_XPStyle As Boolean

Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    picBuffer.BackColor = UserControl.BackColor ' Ambient.BackColor
    m_ForeColor = m_def_ForeColor
    m_ShowBorderOnFocus = m_def_ShowBorderOnFocus
    m_ShowFocusRect = m_def_ShowFocusRect
    Set UserControl.Font = Ambient.Font
    Set g_Font = Ambient.Font
    m_Caption = Ambient.DisplayName
    m_PicturePosition = 1
    m_PictureWidth = 32
    m_PictureHeight = 32
    m_PictureSize = 1
    Set m_PictureHover = LoadPicture("")
    Set m_PictureOriginal = LoadPicture("")
    m_Raised = m_def_Raised
    m_BorderEdged = m_def_BorderEdged
    m_URL = m_def_URL
    m_XPStyle = m_def_XPStyle
    m_XPColor_Pressed = m_def_XPColor_Pressed
    m_XPColor_Hover = m_def_XPColor_Hover
    m_XPDefaultColors = m_def_XPDefaultColors
    
    m_SoundOver = m_def_SoundOver
    m_SoundClick = m_def_SoundClick
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    UserControl.BackColor = m_BackColor
    picBuffer.BackColor = m_BackColor
    
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    UserControl.ForeColor = m_ForeColor
    
    m_ShowFocusRect = PropBag.ReadProperty("ShowFocusRect", m_def_ShowFocusRect)
    m_ShowBorderOnFocus = PropBag.ReadProperty("ShowBorderOnFocus", m_def_ShowBorderOnFocus)
    m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    m_PicturePosition = PropBag.ReadProperty("PicturePosition", 1)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    m_PictureWidth = PropBag.ReadProperty("PictureWidth", 32)
    m_PictureHeight = PropBag.ReadProperty("PictureHeight", 32)
    m_PictureSize = PropBag.ReadProperty("PictureSize", 1)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set g_Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set m_PictureHover = PropBag.ReadProperty("PictureHover", Nothing)
    Set m_PictureOriginal = PropBag.ReadProperty("Picture", Nothing)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", Verdadero)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
''''''    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_Raised = PropBag.ReadProperty("Raised", m_def_Raised)
    m_BorderEdged = PropBag.ReadProperty("BorderEdged", m_def_BorderEdged)
    m_URL = PropBag.ReadProperty("URL", m_def_URL)
    m_XPStyle = PropBag.ReadProperty("XPStyle", m_def_XPStyle)
    m_XPColor_Pressed = PropBag.ReadProperty("XPColor_Pressed", m_def_XPColor_Pressed)
    m_XPColor_Hover = PropBag.ReadProperty("XPColor_Hover", m_def_XPColor_Hover)
    m_XPDefaultColors = PropBag.ReadProperty("XPDefaultColors", m_def_XPDefaultColors)
    
    m_SoundOver = PropBag.ReadProperty("SoundOver", m_def_SoundOver)
    m_SoundClick = PropBag.ReadProperty("SoundClick", m_def_SoundClick)
Refresh
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", m_Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("PicturePosition", m_PicturePosition, 1)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("PictureWidth", m_PictureWidth, 32)
    Call PropBag.WriteProperty("PictureHeight", m_PictureHeight, 32)
    Call PropBag.WriteProperty("PictureSize", m_PictureSize, 1)
    Call PropBag.WriteProperty("PictureHover", m_PictureHover, Nothing)
    
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, Verdadero)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("ShowBorderOnFocus", m_ShowBorderOnFocus, m_def_ShowBorderOnFocus)
    Call PropBag.WriteProperty("ShowFocusRect", m_ShowFocusRect, m_def_ShowFocusRect)
 
    Call PropBag.WriteProperty("Raised", m_Raised, m_def_Raised)
    Call PropBag.WriteProperty("BorderEdged", m_BorderEdged, m_def_BorderEdged)
    Call PropBag.WriteProperty("URL", m_URL, m_def_URL)
    Call PropBag.WriteProperty("XPStyle", m_XPStyle, m_def_XPStyle)
    Call PropBag.WriteProperty("XPColor_Pressed", m_XPColor_Pressed, m_def_XPColor_Pressed)
    Call PropBag.WriteProperty("XPColor_Hover", m_XPColor_Hover, m_def_XPColor_Hover)
    Call PropBag.WriteProperty("XPDefaultColors", m_XPDefaultColors, m_def_XPDefaultColors)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    
    Call PropBag.WriteProperty("SoundOver", m_SoundOver, m_def_SoundOver)
    Call PropBag.WriteProperty("SoundClick", m_SoundClick, m_def_SoundClick)
 End Sub
Private Sub CalcRECTs()
    Dim picWidth, picHeight, capWidth, capHeight As Long
    alan.Left = 0
    alan.Top = 0
    alan.Right = UserControl.ScaleWidth - 1
    alan.Bottom = UserControl.ScaleHeight - 1
    
    With mvarClientRect
     .Left = alan.Left + mvarPadding
     .Top = alan.Top + mvarPadding
     .Right = alan.Right - mvarPadding + 1
     .Bottom = alan.Bottom - mvarPadding + 1
    End With
    
    If m_Picture Is Nothing Then
        With mvarCaptionRect
           .Left = mvarClientRect.Left
           .Top = mvarClientRect.Top
           .Right = mvarClientRect.Right
           .Bottom = mvarClientRect.Bottom
        End With
        CalculateCaptionRect 'Local Sub
    Else
        If m_Caption = "" Then
         With mvarPictureRect
            .Left = (((mvarClientRect.Right - mvarClientRect.Left) - m_PictureWidth) \ 2) + mvarClientRect.Left
            .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - m_PictureHeight) \ 2) + mvarClientRect.Top
            .Right = mvarPictureRect.Left + m_PictureWidth
            .Bottom = mvarPictureRect.Top + m_PictureHeight
         End With
            Exit Sub
        End If
        
        With mvarCaptionRect
        .Left = mvarClientRect.Left
        .Top = mvarClientRect.Top
        .Right = mvarClientRect.Right
        .Bottom = mvarClientRect.Bottom
        End With
        CalculateCaptionRect
        'Width and Height of the picture and the caption
        picWidth = m_PictureWidth
        picHeight = m_PictureHeight
        capWidth = mvarCaptionRect.Right - mvarCaptionRect.Left
        capHeight = mvarCaptionRect.Bottom - mvarCaptionRect.Top
        Select Case m_PicturePosition
        Case gbLEFT
            'final values for the picture and caption rectangles
        With mvarPictureRect
            .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - picHeight) \ 2) + mvarClientRect.Top
            .Left = (((mvarClientRect.Right - mvarClientRect.Left) - (picWidth + mvarPadding + capWidth)) \ 2) + mvarClientRect.Left
            .Bottom = mvarPictureRect.Top + picHeight
            .Right = mvarPictureRect.Left + picWidth
        End With
        With mvarCaptionRect
            .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - capHeight) \ 2) + mvarClientRect.Top
            .Left = mvarPictureRect.Right + mvarPadding
            .Bottom = mvarCaptionRect.Top + capHeight
            .Right = mvarCaptionRect.Left + capWidth
        End With
        
        Case gbRIGHT
            'final values for the picture and caption rectangles
        With mvarCaptionRect
            .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - capHeight) \ 2) + mvarClientRect.Top
            .Left = (((mvarClientRect.Right - mvarClientRect.Left) - (picWidth + mvarPadding + capWidth)) \ 2) + mvarClientRect.Left
            .Bottom = mvarCaptionRect.Top + capHeight
            .Right = mvarCaptionRect.Left + capWidth
        End With
        With mvarPictureRect
            .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - picHeight) \ 2) + mvarClientRect.Top
            .Left = mvarCaptionRect.Right + mvarPadding
            .Bottom = mvarPictureRect.Top + picHeight
            .Right = mvarPictureRect.Left + picWidth
        End With
        Case gbTOP
            'final values for the picture and caption rectangles
        With mvarPictureRect
            .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - (picHeight + mvarPadding + capHeight)) \ 2) + mvarClientRect.Top
            .Left = (((mvarClientRect.Right - mvarClientRect.Left) - picWidth) \ 2) + mvarClientRect.Left
            .Bottom = mvarPictureRect.Top + picHeight
            .Right = mvarPictureRect.Left + picWidth
        End With
        With mvarCaptionRect
            .Top = mvarPictureRect.Bottom + mvarPadding
            .Left = (((mvarClientRect.Right - mvarClientRect.Left) - capWidth) \ 2) + mvarClientRect.Left
            .Bottom = mvarCaptionRect.Top + capHeight
            .Right = mvarCaptionRect.Left + capWidth
        End With
        Case gbBOTTOM
            'final values for the picture and caption rectangles
        With mvarCaptionRect
            .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - (picHeight + mvarPadding + capHeight)) \ 2) + mvarClientRect.Top
            .Left = (((mvarClientRect.Right - mvarClientRect.Left) - capWidth) \ 2) + mvarClientRect.Left
            .Bottom = mvarCaptionRect.Top + capHeight
            .Right = mvarCaptionRect.Left + capWidth
        End With
        With mvarPictureRect
            .Top = mvarCaptionRect.Bottom + mvarPadding
            .Left = (((mvarClientRect.Right - mvarClientRect.Left) - picWidth) \ 2) + mvarClientRect.Left
            .Bottom = mvarPictureRect.Top + picHeight
            .Right = mvarPictureRect.Left + picWidth
        End With
        End Select
    End If
End Sub

Private Sub UserControl_Initialize()
    Set g_Font = New StdFont
    l1 = 100
    l2 = 160
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    If Not Me.Enabled Then Exit Sub
    If KeyAscii = 13 Or KeyAscii = 27 Then 'Default / Cancel
        RaiseEvent Click
        GoToURL
    End If
End Sub

'Private Sub UserControl_AmbientChanged(PropertyName As String)
'    Refresh 'Extender.Default changed
'End Sub

Private Sub UserControl_EnterFocus()
    g_HasFocus = True
    Refresh
End Sub

Private Sub UserControl_ExitFocus()
    g_HasFocus = False
    g_MouseDown = False
    Refresh
End Sub

Private Sub UserControl_Resize()
    'Minimum size = 10 x 10 pixels
    If ScaleWidth < 10 Then UserControl.Width = 150
    If ScaleHeight < 10 Then UserControl.Height = 150
    'Set focus rect:
    g_FocusRect.Left = 2
    g_FocusRect.Right = ScaleWidth - 2
    g_FocusRect.Top = 2
    g_FocusRect.Bottom = ScaleHeight - 2
    Refresh
End Sub
Public Sub ConvertRGB()
    P = UserControl.Point(X, Y)
    Blue = Fix((P / 256) / 256)
    BlueS = (Blue * 256) * 256
    Green = Fix((P - BlueS) / 256)
    GreenS = Green * 256
    Red = Fix(P - BlueS - GreenS)
    RGBs = "RGB(" & Red & ", " & Green & ", " & Blue & ")"
End Sub

Public Sub Refresh()
    AutoRedraw = True
    'Clearing everything
    UserControl.Cls
    If XPStyle = True Then
        UserControl.BackColor = BackColor 'UserControl.Ambient.BackColor
        picBuffer.BackColor = BackColor 'UserControl.Ambient.BackColor
        UserControl.ForeColor = ForeColor ' ?vbButtonText
    End If
    
    'If XP then adjust colors:
    If XPStyle = True Then
        If Not g_MouseDown And g_MouseIn Then 'Mouse Over but Not Pressed
                If XPDefaultColors = True Then
                    UserControl.BackColor = vbHighlight
                    ConvertRGB 'Get RGB Colors
                    UserControl.BackColor = RGB(Red + l1, Green + l1, Blue + l1)
                    picBuffer.BackColor = RGB(Red + l1, Green + l1, Blue + l1)
                    UserControl.ForeColor = vbHighlightText
                Else 'Use user defined colors
                    UserControl.BackColor = XPColor_Hover
                    picBuffer.BackColor = XPColor_Hover
                End If
        End If
        
        If g_MouseDown Then   'Mouse Over and Pressed
                If XPDefaultColors = True Then
                    UserControl.BackColor = RGB(Red + l2, Green + l2, Blue + l2)
                    picBuffer.BackColor = RGB(Red + l2, Green + l2, Blue + l2)
                Else 'Use user defined colors
                    UserControl.BackColor = XPColor_Pressed
                    picBuffer.BackColor = XPColor_Pressed
                End If
        End If
    End If
   
    
    'OK continue ...
    CalcRECTs
    DrawPicture
    If g_HasFocus And m_ShowFocusRect Then DrawFocusRect hdc, g_FocusRect
    DrawCaption
    Draw3DEffect
    AutoRedraw = False
End Sub

Private Sub UserControl_DblClick()
    SetCapture hwnd 'Preseve hWnd on DblClick
    UserControl_MouseDown g_Button, g_Shift, g_X, g_Y
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not g_KeyPressed Then ' Not continuous clicking
                             ' If you want it, remove this filter
                             ' ... or create a new property
        Select Case KeyCode
            Case vbKeyReturn
                RaiseEvent Click
                GoToURL
            Case vbKeySpace
                g_MouseDown = True
                Refresh
                RaiseEvent Click
                GoToURL
        End Select
        g_KeyPressed = True
    End If
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        g_MouseDown = False
        Refresh
    End If
    g_KeyPressed = False
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    g_Button = Button: g_Shift = Shift: g_X = X: g_Y = Y
    If Button <> vbRightButton Then
        g_MouseDown = True
        Refresh
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (X >= 0 And Y >= 0) And (X < ScaleWidth And Y < ScaleHeight) Then
        If g_MouseIn = False Then
            OverTimer.Enabled = True
            g_MouseIn = True
            If Not m_PictureHover Is Nothing Then
                Set m_Picture = m_PictureHover
            End If
            RaiseEvent MouseIn(Shift)
            Refresh
            DoEvents
            dd = PlayASound(SoundOver)
        End If
    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    g_MouseDown = False
    If Button <> vbRightButton Then
        Refresh
        If (X >= 0 And Y >= 0) And (X < ScaleWidth And Y < ScaleHeight) Then
            dd = PlayASound(SoundClick)
            RaiseEvent Click
            GoToURL
        End If
    End If
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    Refresh
End Property
Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
    Set Font = g_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    With g_Font
        .Name = New_Font.Name
        .SIZE = New_Font.SIZE
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
    End With
    PropertyChanged "Font"
End Property

Private Sub g_Font_FontChanged(ByVal PropertyName As String)
    Set UserControl.Font = g_Font
    Refresh
End Sub

'?????????????????? LAZIM MI???????????
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As StdPicture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property
Public Property Get ShowBorderOnFocus() As Boolean
    ShowBorderOnFocus = m_ShowBorderOnFocus
End Property

Public Property Let ShowBorderOnFocus(ByVal New_ShowBorderOnFocus As Boolean)
    m_ShowBorderOnFocus = New_ShowBorderOnFocus
    PropertyChanged "ShowBorderOnFocus"
    Refresh
End Property

Public Property Get ShowFocusRect() As Boolean
    ShowFocusRect = m_ShowFocusRect
End Property

Public Property Let ShowFocusRect(ByVal New_ShowFocusRect As Boolean)
    m_ShowFocusRect = New_ShowFocusRect
    PropertyChanged "ShowFocusRect"
    Refresh
End Property
             
Private Sub Draw3DEffect()
    If Not Ambient.UserMode Then
         Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), g_Shadow, B
         Line (0, 0)-(ScaleWidth - 0, ScaleHeight - 0), g_HighLight, B
    End If
    
    If XPStyle = True Then
        If Not g_MouseDown And g_MouseIn Then 'ÜSTÜNDE AMA BASILI DEÐÝL
            Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), UserControl.ForeColor, B
            Line (0, 0)-(ScaleWidth - 0, ScaleHeight - 0), UserControl.ForeColor, B
            Exit Sub
        End If
        If g_MouseDown Then   'ÜSTÜNDE VE BASILI
            Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), UserControl.ForeColor, B
            Line (0, 0)-(ScaleWidth - 0, ScaleHeight - 0), UserControl.ForeColor, B
            Exit Sub
        End If
        If Not g_MouseDown Then  'DIÞARDA VE BASILI DEÐÝL
            Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), g_HighLight, B
            Line (0, 0)-(ScaleWidth - 0, ScaleHeight - 0), g_HighLight, B
            Exit Sub
        End If
    End If
    
    Select Case BorderEdged
    Case Is = False
        If g_MouseDown Then 'BASILDI
            Line (1, 1)-(ScaleWidth - 1, ScaleHeight - 1), g_Shadow, B
            Line (0, 0)-(ScaleWidth - 2, ScaleHeight - 2), g_Light, B
            Line (0, 0)-(ScaleWidth - 0, ScaleHeight - 0), g_DarkShadow, B
            Line (-1, -1)-(ScaleWidth - 1, ScaleHeight - 1), g_HighLight, B
        End If
        If Not g_MouseDown And g_MouseIn Then 'ÜSTÜNDE AMA BASILI DEÐÝL
            Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), g_Shadow, B
            Line (0, 0)-(ScaleWidth - 0, ScaleHeight - 0), g_HighLight, B
        End If
        
        If Not g_MouseDown And Not g_MouseIn And RAISED Then 'DIÞARDA ÝSE VE RAISED ÝSE
            Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), g_Shadow, B
            Line (0, 0)-(ScaleWidth - 0, ScaleHeight - 0), g_HighLight, B
        End If
         'BORDER
          If (g_HasFocus And m_ShowBorderOnFocus And RAISED And Not g_MouseDown) Or Extender.Default Then
           ' Line (1 + g_3DInc, 1 + g_3DInc)-(ScaleWidth - g_3DInc - 1, ScaleHeight - g_3DInc - 1), g_Light, B
            Line (0, 0)-(ScaleWidth - 2, ScaleHeight - 2), g_Shadow, B
            Line (0, 0)-(ScaleWidth - 0, ScaleHeight - g_3DInc - 0), g_HighLight, B
            Line (-1, -1)-(ScaleWidth - 1, ScaleHeight - 1), g_DarkShadow, B
            'Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), vbBlack, B 'DARK BORDER
         End If
         
    Case Is = True
            Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), g_Shadow, B
            Line (1, 1)-(ScaleWidth - 2, ScaleHeight - 2), g_HighLight, B
            Line (0, 0)-(ScaleWidth - 0, ScaleHeight - 0), g_DarkShadow, B
            Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), g_HighLight, B
            Line (0, 0)-(ScaleWidth - 2, ScaleHeight - 2), g_Shadow, B
    
        If g_MouseDown Then 'BASILDI
            Line (1, 1)-(ScaleWidth - 1, ScaleHeight - 1), g_Shadow, B
            Line (0, 0)-(ScaleWidth - 0, ScaleHeight - 0), g_DarkShadow, B
            Line (-1, -1)-(ScaleWidth - 1, ScaleHeight - 1), g_HighLight, B
            Line (1.5, 1.5)-(ScaleWidth - 2, ScaleHeight - 2), g_DarkShadow, B 'DARK BORDER
        End If
        
        If Not g_MouseDown And (g_MouseIn Or g_HasFocus) Then 'ÜSTÜNDE AMA BASILI DEÐÝL
            Line (2, 2)-(ScaleWidth - 4, 2), g_HighLight
            Line (2, 2)-(2, ScaleHeight - 3), g_HighLight
            Line (0, 0)-(ScaleWidth - 3, ScaleHeight - 3), g_DarkShadow, B
        End If
'        If Not g_MouseDown And Not g_MouseIn Then  'DIÞARDA ÝSE
'
'        End If
    End Select
End Sub

Private Sub OverTimer_Timer()
    Dim P As POINTAPI
    GetCursorPos P
    If hwnd <> WindowFromPoint(P.X, P.Y) Then
        OverTimer.Enabled = False
        g_MouseIn = False
        Set m_Picture = m_PictureOriginal
        RaiseEvent MouseOut(g_Shift)
        Refresh                     'Refresh picture
        If g_MouseDown = True Then  'Resfresh state
            g_MouseDown = False
            Refresh
            g_MouseDown = True
        End If
    End If
End Sub

Public Property Get RAISED() As Boolean
    RAISED = m_Raised
End Property

Public Property Let RAISED(ByVal New_Raised As Boolean)
    m_Raised = New_Raised
    PropertyChanged "Raised"
End Property

Public Property Get BorderEdged() As Boolean
    BorderEdged = m_BorderEdged
End Property

Public Property Let BorderEdged(ByVal New_BorderEdged As Boolean)
    m_BorderEdged = New_BorderEdged
    PropertyChanged "BorderEdged"
    Refresh
End Property

Public Sub GoToURL()
    'On Error Resume Next
    If Left(m_URL, 7) = "mailto:" Then
        Navigate UserControl.Parent, m_URL
        Exit Sub
    End If
        If Not m_URL = "" Then UserControl.Hyperlink.NavigateTo m_URL
End Sub
Private Sub Navigate(frm As Form, ByVal WebPageURL As String)
Dim hBrowse As Long
hBrowse = ShellExecute(frm.hwnd, "open", WebPageURL, "", "", 1)
End Sub
Public Property Get URL() As String
    URL = m_URL
End Property

Public Property Let URL(ByVal New_URL As String)
    m_URL = New_URL
    PropertyChanged "URL"
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    Refresh
End Property
Public Property Get PicturePosition() As XBPicturePosition
    PicturePosition = m_PicturePosition
End Property
Public Property Let PicturePosition(ByVal New_PicturePosition As XBPicturePosition)
    m_PicturePosition = New_PicturePosition
    PropertyChanged "PicturePosition"
    Refresh
End Property
Public Property Get Picture() As Picture
    Set Picture = m_Picture
End Property
Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    Set m_PictureOriginal = New_Picture
    PropertyChanged "Picture"
    If m_PictureSize = sizeDefault Then
        m_PictureWidth = UserControl.ScaleX(m_Picture.Width, vbHimetric, UserControl.ScaleMode)
        m_PictureHeight = UserControl.ScaleY(m_Picture.Height, vbHimetric, UserControl.ScaleMode)
    End If
    Refresh
End Property

Public Property Get PictureWidth() As Long
    PictureWidth = m_PictureWidth
End Property
Public Property Let PictureWidth(ByVal New_PictureWidth As Long)
    m_PictureWidth = New_PictureWidth
    PropertyChanged "PictureWidth"
    Refresh
End Property
Public Property Get PictureHeight() As Long
    PictureHeight = m_PictureHeight
End Property
Public Property Let PictureHeight(ByVal New_PictureHeight As Long)
    m_PictureHeight = New_PictureHeight
    PropertyChanged "PictureHeight"
    Refresh
End Property
Public Property Get PictureSize() As XBPictureSize
    PictureSize = m_PictureSize
End Property
Public Property Let PictureSize(ByVal New_PictureSize As XBPictureSize)
    m_PictureSize = New_PictureSize
    PropertyChanged "PictureSize"
    Select Case New_PictureSize
    Case size16x16
        m_PictureWidth = 16
        m_PictureHeight = 16
    Case size32x32
        m_PictureWidth = 32
        m_PictureHeight = 32
    Case sizeDefault
        If Not (m_Picture Is Nothing) Then
            m_PictureWidth = UserControl.ScaleX(m_Picture.Width, vbHimetric, UserControl.ScaleMode)
            m_PictureHeight = UserControl.ScaleY(m_Picture.Height, vbHimetric, UserControl.ScaleMode)
        Else
            m_PictureWidth = 32
            m_PictureHeight = 32
        End If
    End Select
    Refresh
End Property

Private Sub CalculateCaptionRect()
    Dim mvarWidth, mvarHeight As Long
    Dim mvarFormat As Long
    With mvarDrawTextParams
        .iLeftMargin = 1
        .iRightMargin = 1
        .iTabLength = 1
        .cbSize = Len(mvarDrawTextParams)
    End With
    mvarFormat = &H400 Or &H10 Or &H4 Or &H1
    DrawTextEx UserControl.hdc, m_Caption, Len(m_Caption), mvarCaptionRect, mvarFormat, mvarDrawTextParams
    mvarWidth = mvarCaptionRect.Right - mvarCaptionRect.Left
    mvarHeight = mvarCaptionRect.Bottom - mvarCaptionRect.Top
    With mvarCaptionRect
        .Left = mvarClientRect.Left + (((mvarClientRect.Right - mvarClientRect.Left) - (mvarCaptionRect.Right - mvarCaptionRect.Left)) \ 2)
        .Top = mvarClientRect.Top + (((mvarClientRect.Bottom - mvarClientRect.Top) - (mvarCaptionRect.Bottom - mvarCaptionRect.Top)) \ 2)
        .Right = mvarCaptionRect.Left + mvarWidth
        .Bottom = mvarCaptionRect.Top + mvarHeight
    End With
End Sub

Private Sub DrawCaption()
    If m_Caption = "" Then Exit Sub
    Dim mvarForeColor As OLE_COLOR
    mvarTempRect = mvarCaptionRect
    If g_MouseDown Then
       With mvarCaptionRect
        .Left = mvarCaptionRect.Left + 1
        .Top = mvarCaptionRect.Top + 1
        .Right = mvarCaptionRect.Right + 1
        .Bottom = mvarCaptionRect.Bottom + 1
       End With
    End If
    
    If Not Enabled Then
        Dim g_tmpFontColor As OLE_COLOR
        g_tmpFontColor = UserControl.ForeColor
        
        'AÇIK DISABLED YAZI
        UserControl.ForeColor = g_HighLight
        Dim mvarCaptionRect_Iki As RECT
        With mvarCaptionRect_Iki
            .Bottom = mvarCaptionRect.Bottom
            .Left = mvarCaptionRect.Left + 1
            .Right = mvarCaptionRect.Right + 1
            .Top = mvarCaptionRect.Top + 1
        End With
        DrawTextEx UserControl.hdc, m_Caption, Len(m_Caption), mvarCaptionRect_Iki, &H10 Or &H4 Or &H1, mvarDrawTextParams
        
        'KOYU DISABLED YAZI
        UserControl.ForeColor = g_Shadow
        DrawTextEx UserControl.hdc, m_Caption, Len(m_Caption), mvarCaptionRect, &H10 Or &H4 Or &H1, mvarDrawTextParams
        
        'Normale çevir
        UserControl.ForeColor = g_tmpFontColor
        Exit Sub
    End If
    
    DrawTextEx UserControl.hdc, m_Caption, Len(m_Caption), mvarCaptionRect, &H10 Or &H4 Or &H1, mvarDrawTextParams
    mvarCaptionRect = mvarTempRect
End Sub


Private Sub DrawPicture()
    Dim mvarImageType As Long
    Dim mvarImageState As Long
    Dim mvarImageFlag As Long
    If m_Picture Is Nothing Then Exit Sub
    Select Case m_Picture.Type
    Case vbPicTypeBitmap
        mvarImageType = &H4
    Case vbPicTypeIcon
        mvarImageType = &H3
    End Select
    If Not Enabled Then
        mvarImageState = &H20
    Else
        mvarImageState = &H0
    End If
    mvarTempRect = mvarPictureRect
    If g_MouseDown Then
        With mvarPictureRect
        .Left = mvarPictureRect.Left + 1
        .Top = mvarPictureRect.Top + 1
        .Right = mvarPictureRect.Right + 1
        .Bottom = mvarPictureRect.Bottom + 1
        End With
    End If
    mvarImageFlag = mvarImageType Or mvarImageState
    picBuffer.Width = UserControl.ScaleX(m_Picture.Width, vbHimetric, UserControl.ScaleMode)
    picBuffer.Height = UserControl.ScaleY(m_Picture.Height, vbHimetric, UserControl.ScaleMode)
    picBuffer.ScaleMode = 3
    picBuffer.Cls
    DrawState picBuffer.hdc, 0, 0, m_Picture, 0, 0, 0, 0, 0, mvarImageFlag
    StretchBlt UserControl.hdc, mvarPictureRect.Left, mvarPictureRect.Top, mvarPictureRect.Right - mvarPictureRect.Left, mvarPictureRect.Bottom - mvarPictureRect.Top, picBuffer.hdc, picBuffer.ScaleLeft, picBuffer.ScaleTop, picBuffer.ScaleWidth, picBuffer.ScaleHeight, &HCC0020
    mvarPictureRect = mvarTempRect
End Sub

Public Property Get PictureHover() As Picture
    Set PictureHover = m_PictureHover
End Property

Public Property Set PictureHover(ByVal New_PictureHover As Picture)
    Set m_PictureHover = New_PictureHover
    PropertyChanged "PictureHover"
End Property
Public Property Get XPStyle() As Boolean
    XPStyle = m_XPStyle
End Property

Public Property Let XPStyle(ByVal New_XPStyle As Boolean)
    m_XPStyle = New_XPStyle
    PropertyChanged "XPStyle"
    Refresh
End Property
Public Property Get XPColor_Pressed() As OLE_COLOR
    XPColor_Pressed = m_XPColor_Pressed
End Property

Public Property Let XPColor_Pressed(ByVal New_XPColor_Pressed As OLE_COLOR)
    m_XPColor_Pressed = New_XPColor_Pressed
    PropertyChanged "XPColor_Pressed"
End Property
Public Property Get XPColor_Hover() As OLE_COLOR
    XPColor_Hover = m_XPColor_Hover
End Property

Public Property Let XPColor_Hover(ByVal New_XPColor_Hover As OLE_COLOR)
    m_XPColor_Hover = New_XPColor_Hover
    PropertyChanged "XPColor_Hover"
End Property
Public Property Get XPDefaultColors() As Boolean
    XPDefaultColors = m_XPDefaultColors
End Property
Public Property Let XPDefaultColors(ByVal New_XPDefaultColors As Boolean)
    m_XPDefaultColors = New_XPDefaultColors
    PropertyChanged "XPDefaultColors"
End Property
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    UserControl.BackColor = m_BackColor
    picBuffer.BackColor = m_BackColor
    Refresh
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    UserControl.ForeColor = m_ForeColor
    Refresh
End Property
Public Property Get SoundOver() As Variant
    SoundOver = m_SoundOver
End Property
Public Property Let SoundOver(ByVal New_SoundOver As Variant)
    m_SoundOver = New_SoundOver
    PropertyChanged "SoundOver"
End Property
Public Property Get SoundClick() As String
    SoundClick = m_SoundClick
End Property
Public Property Let SoundClick(ByVal New_SoundClick As String)
    m_SoundClick = New_SoundClick
    PropertyChanged "SoundClick"
End Property
Public Property Get version() As String
Attribute version.VB_Description = "FileVersion"
    version = UserControl.Tag
End Property
Public Property Let version(ByVal New_version As String)
End Property
Private Function PlayASound(SoundFile As String) As Boolean
    Dim bSuccess As Boolean
    'ESKÝ HALÝ(ORJÝNAL)
'    bSuccess = PlaySound(SoundFile, vbNull, SND_FILENAME _
'    + SND_SYNC + SND_NOSTOP + SND_NODEFAULT)
'    PlayASound = bSuccess

    'KULLANDIÐIM:
    bSuccess = PlaySound(SoundFile, vbNull, SND_FILENAME _
    + SND_SYNC + SND_ASYNC + SND_NODEFAULT)
    PlayASound = bSuccess
End Function
