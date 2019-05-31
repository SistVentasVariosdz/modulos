VERSION 5.00
Begin VB.UserControl ucProgressCircular 
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   960
   ScaleHeight     =   960
   ScaleWidth      =   960
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "ucProgressCircular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'---------------------------------------
'Autor: Leandro Ascierto
'Web:   www.leandroascierto.com.ar
'Date:  23/11/2010
'---------------------------------------

Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, ByRef graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mSmoothingMode As Long) As Long
Private Declare Function GdipDrawLine Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Single, ByVal mY1 As Single, ByVal mX2 As Single, ByVal mY2 As Single) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipSetPenStartCap Lib "gdiplus" (ByVal pen As Long, ByVal startCap As LineCap) As Long
Private Declare Function GdipSetPenEndCap Lib "gdiplus" (ByVal pen As Long, ByVal endCap As LineCap) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'---------------------------GDI PLUS SAFE MODE (By LaVolpe)
Private Declare Function CreateWindowExA Lib "user32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Const GWL_WNDPROC       As Long = -4
Private Const GW_OWNER          As Long = 4
Private Const WS_CHILD          As Long = &H40000000
'------------------------------------------------------------

Private Type GDIPlusStartupInput
    GdiPlusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type
 
Public Enum LineCap
    LineCapFlat = &H0
    LineCapSquare = &H1
    LineCapRound = &H2
    LineCapTriangle = &H3
    LineCapNoAnchor = &H10
    LineCapSquareAnchor = &H11
    LineCapRoundAnchor = &H12
    LineCapDiamondAnchor = &H13
    LineCapArrowAnchor = &H14
End Enum

Public Enum enuNumberOfLines
    FortyEightLines = &H0
    TwentyFourLines = &H1
    TwelveLines = &H2
    EightLine = &H3
    SixtLine = &H4
    FourLine = &H5
End Enum

Private Const SmoothingModeAntiAlias    As Long = &H4
Private Const UnitPixel As Long = &H2
Private Const PI180 = 3.14159 / 180

'Private GdipToken          As Long
Private CurrentPos          As Long
Private mDrawWidth          As Long
Private mBackColor          As OLE_COLOR
Private mForeColor          As OLE_COLOR
Private mLineStart          As LineCap
Private mLineEnd            As LineCap
Private mTotalLines         As Long
Private mNumberOfLines      As enuNumberOfLines
Private mInterval           As Long


Public Property Get NumberOfLines() As enuNumberOfLines
    NumberOfLines = mNumberOfLines
End Property


Public Property Let NumberOfLines(ByVal lngNumber As enuNumberOfLines)
    Select Case lngNumber
        Case FortyEightLines: mTotalLines = 7.5
        Case TwentyFourLines: mTotalLines = 15
        Case TwelveLines: mTotalLines = 30
        Case EightLine: mTotalLines = 45
        Case SixtLine: mTotalLines = 60
        Case FourLine: mTotalLines = 90
        Case Else
            lngNumber = TwelveLines
            mTotalLines = 12
    End Select
    
    mNumberOfLines = lngNumber

    PropertyChanged "NumberOfLines"
    Call Draw
End Property


Public Property Get Interval() As Long
    Interval = mInterval
End Property


Public Property Let Interval(ByVal lngValue As Long)
    mInterval = lngValue
    PropertyChanged "Interval"
    If Ambient.UserMode Then
        Timer1.Interval = lngValue
    End If
End Property


Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal lngForeColor As OLE_COLOR)
    mForeColor = lngForeColor
    PropertyChanged "ForeColor"
    Call Draw
End Property


Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal lngBackColor As OLE_COLOR)
    mBackColor = lngBackColor
    UserControl.BackColor = mBackColor
    PropertyChanged "BackColor"
    Call Draw
End Property


Public Property Get LineStart() As LineCap
    LineStart = mLineStart
End Property


Public Property Let LineStart(ByVal enuLineStart As LineCap)
    mLineStart = enuLineStart
    PropertyChanged "LineStart"
    Call Draw
End Property


Public Property Get LineEnd() As LineCap
    LineEnd = mLineEnd
End Property


Public Property Let LineEnd(ByVal enuLineEnd As LineCap)
    mLineEnd = enuLineEnd
    PropertyChanged "LineEnd"
    Call Draw
End Property

Public Property Get DrawWidth() As Long
    DrawWidth = mDrawWidth
End Property


Public Property Let DrawWidth(ByVal lDrawWidth As Long)
    mDrawWidth = lDrawWidth
    PropertyChanged "DrawWidth"
    Call Draw
End Property


Private Sub UserControl_Initialize()
    'InitGDI
    CurrentPos = 360
    UserControl.ScaleMode = vbPixels
    UserControl.AutoRedraw = True
End Sub


Private Sub UserControl_InitProperties()
    mLineStart = LineCapRound
    mLineEnd = LineCapRound
    mBackColor = Ambient.BackColor
    UserControl.BackColor = mBackColor
    mForeColor = Ambient.ForeColor
    mInterval = 100
    mDrawWidth = 6
    Me.NumberOfLines = TwelveLines
    Call ManageGDIToken(UserControl.ContainerHwnd)
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 
    Call ManageGDIToken(UserControl.ContainerHwnd)
    
    With PropBag
        mForeColor = .ReadProperty("ForeColor", vbWindowText)
        mBackColor = .ReadProperty("BackColor", vbButtonFace)
        mLineStart = .ReadProperty("LineStart", LineCapRound)
        mLineEnd = .ReadProperty("Lineend", LineCapRound)
        mDrawWidth = .ReadProperty("DrawWidth", 6)
        mInterval = .ReadProperty("Interval", 100)
        UserControl.BackColor = mBackColor
        Me.NumberOfLines = .ReadProperty("NumberOfLines", TwelveLines) 'And call Draw
    End With

    If Ambient.UserMode Then
        Timer1.Interval = mInterval
    End If
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BackColor", mBackColor, vbButtonFace
        .WriteProperty "ForeColor", mForeColor, vbWindowText
        .WriteProperty "LineStart", mLineStart, LineCapRound
        .WriteProperty "LineEnd", mLineEnd, LineCapRound
        .WriteProperty "DrawWidth", mDrawWidth, 6
        .WriteProperty "NumberOfLines", mNumberOfLines, TwelveLines
        .WriteProperty "Interval", mInterval, 100
    End With
End Sub


Private Sub UserControl_Resize()
    
    If UserControl.ScaleWidth > UserControl.ScaleHeight Then
        UserControl.Height = UserControl.Width
    Else
        UserControl.Width = UserControl.Height
    End If

    Draw
End Sub


Private Sub UserControl_Terminate()
    'TerminateGDI
End Sub


Private Sub Draw()
    Dim lPercent    As Long
    Dim hGraphics   As Long
    Dim hPen        As Long
    Dim i           As Long

    Dim SL As Single, ST As Single
    Dim S As Single, C As Single
    Dim MidSize As Single, Size As Single
    
    UserControl.Cls
    
    If GdipCreateFromHDC(UserControl.hdc, hGraphics) = 0 Then

        Call GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias)

        If mLineEnd = LineCapDiamondAnchor Or mLineEnd = LineCapRoundAnchor Then
            MidSize = mDrawWidth
        Else
            MidSize = (mDrawWidth / 2)
        End If

        Size = (UserControl.ScaleWidth / 2) - MidSize - 1
        SL = Size + MidSize
        ST = Size + MidSize
        
        MidSize = (mDrawWidth * 2)

        For i = 360 To mTotalLines Step -mTotalLines
            
            S = Sin(i * PI180)
            C = Cos(i * PI180)
            
            lPercent = ((CurrentPos + i + 20) Mod 360) * 100 / 360
            lPercent = lPercent * 255 / 100

            GdipCreatePen1 CombineColors(mForeColor, mBackColor, lPercent), mDrawWidth, UnitPixel, hPen
            GdipSetPenStartCap hPen, mLineStart
            GdipSetPenEndCap hPen, mLineEnd

            Call GdipDrawLine(hGraphics, hPen, SL + (S * MidSize), ST - (C * MidSize), S * Size + SL, -C * Size + ST)

            GdipDeletePen hPen
        Next i

        GdipDeleteGraphics hGraphics
    End If
    
    UserControl.Refresh
End Sub
 
 
'Función para combinar dos colores y asignar el color alpha.
Private Function CombineColors(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lPercent As Long, Optional ByVal lAlpha As Long = 255) As Long
 
    Dim clrFore(3)         As Byte
    Dim clrBack(3)         As Byte
    Dim clrFinal(3)        As Byte
 
    OleTranslateColor clrFirst, 0, VarPtr(clrFore(0))
    OleTranslateColor clrSecond, 0, VarPtr(clrBack(0))
  
    clrFinal(0) = (clrFore(2) * lPercent + clrBack(2) * (255 - lPercent)) / 255
    clrFinal(1) = (clrFore(1) * lPercent + clrBack(1) * (255 - lPercent)) / 255
    clrFinal(2) = (clrFore(0) * lPercent + clrBack(0) * (255 - lPercent)) / 255
    clrFinal(3) = lAlpha
    
    CopyMemory CombineColors, clrFinal(0), 4
 
End Function
 
 
Private Sub Timer1_Timer()
    CurrentPos = CurrentPos - 30
    If CurrentPos <= 0 Then CurrentPos = 360
    Draw
End Sub


'Private Sub InitGDI()
'    Dim GdipStartupInput As GDIPlusStartupInput
'    GdipStartupInput.GdiPlusVersion = 1&
'    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
'End Sub
 
'-----------------------------------------------------------------

'Private Sub TerminateGDI()
'    If GdipToken Then Call GdiplusShutdown(GdipToken)
'End Sub
'-----------------------------------------------------------------

'GDI Plus Safe Mode (By LaVolpe)
'Avira Antivir detecta este codigo como malicioso, asi que no preocuparse porque es inofencivo.
'puede ser subplantado por las funciones InitGDI y TerminateGDI pero es recomendable avilitarlas solo cuando se compile el proyecto.
Private Function ManageGDIToken(ByVal projectHwnd As Long) As Long
    If projectHwnd = 0& Then Exit Function
    
    Dim hwndGDIsafe     As Long                 'API window to monitor IDE shutdown
    
    Do
        hwndGDIsafe = GetParent(projectHwnd)
        If Not hwndGDIsafe = 0& Then projectHwnd = hwndGDIsafe
    Loop Until hwndGDIsafe = 0&
    ' ok, got the highest level parent, now find highest level owner
    Do
        hwndGDIsafe = GetWindow(projectHwnd, GW_OWNER)
        If Not hwndGDIsafe = 0& Then projectHwnd = hwndGDIsafe
    Loop Until hwndGDIsafe = 0&
    
    hwndGDIsafe = FindWindowEx(projectHwnd, 0&, "Static", "GDI+Safe Patch")
    If hwndGDIsafe Then
        ManageGDIToken = hwndGDIsafe    ' we already have a manager running for this VB instance
        Exit Function                   ' can abort
    End If
    
    Dim gdiSI           As GDIPlusStartupInput  'GDI+ startup info
    Dim gToken          As Long                 'GDI+ instance token
    
    On Error Resume Next
    gdiSI.GdiPlusVersion = 1                    ' attempt to start GDI+
    GdiplusStartup gToken, gdiSI
    If gToken = 0& Then                         ' failed to start
        If Err Then Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    Dim z_ScMem         As Long                 'Thunk base address
    Dim z_Code()        As Long                 'Thunk machine-code initialised here
    Dim nAddr           As Long                 'hwndGDIsafe prev window procedure

    Const WNDPROC_OFF   As Long = &H30          'Offset where window proc starts from z_ScMem
    Const PAGE_RWX      As Long = &H40&         'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&       'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&       'Release allocated memory flag
    Const MEM_LEN       As Long = &HD4          'Byte length of thunk
        
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
    If z_ScMem <> 0 Then                                     'Ensure the allocation succeeded
        ' we make the api window a child so we can use FindWindowEx to locate it easily
        hwndGDIsafe = CreateWindowExA(0&, "Static", "GDI+Safe Patch", WS_CHILD, 0&, 0&, 0&, 0&, projectHwnd, 0&, App.hInstance, ByVal 0&)
        If hwndGDIsafe <> 0 Then
        
            ReDim z_Code(0 To MEM_LEN \ 4 - 1)
        
            z_Code(12) = &HD231C031: z_Code(13) = &HBBE58960: z_Code(14) = &H12345678: z_Code(15) = &H3FFF631: z_Code(16) = &H74247539: z_Code(17) = &H3075FF5B: z_Code(18) = &HFF2C75FF: z_Code(19) = &H75FF2875
            z_Code(20) = &H2C73FF24: z_Code(21) = &H890853FF: z_Code(22) = &HBFF1C45: z_Code(23) = &H2287D81: z_Code(24) = &H75000000: z_Code(25) = &H443C707: z_Code(26) = &H2&: z_Code(27) = &H2C753339: z_Code(28) = &H2047B81: z_Code(29) = &H75000000
            z_Code(30) = &H2C73FF23: z_Code(31) = &HFFFFFC68: z_Code(32) = &H2475FFFF: z_Code(33) = &H681C53FF: z_Code(34) = &H12345678: z_Code(35) = &H3268&: z_Code(36) = &HFF565600: z_Code(37) = &H43892053: z_Code(38) = &H90909020: z_Code(39) = &H10C261
            z_Code(40) = &H562073FF: z_Code(41) = &HFF2453FF: z_Code(42) = &H53FF1473: z_Code(43) = &H2873FF18: z_Code(44) = &H581053FF: z_Code(45) = &H89285D89: z_Code(46) = &H45C72C75: z_Code(47) = &H800030: z_Code(48) = &H20458B00: z_Code(49) = &H89145D89
            z_Code(50) = &H81612445: z_Code(51) = &H4C4&: z_Code(52) = &HC63FF00

            z_Code(1) = 0                                                   ' shutDown mode; used internally by ASM
            z_Code(2) = zFnAddr("user32", "CallWindowProcA")                ' function pointer CallWindowProc
            z_Code(3) = zFnAddr("kernel32", "VirtualFree")                  ' function pointer VirtualFree
            z_Code(4) = zFnAddr("kernel32", "FreeLibrary")                  ' function pointer FreeLibrary
            z_Code(5) = gToken                                              ' Gdi+ token
            z_Code(10) = LoadLibrary("gdiplus")                             ' library pointer (add reference)
            z_Code(6) = GetProcAddress(z_Code(10), "GdiplusShutdown")       ' function pointer GdiplusShutdown
            z_Code(7) = zFnAddr("user32", "SetWindowLongA")                 ' function pointer SetWindowLong
            z_Code(8) = zFnAddr("user32", "SetTimer")                       ' function pointer SetTimer
            z_Code(9) = zFnAddr("user32", "KillTimer")                      ' function pointer KillTimer
        
            z_Code(14) = z_ScMem                                            ' ASM ebx start point
            z_Code(34) = z_ScMem + WNDPROC_OFF                              ' subclass window procedure location
        
            RtlMoveMemory z_ScMem, VarPtr(z_Code(0)), MEM_LEN               'Copy the thunk code/data to the allocated memory
        
            nAddr = SetWindowLong(hwndGDIsafe, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Subclass our API window
            RtlMoveMemory z_ScMem + 44, VarPtr(nAddr), 4& ' Add prev window procedure to the thunk
            gToken = 0& ' zeroize so final check below does not release it
            
            ManageGDIToken = hwndGDIsafe    ' return handle of our GDI+ manager
        Else
            VirtualFree z_ScMem, 0, MEM_RELEASE     ' failure - release memory
            z_ScMem = 0&
        End If
    Else
        VirtualFree z_ScMem, 0, MEM_RELEASE           ' failure - release memory
        z_ScMem = 0&
    End If
    
    If gToken Then GdiplusShutdown gToken       ' release token if error occurred
    
End Function


Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)  'Get the specified procedure address
End Function
